/* =====================================================================
   SQL Wait Monitor - CENTRAL install (monitor many servers through linked servers)

   Run this whole script once in SSMS on the CENTRAL server (the one that
   stores the data and runs the SQL Agent jobs). Normal SSMS mode. Safe to re-run.

   How it works
     Every 10 seconds the central Agent job loops over dbo.MonitoredServers.
     For each server it runs the DMV snapshot ON that server
        - local server:  directly
        - remote server: EXEC (...) AT [linked server]   (needs RPC Out)
     and stores the rows centrally, tagged with ServerName. Nothing is
     installed on the remote servers except a login with read permissions.

   Requirements
     Central: SQL Server 2017+, SQL Agent running, sysadmin to install.
     Remote:  SQL Server 2016+ (sys.dm_exec_input_buffer), and the linked
              server login needs VIEW SERVER STATE (+ VIEW ANY DEFINITION
              for proc names). See section 0.

   Steps
     0. On EACH remote server: create the monitoring login (section 0, run there).
     1. On the central server: create a linked server per remote server (section 0 example).
     2. Run this script on the central server (edit settings in section 2).
     3. Register servers:  EXEC dbo.usp_AddMonitoredServer @ServerName = N'PRODSQL01', @LinkedServer = N'PRODSQL01';
     4. Test:              EXEC dbo.usp_TestMonitoredServer @ServerName = N'PRODSQL01';
     5. View it with Start-SqlWaitMonitorCentral.cmd (pick the server at the top).
   ===================================================================== */

/* =====================================================================
   0. SETUP EXAMPLES (commented out - copy, edit and run where noted)
   ---------------------------------------------------------------------
   -- 0a. Run on EACH REMOTE server: a read-only monitoring login
   USE master;
   CREATE LOGIN dbawait_monitor WITH PASSWORD = N'<strong password here>', CHECK_POLICY = ON;
   GRANT VIEW SERVER STATE    TO dbawait_monitor;   -- DMVs: requests, sessions, schedulers, plans
   GRANT VIEW ANY DEFINITION  TO dbawait_monitor;   -- proc names and "Show whole procedure"

   -- 0b. Run on the CENTRAL server: one linked server per remote server
   EXEC master.dbo.sp_addlinkedserver
        @server = N'PRODSQL01', @srvproduct = N'', @provider = N'MSOLEDBSQL', @datasrc = N'PRODSQL01',   -- datasrc = host\instance,port
        @provstr = N'Encrypt=Optional';   -- MSOLEDBSQL 19 encrypts by default; use a trusted cert in production
   EXEC master.dbo.sp_addlinkedsrvlogin
        @rmtsrvname = N'PRODSQL01', @useself = N'False', @locallogin = NULL,
        @rmtuser = N'dbawait_monitor', @rmtpassword = N'<same password>';
   -- (usp_AddMonitoredServer below turns on RPC Out and turns off DTC promotion for you)

   -- Test loop-back on a sandbox: a linked server pointing at the central server itself
   -- EXEC master.dbo.sp_addlinkedserver @server = N'LOOPBACK', @srvproduct = N'', @provider = N'MSOLEDBSQL', @datasrc = @@SERVERNAME, @provstr = N'Encrypt=Optional';
   -- EXEC DBAWaitCentral.dbo.usp_AddMonitoredServer @ServerName = N'LOOPBACK test', @LinkedServer = N'LOOPBACK';
   ===================================================================== */
SET NOCOUNT ON;
USE master;
GO
IF DB_ID(N'DBAWaitCentral') IS NULL
BEGIN
    CREATE DATABASE DBAWaitCentral;
    ALTER DATABASE DBAWaitCentral SET RECOVERY SIMPLE;
END
GO
USE DBAWaitCentral;
GO

/* =====================================================================
   1. Tables
   ===================================================================== */
IF OBJECT_ID(N'dbo.MonitorConfig') IS NULL
CREATE TABLE dbo.MonitorConfig (
    ConfigId          INT            NOT NULL CONSTRAINT PK_MonitorConfig PRIMARY KEY CONSTRAINT CK_MonitorConfig_One CHECK (ConfigId = 1),
    IntervalSec       INT            NOT NULL,
    WindowMin         INT            NOT NULL,
    WaitThresholdSec  INT            NOT NULL,
    BlockedThreshold  INT            NOT NULL,
    CooldownMin       INT            NOT NULL,   -- per server
    MailProfile       SYSNAME        NULL,
    Recipients        NVARCHAR(400)  NULL,
    RetentionDays     INT            NOT NULL
);
GO

IF OBJECT_ID(N'dbo.MonitoredServers') IS NULL
CREATE TABLE dbo.MonitoredServers (
    ServerName        SYSNAME        NOT NULL CONSTRAINT PK_MonitoredServers PRIMARY KEY,  -- display name
    LinkedServer      SYSNAME        NULL,       -- NULL = the central server itself
    IsEnabled         BIT            NOT NULL CONSTRAINT DF_MonitoredServers_Enabled DEFAULT 1,
    LastCollected     DATETIME2(0)   NULL,       -- UTC
    LastDurationMs    INT            NULL,
    LastError         NVARCHAR(4000) NULL,
    LastErrorTime     DATETIME2(0)   NULL
);
GO

IF OBJECT_ID(N'dbo.WaitSamples') IS NULL
CREATE TABLE dbo.WaitSamples (
    SampleId          BIGINT IDENTITY CONSTRAINT PK_WaitSamples PRIMARY KEY,
    ServerName        SYSNAME        NOT NULL,
    SampleTime        DATETIME2(0)   NOT NULL,   -- UTC (central clock)
    SessionId         INT            NOT NULL,
    DatabaseName      SYSNAME        NULL,
    LoginName         SYSNAME        NULL,
    HostName          NVARCHAR(128)  NULL,
    ProgramName       NVARCHAR(128)  NULL,
    WaitCategory      NVARCHAR(60)   NOT NULL,
    WaitResource      NVARCHAR(256)  NULL,
    BlockingSessionId INT            NULL,
    WeightSec         INT            NOT NULL,
    QueryHash         BINARY(8)      NULL,
    StatementId       INT            NULL
);
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.WaitSamples') AND name = N'IX_WaitSamples_Server_Time')
    CREATE INDEX IX_WaitSamples_Server_Time ON dbo.WaitSamples (ServerName, SampleTime)
        INCLUDE (WaitCategory, WeightSec, StatementId, SessionId, BlockingSessionId);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.WaitSamples') AND name = N'IX_WaitSamples_Stmt')
    CREATE INDEX IX_WaitSamples_Stmt ON dbo.WaitSamples (StatementId, SampleTime)
        INCLUDE (WaitCategory, WeightSec, BlockingSessionId, SessionId, ServerName);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.WaitSamples') AND name = N'IX_WaitSamples_Time')
    CREATE INDEX IX_WaitSamples_Time ON dbo.WaitSamples (SampleTime);
GO

IF OBJECT_ID(N'dbo.SqlStatements') IS NULL
CREATE TABLE dbo.SqlStatements (
    StatementId       INT IDENTITY   CONSTRAINT PK_SqlStatements PRIMARY KEY,
    ServerName        SYSNAME        NOT NULL,
    QueryHash         BINARY(8)      NOT NULL,
    DbId              INT            NOT NULL,
    ObjectId          INT            NOT NULL,   -- 0 = ad hoc / dynamic SQL
    DatabaseName      SYSNAME        NULL,
    ObjectName        NVARCHAR(300)  NULL,
    StartLine         INT            NULL,
    StatementText     NVARCHAR(MAX)  NULL,
    LastCalledFrom    NVARCHAR(4000) NULL,
    LastPlanHandle    VARBINARY(64)  NULL,
    LastQueryPlanHash BINARY(8)      NULL,
    FirstSeen         DATETIME2(0)   NOT NULL,
    LastSeen          DATETIME2(0)   NOT NULL,
    CONSTRAINT UQ_SqlStatements UNIQUE (ServerName, QueryHash, DbId, ObjectId)
);
GO

IF OBJECT_ID(N'dbo.BlockerSamples') IS NULL
CREATE TABLE dbo.BlockerSamples (
    ServerName        SYSNAME        NOT NULL,
    SampleTime        DATETIME2(0)   NOT NULL,
    SessionId         INT            NOT NULL,
    Status            NVARCHAR(30)   NULL,
    LoginName         SYSNAME        NULL,
    HostName          NVARCHAR(128)  NULL,
    ProgramName       NVARCHAR(128)  NULL,
    OpenTranCount     INT            NULL,
    BlockedCount      INT            NULL,
    LastSqlText       NVARCHAR(4000) NULL,
    CONSTRAINT PK_BlockerSamples PRIMARY KEY (ServerName, SampleTime, SessionId)
);
GO

IF OBJECT_ID(N'dbo.SchedulerSamples') IS NULL
CREATE TABLE dbo.SchedulerSamples (
    ServerName        SYSNAME        NOT NULL,
    SampleTime        DATETIME2(0)   NOT NULL,
    CurrentWorkers    INT, ActiveWorkers INT, RunnableTasks INT, WorkQueueCount INT, MaxWorkers INT,
    CONSTRAINT PK_SchedulerSamples PRIMARY KEY (ServerName, SampleTime)
);
GO

IF OBJECT_ID(N'dbo.AlertLog') IS NULL
CREATE TABLE dbo.AlertLog (
    AlertId           INT IDENTITY   CONSTRAINT PK_AlertLog PRIMARY KEY,
    ServerName        SYSNAME        NOT NULL,
    AlertTime         DATETIME2(0)   NOT NULL CONSTRAINT DF_AlertLog_Time DEFAULT SYSUTCDATETIME(),
    AlertType         NVARCHAR(50),
    TotalWaitSec      INT,
    BlockedCount      INT,
    HeadBlocker       INT,
    Details           NVARCHAR(MAX)
);
GO

/* =====================================================================
   2. SETTINGS - edit these values, then run the script
   ===================================================================== */
MERGE dbo.MonitorConfig AS t
USING (SELECT
        ConfigId         = 1,
        IntervalSec      = 10,
        WindowMin        = 10,
        WaitThresholdSec = 10000,    -- per server
        BlockedThreshold = 20,       -- per server
        CooldownMin      = 15,       -- per server
        MailProfile      = CAST(NULL AS SYSNAME),            -- e.g. N'DBA Mail'  (NULL = log only)
        Recipients       = CAST(NULL AS NVARCHAR(400)),      -- e.g. N'dba-team@yourcompany.com'
        RetentionDays    = 14
      ) AS s ON t.ConfigId = s.ConfigId
WHEN MATCHED THEN UPDATE SET
    IntervalSec = s.IntervalSec, WindowMin = s.WindowMin, WaitThresholdSec = s.WaitThresholdSec,
    BlockedThreshold = s.BlockedThreshold, CooldownMin = s.CooldownMin, MailProfile = s.MailProfile,
    Recipients = s.Recipients, RetentionDays = s.RetentionDays
WHEN NOT MATCHED THEN INSERT (ConfigId, IntervalSec, WindowMin, WaitThresholdSec, BlockedThreshold, CooldownMin, MailProfile, Recipients, RetentionDays)
    VALUES (s.ConfigId, s.IntervalSec, s.WindowMin, s.WaitThresholdSec, s.BlockedThreshold, s.CooldownMin, s.MailProfile, s.Recipients, s.RetentionDays);
GO

-- the central server monitors itself by default (disable with: UPDATE dbo.MonitoredServers SET IsEnabled = 0 WHERE LinkedServer IS NULL)
IF NOT EXISTS (SELECT 1 FROM dbo.MonitoredServers WHERE LinkedServer IS NULL)
    INSERT dbo.MonitoredServers (ServerName, LinkedServer) VALUES (@@SERVERNAME, NULL);
GO

/* =====================================================================
   3. Readable statement label
   ===================================================================== */
CREATE OR ALTER VIEW dbo.vStatements
AS
SELECT s.StatementId, s.ServerName, s.DatabaseName, s.ObjectName, s.StartLine, s.StatementText,
       s.LastCalledFrom, s.LastPlanHandle, s.FirstSeen, s.LastSeen,
       Label = CONCAT(N'#', s.StatementId, N' ',
               CASE WHEN s.ObjectId <> 0
                    THEN CONCAT(ISNULL(s.ObjectName, N'(object ' + CAST(s.ObjectId AS NVARCHAR(20)) + N')'), N' line ', s.StartLine)
                    ELSE CONCAT(N'Ad hoc: ',
                         ISNULL(NULLIF(LEFT(LTRIM(REPLACE(REPLACE(REPLACE(s.StatementText,
                                NCHAR(13), N' '), NCHAR(10), N' '), NCHAR(9), N' ')), 50), N''),
                                N'(no text) ' + LEFT(LTRIM(REPLACE(REPLACE(s.LastCalledFrom, NCHAR(13), N' '), NCHAR(10), N' ')), 40)))
               END)
FROM dbo.SqlStatements s;
GO

/* =====================================================================
   4. Collector for ONE server (runs the snapshot there, stores it here)
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_CollectOne
    @ServerName   SYSNAME,
    @LinkedServer SYSNAME = NULL,
    @IntervalSec  INT = 10
AS
SET LOCK_TIMEOUT 500
BEGIN
    SET NOCOUNT ON;
    DECLARE @now DATETIME2(0) = SYSUTCDATETIME();
    DECLARE @runner NVARCHAR(400) =
        CASE WHEN @LinkedServer IS NULL THEN N'EXEC sys.sp_executesql @q;'
             ELSE N'EXEC (@q) AT ' + QUOTENAME(@LinkedServer) + N';' END;
    DECLARE @q NVARCHAR(MAX);

    CREATE TABLE #req (
        SessionId INT, DatabaseName SYSNAME NULL, LoginName SYSNAME NULL, HostName NVARCHAR(128) NULL,
        ProgramName NVARCHAR(128) NULL, WaitCategory NVARCHAR(60), WaitResource NVARCHAR(256) NULL,
        BlockingSessionId INT NULL, QueryHash BINARY(8) NULL, DbId INT, ObjectId INT,
        StmtDatabaseName SYSNAME NULL, ObjectName NVARCHAR(300) NULL, StartLine INT NULL,
        StmtText NVARCHAR(MAX) NULL, CalledFrom NVARCHAR(4000) NULL,
        PlanHandle VARBINARY(64) NULL, QueryPlanHash BINARY(8) NULL);

    /* --- snapshot of active requests, evaluated ON the monitored server --- */
    SET @q = N'
SET NOCOUNT ON;
SET LOCK_TIMEOUT 500;
SELECT
    r.session_id,
    DB_NAME(r.database_id),
    s.login_name,
    s.host_name,
    s.program_name,
    CASE WHEN r.wait_type IS NULL THEN N''CPU'' ELSE r.wait_type END,
    r.wait_resource,
    NULLIF(r.blocking_session_id, 0),
    ISNULL(r.query_hash, CAST(HASHBYTES(''SHA2_256'', ISNULL(x.StmtText, r.command)) AS BINARY(8))),
    ISNULL(t.dbid, r.database_id),
    ISNULL(t.objectid, 0),
    DB_NAME(ISNULL(t.dbid, r.database_id)),
    CASE WHEN t.objectid IS NOT NULL THEN CONCAT(OBJECT_SCHEMA_NAME(t.objectid, t.dbid), N''.'', OBJECT_NAME(t.objectid, t.dbid)) END,
    x.StartLine,
    x.StmtText,
    LEFT(ib.event_info, 4000),
    r.plan_handle,
    r.query_plan_hash
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
OUTER APPLY (
    SELECT StmtText  = SUBSTRING(t.text, r.statement_start_offset / 2 + 1,
                           (CASE WHEN r.statement_end_offset = -1 THEN DATALENGTH(t.text)
                                 ELSE r.statement_end_offset END - r.statement_start_offset) / 2 + 1),
           StartLine = (DATALENGTH(LEFT(t.text, r.statement_start_offset / 2))
                      - DATALENGTH(REPLACE(LEFT(t.text, r.statement_start_offset / 2), NCHAR(10), N''''))) / 2 + 1
) x
OUTER APPLY sys.dm_exec_input_buffer(r.session_id, r.request_id) ib
WHERE s.is_user_process = 1
  AND r.session_id <> @@SPID
  AND ISNULL(r.wait_type, N'''') NOT IN (N''WAITFOR'', N''SP_SERVER_DIAGNOSTICS_SLEEP'', N''BROKER_RECEIVE_WAITFOR'');';

    INSERT INTO #req EXEC sys.sp_executesql @runner, N'@q NVARCHAR(MAX)', @q = @q;

    /* --- statements --- */
    SELECT QueryHash, DbId, ObjectId,
           StmtDatabaseName = MAX(StmtDatabaseName), ObjectName = MAX(ObjectName),
           StartLine = MIN(StartLine), StmtText = MAX(StmtText), CalledFrom = MAX(CalledFrom),
           PlanHandle = MAX(PlanHandle), QueryPlanHash = MAX(QueryPlanHash)
    INTO #stmt
    FROM #req
    WHERE QueryHash IS NOT NULL
    GROUP BY QueryHash, DbId, ObjectId;

    UPDATE s
       SET LastSeen          = @now,
           LastCalledFrom    = ISNULL(n.CalledFrom, s.LastCalledFrom),
           LastPlanHandle    = ISNULL(n.PlanHandle, s.LastPlanHandle),
           LastQueryPlanHash = ISNULL(n.QueryPlanHash, s.LastQueryPlanHash)
    FROM dbo.SqlStatements s
    JOIN #stmt n ON s.ServerName = @ServerName AND n.QueryHash = s.QueryHash AND n.DbId = s.DbId AND n.ObjectId = s.ObjectId;

    INSERT dbo.SqlStatements (ServerName, QueryHash, DbId, ObjectId, DatabaseName, ObjectName, StartLine,
                              StatementText, LastCalledFrom, LastPlanHandle, LastQueryPlanHash, FirstSeen, LastSeen)
    SELECT @ServerName, n.QueryHash, n.DbId, n.ObjectId, n.StmtDatabaseName, n.ObjectName, n.StartLine,
           n.StmtText, n.CalledFrom, n.PlanHandle, n.QueryPlanHash, @now, @now
    FROM #stmt n
    WHERE NOT EXISTS (SELECT 1 FROM dbo.SqlStatements s
                      WHERE s.ServerName = @ServerName AND s.QueryHash = n.QueryHash
                        AND s.DbId = n.DbId AND s.ObjectId = n.ObjectId);

    /* --- samples --- */
    INSERT dbo.WaitSamples (ServerName, SampleTime, SessionId, DatabaseName, LoginName, HostName, ProgramName,
                            WaitCategory, WaitResource, BlockingSessionId, WeightSec, QueryHash, StatementId)
    SELECT @ServerName, @now, r.SessionId, r.DatabaseName, r.LoginName, r.HostName, r.ProgramName,
           r.WaitCategory, r.WaitResource, r.BlockingSessionId, @IntervalSec, r.QueryHash, s.StatementId
    FROM #req r
    LEFT JOIN dbo.SqlStatements s
           ON s.ServerName = @ServerName AND s.QueryHash = r.QueryHash AND s.DbId = r.DbId AND s.ObjectId = r.ObjectId;

    /* --- head blockers (only when something is blocked) --- */
    IF EXISTS (SELECT 1 FROM #req WHERE BlockingSessionId > 0)
    BEGIN
        CREATE TABLE #blk (SessionId INT, Status NVARCHAR(30) NULL, LoginName SYSNAME NULL, HostName NVARCHAR(128) NULL,
                           ProgramName NVARCHAR(128) NULL, OpenTranCount INT NULL, BlockedCount INT NULL, LastSqlText NVARCHAR(4000) NULL);
        SET @q = N'
SET NOCOUNT ON;
SET LOCK_TIMEOUT 500;
WITH r AS (SELECT session_id, blocking_session_id FROM sys.dm_exec_requests WHERE blocking_session_id > 0)
SELECT b.blocking_session_id, s.status, s.login_name, s.host_name, s.program_name,
       s.open_transaction_count, b.BlockedCount, LEFT(t.text, 4000)
FROM (SELECT blocking_session_id, BlockedCount = COUNT(*) FROM r GROUP BY blocking_session_id) b
JOIN sys.dm_exec_sessions s ON s.session_id = b.blocking_session_id
OUTER APPLY (SELECT TOP (1) c.most_recent_sql_handle FROM sys.dm_exec_connections c WHERE c.session_id = b.blocking_session_id) c
OUTER APPLY sys.dm_exec_sql_text(c.most_recent_sql_handle) t
WHERE b.blocking_session_id NOT IN (SELECT session_id FROM r);';
        INSERT INTO #blk EXEC sys.sp_executesql @runner, N'@q NVARCHAR(MAX)', @q = @q;

        INSERT dbo.BlockerSamples (ServerName, SampleTime, SessionId, Status, LoginName, HostName, ProgramName,
                                   OpenTranCount, BlockedCount, LastSqlText)
        SELECT DISTINCT @ServerName, @now, SessionId, Status, LoginName, HostName, ProgramName, OpenTranCount, BlockedCount, LastSqlText
        FROM #blk;
    END

    /* --- worker threads --- */
    CREATE TABLE #sch (CurrentWorkers INT, ActiveWorkers INT, RunnableTasks INT, WorkQueueCount INT, MaxWorkers INT);
    SET @q = N'
SET NOCOUNT ON;
SELECT SUM(current_workers_count), SUM(active_workers_count), SUM(runnable_tasks_count), SUM(work_queue_count),
       (SELECT max_workers_count FROM sys.dm_os_sys_info)
FROM sys.dm_os_schedulers
WHERE status = N''VISIBLE ONLINE'';';
    INSERT INTO #sch EXEC sys.sp_executesql @runner, N'@q NVARCHAR(MAX)', @q = @q;
    INSERT dbo.SchedulerSamples (ServerName, SampleTime, CurrentWorkers, ActiveWorkers, RunnableTasks, WorkQueueCount, MaxWorkers)
    SELECT TOP (1) @ServerName, @now, CurrentWorkers, ActiveWorkers, RunnableTasks, WorkQueueCount, MaxWorkers FROM #sch;
END
GO

/* =====================================================================
   5. Collector for ALL enabled servers (the Agent job runs this)
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_CollectWaits
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @IntervalSec INT, @keepDays INT, @srv SYSNAME, @ls SYSNAME, @t0 DATETIME2(3);
    SELECT @IntervalSec = IntervalSec, @keepDays = RetentionDays FROM dbo.MonitorConfig;
    SET @IntervalSec = ISNULL(@IntervalSec, 10);
    SET @keepDays    = ISNULL(@keepDays, 14);

    DECLARE srv CURSOR LOCAL FAST_FORWARD FOR
        SELECT ServerName, LinkedServer FROM dbo.MonitoredServers WHERE IsEnabled = 1 ORDER BY ServerName;
    OPEN srv;
    FETCH NEXT FROM srv INTO @srv, @ls;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @t0 = SYSUTCDATETIME();
        BEGIN TRY
            EXEC dbo.usp_CollectOne @ServerName = @srv, @LinkedServer = @ls, @IntervalSec = @IntervalSec;
            UPDATE dbo.MonitoredServers
               SET LastCollected = SYSUTCDATETIME(), LastDurationMs = DATEDIFF(MILLISECOND, @t0, SYSUTCDATETIME()), LastError = NULL
             WHERE ServerName = @srv;
        END TRY
        BEGIN CATCH
            IF XACT_STATE() <> 0 ROLLBACK;
            UPDATE dbo.MonitoredServers
               SET LastError = LEFT(CONCAT(N'Error ', ERROR_NUMBER(), N': ', ERROR_MESSAGE()), 4000), LastErrorTime = SYSUTCDATETIME(),
                   LastDurationMs = DATEDIFF(MILLISECOND, @t0, SYSUTCDATETIME())
             WHERE ServerName = @srv;
        END CATCH
        FETCH NEXT FROM srv INTO @srv, @ls;
    END
    CLOSE srv; DEALLOCATE srv;

    /* retention (small batches) */
    DECLARE @now DATETIME2(0) = SYSUTCDATETIME();
    DECLARE @cut DATETIME2(0) = DATEADD(DAY, -@keepDays, @now);
    DELETE TOP (20000) FROM dbo.WaitSamples      WHERE SampleTime < @cut;
    DELETE TOP (5000) FROM  dbo.SchedulerSamples WHERE SampleTime < @cut;
    DELETE TOP (5000) FROM  dbo.BlockerSamples   WHERE SampleTime < @cut;
    IF DATEPART(MINUTE, @now) = 0 AND DATEPART(SECOND, @now) < @IntervalSec
        DELETE s FROM dbo.SqlStatements s
        WHERE s.LastSeen < DATEADD(DAY, -(@keepDays + 1), @now)
          AND NOT EXISTS (SELECT 1 FROM dbo.WaitSamples w WHERE w.StatementId = s.StatementId);
END
GO

/* =====================================================================
   6. Alert check, per server
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_CheckWaitAlert
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @WindowMin INT, @WaitThresholdSec INT, @BlockedThreshold INT, @CooldownMin INT,
            @MailProfile SYSNAME, @Recipients NVARCHAR(400);
    SELECT @WindowMin = WindowMin, @WaitThresholdSec = WaitThresholdSec, @BlockedThreshold = BlockedThreshold,
           @CooldownMin = CooldownMin, @MailProfile = MailProfile, @Recipients = Recipients
    FROM dbo.MonitorConfig;

    DECLARE @from DATETIME2(0) = DATEADD(MINUTE, -@WindowMin, SYSUTCDATETIME());

    /* per-server totals and blocked count in that server's latest sample */
    SELECT m.ServerName,
           TotalWait = ISNULL((SELECT SUM(w.WeightSec) FROM dbo.WaitSamples w
                               WHERE w.ServerName = m.ServerName AND w.SampleTime >= @from), 0),
           Blocked   = ISNULL((SELECT COUNT(DISTINCT w.SessionId) FROM dbo.WaitSamples w
                               WHERE w.ServerName = m.ServerName AND w.BlockingSessionId IS NOT NULL
                                 AND w.SampleTime = (SELECT MAX(x.SampleTime) FROM dbo.WaitSamples x
                                                     WHERE x.ServerName = m.ServerName AND x.SampleTime >= @from)), 0)
    INTO #s
    FROM dbo.MonitoredServers m
    WHERE m.IsEnabled = 1;

    DECLARE @srv SYSNAME, @total INT, @blocked INT, @head INT, @hb NVARCHAR(MAX), @top NVARCHAR(MAX), @body NVARCHAR(MAX), @alertId INT;
    DECLARE a CURSOR LOCAL FAST_FORWARD FOR
        SELECT ServerName, TotalWait, Blocked FROM #s
        WHERE (TotalWait >= @WaitThresholdSec OR Blocked >= @BlockedThreshold)
          AND NOT EXISTS (SELECT 1 FROM dbo.AlertLog l
                          WHERE l.ServerName = #s.ServerName AND l.AlertTime > DATEADD(MINUTE, -@CooldownMin, SYSUTCDATETIME()));
    OPEN a;
    FETCH NEXT FROM a INTO @srv, @total, @blocked;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SELECT @top = NULL, @hb = NULL, @head = NULL;

        SELECT @top = STRING_AGG(CAST(CONCAT(x.Label, N'  |  ', x.WaitSec, N' s  |  mostly ', x.TopWait) AS NVARCHAR(MAX)), NCHAR(10))
                      WITHIN GROUP (ORDER BY x.WaitSec DESC)
        FROM (
            SELECT TOP (10) v.Label, WaitSec = SUM(w.WeightSec),
                   TopWait = (SELECT TOP (1) w2.WaitCategory FROM dbo.WaitSamples w2
                              WHERE w2.StatementId = w.StatementId AND w2.SampleTime >= @from
                              GROUP BY w2.WaitCategory ORDER BY SUM(w2.WeightSec) DESC)
            FROM dbo.WaitSamples w
            JOIN dbo.vStatements v ON v.StatementId = w.StatementId
            WHERE w.ServerName = @srv AND w.SampleTime >= @from
            GROUP BY w.StatementId, v.Label
            ORDER BY SUM(w.WeightSec) DESC
        ) x;

        SELECT TOP (1) @head = SessionId,
               @hb = CONCAT(N'SPID ', SessionId, N'  ', LoginName, N' @ ', HostName, N'  (', Status, N', open tran ', OpenTranCount,
                            N', blocking ', BlockedCount, N')', NCHAR(10), N'Head blocker SQL: ', LEFT(LastSqlText, 1000))
        FROM dbo.BlockerSamples
        WHERE ServerName = @srv AND SampleTime >= @from
        ORDER BY SampleTime DESC, BlockedCount DESC;

        SET @body = CONCAT(
            N'Server: ', @srv, NCHAR(10),
            N'Total wait last ', @WindowMin, N' min: ', @total, N' s (threshold ', @WaitThresholdSec, N')', NCHAR(10),
            N'Blocked sessions in latest sample: ', @blocked, N' (threshold ', @BlockedThreshold, N')', NCHAR(10),
            N'Head blocker: ', ISNULL(@hb, N'none captured'), NCHAR(10), NCHAR(10),
            N'Top statements by wait (#id = statement in the viewer):', NCHAR(10), @top);

        INSERT dbo.AlertLog (ServerName, AlertType, TotalWaitSec, BlockedCount, HeadBlocker, Details)
        VALUES (@srv, CASE WHEN @blocked >= @BlockedThreshold THEN N'Blocking' ELSE N'WaitTime' END, @total, @blocked, @head, @body);
        SET @alertId = SCOPE_IDENTITY();

        IF @MailProfile IS NOT NULL AND @Recipients IS NOT NULL
        BEGIN
            BEGIN TRY
                DECLARE @subject NVARCHAR(255) = CONCAT(N'SQL wait alert on ', @srv, N': total wait / blocking threshold exceeded');
                EXEC msdb.dbo.sp_send_dbmail
                    @profile_name = @MailProfile, @recipients = @Recipients, @subject = @subject, @body = @body;
            END TRY
            BEGIN CATCH
                UPDATE dbo.AlertLog SET Details = CONCAT(Details, NCHAR(10), N'[email failed: ', ERROR_MESSAGE(), N']')
                WHERE AlertId = @alertId;
            END CATCH
        END

        FETCH NEXT FROM a INTO @srv, @total, @blocked;
    END
    CLOSE a; DEALLOCATE a;
END
GO

/* =====================================================================
   7. Helpers: register and test servers
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_AddMonitoredServer
    @ServerName   SYSNAME,          -- name shown in the viewer
    @LinkedServer SYSNAME = NULL,   -- existing linked server; NULL = the central server itself
    @IsEnabled    BIT = 1
AS
BEGIN
    SET NOCOUNT ON;
    IF @LinkedServer IS NOT NULL
    BEGIN
        IF NOT EXISTS (SELECT 1 FROM sys.servers WHERE name = @LinkedServer AND is_linked = 1)
            THROW 50001, N'Linked server not found. Create it first with sp_addlinkedserver / sp_addlinkedsrvlogin (section 0b).', 1;
        -- EXEC ... AT needs RPC Out; INSERT ... EXEC AT must not escalate to a distributed (DTC) transaction
        EXEC master.dbo.sp_serveroption @server = @LinkedServer, @optname = N'rpc out', @optvalue = N'true';
        EXEC master.dbo.sp_serveroption @server = @LinkedServer, @optname = N'remote proc transaction promotion', @optvalue = N'false';
        EXEC master.dbo.sp_serveroption @server = @LinkedServer, @optname = N'connect timeout', @optvalue = N'5';
        EXEC master.dbo.sp_serveroption @server = @LinkedServer, @optname = N'query timeout', @optvalue = N'8';
    END

    MERGE dbo.MonitoredServers AS t
    USING (SELECT ServerName = @ServerName, LinkedServer = @LinkedServer, IsEnabled = @IsEnabled) AS s
       ON t.ServerName = s.ServerName
    WHEN MATCHED THEN UPDATE SET LinkedServer = s.LinkedServer, IsEnabled = s.IsEnabled, LastError = NULL
    WHEN NOT MATCHED THEN INSERT (ServerName, LinkedServer, IsEnabled) VALUES (s.ServerName, s.LinkedServer, s.IsEnabled);

    SELECT * FROM dbo.MonitoredServers WHERE ServerName = @ServerName;
END
GO

CREATE OR ALTER PROCEDURE dbo.usp_TestMonitoredServer
    @ServerName SYSNAME
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @ls SYSNAME, @runner NVARCHAR(400), @q NVARCHAR(MAX);
    SELECT @ls = LinkedServer FROM dbo.MonitoredServers WHERE ServerName = @ServerName;
    IF @@ROWCOUNT = 0 THROW 50002, N'Server is not registered. Run dbo.usp_AddMonitoredServer first.', 1;
    SET @runner = CASE WHEN @ls IS NULL THEN N'EXEC sys.sp_executesql @q;' ELSE N'EXEC (@q) AT ' + QUOTENAME(@ls) + N';' END;

    CREATE TABLE #t (RemoteServer SYSNAME NULL, Version NVARCHAR(200) NULL, LoginName SYSNAME NULL,
                     HasViewServerState INT NULL, HasViewAnyDefinition INT NULL, ActiveRequests INT NULL);
    SET @q = N'SET NOCOUNT ON;
SELECT @@SERVERNAME, LEFT(@@VERSION, 200), SUSER_SNAME(),
       HAS_PERMS_BY_NAME(NULL, NULL, N''VIEW SERVER STATE''),
       HAS_PERMS_BY_NAME(NULL, NULL, N''VIEW ANY DEFINITION''),
       (SELECT COUNT(*) FROM sys.dm_exec_requests);';
    INSERT INTO #t EXEC sys.sp_executesql @runner, N'@q NVARCHAR(MAX)', @q = @q;
    SELECT Monitored = @ServerName, LinkedServer = @ls, * FROM #t;

    -- full collection test (writes one sample)
    EXEC dbo.usp_CollectOne @ServerName = @ServerName, @LinkedServer = @ls;
    SELECT CollectedRows = COUNT(*) FROM dbo.WaitSamples
    WHERE ServerName = @ServerName AND SampleTime >= DATEADD(SECOND, -5, SYSUTCDATETIME());
END
GO

/* =====================================================================
   8. SQL Agent jobs (recreated on every run)
   ===================================================================== */
USE msdb;
GO
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBAWaitCentral - Collect') EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWaitCentral - Collect';
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBAWaitCentral - Alert')   EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWaitCentral - Alert';

EXEC msdb.dbo.sp_add_job        @job_name = N'DBAWaitCentral - Collect', @description = N'SQL Wait Monitor central: sample all monitored servers every 10 seconds';
EXEC msdb.dbo.sp_add_jobstep    @job_name = N'DBAWaitCentral - Collect', @step_name = N'Collect', @subsystem = N'TSQL',
                                @database_name = N'DBAWaitCentral', @command = N'EXEC dbo.usp_CollectWaits;';
EXEC msdb.dbo.sp_add_jobschedule @job_name = N'DBAWaitCentral - Collect', @name = N'DBAWaitCentral every 10 seconds',
                                @freq_type = 4, @freq_interval = 1, @freq_subday_type = 2, @freq_subday_interval = 10;
EXEC msdb.dbo.sp_add_jobserver  @job_name = N'DBAWaitCentral - Collect';

EXEC msdb.dbo.sp_add_job        @job_name = N'DBAWaitCentral - Alert', @description = N'SQL Wait Monitor central: check thresholds per server every minute';
EXEC msdb.dbo.sp_add_jobstep    @job_name = N'DBAWaitCentral - Alert', @step_name = N'Check', @subsystem = N'TSQL',
                                @database_name = N'DBAWaitCentral', @command = N'EXEC dbo.usp_CheckWaitAlert;';
EXEC msdb.dbo.sp_add_jobschedule @job_name = N'DBAWaitCentral - Alert', @name = N'DBAWaitCentral every minute',
                                @freq_type = 4, @freq_interval = 1, @freq_subday_type = 4, @freq_subday_interval = 1;
EXEC msdb.dbo.sp_add_jobserver  @job_name = N'DBAWaitCentral - Alert';
GO

/* =====================================================================
   9. Check
   ===================================================================== */
USE DBAWaitCentral;
GO
EXEC dbo.usp_CollectWaits;
SELECT ServerName, LinkedServer, IsEnabled, LastCollected, LastDurationMs, LastError FROM dbo.MonitoredServers;
GO

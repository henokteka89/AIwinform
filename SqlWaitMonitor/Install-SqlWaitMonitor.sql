/* =====================================================================
   SQL Wait Monitor - SINGLE SERVER install (DPA-style wait monitoring)

   Run this whole script once in SSMS on the server you want to monitor.
   - Normal SSMS mode (no SQLCMD mode needed). Safe to run again: it
     upgrades an existing DBAWait database and keeps the data.
   - Needs SQL Server 2017 or later, sysadmin to install, SQL Agent running.
   - Edit the settings in section 2 (alert thresholds, Database Mail).

   What it creates
     DBAWait database
       WaitSamples       every 10 s: each active request, its wait and statement
       SqlStatements     one row per distinct statement (proc + line, or ad hoc text)
       BlockerSamples    head blockers, including idle sessions with open transactions
       SchedulerSamples  worker threads / runnable tasks
       AlertLog          alerts raised
       MonitorConfig     settings (thresholds, mail, retention)
       vStatements       readable statement labels
       usp_CollectWaits  collector
       usp_CheckWaitAlert alert check + Database Mail
     SQL Agent jobs
       DBAWait - Collect (every 10 seconds)
       DBAWait - Alert   (every minute)

   View it with SqlWaitMonitor.ps1 (Start-SqlWaitMonitor.cmd).
   ===================================================================== */
SET NOCOUNT ON;
USE master;
GO
IF DB_ID(N'DBAWait') IS NULL
BEGIN
    CREATE DATABASE DBAWait;
    ALTER DATABASE DBAWait SET RECOVERY SIMPLE;
END
GO
USE DBAWait;
GO

/* =====================================================================
   1. Tables
   ===================================================================== */
IF OBJECT_ID(N'dbo.MonitorConfig') IS NULL
CREATE TABLE dbo.MonitorConfig (
    ConfigId          INT            NOT NULL CONSTRAINT PK_MonitorConfig PRIMARY KEY CONSTRAINT CK_MonitorConfig_One CHECK (ConfigId = 1),
    IntervalSec       INT            NOT NULL,   -- must match the Collect job schedule
    WindowMin         INT            NOT NULL,   -- alert looks at the last N minutes
    WaitThresholdSec  INT            NOT NULL,   -- total wait in the window that raises an alert
    BlockedThreshold  INT            NOT NULL,   -- blocked sessions in the latest sample that raise an alert
    CooldownMin       INT            NOT NULL,   -- minimum minutes between alert emails
    MailProfile       SYSNAME        NULL,       -- Database Mail profile (NULL = log only, no email)
    Recipients        NVARCHAR(400)  NULL,
    RetentionDays     INT            NOT NULL
);
GO

IF OBJECT_ID(N'dbo.WaitSamples') IS NULL
CREATE TABLE dbo.WaitSamples (
    SampleId          BIGINT IDENTITY CONSTRAINT PK_WaitSamples PRIMARY KEY,
    SampleTime        DATETIME2(0)   NOT NULL,   -- UTC
    SessionId         INT            NOT NULL,
    DatabaseName      SYSNAME        NULL,
    LoginName         SYSNAME        NULL,
    HostName          NVARCHAR(128)  NULL,
    ProgramName       NVARCHAR(128)  NULL,
    WaitCategory      NVARCHAR(60)   NOT NULL,   -- CPU, or the wait type
    WaitResource      NVARCHAR(256)  NULL,
    BlockingSessionId INT            NULL,
    WeightSec         INT            NOT NULL,   -- seconds this sample represents
    QueryHash         BINARY(8)      NULL,
    SqlText           NVARCHAR(4000) NULL,
    StatementId       INT            NULL
);
GO
-- upgrade older DBAWait versions
IF COL_LENGTH(N'dbo.WaitSamples', N'StatementId') IS NULL ALTER TABLE dbo.WaitSamples ADD StatementId INT NULL;
IF COL_LENGTH(N'dbo.WaitSamples', N'ProgramName') IS NULL ALTER TABLE dbo.WaitSamples ADD ProgramName NVARCHAR(128) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.WaitSamples') AND name = N'IX_WaitSamples_Time')
    CREATE INDEX IX_WaitSamples_Time ON dbo.WaitSamples (SampleTime) INCLUDE (WaitCategory, WeightSec, QueryHash);
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.WaitSamples') AND name = N'IX_WaitSamples_Stmt')
    CREATE INDEX IX_WaitSamples_Stmt ON dbo.WaitSamples (StatementId, SampleTime)
        INCLUDE (WaitCategory, WeightSec, BlockingSessionId, SessionId);
GO

IF OBJECT_ID(N'dbo.SqlStatements') IS NULL
CREATE TABLE dbo.SqlStatements (
    StatementId       INT IDENTITY   CONSTRAINT PK_SqlStatements PRIMARY KEY,
    QueryHash         BINARY(8)      NOT NULL,
    DbId              INT            NOT NULL,
    ObjectId          INT            NOT NULL,   -- 0 = ad hoc / dynamic SQL
    DatabaseName      SYSNAME        NULL,
    ObjectName        NVARCHAR(300)  NULL,       -- schema.proc (NULL for ad hoc)
    StartLine         INT            NULL,       -- line inside the proc / batch
    StatementText     NVARCHAR(MAX)  NULL,
    LastCalledFrom    NVARCHAR(4000) NULL,       -- outer call (shows the proc behind dynamic SQL)
    LastPlanHandle    VARBINARY(64)  NULL,
    LastQueryPlanHash BINARY(8)      NULL,
    FirstSeen         DATETIME2(0)   NOT NULL,
    LastSeen          DATETIME2(0)   NOT NULL,
    CONSTRAINT UQ_SqlStatements UNIQUE (QueryHash, DbId, ObjectId)
);
GO

IF OBJECT_ID(N'dbo.BlockerSamples') IS NULL
CREATE TABLE dbo.BlockerSamples (
    SampleTime        DATETIME2(0)   NOT NULL,
    SessionId         INT            NOT NULL,
    Status            NVARCHAR(30)   NULL,       -- 'sleeping' = idle with open transaction
    LoginName         SYSNAME        NULL,
    HostName          NVARCHAR(128)  NULL,
    ProgramName       NVARCHAR(128)  NULL,
    OpenTranCount     INT            NULL,
    BlockedCount      INT            NULL,
    LastSqlText       NVARCHAR(4000) NULL,
    CONSTRAINT PK_BlockerSamples PRIMARY KEY (SampleTime, SessionId)
);
GO

IF OBJECT_ID(N'dbo.SchedulerSamples') IS NULL
CREATE TABLE dbo.SchedulerSamples (
    SampleTime        DATETIME2(0)   NOT NULL CONSTRAINT PK_SchedulerSamples PRIMARY KEY,
    CurrentWorkers    INT, ActiveWorkers INT, RunnableTasks INT, WorkQueueCount INT, MaxWorkers INT
);
GO

IF OBJECT_ID(N'dbo.AlertLog') IS NULL
CREATE TABLE dbo.AlertLog (
    AlertId           INT IDENTITY   CONSTRAINT PK_AlertLog PRIMARY KEY,
    AlertTime         DATETIME2(0)   NOT NULL CONSTRAINT DF_AlertLog_Time DEFAULT SYSUTCDATETIME(),
    AlertType         NVARCHAR(50),
    TotalWaitSec      INT,
    BlockedCount      INT,
    HeadBlocker       INT,
    Details           NVARCHAR(MAX)
);
GO
-- older versions defaulted AlertTime to local time; make it UTC like the rest
DECLARE @dc SYSNAME, @def NVARCHAR(MAX), @cmd NVARCHAR(400);
SELECT @dc = dc.name, @def = dc.definition
FROM sys.default_constraints dc
JOIN sys.columns c ON c.object_id = dc.parent_object_id AND c.column_id = dc.parent_column_id
WHERE dc.parent_object_id = OBJECT_ID(N'dbo.AlertLog') AND c.name = N'AlertTime';
IF @dc IS NOT NULL AND @def NOT LIKE N'%sysutcdatetime%'
BEGIN
    SET @cmd = N'ALTER TABLE dbo.AlertLog DROP CONSTRAINT ' + QUOTENAME(@dc) + N';';
    EXEC (@cmd);
    SET @dc = NULL;
END
IF @dc IS NULL
    ALTER TABLE dbo.AlertLog ADD CONSTRAINT DF_AlertLog_Time DEFAULT SYSUTCDATETIME() FOR AlertTime;
GO

/* =====================================================================
   2. SETTINGS - edit these values, then run the script
   ===================================================================== */
MERGE dbo.MonitorConfig AS t
USING (SELECT
        ConfigId         = 1,
        IntervalSec      = 10,
        WindowMin        = 10,
        WaitThresholdSec = 10000,
        BlockedThreshold = 20,
        CooldownMin      = 15,
        MailProfile      = CAST(NULL AS SYSNAME),            -- e.g. N'DBA Mail'  (NULL = no email, alerts only logged)
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

/* =====================================================================
   3. Readable statement label
   ===================================================================== */
CREATE OR ALTER VIEW dbo.vStatements
AS
SELECT s.StatementId, s.DatabaseName, s.ObjectName, s.StartLine, s.StatementText,
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
   4. Collector
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_CollectWaits
    @IntervalSec INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET LOCK_TIMEOUT 500;   -- never wait behind blocking: give up after 0.5 s and try again next run
    DECLARE @now DATETIME2(0) = SYSUTCDATETIME(), @keepDays INT;
    SELECT @IntervalSec = ISNULL(@IntervalSec, IntervalSec), @keepDays = RetentionDays FROM dbo.MonitorConfig;
    SET @IntervalSec = ISNULL(@IntervalSec, 10);
    SET @keepDays    = ISNULL(@keepDays, 14);

    /* active requests, cut down to the running statement */
    SELECT
        SessionId         = r.session_id,
        DatabaseName      = DB_NAME(r.database_id),
        LoginName         = s.login_name,
        HostName          = s.host_name,
        ProgramName       = s.program_name,
        WaitCategory      = CASE WHEN r.wait_type IS NULL THEN N'CPU' ELSE r.wait_type END,
        WaitResource      = r.wait_resource,
        BlockingSessionId = NULLIF(r.blocking_session_id, 0),
        QueryHash         = ISNULL(r.query_hash, CAST(HASHBYTES('SHA2_256', ISNULL(x.StmtText, r.command)) AS BINARY(8))),
        DbId              = ISNULL(t.dbid, r.database_id),
        ObjectId          = ISNULL(t.objectid, 0),
        StartLine         = x.StartLine,
        StmtText          = x.StmtText,
        CalledFrom        = LEFT(ib.event_info, 4000),
        PlanHandle        = r.plan_handle,
        QueryPlanHash     = r.query_plan_hash
    INTO #req
    FROM sys.dm_exec_requests r
    JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
    OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
    OUTER APPLY (
        SELECT StmtText  = SUBSTRING(t.text, r.statement_start_offset / 2 + 1,
                               (CASE WHEN r.statement_end_offset = -1 THEN DATALENGTH(t.text)
                                     ELSE r.statement_end_offset END - r.statement_start_offset) / 2 + 1),
               StartLine = (DATALENGTH(LEFT(t.text, r.statement_start_offset / 2))
                          - DATALENGTH(REPLACE(LEFT(t.text, r.statement_start_offset / 2), NCHAR(10), N''))) / 2 + 1
    ) x
    OUTER APPLY sys.dm_exec_input_buffer(r.session_id, r.request_id) ib
    WHERE s.is_user_process = 1
      AND r.session_id <> @@SPID
      AND ISNULL(r.wait_type, N'') NOT IN (N'WAITFOR', N'SP_SERVER_DIAGNOSTICS_SLEEP', N'BROKER_RECEIVE_WAITFOR');

    /* distinct statements in this snapshot */
    SELECT QueryHash, DbId, ObjectId,
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
    JOIN #stmt n ON n.QueryHash = s.QueryHash AND n.DbId = s.DbId AND n.ObjectId = s.ObjectId;

    INSERT dbo.SqlStatements (QueryHash, DbId, ObjectId, DatabaseName, ObjectName, StartLine,
                              StatementText, LastCalledFrom, LastPlanHandle, LastQueryPlanHash, FirstSeen, LastSeen)
    SELECT n.QueryHash, n.DbId, n.ObjectId, DB_NAME(n.DbId),
           CASE WHEN n.ObjectId <> 0 THEN CONCAT(OBJECT_SCHEMA_NAME(n.ObjectId, n.DbId), N'.', OBJECT_NAME(n.ObjectId, n.DbId)) END,
           n.StartLine, n.StmtText, n.CalledFrom, n.PlanHandle, n.QueryPlanHash, @now, @now
    FROM #stmt n
    WHERE NOT EXISTS (SELECT 1 FROM dbo.SqlStatements s
                      WHERE s.QueryHash = n.QueryHash AND s.DbId = n.DbId AND s.ObjectId = n.ObjectId);

    /* samples: each active request = @IntervalSec seconds of its current wait */
    INSERT dbo.WaitSamples (SampleTime, SessionId, DatabaseName, LoginName, HostName, ProgramName,
                            WaitCategory, WaitResource, BlockingSessionId, WeightSec, QueryHash, SqlText, StatementId)
    SELECT @now, r.SessionId, r.DatabaseName, r.LoginName, r.HostName, r.ProgramName,
           r.WaitCategory, r.WaitResource, r.BlockingSessionId, @IntervalSec, r.QueryHash,
           LEFT(r.StmtText, 4000), s.StatementId
    FROM #req r
    LEFT JOIN dbo.SqlStatements s ON s.QueryHash = r.QueryHash AND s.DbId = r.DbId AND s.ObjectId = r.ObjectId;

    /* head blockers: blocking others, not blocked themselves (may be idle with an open transaction) */
    INSERT dbo.BlockerSamples (SampleTime, SessionId, Status, LoginName, HostName, ProgramName, OpenTranCount, BlockedCount, LastSqlText)
    SELECT @now, b.BlockingSessionId, s.status, s.login_name, s.host_name, s.program_name,
           s.open_transaction_count, b.BlockedCount, LEFT(t.text, 4000)
    FROM (SELECT BlockingSessionId, BlockedCount = COUNT(*)
          FROM #req WHERE BlockingSessionId IS NOT NULL GROUP BY BlockingSessionId) b
    JOIN sys.dm_exec_sessions s ON s.session_id = b.BlockingSessionId
    OUTER APPLY (SELECT TOP (1) c.most_recent_sql_handle FROM sys.dm_exec_connections c WHERE c.session_id = b.BlockingSessionId) c
    OUTER APPLY sys.dm_exec_sql_text(c.most_recent_sql_handle) t
    WHERE NOT EXISTS (SELECT 1 FROM #req r2 WHERE r2.SessionId = b.BlockingSessionId AND r2.BlockingSessionId IS NOT NULL);

    /* worker threads / scheduler pressure */
    INSERT dbo.SchedulerSamples (SampleTime, CurrentWorkers, ActiveWorkers, RunnableTasks, WorkQueueCount, MaxWorkers)
    SELECT @now, SUM(current_workers_count), SUM(active_workers_count), SUM(runnable_tasks_count), SUM(work_queue_count),
           (SELECT max_workers_count FROM sys.dm_os_sys_info)
    FROM sys.dm_os_schedulers
    WHERE status = N'VISIBLE ONLINE';

    /* retention (small batches so the 10-second job stays quick) */
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
   5. Alert check (total wait in the window, or blocked sessions now)
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
    DECLARE @last DATETIME2(0) = (SELECT MAX(SampleTime) FROM dbo.WaitSamples);
    DECLARE @total INT, @blocked INT, @head INT, @body NVARCHAR(MAX), @top NVARCHAR(MAX);

    SELECT @total = ISNULL(SUM(WeightSec), 0) FROM dbo.WaitSamples WHERE SampleTime >= @from;
    SELECT @blocked = COUNT(DISTINCT SessionId) FROM dbo.WaitSamples WHERE SampleTime = @last AND BlockingSessionId IS NOT NULL;

    IF @total < @WaitThresholdSec AND @blocked < @BlockedThreshold RETURN;
    IF EXISTS (SELECT 1 FROM dbo.AlertLog WHERE AlertTime > DATEADD(MINUTE, -@CooldownMin, SYSUTCDATETIME())) RETURN;

    SELECT @top = STRING_AGG(CAST(CONCAT(x.Label, N'  |  ', x.WaitSec, N' s  |  mostly ', x.TopWait) AS NVARCHAR(MAX)), NCHAR(10))
                  WITHIN GROUP (ORDER BY x.WaitSec DESC)
    FROM (
        SELECT TOP (10) v.Label, WaitSec = SUM(w.WeightSec),
               TopWait = (SELECT TOP (1) w2.WaitCategory FROM dbo.WaitSamples w2
                          WHERE w2.StatementId = w.StatementId AND w2.SampleTime >= @from
                          GROUP BY w2.WaitCategory ORDER BY SUM(w2.WeightSec) DESC)
        FROM dbo.WaitSamples w
        JOIN dbo.vStatements v ON v.StatementId = w.StatementId
        WHERE w.SampleTime >= @from
        GROUP BY w.StatementId, v.Label
        ORDER BY SUM(w.WeightSec) DESC
    ) x;

    DECLARE @hb NVARCHAR(MAX);
    SELECT TOP (1) @head = SessionId,
           @hb = CONCAT(N'SPID ', SessionId, N'  ', LoginName, N' @ ', HostName, N'  (', Status, N', open tran ', OpenTranCount,
                        N', blocking ', BlockedCount, N')', NCHAR(10), N'Head blocker SQL: ', LEFT(LastSqlText, 1000))
    FROM dbo.BlockerSamples
    WHERE SampleTime >= @from
    ORDER BY SampleTime DESC, BlockedCount DESC;

    SET @body = CONCAT(
        N'Server: ', @@SERVERNAME, NCHAR(10),
        N'Total wait last ', @WindowMin, N' min: ', @total, N' s (threshold ', @WaitThresholdSec, N')', NCHAR(10),
        N'Blocked sessions in latest sample: ', @blocked, N' (threshold ', @BlockedThreshold, N')', NCHAR(10),
        N'Head blocker: ', ISNULL(@hb, N'none captured'), NCHAR(10), NCHAR(10),
        N'Top statements by wait (#id = statement in the viewer):', NCHAR(10), @top);

    INSERT dbo.AlertLog (AlertType, TotalWaitSec, BlockedCount, HeadBlocker, Details)
    VALUES (CASE WHEN @blocked >= @BlockedThreshold THEN N'Blocking' ELSE N'WaitTime' END, @total, @blocked, @head, @body);

    DECLARE @alertId INT = SCOPE_IDENTITY();
    IF @MailProfile IS NOT NULL AND @Recipients IS NOT NULL
    BEGIN
        BEGIN TRY
            EXEC msdb.dbo.sp_send_dbmail
                @profile_name = @MailProfile,
                @recipients   = @Recipients,
                @subject      = N'SQL wait alert: total wait / blocking threshold exceeded',
                @body         = @body;
        END TRY
        BEGIN CATCH
            UPDATE dbo.AlertLog SET Details = CONCAT(Details, NCHAR(10), N'[email failed: ', ERROR_MESSAGE(), N']')
            WHERE AlertId = @alertId;
        END CATCH
    END
END
GO

/* =====================================================================
   6. SQL Agent jobs (recreated on every run)
   ===================================================================== */
USE msdb;
GO
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBAWait - Collect') EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWait - Collect';
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBAWait - Alert')   EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWait - Alert';

EXEC msdb.dbo.sp_add_job        @job_name = N'DBAWait - Collect', @description = N'SQL Wait Monitor: sample waits every 10 seconds';
EXEC msdb.dbo.sp_add_jobstep    @job_name = N'DBAWait - Collect', @step_name = N'Collect', @subsystem = N'TSQL',
                                @database_name = N'DBAWait', @command = N'EXEC dbo.usp_CollectWaits;';
EXEC msdb.dbo.sp_add_jobschedule @job_name = N'DBAWait - Collect', @name = N'DBAWait every 10 seconds',
                                @freq_type = 4, @freq_interval = 1, @freq_subday_type = 2, @freq_subday_interval = 10;
EXEC msdb.dbo.sp_add_jobserver  @job_name = N'DBAWait - Collect';

EXEC msdb.dbo.sp_add_job        @job_name = N'DBAWait - Alert', @description = N'SQL Wait Monitor: check thresholds every minute';
EXEC msdb.dbo.sp_add_jobstep    @job_name = N'DBAWait - Alert', @step_name = N'Check', @subsystem = N'TSQL',
                                @database_name = N'DBAWait', @command = N'EXEC dbo.usp_CheckWaitAlert;';
EXEC msdb.dbo.sp_add_jobschedule @job_name = N'DBAWait - Alert', @name = N'DBAWait every minute',
                                @freq_type = 4, @freq_interval = 1, @freq_subday_type = 4, @freq_subday_interval = 1;
EXEC msdb.dbo.sp_add_jobserver  @job_name = N'DBAWait - Alert';
GO

/* =====================================================================
   7. Check
   ===================================================================== */
USE DBAWait;
GO
EXEC dbo.usp_CollectWaits;
SELECT Installed = N'SQL Wait Monitor (single server) installed',
       Samples   = (SELECT COUNT(*) FROM dbo.WaitSamples),
       Jobs      = (SELECT COUNT(*) FROM msdb.dbo.sysjobs WHERE name LIKE N'DBAWait - %');
GO

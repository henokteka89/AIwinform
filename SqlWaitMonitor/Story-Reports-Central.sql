/* =====================================================================
   SQL Wait Monitor - STORY reports  (central server, database DBAWaitCentral)

   Plain-English answers from the monitoring tables. Run once in SSMS
   (after the install script). Safe to run again.

   How to read the stories
     In SSMS press Ctrl+T (Results to Text) before running, then F5.
     Each row is one sentence or paragraph, grouped into chapters.

   1) What happened in a period?
        EXEC DBAWaitCentral.dbo.usp_WaitStory;                                   -- last 24 hours
        EXEC DBAWaitCentral.dbo.usp_WaitStory @From = '2026-11-02 09:00', @To = '2026-11-02 17:00';
        EXEC DBAWaitCentral.dbo.usp_WaitStory @ServerName = N'PRODSQL01';       -- pick the server (optional when only one is monitored)
        EXEC DBAWaitCentral.dbo.usp_WaitStory @From = '2026-10-26', @To = '2026-11-02';   -- a week, with a day-by-day chapter
      Times are the central server's local time.

   2) Tell me about one query (use the #id from the viewer or from the story):
        EXEC DBAWaitCentral.dbo.usp_WaitStoryForStatement @StatementId = 1;              -- last 7 days
        EXEC DBAWaitCentral.dbo.usp_WaitStoryForStatement @StatementId = 1, @Days = 30;

   3) What does a wait type mean?
        SELECT * FROM DBAWaitCentral.dbo.WaitGlossary ORDER BY SortOrder;
        SELECT * FROM DBAWaitCentral.dbo.fn_WaitGroup(N'LCK_M_IS');
   ===================================================================== */
USE DBAWaitCentral;
GO

/* =====================================================================
   1. Glossary: wait types in plain English (edit the wording freely)
      The first pattern (lowest SortOrder) that matches a wait type wins.
   ===================================================================== */
IF OBJECT_ID(N'dbo.WaitGlossary') IS NULL
CREATE TABLE dbo.WaitGlossary (
    SortOrder  INT            NOT NULL CONSTRAINT PK_WaitGlossary PRIMARY KEY,
    Pattern    NVARCHAR(60)   NOT NULL,   -- LIKE pattern on the wait type
    GroupName  NVARCHAR(60)   NOT NULL,   -- short plain-English name
    Meaning    NVARCHAR(400)  NOT NULL,   -- what the query was doing
    Advice     NVARCHAR(600)  NOT NULL    -- where to look when this is high
);
GO
DELETE dbo.WaitGlossary;
INSERT dbo.WaitGlossary (SortOrder, Pattern, GroupName, Meaning, Advice) VALUES
 (10,  N'CPU',                     N'Working (CPU)',
       N'the query was actually running on a processor (or ready to run)',
       N'Normal in itself. If it dominates, the heaviest queries are doing a lot of work: look for missing indexes, large scans and big sorts in their plans.'),
 (20,  N'LCK[_]M[_]SCH%',          N'Blocked by a table change',
       N'the query was waiting because someone was changing the table structure or rebuilding an index (schema lock)',
       N'Run index rebuilds ONLINE or in a maintenance window, never inside an open transaction; avoid schema changes during business hours.'),
 (30,  N'LCK[_]M[_]%',             N'Blocked by another session',
       N'the query was waiting for another session to finish and release its lock',
       N'Find the head blocker (chapter "Blocking incidents"). Keep transactions short, make sure applications always commit or roll back, and consider READ_COMMITTED_SNAPSHOT so readers stop waiting for writers.'),
 (40,  N'PAGEIOLATCH[_]%',         N'Reading from disk',
       N'the query was waiting for data to be read from disk into memory',
       N'The query reads a lot of data (often a scan that an index could avoid), or the server is short of memory, or the disk is slow. Check the plans of the top queries and the disk latency.'),
 (50,  N'WRITELOG',                N'Saving changes to the log',
       N'the query was waiting for its changes to be written to the transaction log on commit',
       N'Many tiny transactions (commit in batches instead), or a slow log disk.'),
 (60,  N'ASYNC[_]NETWORK[_]IO',    N'Waiting for the application',
       N'SQL Server had the results ready but the application was slow to read them',
       N'Usually the application: it reads rows one by one while doing other work, or asks for far more rows than it needs. Sometimes the network.'),
 (70,  N'CX%',                     N'Parallel query coordination',
       N'parts of a parallel query were waiting for each other',
       N'A symptom of big queries going parallel. Tune the top queries first; then review MAXDOP and "cost threshold for parallelism".'),
 (80,  N'SOS[_]SCHEDULER[_]YIELD', N'Queuing for CPU',
       N'the query gave up the processor and waited for its next turn',
       N'The CPUs are busy. Reduce the work of the top CPU queries.'),
 (90,  N'RESOURCE[_]SEMAPHORE%',   N'Waiting for memory',
       N'the query was waiting for enough memory to run a sort or hash join',
       N'Queries are asking for large memory grants - often because of bad row estimates. Update statistics and check the plans of the top queries.'),
 (100, N'THREADPOOL',              N'Out of worker threads',
       N'the query could not even start because no worker thread was free',
       N'Almost always the result of heavy blocking or a flood of connections. Fix the blocking first.'),
 (110, N'PAGELATCH[_]%',           N'Hot spot in memory',
       N'many sessions were competing for the same data page in memory (often tempdb, or inserts at the end of a table)',
       N'Check tempdb configuration (number of data files) and tables with heavy inserts on an ever-increasing key.'),
 (120, N'HADR[_]SYNC[_]COMMIT',    N'Waiting for the secondary replica',
       N'the commit waited for an Availability Group secondary to confirm it',
       N'Network or disk speed on the secondary, or a large volume of changes.'),
 (130, N'OLEDB',                   N'Waiting for a linked server',
       N'the query was waiting for a linked server or external data source',
       N'Look at the remote query: it may return many rows or the remote server may be slow.'),
 (140, N'%MEMORY[_]ALLOCATION%',   N'Getting memory',
       N'the query was allocating memory (usually very short)',
       N'Only a concern when it is large; then check memory pressure on the server.'),
 (150, N'%IO[_]COMPLETION',        N'Other disk work',
       N'the query was waiting for other disk I/O (sorts or hashes spilling to disk, backups, file growth)',
       N'Check tempdb spills in plans, and file autogrowth settings.'),
 (160, N'LATCH[_]%',               N'Internal contention',
       N'SQL Server internal structures were busy',
       N'Usually small. If large, look up the specific wait type.'),
 (165, N'PREEMPTIVE[_]%',          N'Waiting on Windows',
       N'SQL Server was waiting for Windows to finish a call (files, network, security or registry)',
       N'Usually small. If large, look at what the query calls outside SQL Server (xp_cmdshell, linked servers, file access, encryption).'),
 (170, N'BACKUP%',                 N'Backups',
       N'waiting on backup activity',
       N'Check backup schedules against busy hours.'),
 (999, N'%',                       N'Other',
       N'a less common internal wait',
       N'Look up the specific wait type if it is large.');
GO

CREATE OR ALTER FUNCTION dbo.fn_WaitGroup (@WaitType NVARCHAR(60))
RETURNS TABLE
AS RETURN
    SELECT TOP (1) GroupName, Meaning, Advice, SortOrder
    FROM dbo.WaitGlossary
    WHERE @WaitType LIKE Pattern
    ORDER BY SortOrder;
GO

/* =====================================================================
   2. Small helpers that turn numbers and names into words
   ===================================================================== */
CREATE OR ALTER FUNCTION dbo.fn_Duration (@Sec BIGINT)
RETURNS NVARCHAR(40)
AS
BEGIN
    RETURN CASE
        WHEN @Sec IS NULL OR @Sec <= 0 THEN N'0 s'
        WHEN @Sec < 60    THEN CONCAT(@Sec, N' s')
        WHEN @Sec < 600   THEN CONCAT(@Sec / 60, N' min', CASE WHEN @Sec % 60 > 0 THEN CONCAT(N' ', @Sec % 60, N' s') ELSE N'' END)
        WHEN @Sec < 3600  THEN CONCAT(@Sec / 60, N' min')
        WHEN @Sec < 86400 THEN CONCAT(@Sec / 3600, N' h', CASE WHEN (@Sec % 3600) / 60 > 0 THEN CONCAT(N' ', (@Sec % 3600) / 60, N' min') ELSE N'' END)
        ELSE CONCAT(@Sec / 86400, N' d', CASE WHEN (@Sec % 86400) / 3600 > 0 THEN CONCAT(N' ', (@Sec % 86400) / 3600, N' h') ELSE N'' END)
    END;
END
GO

CREATE OR ALTER FUNCTION dbo.fn_OneLine (@Text NVARCHAR(MAX), @Max INT)
RETURNS NVARCHAR(4000)
AS
BEGIN
    DECLARE @t NVARCHAR(MAX) = LTRIM(RTRIM(REPLACE(REPLACE(REPLACE(ISNULL(@Text, N''), NCHAR(13), N' '), NCHAR(10), N' '), NCHAR(9), N' ')));
    WHILE CHARINDEX(N'  ', @t) > 0
    BEGIN
        SET @t = REPLACE(@t, N'  ', N' ');
    END
    RETURN CASE WHEN LEN(@t) > @Max THEN LEFT(@t, @Max - 3) + N'...' ELSE @t END;
END
GO

CREATE OR ALTER FUNCTION dbo.fn_ProgramName (@Program NVARCHAR(128))
RETURNS NVARCHAR(300)
AS
BEGIN
    RETURN CASE
        WHEN @Program IS NULL OR @Program = N'' THEN N'an unnamed application'
        WHEN @Program LIKE N'SQLAgent - TSQL JobStep%' THEN N'a SQL Agent job'
        WHEN @Program LIKE N'SQLAgent%' THEN N'SQL Agent'
        WHEN @Program LIKE N'Microsoft SQL Server Management Studio%' THEN N'SSMS'
        WHEN @Program LIKE N'%Azure Data Studio%' THEN N'Azure Data Studio'
        WHEN @Program LIKE N'SQL Wait Monitor%' THEN N'the wait monitor viewer'
        ELSE CONCAT(N'"', @Program, N'"')
    END;
END
GO

CREATE OR ALTER FUNCTION dbo.fn_StatementName (@StatementId INT)
RETURNS NVARCHAR(400)
AS
BEGIN
    RETURN (
        SELECT CONCAT(
                 CASE WHEN s.ObjectId <> 0
                      THEN CONCAT(ISNULL(s.ObjectName, N'a stored procedure'), N' (line ', s.StartLine, N', database ', s.DatabaseName, N')')
                      ELSE CONCAT(N'the ad hoc query "', dbo.fn_OneLine(s.StatementText, 70), N'" (database ', s.DatabaseName, N')')
                 END,
                 N' [#', s.StatementId, N']')
        FROM dbo.SqlStatements s
        WHERE s.StatementId = @StatementId);
END
GO

-- UTC <-> local time, daylight saving aware (@TzName = Windows time zone name, NULL = use today's offset)
CREATE OR ALTER FUNCTION dbo.fn_Local (@TzName SYSNAME, @Utc DATETIME2(0))
RETURNS DATETIME2(0)
AS
BEGIN
    RETURN CASE WHEN @TzName IS NULL THEN DATEADD(MINUTE, DATEDIFF(MINUTE, SYSUTCDATETIME(), SYSDATETIME()), @Utc)
                ELSE CONVERT(DATETIME2(0), (@Utc AT TIME ZONE 'UTC') AT TIME ZONE @TzName) END;
END
GO

CREATE OR ALTER FUNCTION dbo.fn_Utc (@TzName SYSNAME, @Local DATETIME2(0))
RETURNS DATETIME2(0)
AS
BEGIN
    RETURN CASE WHEN @TzName IS NULL THEN DATEADD(MINUTE, -DATEDIFF(MINUTE, SYSUTCDATETIME(), SYSDATETIME()), @Local)
                ELSE CONVERT(DATETIME2(0), (@Local AT TIME ZONE @TzName) AT TIME ZONE 'UTC') END;
END
GO

-- the server's Windows time zone (SQL Server 2022: CURRENT_TIMEZONE_ID; older: registry; else NULL)
CREATE OR ALTER PROCEDURE dbo.usp_GetTimeZone
    @TzName SYSNAME OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @TzName = NULL;
    BEGIN TRY
        EXEC sys.sp_executesql N'SELECT @n = CURRENT_TIMEZONE_ID();', N'@n SYSNAME OUTPUT', @n = @TzName OUTPUT;
    END TRY BEGIN CATCH END CATCH
    IF @TzName IS NULL
    BEGIN TRY
        EXEC master.dbo.xp_regread N'HKEY_LOCAL_MACHINE', N'SYSTEM\CurrentControlSet\Control\TimeZoneInformation',
                                   N'TimeZoneKeyName', @TzName OUTPUT;
    END TRY BEGIN CATCH END CATCH
    IF @TzName IS NOT NULL AND NOT EXISTS (SELECT 1 FROM sys.time_zone_info WHERE name = @TzName) SET @TzName = NULL;
END
GO

/* =====================================================================
   3. The story of a period
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_WaitStory
    @From DATETIME2(0) = NULL,      -- central server's local time; default = 24 hours before @To
    @To   DATETIME2(0) = NULL,
    @ServerName SYSNAME = NULL       -- which monitored server; optional when only one is enabled
AS
BEGIN
    SET NOCOUNT ON;
    IF @ServerName IS NULL AND (SELECT COUNT(*) FROM dbo.MonitoredServers WHERE IsEnabled = 1) = 1
        SELECT @ServerName = ServerName FROM dbo.MonitoredServers WHERE IsEnabled = 1;
    IF @ServerName IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.MonitoredServers WHERE ServerName = @ServerName)
    BEGIN
        SELECT Chapter = N'Which server?',
               Story   = CONCAT(N'Please pass @ServerName. Monitored servers: ',
                                (SELECT STRING_AGG(ServerName, N', ') FROM dbo.MonitoredServers), N'.');
        RETURN;
    END
    DECLARE @srvLabel NVARCHAR(128) = @ServerName;

    /* ---------- the period ---------- */
    DECLARE @tzName SYSNAME;
    EXEC dbo.usp_GetTimeZone @TzName = @tzName OUTPUT;
    IF @To   IS NULL SET @To   = CAST(SYSDATETIME() AS DATETIME2(0));
    IF @From IS NULL SET @From = DATEADD(HOUR, -24, @To);
    IF @From >= @To
    BEGIN
        SELECT Chapter = N'Oops', Story = N'@From must be earlier than @To.';
        RETURN;
    END
    DECLARE @fromU  DATETIME2(0) = dbo.fn_Utc(@tzName, @From),
            @toU    DATETIME2(0) = dbo.fn_Utc(@tzName, @To);
    DECLARE @lenSec BIGINT = DATEDIFF(SECOND, @fromU, @toU);
    DECLARE @prevFromU DATETIME2(0) = DATEADD(SECOND, -@lenSec, @fromU);
    DECLARE @iv INT = ISNULL((SELECT IntervalSec FROM dbo.MonitorConfig), 10);
    DECLARE @win INT = CASE WHEN @lenSec <= 6 * 3600 THEN 300 WHEN @lenSec <= 2 * 86400 THEN 900 ELSE 3600 END;
    DECLARE @fmt NVARCHAR(30) = CASE WHEN CAST(@From AS DATE) = CAST(DATEADD(SECOND, -1, @To) AS DATE) THEN N'HH:mm' ELSE N'ddd d MMM HH:mm' END;
    DECLARE @periodText NVARCHAR(200) = CONCAT(FORMAT(@From, N'ddd d MMM HH:mm'), N' and ', FORMAT(@To, @fmt), N' (', dbo.fn_Duration(@lenSec), N')');

    CREATE TABLE #story (Seq INT IDENTITY, Chapter NVARCHAR(60), Story NVARCHAR(MAX));

    /* ---------- the raw material ---------- */
    SELECT w.SampleTime, w.SessionId, w.WaitCategory, w.WeightSec, w.StatementId,
           w.BlockingSessionId, w.LoginName, w.HostName, w.ProgramName
    INTO #w
    FROM dbo.WaitSamples w
    WHERE w.SampleTime >= @fromU AND w.SampleTime < @toU AND w.ServerName = @ServerName;

    DECLARE @total BIGINT = ISNULL((SELECT SUM(WeightSec) FROM #w), 0);
    DECLARE @prev  BIGINT = ISNULL((SELECT SUM(WeightSec) FROM dbo.WaitSamples
                                    WHERE SampleTime >= @prevFromU AND SampleTime < @fromU AND ServerName = @ServerName), 0);
    DECLARE @snapshots INT = (SELECT COUNT(*) FROM dbo.SchedulerSamples WHERE SampleTime >= @fromU AND SampleTime < @toU AND ServerName = @ServerName);
    DECLARE @prevSnapshots INT = (SELECT COUNT(*) FROM dbo.SchedulerSamples WHERE SampleTime >= @prevFromU AND SampleTime < @fromU AND ServerName = @ServerName);
    DECLARE @coverage DECIMAL(9,1) = CASE WHEN @lenSec > 0 THEN 100.0 * @snapshots * @iv / @lenSec END;
    IF @coverage > 100 SET @coverage = 100;

    IF @snapshots = 0
    BEGIN
        INSERT #story (Chapter, Story) VALUES (N'1. In short',
            CONCAT(N'Nothing was recorded for ', @srvLabel, N' between ', @periodText,
                   N'. Either the server was off, or the collector job ("DBAWaitCentral - Collect") was not running.'));
        SELECT Chapter, Story FROM #story ORDER BY Seq;
        RETURN;
    END

    /* wait time per plain-English group */
    SELECT g.GroupName, g.Meaning, g.Advice, g.SortOrder,
           Sec = SUM(x.Sec),
           Types = STRING_AGG(CAST(CONCAT(x.WaitCategory, N' ', dbo.fn_Duration(x.Sec)) AS NVARCHAR(MAX)), N', ')
                   WITHIN GROUP (ORDER BY x.Sec DESC)
    INTO #groups
    FROM (SELECT WaitCategory, Sec = SUM(WeightSec) FROM #w GROUP BY WaitCategory) x
    CROSS APPLY dbo.fn_WaitGroup(x.WaitCategory) g
    GROUP BY g.GroupName, g.Meaning, g.Advice, g.SortOrder;

    DECLARE @topGroup NVARCHAR(60), @topGroupMeaning NVARCHAR(400), @topGroupPct INT;
    SELECT TOP (1) @topGroup = GroupName, @topGroupMeaning = Meaning, @topGroupPct = CAST(100.0 * Sec / NULLIF(@total, 0) AS INT)
    FROM #groups ORDER BY Sec DESC;

    /* busiest single moment (most queries active in one snapshot) */
    DECLARE @peakTime DATETIME2(0), @peakCount INT;
    SELECT TOP (1) @peakTime = dbo.fn_Local(@tzName, SampleTime), @peakCount = COUNT(*)
    FROM #w GROUP BY SampleTime ORDER BY COUNT(*) DESC, SampleTime;

    /* ---------- chapter 1: in short ---------- */
    INSERT #story (Chapter, Story) VALUES (N'1. In short',
        CONCAT(N'Between ', @periodText, N', the queries on ', @srvLabel, N' were busy for a combined ',
               dbo.fn_Duration(@total), N' (running or waiting). That is like ',
               FORMAT(1.0 * @total / NULLIF(@lenSec, 0), N'0.0'), N' queries being active every single second.'));

    INSERT #story (Chapter, Story) VALUES (N'1. In short',
        CASE WHEN @prevSnapshots = 0
                  THEN N'There is no data for the period just before, so there is nothing to compare with yet.'
             WHEN @prev = 0
                  THEN N'In the same length of time just before, the server was completely quiet, so this period was busier.'
             WHEN 1.0 * @total / @prev >= 1.5
                  THEN CONCAT(N'That is ', FORMAT(1.0 * @total / @prev, N'0.0'), N' times as much as the ', dbo.fn_Duration(@lenSec),
                              N' before it (', dbo.fn_Duration(@prev), N') - a noticeably busier or slower period.')
             WHEN 1.0 * @total / @prev <= 0.67
                  THEN CONCAT(N'That is less than the ', dbo.fn_Duration(@lenSec), N' before it (', dbo.fn_Duration(@prev),
                              N') - a quieter period.')
             ELSE CONCAT(N'That is about the same as the ', dbo.fn_Duration(@lenSec), N' before it (', dbo.fn_Duration(@prev), N').')
        END);

    IF @topGroup IS NOT NULL
        INSERT #story (Chapter, Story) VALUES (N'1. In short',
            CONCAT(N'The biggest share of that time (', @topGroupPct, N'%) was "', @topGroup, N'": ', @topGroupMeaning, N'.'));

    IF @peakCount IS NOT NULL
        INSERT #story (Chapter, Story) VALUES (N'1. In short',
            CONCAT(N'The busiest single moment was ', FORMAT(@peakTime, @fmt), N', when ', @peakCount,
                   CASE WHEN @peakCount = 1 THEN N' query was' ELSE N' queries were' END, N' active at the same time.'));

    IF @coverage < 90
        INSERT #story (Chapter, Story) VALUES (N'1. In short',
            CONCAT(N'Note: the collector only covered about ', CAST(@coverage AS INT), N'% of this period (the server or SQL Agent was off at times), ',
                   N'so the real totals are higher than shown.'));

    /* ---------- chapter 2: where the time went ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (6) N'2. Where the time went',
           CONCAT(CAST(100.0 * Sec / NULLIF(@total, 0) AS INT), N'%  ', GroupName, N' - ', dbo.fn_Duration(Sec), N'. ',
                  UPPER(LEFT(Meaning, 1)), SUBSTRING(Meaning, 2, 400), N'. (', Types, N')')
    FROM #groups
    ORDER BY Sec DESC;

    /* ---------- chapter 3: busiest moments ---------- */
    SELECT Bucket = DATEADD(SECOND, DATEDIFF(SECOND, '20000101', SampleTime) / @win * @win, CAST('20000101' AS DATETIME2(0))),
           WaitCategory, StatementId, WeightSec
    INTO #b
    FROM #w;

    DECLARE @avgWin DECIMAL(18,2) = 1.0 * @total / NULLIF(1.0 * @lenSec / @win, 0);

    INSERT #story (Chapter, Story)
    SELECT N'3. Busiest moments',
           CONCAT(FORMAT(dbo.fn_Local(@tzName, t.Bucket), @fmt), N' - ', FORMAT(DATEADD(SECOND, @win, dbo.fn_Local(@tzName, t.Bucket)), N'HH:mm'), N': ',
                  dbo.fn_Duration(t.Sec), N' of query time in ', @win / 60, N' minutes',
                  CASE WHEN @avgWin > 0 AND t.Sec / @avgWin >= 1.5 THEN CONCAT(N' (', FORMAT(t.Sec / @avgWin, N'0.0'), N' times the average for this period)') ELSE N'' END,
                  N'. Mostly "', tg.GroupName, N'" (', CAST(100.0 * tg.Sec / t.Sec AS INT), N'%).',
                  CASE WHEN ts.StatementId IS NOT NULL THEN CONCAT(N' Hardest hit: ', dbo.fn_StatementName(ts.StatementId), N'.') ELSE N'' END)
    FROM (SELECT TOP (3) Bucket, Sec = SUM(WeightSec) FROM #b GROUP BY Bucket ORDER BY SUM(WeightSec) DESC) t
    OUTER APPLY (SELECT TOP (1) g.GroupName, Sec = SUM(b.WeightSec)
                 FROM #b b CROSS APPLY dbo.fn_WaitGroup(b.WaitCategory) g
                 WHERE b.Bucket = t.Bucket GROUP BY g.GroupName ORDER BY SUM(b.WeightSec) DESC) tg
    OUTER APPLY (SELECT TOP (1) b.StatementId
                 FROM #b b WHERE b.Bucket = t.Bucket AND b.StatementId IS NOT NULL
                 GROUP BY b.StatementId ORDER BY SUM(b.WeightSec) DESC) ts
    ORDER BY t.Sec DESC;

    /* ---------- chapter 4: the queries that waited most ---------- */
    INSERT #story (Chapter, Story)
    SELECT N'4. Queries that waited the most',
           CONCAT(ROW_NUMBER() OVER (ORDER BY q.Sec DESC), N'. ', dbo.fn_StatementName(q.StatementId), N': ',
                  dbo.fn_Duration(q.Sec), N' in total (', CAST(100.0 * q.Sec / NULLIF(@total, 0) AS INT), N'% of everything). ',
                  CAST(100.0 * qg.Sec / q.Sec AS INT), N'% of its time was "', qg.GroupName, N'". ',
                  N'It was seen in ', q.Sessions, CASE WHEN q.Sessions = 1 THEN N' session' ELSE N' sessions' END,
                  N', mostly from ', dbo.fn_ProgramName(qp.ProgramName), N' (', qp.LoginName, N').')
    FROM (SELECT TOP (5) StatementId, Sec = SUM(WeightSec), Sessions = COUNT(DISTINCT SessionId)
          FROM #w WHERE StatementId IS NOT NULL
          GROUP BY StatementId ORDER BY SUM(WeightSec) DESC) q
    OUTER APPLY (SELECT TOP (1) g.GroupName, Sec = SUM(w.WeightSec)
                 FROM #w w CROSS APPLY dbo.fn_WaitGroup(w.WaitCategory) g
                 WHERE w.StatementId = q.StatementId GROUP BY g.GroupName ORDER BY SUM(w.WeightSec) DESC) qg
    OUTER APPLY (SELECT TOP (1) w.ProgramName, w.LoginName
                 FROM #w w WHERE w.StatementId = q.StatementId
                 GROUP BY w.ProgramName, w.LoginName ORDER BY SUM(w.WeightSec) DESC) qp
    ORDER BY q.Sec DESC;

    /* ---------- chapter 5: blocking incidents ---------- */
    -- consecutive head-blocker snapshots of the same session (gaps of up to 3 snapshots allowed) = one incident
    SELECT b.SessionId, b.SampleTime, b.Status, b.LoginName, b.HostName, b.ProgramName, b.OpenTranCount, b.BlockedCount, b.LastSqlText,
           NewIncident = CASE WHEN DATEDIFF(SECOND, LAG(b.SampleTime) OVER (PARTITION BY b.SessionId ORDER BY b.SampleTime), b.SampleTime) <= 3 * @iv
                              THEN 0 ELSE 1 END
    INTO #bs
    FROM dbo.BlockerSamples b
    WHERE b.SampleTime >= @fromU AND b.SampleTime < @toU AND b.ServerName = @ServerName;

    SELECT SessionId, SampleTime, Status, LoginName, HostName, ProgramName, OpenTranCount, BlockedCount, LastSqlText,
           IncidentNo = SUM(NewIncident) OVER (PARTITION BY SessionId ORDER BY SampleTime ROWS UNBOUNDED PRECEDING)
    INTO #bi
    FROM #bs;

    SELECT i.SessionId, i.IncidentNo,
           StartU = MIN(i.SampleTime), EndU = MAX(i.SampleTime),
           MaxBlocked = MAX(i.BlockedCount), MaxOpenTran = MAX(i.OpenTranCount),
           WasIdle = MAX(CASE WHEN i.Status = N'sleeping' THEN 1 ELSE 0 END),
           LoginName = MAX(i.LoginName), HostName = MAX(i.HostName), ProgramName = MAX(i.ProgramName),
           LastSql = MAX(i.LastSqlText)
    INTO #inc0
    FROM #bi i
    GROUP BY i.SessionId, i.IncidentNo;

    -- how much waiting each incident caused (time its direct victims spent blocked by it)
    SELECT n.*, Caused = v.Sec, Victims = v.Sessions
    INTO #inc
    FROM #inc0 n
    OUTER APPLY (SELECT Sec = SUM(w.WeightSec), Sessions = COUNT(DISTINCT w.SessionId)
                 FROM #w w
                 WHERE w.BlockingSessionId = n.SessionId
                   AND w.SampleTime BETWEEN n.StartU AND n.EndU) v;

    IF NOT EXISTS (SELECT 1 FROM #inc)
        INSERT #story (Chapter, Story) VALUES (N'5. Blocking incidents', N'No blocking was recorded in this period. Queries did not get stuck behind each other.');
    ELSE
    BEGIN
        INSERT #story (Chapter, Story)
        SELECT TOP (5) N'5. Blocking incidents',
               CONCAT(FORMAT(dbo.fn_Local(@tzName, n.StartU), @fmt), N' - ', FORMAT(dbo.fn_Local(@tzName, n.EndU), @fmt),
                      N' (', dbo.fn_Duration(DATEDIFF(SECOND, n.StartU, n.EndU) + @iv), N'): session ', n.SessionId,
                      N' (', n.LoginName, N', using ', dbo.fn_ProgramName(n.ProgramName), N' on ', n.HostName, N') was holding locks that other queries needed. ',
                      CASE WHEN n.WasIdle = 1 AND n.MaxOpenTran > 0
                           THEN N'It was sitting idle with an unfinished transaction, so everything behind it had to wait until someone committed or rolled it back. '
                           WHEN n.WasIdle = 1
                           THEN N'It was sitting idle while still holding its locks. '
                           ELSE N'It was busy running its own work while holding them. ' END,
                      N'Up to ', n.MaxBlocked, CASE WHEN n.MaxBlocked = 1 THEN N' query was' ELSE N' queries were' END,
                      N' stuck behind it at once (', ISNULL(n.Victims, 0), N' different sessions), losing ', dbo.fn_Duration(n.Caused), N' in total. ',
                      N'Its last command: "', dbo.fn_OneLine(n.LastSql, 150), N'".')
        FROM #inc n
        ORDER BY ISNULL(n.Caused, 0) DESC, n.MaxBlocked DESC;

        INSERT #story (Chapter, Story)
        SELECT N'5. Blocking incidents',
               CONCAT(N'In total there were ', COUNT(*), CASE WHEN COUNT(*) = 1 THEN N' blocking incident' ELSE N' blocking incidents' END,
                      N', costing ', dbo.fn_Duration(SUM(ISNULL(Caused, 0))), N' of waiting',
                      CASE WHEN @total > 0 THEN CONCAT(N' (', CAST(100.0 * SUM(ISNULL(Caused, 0)) / @total AS INT), N'% of all query time)') ELSE N'' END,
                      CASE WHEN COUNT(*) > 5 THEN N'. The five worst are listed above.' ELSE N'.' END)
        FROM #inc;
    END

    /* ---------- chapter 6: server capacity ---------- */
    DECLARE @maxActive INT, @maxWorkers INT, @maxRunnable INT, @avgRunnable DECIMAL(9,2), @maxRunnableAt DATETIME2(0);
    SELECT @maxActive = MAX(ActiveWorkers), @maxWorkers = MAX(MaxWorkers), @avgRunnable = AVG(1.0 * RunnableTasks), @maxRunnable = MAX(RunnableTasks)
    FROM dbo.SchedulerSamples WHERE SampleTime >= @fromU AND SampleTime < @toU AND ServerName = @ServerName;
    SELECT TOP (1) @maxRunnableAt = dbo.fn_Local(@tzName, SampleTime)
    FROM dbo.SchedulerSamples WHERE SampleTime >= @fromU AND SampleTime < @toU AND ServerName = @ServerName
    ORDER BY RunnableTasks DESC, SampleTime;

    INSERT #story (Chapter, Story) VALUES (N'6. Server capacity',
        CONCAT(N'Worker threads (the server''s "hands" for running queries): at most ', @maxActive, N' of ', @maxWorkers, N' were busy (',
               CAST(100.0 * @maxActive / NULLIF(@maxWorkers, 0) AS INT), N'%) - ',
               CASE WHEN 100.0 * @maxActive / NULLIF(@maxWorkers, 0) < 50 THEN N'plenty of room.'
                    WHEN 100.0 * @maxActive / NULLIF(@maxWorkers, 0) < 80 THEN N'fine, but keep an eye on it.'
                    ELSE N'close to the limit: new queries may not get a thread and simply hang (THREADPOOL). This usually follows heavy blocking.' END));

    INSERT #story (Chapter, Story) VALUES (N'6. Server capacity',
        CASE WHEN ISNULL(@maxRunnable, 0) = 0
             THEN N'Processors: queries never had to queue for a CPU - the processors kept up.'
             ELSE CONCAT(N'Processors: at the worst moment (', FORMAT(@maxRunnableAt, @fmt), N') ', @maxRunnable,
                         CASE WHEN @maxRunnable = 1 THEN N' task was' ELSE N' tasks were' END, N' queuing for a CPU; on average ',
                         FORMAT(@avgRunnable, N'0.0'), N'. ',
                         CASE WHEN @avgRunnable >= 1 THEN N'That means the CPUs were regularly overloaded.'
                              WHEN @maxRunnable >= 5 THEN N'Short CPU bursts, not a constant problem.'
                              ELSE N'Only brief, minor queuing.' END) END);

    /* ---------- chapter 7: alerts ---------- */
    INSERT #story (Chapter, Story)
    SELECT N'7. Alerts',
           CASE WHEN COUNT(*) = 0 THEN N'No alerts were raised in this period.'
                ELSE CONCAT(COUNT(*), CASE WHEN COUNT(*) = 1 THEN N' alert was' ELSE N' alerts were' END, N' raised: ',
                            STRING_AGG(CAST(CONCAT(FORMAT(dbo.fn_Local(@tzName, AlertTime), @fmt), N' (', AlertType, N', ',
                                                   dbo.fn_Duration(TotalWaitSec), N' total wait, ', BlockedCount, N' blocked)') AS NVARCHAR(MAX)), N'; ')
                                       WITHIN GROUP (ORDER BY AlertTime), N'.') END
    FROM dbo.AlertLog
    WHERE AlertTime >= @fromU AND AlertTime < @toU AND ServerName = @ServerName;

    /* ---------- chapter 8: day by day (periods longer than 2 days) ---------- */
    IF @lenSec > 2 * 86400
    BEGIN
        INSERT #story (Chapter, Story)
        SELECT N'8. Day by day',
               CONCAT(FORMAT(d.Day, N'ddd d MMM'), N': ', dbo.fn_Duration(d.Sec), N' of query time',
                      CASE WHEN d.Sec >= 1.5 * a.AvgSec THEN N' (a heavy day)' WHEN d.Sec <= 0.5 * a.AvgSec THEN N' (a light day)' ELSE N'' END,
                      N', mostly "', dg.GroupName, N'"',
                      CASE WHEN db.Sec > 0 THEN CONCAT(N', with ', dbo.fn_Duration(db.Sec), N' lost to blocking') ELSE N'' END, N'.')
        FROM (SELECT Day = CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE), Sec = SUM(WeightSec)
              FROM #w GROUP BY CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) d
        CROSS JOIN (SELECT AvgSec = AVG(1.0 * Sec) FROM (SELECT Sec = SUM(WeightSec) FROM #w GROUP BY CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) z) a
        OUTER APPLY (SELECT TOP (1) g.GroupName FROM #w w CROSS APPLY dbo.fn_WaitGroup(w.WaitCategory) g
                     WHERE CAST(dbo.fn_Local(@tzName, w.SampleTime) AS DATE) = d.Day
                     GROUP BY g.GroupName ORDER BY SUM(w.WeightSec) DESC) dg
        OUTER APPLY (SELECT Sec = SUM(w.WeightSec) FROM #w w
                     WHERE CAST(dbo.fn_Local(@tzName, w.SampleTime) AS DATE) = d.Day AND w.BlockingSessionId IS NOT NULL) db
        ORDER BY d.Day;
    END

    /* ---------- chapter 9: what to look at ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (2) N'9. What to look at',
           CONCAT(N'Because "', GroupName, N'" took ', CAST(100.0 * Sec / NULLIF(@total, 0) AS INT), N'% of the time: ', Advice)
    FROM #groups
    WHERE 100.0 * Sec / NULLIF(@total, 0) >= 10
    ORDER BY Sec DESC;

    IF EXISTS (SELECT 1 FROM #inc WHERE WasIdle = 1 AND MaxOpenTran > 0)
        INSERT #story (Chapter, Story) VALUES (N'9. What to look at',
            N'At least one session blocked others while sitting idle with an unfinished transaction. That is almost always a forgotten BEGIN TRAN in SSMS or an application that does not commit or roll back after an error. Ask the owner (see chapter 5) to finish it, and check the application''s error handling.');

    IF NOT EXISTS (SELECT 1 FROM #story WHERE Chapter = N'9. What to look at')
        INSERT #story (Chapter, Story) VALUES (N'9. What to look at', N'Nothing stands out. The time is spread over many kinds of work.');

    INSERT #story (Chapter, Story) VALUES (N'Next',
        N'For the full story of one query, run: EXEC dbo.usp_WaitStoryForStatement @StatementId = <the number after # above>;');

    SELECT Chapter, Story FROM #story ORDER BY Seq;
END
GO

/* =====================================================================
   4. The story of one query
   ===================================================================== */
CREATE OR ALTER PROCEDURE dbo.usp_WaitStoryForStatement
    @StatementId INT,
    @Days        INT = 7
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @tzName SYSNAME;
    EXEC dbo.usp_GetTimeZone @TzName = @tzName OUTPUT;
    DECLARE @toU DATETIME2(0) = SYSUTCDATETIME();
    DECLARE @fromU DATETIME2(0) = DATEADD(DAY, -@Days, @toU);
    DECLARE @iv INT = ISNULL((SELECT IntervalSec FROM dbo.MonitorConfig), 10);
    CREATE TABLE #story (Seq INT IDENTITY, Chapter NVARCHAR(60), Story NVARCHAR(MAX));

    IF NOT EXISTS (SELECT 1 FROM dbo.SqlStatements WHERE StatementId = @StatementId)
    BEGIN
        SELECT Chapter = N'Oops', Story = CONCAT(N'There is no statement #', @StatementId, N'.');
        RETURN;
    END

    SELECT w.SampleTime, w.SessionId, w.WaitCategory, w.WeightSec, w.BlockingSessionId, w.LoginName, w.HostName, w.ProgramName
    INTO #w
    FROM dbo.WaitSamples w
    WHERE w.StatementId = @StatementId AND w.SampleTime >= @fromU AND w.SampleTime < @toU;

    DECLARE @total BIGINT = ISNULL((SELECT SUM(WeightSec) FROM #w), 0);

    /* ---------- who is it ---------- */
    INSERT #story (Chapter, Story)
    SELECT N'1. Which query',
           CONCAT(N'This is ', dbo.fn_StatementName(s.StatementId), N' on server ', s.ServerName, N'. ',
                  CASE WHEN s.ObjectId = 0 AND s.LastCalledFrom IS NOT NULL
                       THEN CONCAT(N'It is not part of a stored procedure; the last time it ran, the batch sent to the server started with: "', dbo.fn_OneLine(s.LastCalledFrom, 120), N'". ')
                       WHEN s.LastCalledFrom IS NOT NULL
                       THEN CONCAT(N'The last call that ran it was: "', dbo.fn_OneLine(s.LastCalledFrom, 120), N'". ')
                       ELSE N'' END,
                  N'It was first seen ', FORMAT(dbo.fn_Local(@tzName, s.FirstSeen), N'ddd d MMM HH:mm'),
                  N' and last seen ', FORMAT(dbo.fn_Local(@tzName, s.LastSeen), N'ddd d MMM HH:mm'), N'.')
    FROM dbo.SqlStatements s WHERE s.StatementId = @StatementId;

    INSERT #story (Chapter, Story)
    SELECT N'1. Which query', CONCAT(N'The exact statement: ', dbo.fn_OneLine(StatementText, 600))
    FROM dbo.SqlStatements WHERE StatementId = @StatementId;

    IF @total = 0
    BEGIN
        INSERT #story (Chapter, Story) VALUES (N'2. How much', CONCAT(N'It was not caught running in the last ', @Days, N' days.'));
        SELECT Chapter, Story FROM #story ORDER BY Seq;
        RETURN;
    END

    /* ---------- how much ---------- */
    DECLARE @daysSeen INT = (SELECT COUNT(DISTINCT CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) FROM #w);
    DECLARE @allTotal BIGINT = (SELECT SUM(WeightSec) FROM dbo.WaitSamples w2
                                WHERE w2.SampleTime >= @fromU AND w2.SampleTime < @toU
                                  AND w2.ServerName = (SELECT ServerName FROM dbo.SqlStatements WHERE StatementId = @StatementId));
    INSERT #story (Chapter, Story) VALUES (N'2. How much',
        CONCAT(N'In the last ', @Days, N' days it was caught running on ', @daysSeen, CASE WHEN @daysSeen = 1 THEN N' day' ELSE N' days' END,
               N', for ', dbo.fn_Duration(@total), N' in total - ', FORMAT(100.0 * @total / NULLIF(@allTotal, 0), N'0.0'),
               N'% of all query time on the server in that period.'));

    /* ---------- what it waits on ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (4) N'3. What it spends its time on',
           CONCAT(CAST(100.0 * x.Sec / @total AS INT), N'%  ', g.GroupName, N' - ', dbo.fn_Duration(x.Sec), N'. ',
                  UPPER(LEFT(g.Meaning, 1)), SUBSTRING(g.Meaning, 2, 400), N'.')
    FROM (SELECT GroupName = g0.GroupName, Sec = SUM(w.WeightSec)
          FROM #w w CROSS APPLY dbo.fn_WaitGroup(w.WaitCategory) g0 GROUP BY g0.GroupName) x
    CROSS APPLY (SELECT TOP (1) GroupName, Meaning FROM dbo.WaitGlossary WHERE GroupName = x.GroupName ORDER BY SortOrder) g
    ORDER BY x.Sec DESC;

    /* ---------- when ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (1) N'4. When',
           CONCAT(N'It is busiest around ', FORMAT(DATEADD(HOUR, h.Hr, CAST('20000101' AS DATETIME2(0))), N'HH:mm'), N' - ',
                  FORMAT(DATEADD(HOUR, h.Hr + 1, CAST('20000101' AS DATETIME2(0))), N'HH:mm'),
                  N' (', CAST(100.0 * h.Sec / @total AS INT), N'% of its time falls in that hour of the day).')
    FROM (SELECT Hr = DATEPART(HOUR, dbo.fn_Local(@tzName, SampleTime)), Sec = SUM(WeightSec) FROM #w GROUP BY DATEPART(HOUR, dbo.fn_Local(@tzName, SampleTime))) h
    ORDER BY h.Sec DESC;

    INSERT #story (Chapter, Story)
    SELECT N'4. When',
           CONCAT(N'Day by day: ',
                  STRING_AGG(CAST(CONCAT(FORMAT(d.Day, N'ddd d MMM'), N' ', dbo.fn_Duration(d.Sec)) AS NVARCHAR(MAX)), N', ') WITHIN GROUP (ORDER BY d.Day),
                  N'.')
    FROM (SELECT Day = CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE), Sec = SUM(WeightSec) FROM #w GROUP BY CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) d;

    -- trend: last day it ran vs the average of the other days
    DECLARE @lastDay DATE, @lastSec BIGINT, @otherAvg DECIMAL(18,2);
    SELECT TOP (1) @lastDay = Day, @lastSec = Sec
    FROM (SELECT Day = CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE), Sec = SUM(WeightSec) FROM #w GROUP BY CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) d
    ORDER BY Day DESC;
    SELECT @otherAvg = AVG(1.0 * Sec)
    FROM (SELECT Day = CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE), Sec = SUM(WeightSec) FROM #w GROUP BY CAST(dbo.fn_Local(@tzName, SampleTime) AS DATE)) d
    WHERE Day <> @lastDay;
    IF @otherAvg > 0
        INSERT #story (Chapter, Story) VALUES (N'4. When',
            CASE WHEN @lastSec >= 1.5 * @otherAvg THEN CONCAT(N'Getting worse: on ', FORMAT(@lastDay, N'ddd d MMM'), N' it took ', FORMAT(@lastSec / @otherAvg, N'0.0'), N' times its usual daily time.')
                 WHEN @lastSec <= 0.67 * @otherAvg THEN CONCAT(N'Getting better: on ', FORMAT(@lastDay, N'ddd d MMM'), N' it took less time than on its usual day.')
                 ELSE N'Steady: its most recent day looks like its usual day.' END);

    /* ---------- blocked by ---------- */
    IF EXISTS (SELECT 1 FROM #w WHERE BlockingSessionId IS NOT NULL)
        INSERT #story (Chapter, Story)
        SELECT TOP (3) N'5. Who got in its way',
               CONCAT(N'Session ', x.BlockingSessionId, N' blocked it for ', dbo.fn_Duration(x.Sec),
                      N' (between ', FORMAT(dbo.fn_Local(@tzName, x.FirstAt), N'ddd HH:mm'), N' and ', FORMAT(dbo.fn_Local(@tzName, x.LastAt), N'ddd HH:mm'), N')',
                      CASE WHEN bs.SessionId IS NOT NULL
                           THEN CONCAT(N': ', bs.LoginName, N' using ', dbo.fn_ProgramName(bs.ProgramName),
                                       CASE WHEN bs.Status = N'sleeping' THEN N', idle with an unfinished transaction' ELSE N'' END,
                                       N'. Its last command: "', dbo.fn_OneLine(bs.LastSqlText, 120), N'".')
                           ELSE N'.' END)
        FROM (SELECT BlockingSessionId, Sec = SUM(WeightSec), FirstAt = MIN(SampleTime), LastAt = MAX(SampleTime)
              FROM #w WHERE BlockingSessionId IS NOT NULL GROUP BY BlockingSessionId) x
        OUTER APPLY (SELECT TOP (1) b.SessionId, b.LoginName, b.ProgramName, b.Status, b.LastSqlText
                     FROM dbo.BlockerSamples b
                     WHERE b.SessionId = x.BlockingSessionId AND b.SampleTime BETWEEN x.FirstAt AND x.LastAt
                       AND b.ServerName = (SELECT ServerName FROM dbo.SqlStatements WHERE StatementId = @StatementId)
                     ORDER BY b.SampleTime DESC) bs
        ORDER BY x.Sec DESC;
    ELSE
        INSERT #story (Chapter, Story) VALUES (N'5. Who got in its way', N'Nobody: it was never caught waiting behind another session.');

    /* ---------- who runs it ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (3) N'6. Who runs it',
           CONCAT(dbo.fn_ProgramName(ProgramName), N' as ', LoginName, N' from ', HostName, N': ', dbo.fn_Duration(SUM(WeightSec)),
                  N' (', COUNT(DISTINCT SessionId), CASE WHEN COUNT(DISTINCT SessionId) = 1 THEN N' session)' ELSE N' sessions)' END, N'.')
    FROM #w
    GROUP BY ProgramName, LoginName, HostName
    ORDER BY SUM(WeightSec) DESC;

    /* ---------- advice ---------- */
    INSERT #story (Chapter, Story)
    SELECT TOP (1) N'7. What to look at', CONCAT(N'Mostly "', g.GroupName, N'": ', g.Advice)
    FROM #w w CROSS APPLY dbo.fn_WaitGroup(w.WaitCategory) g
    GROUP BY g.GroupName, g.Advice
    ORDER BY SUM(w.WeightSec) DESC;

    SELECT Chapter, Story FROM #story ORDER BY Seq;
END
GO

PRINT 'Story reports installed. Try:  EXEC dbo.usp_WaitStory;   (press Ctrl+T first for text output)';
GO

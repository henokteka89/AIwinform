/* =====================================================================
   SQL Server Patching Pipeline - control database
   Run on the central node. Creates DB [DBAPatch] with inventory,
   baselines, schedule, run log, compliance view and claim proc.
   ===================================================================== */
IF DB_ID('DBAPatch') IS NULL CREATE DATABASE DBAPatch;
GO
USE DBAPatch;
GO

/* ---------- Inventory: one row per SQL instance (= AG replica) ------- */
CREATE TABLE dbo.PatchInstance (
    InstanceId        int IDENTITY PRIMARY KEY,
    SqlInstance       sysname       NOT NULL UNIQUE,   -- what you connect to: HOST\INST or FCI VNN\INST
    InstanceType      varchar(20)   NOT NULL CHECK (InstanceType IN ('Standalone','FCI')),
    FciRoleName       nvarchar(256) NULL,              -- WSFC role, e.g. 'SQL Server (MSSQLSERVER)'
    PatchGroup        sysname       NOT NULL,          -- AG or cluster name; only ONE instance per group patches at a time
    Environment       varchar(20)   NOT NULL DEFAULT 'Prod',   -- Dev/Test/Prod
    Wave              tinyint       NOT NULL DEFAULT 1,
    ExecutionMode     varchar(10)   NOT NULL DEFAULT 'Remote' CHECK (ExecutionMode IN ('Remote','Local')),
    IsEnabled         bit           NOT NULL DEFAULT 1,
    PreferredPrimary  bit           NOT NULL DEFAULT 0, -- fail AG back here after patching
    MajorVersion      varchar(10)   NULL,               -- '2019','2022' (filled by discovery)
    CurrentBuild      varchar(20)   NULL,
    WinRmOk           bit           NULL,
    LastCheckedAt     datetime2(0)  NULL,
    Notes             nvarchar(500) NULL
);

/* ---------- Nodes: Windows hosts behind an instance ------------------ */
-- Standalone = 1 node. FCI = every possible owner node.
CREATE TABLE dbo.PatchNode (
    NodeId        int IDENTITY PRIMARY KEY,
    InstanceId    int     NOT NULL REFERENCES dbo.PatchInstance(InstanceId),
    ComputerName  sysname NOT NULL,
    IsClustered   bit     NOT NULL DEFAULT 1,   -- WSFC member (AG or FCI) -> drain/resume around reboot
    IsEnabled     bit     NOT NULL DEFAULT 1,
    CONSTRAINT UQ_PatchNode UNIQUE (InstanceId, ComputerName)
);

/* ---------- Approved target build per major version ------------------ */
CREATE TABLE dbo.PatchBaseline (
    BaselineId    int IDENTITY PRIMARY KEY,
    MajorVersion  varchar(10)   NOT NULL,       -- '2019'
    TargetKB      varchar(20)   NOT NULL,       -- 'KB5046365'
    TargetBuild   varchar(20)   NOT NULL,       -- '15.0.4405.4'
    MediaFile     nvarchar(400) NULL,           -- filled by Save-PatchMedia
    ApprovedBy    sysname       NOT NULL DEFAULT SUSER_SNAME(),
    ApprovedOn    datetime2(0)  NOT NULL DEFAULT SYSDATETIME(),
    IsActive      bit           NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX UX_PatchBaseline_Active ON dbo.PatchBaseline(MajorVersion) WHERE IsActive = 1;

/* ---------- Schedule: what to patch, when ---------------------------- */
CREATE TABLE dbo.PatchSchedule (
    ScheduleId    int IDENTITY PRIMARY KEY,
    InstanceId    int          NOT NULL REFERENCES dbo.PatchInstance(InstanceId),
    PlannedStart  datetime2(0) NOT NULL,
    AllowFailover bit          NOT NULL DEFAULT 1,  -- may fail AG over if this replica is primary
    Status        varchar(20)  NOT NULL DEFAULT 'Planned'
                  CHECK (Status IN ('Planned','Claimed','InProgress','Succeeded','Failed','Skipped','Cancelled')),
    ClaimedBy     sysname      NULL,
    ClaimedAt     datetime2(0) NULL,
    LastRunId     int          NULL
);

/* ---------- Run + step log ------------------------------------------- */
CREATE TABLE dbo.PatchRun (
    RunId       int IDENTITY PRIMARY KEY,
    ScheduleId  int           NULL REFERENCES dbo.PatchSchedule(ScheduleId),
    InstanceId  int           NOT NULL REFERENCES dbo.PatchInstance(InstanceId),
    RunHost     sysname       NOT NULL,
    RunAs       sysname       NOT NULL,
    Mode        varchar(10)   NOT NULL,     -- Remote/Local
    DryRun      bit           NOT NULL,
    FromBuild   varchar(20)   NULL,
    ToBuild     varchar(20)   NULL,
    Status      varchar(20)   NOT NULL DEFAULT 'InProgress',
    StartTime   datetime2(0)  NOT NULL DEFAULT SYSDATETIME(),
    EndTime     datetime2(0)  NULL
);

CREATE TABLE dbo.PatchRunLog (
    LogId     bigint IDENTITY PRIMARY KEY,
    RunId     int           NOT NULL REFERENCES dbo.PatchRun(RunId),
    LogTime   datetime2(3)  NOT NULL DEFAULT SYSDATETIME(),
    Stage     varchar(30)   NOT NULL,   -- Precheck/Media/Patch/Failover/Validate/...
    Target    sysname       NULL,       -- node or instance the step touched
    Level     varchar(10)   NOT NULL,   -- Info/Warn/Error
    Message   nvarchar(max) NOT NULL
);
CREATE INDEX IX_PatchRunLog_Run ON dbo.PatchRunLog(RunId, LogId);
GO

/* ---------- Compliance view ------------------------------------------ */
CREATE OR ALTER VIEW dbo.vw_PatchCompliance AS
SELECT  i.SqlInstance, i.PatchGroup, i.Environment, i.Wave, i.MajorVersion,
        i.CurrentBuild, b.TargetBuild, b.TargetKB,
        CASE WHEN i.CurrentBuild IS NULL OR b.TargetBuild IS NULL THEN 'Unknown'
             WHEN TRY_CAST(PARSENAME(i.CurrentBuild,2) AS int) >= TRY_CAST(PARSENAME(b.TargetBuild,2) AS int) THEN 'Compliant'
             ELSE 'Behind' END AS Compliance,
        i.WinRmOk, i.ExecutionMode, i.LastCheckedAt
FROM dbo.PatchInstance i
LEFT JOIN dbo.PatchBaseline b ON b.MajorVersion = i.MajorVersion AND b.IsActive = 1
WHERE i.IsEnabled = 1;
GO

/* ---------- Claim proc: safe for concurrent runners ------------------ */
-- Claims the next due schedule row. Refuses a row if another instance in the
-- same PatchGroup is already Claimed/InProgress (never two replicas of one AG at once).
CREATE OR ALTER PROCEDURE dbo.usp_ClaimNextSchedule
    @RunHost    sysname,
    @ScheduleId int = NULL,          -- claim a specific row
    @SqlInstance sysname = NULL      -- or the next due row for one instance (local mode)
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    BEGIN TRAN;
    DECLARE @id int;

    SELECT TOP (1) @id = s.ScheduleId
    FROM dbo.PatchSchedule s WITH (UPDLOCK, READPAST, ROWLOCK)
    JOIN dbo.PatchInstance i ON i.InstanceId = s.InstanceId
    WHERE s.Status = 'Planned'
      AND i.IsEnabled = 1
      AND s.PlannedStart <= SYSDATETIME()
      AND (@ScheduleId  IS NULL OR s.ScheduleId  = @ScheduleId)
      AND (@SqlInstance IS NULL OR i.SqlInstance = @SqlInstance)
      AND NOT EXISTS (SELECT 1 FROM dbo.PatchSchedule s2
                      JOIN dbo.PatchInstance i2 ON i2.InstanceId = s2.InstanceId
                      WHERE i2.PatchGroup = i.PatchGroup
                        AND s2.Status IN ('Claimed','InProgress'))
    ORDER BY s.PlannedStart, i.Wave;

    IF @id IS NOT NULL
        UPDATE dbo.PatchSchedule
           SET Status = 'Claimed', ClaimedBy = @RunHost, ClaimedAt = SYSDATETIME()
         WHERE ScheduleId = @id;
    COMMIT;

    SELECT s.ScheduleId, s.AllowFailover, i.*
    FROM dbo.PatchSchedule s JOIN dbo.PatchInstance i ON i.InstanceId = s.InstanceId
    WHERE s.ScheduleId = @id;
END
GO

/* =====================================================================
   02_SampleInventory.sql  -  load YOUR environment into DBAPatch
   Run on the control instance, top to bottom, after 01_PatchControl.sql.

   Order matters because of foreign keys:
     1. PatchInstance  - the SQL instances (one row per AG replica)
     2. PatchNode      - the Windows machine(s) behind each instance
     3. PatchBaseline  - the CU you approved per SQL major version
     4. PatchSchedule  - when each instance gets patched
   ===================================================================== */
USE DBAPatch;
GO

/* ---------------------------------------------------------------------
   0. Optional reset (lab only). Set @Reset = 1 to wipe and reload.
   --------------------------------------------------------------------- */
DECLARE @Reset bit = 1;
IF @Reset = 1
BEGIN
    DELETE dbo.PatchRunLog;  DELETE dbo.PatchRun;  DELETE dbo.PatchSchedule;
    DELETE dbo.PatchNode;    DELETE dbo.PatchInstance;  DELETE dbo.PatchBaseline;
END
GO

/* ---------------------------------------------------------------------
   1. PatchInstance - one row per SQL Server instance

   SqlInstance   What you connect to. MUST equal replica_server_name in the AG.
                 Default instance = host name; named = HOST\INST; FCI = VNN\INST.
                 Check:  SELECT replica_server_name FROM sys.availability_replicas;

   InstanceType  'Standalone' = runs on one Windows box (normal AG replica).
                 'FCI'        = clustered instance that moves between nodes.

   FciRoleName   FCI only: the WSFC role name (Failover Cluster Manager > Roles),
                 e.g. 'SQL Server (MSSQLSERVER)'. NULL for standalone.

   PatchGroup    A label for "things that must never be patched at the same time".
                 Use your REAL AG name. The claim proc refuses to start a second
                 instance of the same group while one is in progress, so two
                 replicas of one AG are never down together.
                 Find it:  SELECT name FROM sys.availability_groups;
                 ('AG_Sales' in the old sample was only a placeholder.)

   Environment   Free text for reporting/filtering: Dev / Test / Prod.

   Wave          Rollout ring: 0 = dev/lab, 1 = test or first prod replicas,
                 2 = remaining prod. Today it is a TIEBREAKER only - when two rows
                 are due at the same PlannedStart, lower Wave goes first. The real
                 order is PlannedStart in PatchSchedule. It's there so you can later
                 add a rule like "don't start wave 2 until wave 1 succeeded".

   ExecutionMode 'Remote' = central node patches it over WinRM.
                 'Local'  = WinRM blocked; run the script ON that server with -Local.
                 Run -Action Discover first: it fills WinRmOk so you know which.

   PreferredPrimary 1 = after patching this replica, fail the AG back to it.
                 Set it on the replica you normally want as primary, 0 elsewhere.
   --------------------------------------------------------------------- */
DECLARE @AgName sysname = N'AG_Sales';   -- <<< replace with your real AG name

INSERT dbo.PatchInstance
       (SqlInstance,        InstanceType, FciRoleName, PatchGroup, Environment, Wave, ExecutionMode, PreferredPrimary)
VALUES (N'WIN-Q32QIO9QAL5', 'Standalone', NULL,        @AgName,    'Prod',      1,    'Remote',      0),
       (N'WIN-R27C4ESS9UJ', 'Standalone', NULL,        @AgName,    'Prod',      1,    'Local',       0),  -- also the control instance; WinRM blocked
       (N'WIN-HEMKJKL4CJO', 'Standalone', NULL,        @AgName,    'Prod',      1,    'Remote',      0);

-- FCI example (keep for later):
-- (N'SQLFCI01\PROD', 'FCI', N'SQL Server (PROD)', @AgName, 'Prod', 2, 'Remote', 1)
GO

/* ---------------------------------------------------------------------
   2. PatchNode - the Windows machine(s) the setup.exe actually runs on

   Standalone instance -> exactly 1 row (its host).
   FCI                 -> 1 row per possible owner node (all get patched in one run,
                          one after another: passive, move role, old owner).

   ComputerName  Windows host name (FQDN ok). For a default standalone instance
                 it's the same as SqlInstance.
   IsClustered   1 if the box is a WSFC member (every AG replica is). The pipeline
                 drains the node before reboot and resumes it after.
   --------------------------------------------------------------------- */
INSERT dbo.PatchNode (InstanceId, ComputerName, IsClustered)
SELECT i.InstanceId, n.ComputerName, 1
FROM (VALUES (N'WIN-Q32QIO9QAL5', N'WIN-Q32QIO9QAL5'),
             (N'WIN-R27C4ESS9UJ', N'WIN-R27C4ESS9UJ'),
             (N'WIN-HEMKJKL4CJO', N'WIN-HEMKJKL4CJO')
     ) n(SqlInstance, ComputerName)
JOIN dbo.PatchInstance i ON i.SqlInstance = n.SqlInstance;

-- FCI example: two nodes behind one instance
-- INSERT dbo.PatchNode (InstanceId, ComputerName, IsClustered)
-- SELECT InstanceId, n.c, 1 FROM dbo.PatchInstance
-- CROSS APPLY (VALUES (N'NODE01'), (N'NODE02')) n(c) WHERE SqlInstance = N'SQLFCI01\PROD';
GO

/* ---------------------------------------------------------------------
   3. PatchBaseline - the CU you approve, one ACTIVE row per major version

   MajorVersion  Must match what Discover writes: '2016','2017','2019','2022','2025'.
                 Check after Discover:  SELECT SqlInstance, MajorVersion, CurrentBuild FROM dbo.PatchInstance;
   TargetKB      KB of the CU, e.g. 'KB5xxxxxx'. Used to download the media.
   TargetBuild   Build that CU produces, e.g. '16.0.4xxx.x'. Used for compliance + validation.

   Look both up in PowerShell (on the central node):
     Test-DbaBuild -SqlInstance WIN-Q32QIO9QAL5 -Latest -Update | Select SqlInstance, Build, BuildTarget, CUTarget
     Get-DbaBuild  -Build <BuildTarget from above> | Select Build, KBLevel, NameLevel, CULevel
   MediaFile     Leave NULL - Save-PatchMedia (-Action Download) fills it.

   New CU next month: UPDATE ... SET IsActive = 0 on the old row, insert the new one.
   --------------------------------------------------------------------- */
INSERT dbo.PatchBaseline (MajorVersion, TargetKB, TargetBuild)
VALUES ('2022', 'KB0000000', '16.0.0000.0');   -- <<< replace all three values
GO

/* ---------------------------------------------------------------------
   4. PatchSchedule - one row per instance per patch cycle

   PlannedStart  Earliest time the run may start. The runner only claims rows
                 whose PlannedStart has passed, so put each replica on its own night.
   AllowFailover 1 = if this replica is PRIMARY when its turn comes, fail the AG over
                 to a patched sync secondary first. 0 = refuse to patch a primary.
   Status        Starts 'Planned'; the pipeline moves it to Claimed/InProgress/
                 Succeeded/Failed. Re-plan a failed row:  UPDATE ... SET Status='Planned'

   ORDER RULE: secondaries first, current primary LAST. The primary's run will
   fail its pre-check until every other replica is at the target build.
   Who is primary right now:
     SELECT ar.replica_server_name, ars.role_desc
     FROM sys.dm_hadr_availability_replica_states ars
     JOIN sys.availability_replicas ar ON ar.replica_id = ars.replica_id;
   Adjust the dates below so the primary is last.
   --------------------------------------------------------------------- */
INSERT dbo.PatchSchedule (InstanceId, PlannedStart, AllowFailover)
SELECT i.InstanceId, s.PlannedStart, 1
FROM (VALUES (N'WIN-HEMKJKL4CJO', CAST('2026-11-03T22:00:00' AS datetime2(0))),   -- secondary
             (N'WIN-Q32QIO9QAL5', CAST('2026-11-05T22:00:00' AS datetime2(0))),   -- secondary
             (N'WIN-R27C4ESS9UJ', CAST('2026-11-08T22:00:00' AS datetime2(0)))    -- primary (if it is) - last
     ) s(SqlInstance, PlannedStart)
JOIN dbo.PatchInstance i ON i.SqlInstance = s.SqlInstance;
GO

/* ---------------------------------------------------------------------
   5. Check what you loaded
   --------------------------------------------------------------------- */
SELECT i.SqlInstance, i.PatchGroup, i.Wave, i.ExecutionMode, n.ComputerName, s.PlannedStart, s.Status
FROM dbo.PatchInstance i
LEFT JOIN dbo.PatchNode n     ON n.InstanceId = i.InstanceId
LEFT JOIN dbo.PatchSchedule s ON s.InstanceId = i.InstanceId
ORDER BY s.PlannedStart;
SELECT * FROM dbo.PatchBaseline WHERE IsActive = 1;

# SqlPatch: SQL Server patching pipeline

PowerShell + dbatools, driven by a control database (`DBAPatch`) on the central node.
**The unit of work is one SQL instance (one AG replica) per run.** If that instance is an FCI, all of its nodes are patched in the same run.

```
SqlPatch/
├─ Invoke-SqlPatch.ps1          entry point (Discover / Download / Plan / Patch)
├─ module/SqlPatch.psm1         pipeline logic
└─ sql/
   ├─ 01_PatchControl.sql       control DB: tables, compliance view, claim proc
   └─ 02_SampleInventory.sql    example AG + FCI topology and schedule
```

## Setup (once)
1. Run `sql/01_PatchControl.sql` on the central node.
2. Load your inventory (see `02_SampleInventory.sql`):
   - `PatchInstance.SqlInstance` must match `replica_server_name` in the AG exactly (FCI = `VNN\INSTANCE`).
   - `PatchGroup` = the AG (or cluster) name. Only one instance per group is ever patched at a time.
   - `PatchNode`: one row per Windows node. An FCI lists every possible owner node.
3. Install dbatools on the central node (`Install-Module dbatools`). Create a share for media, e.g. `\\CENTRAL01\SQLPatch\Media`.
4. Run as an account that is sysadmin on every instance and local admin on every node. A gMSA is ideal for scheduled runs.

## Monthly flow
```powershell
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Action Discover            # builds + WinRM reachability
# approve the CU: INSERT/UPDATE dbo.PatchBaseline (MajorVersion, TargetKB, TargetBuild)
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Action Download -MediaShare \\CENTRAL01\SQLPatch\Media
# schedule: INSERT dbo.PatchSchedule (InstanceId, PlannedStart) - one row per replica, any night you like
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Action Plan
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -DryRun                     # pre-checks only, changes nothing
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01                             # patches the next due row
```
To automate, create a Task Scheduler (or Agent) job on the central node that runs the last command every 30 minutes during your window. It does nothing until a row is due.

Ad hoc: `-SqlInstance SQLAG02 -Now`. Compliance report: `SELECT * FROM DBAPatch.dbo.vw_PatchCompliance`.

## What one run does
1. **Claim**: `usp_ClaimNextSchedule` takes the next due row. It refuses the row if another instance in the same `PatchGroup` is in progress.
2. **Pre-checks**: any failure stops the run before anything changes.
   - SQL is reachable, and all AGs the instance belongs to are HEALTHY.
   - Last full/diff backup within 26h, and log backup within 60 min for FULL recovery. These are checked across **all replicas**, because AG backups may run on a secondary.
   - At least 6 GB free on C:, and no reboot already pending on any node.
   - **If the instance is an AG primary, every other replica must already be at the target build.** Otherwise the run fails with "patch replica X first".
   - Warning only: transactions open more than 5 minutes.
3. **Failover** (only if the instance is primary): manual failover to a synchronous, HEALTHY secondary that is already patched. The run then waits for the AG to be healthy again.
4. **Patch**
   - Standalone: copy the CU to `C:\SQLPatch` on the node (this avoids double-hop and CredSSP), drain the cluster node, run setup `/quiet /Action=Patch /AllInstances`, reboot, then resume the node.
   - FCI: patch the passive node(s), move the role to a patched node, wait for SQL (script-upgrade mode), then patch the old owner.
5. **Validate**: build is at least the target, auto-start services are running, non-AG databases are ONLINE, and AGs are HEALTHY.
6. **Fail back** if `PreferredPrimary = 1`. The run then writes `PatchRun`, `PatchRunLog`, and `PatchInstance.CurrentBuild`.

Already-patched instances skip straight to validation, so reruns are safe.

## Recommended order for an AG
DR/async secondaries first, then sync secondaries, then the primary (which fails over as part of its own run). These can be on different nights. Running with a secondary at a higher build than the primary is supported. The reverse is not: secondaries on a lower build than the primary stop syncing. That's why the primary's run refuses to start until every other replica is patched.

**AG + FCI note:** AG replicas hosted on an FCI can't be automatic-failover targets. This pipeline only does manual failovers, so it works with that limitation.

## WinRM-blocked servers
Set `PatchInstance.ExecutionMode = 'Local'`. The central runner will skip that instance and tell you. Run the script on that server instead:
```powershell
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Local     # installs, logs, reboots
.\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Local     # after boot: resumes and validates
```
Local mode is for standalone instances only (v1). FCIs need Remote mode, because the other nodes must be reachable.

## Going parallel later
The claim proc already locks rows (`UPDLOCK, READPAST`) and enforces one instance per PatchGroup. To run in parallel, launch several runners, or raise `-MaxRuns` and wrap `Invoke-PatchInstance` in `Start-ThreadJob` / `ForEach-Object -Parallel`. Different AGs can then patch the same night without colliding.

## Known gaps / next steps
- `/AllInstances` patches every instance on the node. Model shared hosts as one PatchGroup.
- No email/Teams notification yet: add it at the end of `Invoke-SqlPatch.ps1`, using `PatchRunLog` for the run.
- Rollback = VM snapshot (recommended before prod), or uninstall the CU via Programs and Features.
- Test the first runs in a lab AG+FCI, with `-DryRun` first.

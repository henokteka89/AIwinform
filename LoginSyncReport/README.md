# LoginSyncReport

A read-only check that SQL Server logins are in sync between a **reference** server (AG primary, or the current server in a migration) and one or more **targets** (AG secondaries, or the new server).
- **With a window:** discover servers, see the drift, and review the fixes.
- **Headless:** produces an interactive HTML report, CSV/JSON, and remediation `.sql` scripts.

Remediation T-SQL is generated for you to review. **The tool never runs it.**

## Start
- **Window:** double-click **`Start-LoginSyncReport.cmd`**. No admin rights are needed.
- **Headless / scheduled:**
  ```powershell
  # Discover an AG: the primary becomes the reference, every other replica a target (with its real commit mode)
  .\LoginSyncReport.ps1 -DiscoverAgFrom SQLAGLISTENER01

  # Migration: new server vs current server
  .\LoginSyncReport.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01

  # Mixed modes
  .\LoginSyncReport.ps1 -ReferenceServer SQL01 -TargetServer SQL02 -TargetMode Synchronous -AsyncTarget SQL03,SQL04

  # SQL Agent / scheduler: exit code 2 when High or Critical drift is found
  .\LoginSyncReport.ps1 -DiscoverAgFrom SQLAG01 -NoOpen -FailOnSeverity High -OutputPath \\share\dba\LoginSync

  # Sample report, no server contacted
  .\LoginSyncReport.ps1 -Demo
  ```

## What it compares
- **Logins:** SQL logins, Windows users and Windows groups exist on every server, plus custom server roles.
- **SID match.** A mismatched SID is what orphans database users after a failover or restore.
- **Password hash** for SQL logins, compared by SHA-256 digest. Reading it requires CONTROL SERVER.
- **Login settings:** disabled state, default database and language, CHECK_POLICY / CHECK_EXPIRATION.
- **Server roles and permissions:** role membership and server-level GRANT / DENY.
- **Orphan risk:** every database user on the reference is checked against each server's logins by SID.
- **Target-only logins:** logins that exist only on a target are reported, never dropped.

Drift on a synchronous (automatic-failover) partner or a migration target is rated one level higher than the same drift on an asynchronous DR replica.

## Useful parameters
| Parameter | Purpose |
|---|---|
| `-ReferenceServer` / `-TargetServer` | Compare a specific source against one or more targets |
| `-TargetMode` | `Synchronous`, `Asynchronous` or `Standalone` for `-TargetServer` entries |
| `-AsyncTarget` | Targets that are asynchronous replicas (lower severity) |
| `-DiscoverAgFrom` / `-AvailabilityGroup` / `-SyncReplicasOnly` | Build the comparison from an AG automatically |
| `-ExcludeLogin` | Wildcards to ignore (defaults: `##*`, `NT SERVICE\*`, `NT AUTHORITY\*`, `BUILTIN\*`) |
| `-SkipPasswordCompare` | Don't compare password hashes |
| `-OutputPath` / `-NoOpen` / `-PassThru` | Where reports go, don't open them, return objects |
| `-FailOnSeverity` | Exit code 2 at or above this severity |
| `-Encrypt` / `-TrustServerCertificate` | Connection encryption options |
| `-Gui` / `-Demo` | Open the window / generate a sample report |

## Built-in safety
- **Queries:** only static SELECTs against catalog views. Anything else is refused before it reaches SQL Server.
- **Sessions back off:** `LOCK_TIMEOUT` (3 s by default), `DEADLOCK_PRIORITY LOW`, READ UNCOMMITTED. A locked catalog is skipped, never queued behind.
- **Connections:** one at a time, no pooling, closed straight away. The application name `LoginSyncCheck (read-only)` makes the session easy to spot.
- **Remediation `.sql` files:** each one refuses to run on the wrong server (`@@SERVERNAME` check plus `SET NOEXEC ON`), and destructive steps are commented out.

## Permissions
- **VIEW ANY DEFINITION** covers everything except password hashes, which need **CONTROL SERVER** (sysadmin).
- **VIEW SERVER STATE** is used for AG discovery.

> When CONTROL SERVER was available, the remediation `.sql` files contain password hashes. **Treat them as secrets and delete them after use.** The HTML and JSON reports never contain hashes, only digests.

## Requirements
Windows PowerShell 5.1, or PowerShell 7+. No module dependencies.

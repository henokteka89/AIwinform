# MigrationReadiness

A read-only readiness check for SQL Server migrations and HA/DR partners. It compares a **reference** server (current production, or the AG primary) with one or more **targets** (the new build, or AG secondaries / DR replicas). It covers every area that usually breaks a cut-over or a failover.

For each target you get:
- a readiness verdict and score
- a *needs installation / needs action* list
- a preparation T-SQL script for review

## Start
- **Window:** double-click **`Start-MigrationReadiness.cmd`**. No admin rights are needed to open it.
- **Headless:**
  ```powershell
  .\MigrationReadiness.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01
  .\MigrationReadiness.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01 -PathMap 'D:\=E:\','L:\=F:\'
  .\MigrationReadiness.ps1 -DiscoverAgFrom SQLAGLISTENER01          # primary = reference, real commit modes
  .\MigrationReadiness.ps1 -ReferenceServer SQL01 -TargetServer SQL02 -TargetMode Synchronous -AsyncTarget SQLDR01
  .\MigrationReadiness.ps1 -Demo                                    # sample report, no server needed
  ```
  A headless run writes interactive HTML, Word (`.docx`), plain text, CSV, JSON and `prep_<target>.sql`, and returns an exit code for schedulers.

## What it checks
| Area | Examples |
|---|---|
| Platform & OS | OS build, domain, time zone, locale, power plan, cluster membership |
| Hardware & disks | Cores, RAM, volumes (64K allocation unit), free space per drive (honours `-PathMap`), tempdb layout, system DB file locations |
| Software, drivers, providers | VC++ runtimes, Java, agents, ODBC drivers/DSNs, OLE DB providers and their options, SQL aliases, .NET, TLS settings |
| Windows features & services | Failover Clustering, MPIO, SSIS/SSRS/SSAS/Full-Text, service accounts, firewall rules for SQL / HADR / listener / Browser |
| SQL instance | Version/CU, edition (incl. Enterprise features in use), collation, HADR, FILESTREAM, PolyBase, IFI, locked pages, TCP port, force encryption |
| Certificates | SQL TLS cert (expiry, SAN coverage), machine certs, TDE/backup certificates |
| Configuration & CLR | `sp_configure` parity, trace flags, assemblies, CLR strict security, TRUSTWORTHY |
| Linked servers | Definition, options, provider registered on target, stored logins, code references |
| Agent & automation | Jobs, steps, schedules, proxies, operators, alerts, Database Mail, maintenance plans, files/scripts referenced by job steps |
| SSIS / SSRS / SSAS | SSISDB projects/environments, msdb packages, services |
| Security & server objects | Credentials, TDE certs, audits, server triggers, endpoints, startup procs, XE sessions, Resource Governor |
| Databases & paths | Compat level, owners, CDC/replication/broker, cross-DB dependencies, folders, backup/audit paths |
| Source health & host policy | Backups, CHECKDB, suspect pages, log shipping/mirroring, support lifecycle, pending reboot, patch level, MSDTC, Defender exclusions |
| Identity & hidden dependencies | Kerberos SPNs, DNS, hard-coded server names/IPs, failing jobs, other instances on the host |

Logins are out of scope here; use **LoginSyncReport** for those.

## Useful parameters
| Parameter | Purpose |
|---|---|
| `-ReferenceServer` / `-TargetServer` / `-TargetMode` / `-AsyncTarget` | Manual comparison |
| `-DiscoverAgFrom` / `-AvailabilityGroup` / `-SyncReplicasOnly` | Build the comparison from an AG |
| `-PathMap 'D:\=E:\'` / `-ComputerMap` | Map paths / host names that differ on the target |
| `-Skip` | Skip check areas |
| `-Format` | Output formats (default: Html, Docx, Txt, Csv, Json, Sql) |
| `-OutputPath` / `-Open` | Report folder / open the report when done |
| `-FailOnSeverity` | Exit code 2 at or above this severity (default High) |
| `-CimProtocol` | `Wsman` (default) or `Dcom` for Windows queries |
| `-Encrypt` / `-TrustServerCertificate` | SQL connection encryption options |
| `-LockTimeoutMs` / `-QueryTimeoutSec` / `-MaxDatabases` | Safety limits |

**Exit codes:**
- `0`: ready, or nothing at or above `-FailOnSeverity`
- `2`: findings at or above `-FailOnSeverity`
- `1`: the run failed

## Built-in safety
- **SQL:** only static SELECTs against catalog views, DMVs and msdb tables. Anything else is refused before it reaches SQL Server. Lock timeouts, low deadlock priority, no pooling.
- **Windows:** read-only CIM/WMI queries and read-only registry methods only.
  - Nothing is installed, started, stopped or changed.
  - No `Invoke-Command`, no admin shares.
  - Certificates are parsed in memory; private keys are never read.
- **`prep_<target>.sql`:** it aborts unless run on the target, every change is `IF NOT EXISTS` or commented out, and secrets are placeholders.

## Permissions
- **SQL:** VIEW SERVER STATE plus VIEW ANY DEFINITION. sysadmin gives the fullest picture; msdb/SSISDB reads need db_datareader or sysadmin.
- **Windows:** local admin, or WMI/DCOM read rights (`root\cimv2`, `root\default` StdRegProv) on each host, with WinRM or DCOM reachable.

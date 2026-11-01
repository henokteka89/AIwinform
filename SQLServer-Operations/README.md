# SQL Server Operations Dashboard

A PowerShell monitoring and reporting engine. It collects read-only operational data from one or more SQL Server instances (standalone or Always On nodes) and their Windows hosts, then builds a **single standalone HTML dashboard**.

## Start
1. Install the report module (one time):
   ```powershell
   Install-Module PSWriteHTML -Scope CurrentUser
   ```
2. Run it with your instance names:
   - **Launcher:** `Start-SQLServer-Operations.cmd -SqlInstances SQL01,SQL02`
     - The console stays visible so you can follow progress.
     - When the script succeeds, the report opens in your browser; then the window pauses.
     - Any arguments are passed through to the script.
   - **PowerShell:**
     ```powershell
     .\SQLServer-Operations.ps1 -SqlInstances 'SQLNODE1','SQLNODE2'
     .\SQLServer-Operations.ps1 -SqlInstances 'SQLPROD01\INST1' -OutputFolder 'D:\Reports' -EnableMetricHistory $true -TimestampedFileName
     ```

> The default instance names in the script (`SQLNODE1`, `SQLNODE2`) are placeholders. Always pass `-SqlInstances`, or edit the default in the `param()` block.

**Output:**
- Report: `<script folder>\Output\SQLServer-Operations.html` (change it with `-OutputFolder` / `-ReportFileName`).
- Optional metric history: `<script folder>\Output\History\MetricHistory.csv`.

## Dashboard tabs
1. **Overview:** KPIs, health cards, instance and database inventory, OS, services.
2. **Always On:** AG topology, replica connectivity and sync, database sync, queues, listeners.
3. **Backups:** freshness, compliance across AG replicas, history, restores, gaps.
4. **Capacity & Performance:** CPU history, memory, storage, DB sizes, waits, I/O, log/VLF.
5. **Cluster & Windows:** cluster health, groups, resources, event logs, services, accounts.
6. **SQL Errors:** error log, grouped messages, suspect pages, memory dumps, critical Windows events.
7. **Security & Configuration:** DB security, orphaned users, privileged roles, certificates, HA config differences.
8. **Investigation:** job duration trends, CHECKDB evidence, long-running queries and blocking, findings, collection/permission errors, optional metric history.

## Main parameters
| Parameter | Default | Purpose |
|---|---|---|
| `-SqlInstances` | placeholders | Instances to report on |
| `-EnvironmentName` | `SQL Server Environment` | Title shown in the report |
| `-FullBackupMaxAgeHours` / `-DiffBackupMaxAgeHours` / `-LogBackupMaxAgeMinutes` | 168 / 26 / 60 | Backup freshness thresholds |
| `-DiskFreeWarningPct` / `-DiskFreeCriticalPct` | 20 / 10 | Disk space thresholds |
| `-LongRunningQuerySeconds` / `-BlockingWaitSeconds` | 300 / 30 | Long query / blocking thresholds |
| `-IoLatencyWarningMs` / `-IoLatencyCriticalMs` | 20 / 50 | I/O latency thresholds |
| `-VlfWarningCount` / `-LogUsedWarningPct` | 1000 / 75 | Transaction log thresholds |
| `-CheckDbMaxAgeDays` | 7 | Last good CHECKDB age |
| `-CertificateExpiryWarningDays` | 30 | Certificate expiry warning |
| `-AgSendQueueWarningKB` / `-AgRedoQueueWarningKB` | 10240 | AG queue thresholds |
| `-LookbackHours` / `-HistoryLookbackDays` / `-CpuHistoryMinutes` | 24 / 14 / 240 | How far back to look |
| `-CollectEventViewer` / `-CollectClusterInfo` | `$true` | Turn Windows event / cluster collection on or off |
| `-EnableMetricHistory` / `-MetricHistoryRetentionDays` | `$false` / 30 | Keep a metric history CSV |
| `-OutputFolder` / `-ReportFileName` / `-TimestampedFileName` | Output folder next to the script | Where the report goes |
| `-ConnectionTimeoutSeconds` / `-QueryTimeoutSeconds` | 15 / 120 | Timeouts |

## Read-only
Every SQL statement is either a SELECT against DMVs, catalog views or msdb history, or a read-only system procedure (`xp_readerrorlog`, `DBCC TRACESTATUS`). Nothing is created, altered or deleted, and no configuration is changed.

## Requirements
- **Windows PowerShell 5.1**, with the **PSWriteHTML** module.
- **Authentication:** Windows-integrated authentication to every SQL Server.
- **SQL permissions:** VIEW SERVER STATE, VIEW ANY DEFINITION, and read access to msdb history. sysadmin is simplest; the error log and some security checks need it.
- **Windows permissions:** remote CIM/WMI, Event Log and service queries on each host.
- **Failover cluster sections:** the FailoverClusters module (RSAT).
- **SQL version:** SQL Server 2016 SP2 or later for every section. Older builds still report, and sections that fail are listed under Investigation.

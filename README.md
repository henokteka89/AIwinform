# SQL Server DBA tools (PowerShell + WinForms)

A set of standalone PowerShell tools for SQL Server administration. Each tool lives in its own folder with:
- the script
- a `Start-*.cmd` launcher: double-click it, no console window behind it
- a README

| Tool | What it does | Changes anything? |
|---|---|---|
| [Compare-ExecutionPlans](Compare-ExecutionPlans/) | Compare two or more execution plans, flag the differences, build an AI-ready prompt | No |
| [LoginSyncReport](LoginSyncReport/) | Login / SID / password / permission drift between AG replicas or migration servers; generates remediation scripts for review | No (scripts are never run) |
| [MigrationReadiness](MigrationReadiness/) | Readiness check of a new server or DR replica against the current one, with score, action list and prep script | No |
| [PITR-RestoreBuilder](PITR-RestoreBuilder/) | Point-in-time restore builder: picks the FULL/DIFF/LOG chain from msdb, preflights the destination, generates the restore script and a SQL Agent job. **Test before production use** | No by default (restore, job, credential and TDE actions only after opt-in and confirmation) |
| [PlanGuideGenerator](PlanGuideGenerator/) | Builds a reviewed plan guide or Query Store hint / forced-plan script from a statement plus a hint or plan XML, or straight from saved good/bad plans: decides which plan is good, extracts the T-SQL, handles PSP query variants, and checks exact text matching against the plan cache, module or Query Store | No (generates scripts only; optional read-only server check) |
| [SQLServer-Operations](SQLServer-Operations/) | Multi-instance operations dashboard (HTML): Always On, backups, capacity, errors, security | No |
| [SqlBlockingDeadlockMonitor](SqlBlockingDeadlockMonitor/) | Live blocking chains and deadlock analysis, explained in plain English | No |
| [SqlFleetStudio](SqlFleetStudio/) | SSMS-style query window for many servers at once: cached server list from a central server/CMS, query library, parallel execution with merged results | No by default (read-only guard plus forced rollback; writes need a one-shot opt-in and typed confirmation) |
| [SqlIndexStatsAnalyzer](SqlIndexStatsAnalyzer/) | Index and statistics analyzer for databases from MBs to tens of TB: fuzzy table search, usage, fragmentation (size-capped, on demand), histograms explained in plain English, priorities, and procedure-specific analysis from the cached plan. Generates REBUILD/REORGANIZE/UPDATE STATISTICS scripts for review | No (read-only guard on every query; scripts are never run) |
| [SqlWaitMonitor](SqlWaitMonitor/) | DPA-style wait monitoring: per-statement wait bars you can drill into, blocking capture, alerts, plain-English story reports. Standalone or many servers through linked servers | **Yes** at install: creates a monitoring database and two SQL Agent jobs (the viewer only reads) |
| [SqlPatchAutomation](SqlPatchAutomation/) | Patch one SQL Server instance with a CU/GDR: guardrails, staging, install, reboot, validation | **Yes**, it patches SQL Server (admin required) |

## Getting started
1. Copy a tool's folder to the machine you'll run it from, e.g. `C:\Apps\<Tool>`.
2. Files downloaded from the internet are flagged by Windows. Unblock them once:
   ```powershell
   Get-ChildItem C:\Apps -Recurse -Include *.ps1,*.cmd | Unblock-File
   ```
3. Double-click the tool's `Start-*.cmd`.

**Requirements:** Windows PowerShell 5.1, or PowerShell 7+ on Windows. See each README for module and permission requirements.

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
| [SQLServer-Operations](SQLServer-Operations/) | Multi-instance operations dashboard (HTML): Always On, backups, capacity, errors, security | No |
| [SqlBlockingDeadlockMonitor](SqlBlockingDeadlockMonitor/) | Live blocking chains and deadlock analysis, explained in plain English | No |
| [SqlPatchAutomation](SqlPatchAutomation/) | Patch one SQL Server instance with a CU/GDR: guardrails, staging, install, reboot, validation | **Yes**, it patches SQL Server (admin required) |

## Getting started
1. Copy a tool's folder to the machine you'll run it from, e.g. `C:\Apps\<Tool>`.
2. Files downloaded from the internet are flagged by Windows. Unblock them once:
   ```powershell
   Get-ChildItem C:\Apps -Recurse -Include *.ps1,*.cmd | Unblock-File
   ```
3. Double-click the tool's `Start-*.cmd`.

**Requirements:** Windows PowerShell 5.1, or PowerShell 7+ on Windows. See each README for module and permission requirements.

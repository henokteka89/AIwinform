# SQL Wait Monitor (DPA-style)

Wait-time monitoring for SQL Server, in the style of SolarWinds DPA. Every 10 seconds a SQL Agent job records each running request and what it is waiting on (or CPU). It also records **the exact statement**: the procedure and line number, or the ad hoc / dynamic SQL text.

You can look at the data three ways:
- a **WinForms viewer** with stacked wait bars per statement. Click a bar to drill into that statement.
- **alerts**: total wait above a threshold or a blocking pile-up, written to a log table and optionally emailed through Database Mail.
- **story reports**: stored procedures that explain a period, or one query, in plain English.

There are two ways to install it:

| | Standalone | Central (linked servers) |
|---|---|---|
| Monitors | the server it is installed on | many servers, from one central server |
| Database | `DBAWait` | `DBAWaitCentral` |
| Install | `Install-SqlWaitMonitor.sql` | `Install-SqlWaitMonitorCentral.sql` |
| Story reports | `Story-Reports.sql` | `Story-Reports-Central.sql` |
| Viewer launcher | `Start-SqlWaitMonitor.cmd` | `Start-SqlWaitMonitorCentral.cmd` |
| Agent jobs | `DBAWait - Collect` (10 s), `DBAWait - Alert` (1 min) | `DBAWaitCentral - Collect`, `DBAWaitCentral - Alert` |

The viewer (`SqlWaitMonitor.ps1`) is the same file for both. It detects which database it is connected to. On the central database it adds a server picker and a *Collector status* tab.

## Files
| File | What it is |
|---|---|
| `Install-SqlWaitMonitor.sql` | Standalone install: database, tables, collector, alert check, Agent jobs. Safe to re-run: it upgrades in place and keeps the data. |
| `Install-SqlWaitMonitorCentral.sql` | Central install. Section 0 has the remote login and linked server examples. |
| `Story-Reports.sql` / `Story-Reports-Central.sql` | Plain-English story procedures, the wait-type glossary and helper functions. Run after the install. |
| `SqlWaitMonitor.ps1` | The viewer (WinForms). |
| `Start-SqlWaitMonitor.cmd` / `Start-SqlWaitMonitorCentral.cmd` | Double-click launchers. Edit `SERVER=` inside to point at your server. |
| `Test-BlockingScenario.sql` | Recreates a blocking incident on a **test** server, so you can watch it being captured. |

## Install: standalone
1. Open `Install-SqlWaitMonitor.sql` in SSMS on the server to monitor.
2. Edit section 2 (*SETTINGS*). For email alerts set `MailProfile` (an existing Database Mail profile) and `Recipients`. Optionally change the thresholds.
3. Run it (F5). It ends with a row saying *installed*.
4. Run `Story-Reports.sql`.
5. Edit `SERVER=` in `Start-SqlWaitMonitor.cmd` if needed, then double-click it.

## Install: central, through linked servers
Nothing is installed on the monitored servers except a read-only login. The central collector runs the DMV snapshot **on** each server with `EXEC (...) AT [linked server]` and stores the rows centrally, tagged with the server name.

1. On **each remote** server, create a monitoring login (section 0a of the script):
   ```sql
   CREATE LOGIN dbawait_monitor WITH PASSWORD = N'...';
   GRANT VIEW SERVER STATE   TO dbawait_monitor;
   GRANT VIEW ANY DEFINITION TO dbawait_monitor;
   ```
2. On the **central** server, create a linked server per remote server (section 0b):
   ```sql
   EXEC sp_addlinkedserver   @server = N'PRODSQL01', @srvproduct = N'', @provider = N'MSOLEDBSQL',
                             @datasrc = N'PRODSQL01', @provstr = N'Encrypt=Optional';
   EXEC sp_addlinkedsrvlogin @rmtsrvname = N'PRODSQL01', @useself = N'False', @locallogin = NULL,
                             @rmtuser = N'dbawait_monitor', @rmtpassword = N'...';
   ```
3. Run `Install-SqlWaitMonitorCentral.sql` on the central server (settings are in section 2). The central server registers itself automatically.
4. Register and test each remote server:
   ```sql
   USE DBAWaitCentral;
   EXEC dbo.usp_AddMonitoredServer  @ServerName = N'PRODSQL01', @LinkedServer = N'PRODSQL01';
   EXEC dbo.usp_TestMonitoredServer @ServerName = N'PRODSQL01';
   ```
   `usp_AddMonitoredServer` does three things to the linked server: turns on RPC Out, turns off DTC promotion, and sets short timeouts. The test shows the remote version, the login it connects as, and its permissions, then collects once.
5. Run `Story-Reports-Central.sql`, edit `SERVER=` in `Start-SqlWaitMonitorCentral.cmd` (the **central** server), and double-click it.

**Try it on one machine:** make a loop-back linked server, then register it as above.
```sql
EXEC sp_addlinkedserver @server = N'LOOPBACK', @srvproduct = N'', @provider = N'MSOLEDBSQL',
                        @datasrc = @@SERVERNAME, @provstr = N'Encrypt=Optional';
```

One unreachable server does not stop the others. Its error is stored in `MonitoredServers.LastError` and shown in red in the viewer. Servers are collected one after another, at roughly 50–300 ms each, so one central server handles about 20 servers at a 10-second interval.

## The viewer
- **Time range:** 5 minutes (10-second bars) to 7 days.
- **Auto-refresh:** Off, 5 s, 10 s, 30 s, 1 m, 5 m or 15 m.
- **Bars by:** *Statement* or *Wait type*.
- **Click a bar** to open that statement's detail for the time slice you clicked:
  - the exact SQL text
  - the proc and line, with *Show whole procedure*, which highlights the statement
  - waits over time and a breakdown by wait type
  - who blocked it, and who ran it
  - *Open execution plan*, which opens the cached plan in SSMS
- **In Wait type mode,** clicking a bar lists the statements that had that wait in that slice.
- **Tabs:** top statements, head blockers (including idle sessions with open transactions), worker threads, alert history, and collector status (central only).
- **Header:** total wait over the last 10 minutes, colored green, orange or red against the threshold, plus the number of sessions blocked now.

The viewer can run on your own PC. Point `SERVER=` at the remote server. It uses Windows authentication; add `-SqlLogin` at the end of the `start` line to be prompted for a SQL login instead. Times are shown in your PC's local time.

## Alerts
The *Alert* job runs every minute and fires when **either**:
- the total wait in the last `WindowMin` minutes is at least `WaitThresholdSec` (default 10,000 s in 10 minutes). Each 10-second sample of an active request counts as 10 s.
- at least `BlockedThreshold` sessions are blocked in the latest sample (default 20).

After an alert it stays quiet for `CooldownMin` (15 minutes); in the central install this works separately for each server. Every alert is written to `AlertLog`, which the viewer shows on the *Alert history* tab. It is also emailed if a mail profile is set. The email names the head blocker (even if it is idle with an open transaction) and the top 10 statements with their main wait.

Change the settings at any time, without reinstalling:
```sql
UPDATE DBAWait.dbo.MonitorConfig      -- or DBAWaitCentral
SET MailProfile = N'DBA Mail', Recipients = N'you@yourcompany.com',
    WaitThresholdSec = 10000, BlockedThreshold = 20, WindowMin = 10, CooldownMin = 15;
```

## Story reports: plain English
In SSMS, press **Ctrl+T** (Results to Text) first, then:
```sql
EXEC DBAWait.dbo.usp_WaitStory;                                            -- last 24 hours
EXEC DBAWait.dbo.usp_WaitStory @From = '2026-11-02 09:00', @To = '2026-11-02 17:00';
EXEC DBAWait.dbo.usp_WaitStoryForStatement @StatementId = 12;              -- one query (#id from the viewer)
SELECT * FROM DBAWait.dbo.WaitGlossary;                                    -- wait types in plain English
-- central: DBAWaitCentral.dbo.usp_WaitStory @ServerName = N'PRODSQL01', ...
```

**`usp_WaitStory`** has these chapters:
1. In short
2. Where the time went
3. Busiest moments
4. Queries that waited the most
5. Blocking incidents
6. Server capacity
7. Alerts
8. Day by day (periods over 2 days)
9. What to look at

Example of a blocking incident:
> *11:17 – 12:28 (1 h 11 min): session 57 (DOMAIN\dba, using SSMS on WS01) was holding locks that other queries needed. It was sitting idle with an unfinished transaction, so everything behind it had to wait until someone committed or rolled it back. Up to 5 queries were stuck behind it at once (17 different sessions), losing 4 h 55 min in total. Its last command: "BEGIN TRAN; ALTER INDEX ALL ON dbo.Posts REBUILD;"*

**`usp_WaitStoryForStatement`** covers one query:
- which query it is, and for dynamic SQL, which proc call ran it
- how much time it took
- what it spends its time on
- when it runs, and whether it is getting worse
- who got in its way, and who runs it
- what to look at

Times are the (central) server's local time, and daylight saving is handled. The glossary wording is in `dbo.WaitGlossary` and can be edited freely.

## Gentle on the server
- **Short snapshots:** the collector reads DMVs only and runs with `SET LOCK_TIMEOUT 500`, so it never waits behind blocking.
- **Small cleanup:** data older than `RetentionDays` (default 14) is removed in small batches.
- **Disconnects after each query:** the viewer only reads the monitoring database, and closes its connection after every query.
- Running both installs on the same server works, but each one records the other's snapshot queries. Keep only one in normal use.

## Requirements
- **SQL Server:**
  - Standalone and central: SQL Server 2017 or later.
  - Remote servers: SQL Server 2016 or later.
  - SQL Agent running, and sysadmin to install.
- **Viewer:**
  - Windows PowerShell 5.1. On PowerShell 7 it relaunches itself in 5.1, because the chart library is .NET Framework only.
  - Read access to the monitoring database.
  - `VIEW SERVER STATE` for *Open execution plan*, and `VIEW ANY DEFINITION` for *Show whole procedure*.

## Remove
```sql
EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWait - Collect';   -- or DBAWaitCentral - Collect
EXEC msdb.dbo.sp_delete_job @job_name = N'DBAWait - Alert';     -- or DBAWaitCentral - Alert
DROP DATABASE DBAWait;                                           -- or DBAWaitCentral
```

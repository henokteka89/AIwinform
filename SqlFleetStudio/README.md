# SQL Fleet Studio

An SSMS-style query window for many servers at once. Pick servers from a cached list, run a query on all of them in parallel, and get one merged result grid with a `ServerName` column. No more reconnecting window by window.

## Start
Double-click **`Start-SqlFleetStudio.cmd`**. The window opens with no console behind it.

Or from PowerShell:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\SqlFleetStudio.ps1
```

## First run
1. Open **Settings** and enter the central server, database and server list query.
   - The first column of the query must be the server name. Extra columns (Environment, Group, ...) become filterable columns.
   - The default query reads a Central Management Server (`msdb`). Replace it with your own inventory table if you have one.
2. Click **Save + refresh server list**. The list is cached locally and loads instantly next time; **Refresh** re-queries it.
3. Servers can also be added by hand with **Add...**.

## Tabs
- **Query:**
  - Check servers on the left (filter box: space-separated words, all must match). PROD servers are shown in red.
  - **F5** / **Ctrl+E** runs the selection, or everything if nothing is selected. `GO` and `GO n` work. **Esc** cancels.
  - Results are merged per result set across servers. **Messages** shows PRINT output, row counts and errors; **Servers** shows status, rows and duration per server.
  - Export the current result to CSV, or save the editor contents to the library.
- **Query Library:** searchable, grouped by category. Copy a query, or **Send to editor** (double-click does the same): it is appended and selected, so F5 runs just that query. Comes with starter queries (version, database sizes, disk space, backups, active requests, blocking, failed jobs, AG health).
- **Settings:** central server, authentication (Windows or SQL; the SQL password is asked once per session and never saved), encryption, ApplicationIntent, MultiSubnetFailover, timeouts, row cap, parallelism, PROD pattern.

## Production safety
- **Read-only by default:**
  - Write statements (INSERT/UPDATE/DELETE/MERGE, DDL, SELECT INTO, GRANT, BACKUP/RESTORE, KILL, transaction control, xp_cmdshell, ...) are blocked before anything is sent. Temp tables (`#t`) and table variables (`@t`) are allowed.
  - As a backstop, every run is wrapped in a transaction that is **always rolled back**.
  - EXEC, DBCC, sp_executesql and OPENQUERY ask for confirmation, since effects outside the transaction cannot be undone.
- **Writes are one-shot:** tick **Allow writes (this run only)** and type `EXECUTE` to confirm. The tick clears itself after the run.
- **Polite on the server:** `SET DEADLOCK_PRIORITY LOW` and `SET LOCK_TIMEOUT` (10 s by default) on every session; query timeout, max rows per result set, and a limit on servers in parallel.
- **Confirmation** before a read-only run against many servers (10 by default).
- **Audit log:** every run (who, when, servers, database, mode, full query, outcome) is appended to `%APPDATA%\SqlFleetStudio\execution.log`.

## Data
Everything lives in `%APPDATA%\SqlFleetStudio`: `settings.json`, `servers.json` (cache and checked servers), `library.json` (copy it to share queries), `draft.sql` (editor contents), `execution.log` and `error.log` (full details of any error, including the script line).

## Requirements
- Windows PowerShell 5.1 (uses `System.Data.SqlClient`).
- Permissions on each target are whatever your queries need; the starter library mostly needs **VIEW SERVER STATE** and read access to `msdb`.

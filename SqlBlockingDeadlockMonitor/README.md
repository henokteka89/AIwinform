# SQL Server Blocking & Deadlock Monitor

A WinForms tool that tracks **only blocking and deadlocks**. It explains each one in plain English, with recommended fixes.

## Start
Double-click **`Start-SqlBlockingDeadlockMonitor.cmd`**. The window opens with no console behind it, and no admin rights are needed.

Or from PowerShell:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SqlBlockingDeadlockMonitor.ps1
```

## Tabs
- **Blocking:**
  - Live blocking chains and the head blocker.
  - Who / what / why, with fixes.
  - KILL guidance.
- **Deadlocks:**
  - Pulls deadlocks from the built-in `system_health` session: the ring buffer is fast, the event files are complete.
  - You can also **Open** or **Paste** any deadlock graph (`.xdl` / XML) for a full breakdown and recommendations.
- **AI Prompt:** a ready-to-paste prompt for an AI assistant, if you want a deeper review.

## Gentle on the server
- **Connect, query, disconnect on every refresh.** Connection pooling is off, so the connection really closes between cycles.
- **No overlap.** The next refresh is scheduled only after the previous one finishes. The optional **Stop after** setting ends monitoring automatically.
- **The monitoring queries never cause blocking or wait behind it:**
  - `SET LOCK_TIMEOUT 500`: gives up after 0.5 s if anything is locked.
  - `SET DEADLOCK_PRIORITY LOW`: the monitor is always the preferred deadlock victim.
  - READ UNCOMMITTED plus DMVs only: no shared locks on user data.
  - `MAXDOP 1` and short command timeouts.

## Requirements
- Windows PowerShell 5.1, or PowerShell 7+ on Windows.
- The login needs **VIEW SERVER STATE**. On SQL Server 2022 and later, **VIEW SERVER PERFORMANCE STATE** is enough.

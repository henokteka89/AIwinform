# SqlPatchGui: single-server SQL Server patcher

A single PowerShell file with a Windows GUI. Run it elevated:

Easiest: double-click **Start-SqlPatchGui.cmd** (kept in the same folder as the script, e.g. C:\Apps\PatchAutomation). It asks for elevation (UAC) and opens the window with no console behind it.
Or from an elevated PowerShell:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\SqlPatchGui.ps1
```
Close the tool with the window's X button, not Ctrl+C. Ctrl+C is ignored while the window is open, and if the PowerShell engine stops anyway, the window now closes cleanly instead of looping error dialogs.
Requires dbatools (`Install-Module dbatools`). Run it as an account that is sysadmin on SQL Server and local admin on the servers.

## Where to run it
| Run from | Works when |
|---|---|
| Central server | WinRM from central to the destination is open. |
| Your workstation | WinRM is open, and the workstation can reach the central server's share. |
| The destination itself | Always. Use this when WinRM to the destination is blocked: put the destination's own name in "Destination". |

## Flow
1. **Check**
   - Fills **Server info**: version, edition, build/CU/KB, latest available CU, support end date, SQL start time, last Windows boot and uptime, free memory, C: free space, AG role.
   - Runs **Guardrail checks**:
     - FAIL: pending reboot; < 6 GB free on C:; AG primary; AG unhealthy; FCI (use the pipeline instead); WinRM blocked (install only); not elevated (local); no full/diff backup in 26h (WARN instead when in an AG).
     - WARN: log backups older than 60 min; other instances on the host; open transactions > 5 min; Agent jobs running; databases not ONLINE; out of support; uptime > 90 days.
   - Fills the right-hand list with **every build for that SQL version**: installed (green), newer, older (hidden by default). Click one to select it.
2. **Download & share**
   - Downloads the selected KB into the folder on the central server, e.g. `D:\SQLPatch`. If the tool isn't running on the central server, it writes through `\\central\D$`.
   - Shares that folder as `SQLPatch$` (read: Everyone, full: Administrators).
   - If a share with that name already exists, the tool reuses it and never deletes it.
3. **Install selected**: does steps 1–2 if needed, then:
   - Copies the installer to `C:\SQLPatch` on the destination, with progress.
   - Runs setup silently. The status line shows elapsed time and the latest line of setup's Detail.txt.
   - Reads the exit code and Summary.txt.
   - Reboots Windows if setup requires it (optional), or always (optional).
   - Validates, deletes the installer copy, and removes the share (optional).

## Which reboot?
- **Restart SQL service** restarts only SQL Server and Agent. CU setup already does this by itself, so you rarely need this button.
- **Reboot Windows** restarts the whole server/VM. You only need it when setup says "reboot required" (exit 3010), or when a reboot was already pending before patching. The checks flag pending reboots as FAIL, so reboot first in that case.
- Reboot times are recorded in Server info: requested, offline, back online with the new boot time, and SQL accepting connections.

## Files
- Log: `%ProgramData%\SqlPatchGui\logs\SqlPatchGui_<timestamp>.log`
- Last-used values: `%APPDATA%\SqlPatchGui\settings.json`
- When you close the window, the tool asks whether to remove the share if it's still there.

# SqlPatchGui: single-server SQL Server patcher

A Windows GUI that patches one SQL Server instance at a time with a Cumulative Update or security update (GDR). It does it safely:
- guardrail checks before anything changes
- a temporary download share on a central server
- staging the installer during the day
- a silent install with live progress
- controlled reboots
- validation, and setup logs copied back to you

## Start
1. Put `SqlPatchGui.ps1` and `Start-SqlPatchGui.cmd` in the same folder, for example `C:\Apps\PatchAutomation`.
2. Double-click **`Start-SqlPatchGui.cmd`**. It asks for elevation (UAC), because patching needs admin rights, and opens the window with no console behind it.

**Requirements:**
- [dbatools](https://dbatools.io): `Install-Module dbatools`.
- Run as an account that is sysadmin on SQL Server and local admin on the servers.

Close the tool with the window's X button.

## Where to run it
| Run from | Works when |
|---|---|
| Central server | WinRM from the central server to the destination is open |
| Your workstation | WinRM is open, and the workstation can reach the central server's share |
| The destination itself | Always. Use this when WinRM to the destination is blocked: enter the destination's own name |

## Typical flow
**During the day, with no impact on SQL:**
1. **Check:** server info and guardrails. The update list on the right shows every build for that SQL version, with type: *Cumulative update*, *Security update (CU+GDR)*, *Security update (GDR)*. Select one.
2. **Download, share & stage:**
   - Downloads the update into the folder on the central server, e.g. `D:\SQLPatch`.
   - Shares the folder temporarily.
   - Copies the installer to `C:\SQLPatch` on the destination.
   - Checks the copy is complete and Microsoft-signed.
   - Removes the share (optional).

**In the maintenance window:**

3. **Check**, select the same update, then **Install selected**:
   - Uses the staged installer.
   - Runs setup silently, with elapsed time and the current setup-log line.
   - Reboots Windows only if setup requires it, or always if you tick that option.
   - Validates the result and copies the setup logs back to this machine.

**Afterwards:** **Clear installers** deletes the staged files from the destination when you're done.

## Guardrail checks
| Result | Checks |
|---|---|
| FAIL (blocks Install) | Pending reboot (including pending file renames), less than 6 GB free on C:, AG primary, AG unhealthy, FCI, WinRM blocked (for install), not elevated (local), no full/diff backup in 26h (WARN instead when in an AG) |
| WARN | Log backups older than 60 min, other instances on the host, transactions open more than 5 min, Agent jobs running, databases not ONLINE, out of support, uptime over 90 days |

- **Override failed checks** lets you install despite FAIL items. It won't install an update that's already installed or older than the current build.
- If the server was patched or rebooted after your last Check, Install re-runs the checks itself.

## Options
- **Patch all instances on host** is *off* by default: only the instance you entered is patched (`/InstanceName`). Shared components are always updated by a CU.
- **Auto-reboot Windows if setup requires it** is on by default. **Always reboot Windows after install** is off by default.
- **Remove share when finished** is on by default. **Remove share** also works in a later session: the tool recognises its own shares by their description and never removes shares it didn't create.

## Reboot vs SQL restart
- **Restart SQL service:** restarts SQL Server and Agent only. CU setup already does this itself.
- **Reboot Windows:** restarts the whole server/VM. You only need it when setup reports *reboot required* (exit 3010), or when a reboot was already pending.
  - The restart has to be accepted by the server. If one method fails, it tries the others: WinRM, CIM, WMI, `shutdown /m`.
  - If the server hasn't gone down within 3 minutes, it retries with another method.
  - Times are recorded in Server info: requested, back online with the new boot time, and SQL accepting connections.

## Setup logs and exit codes
After every install, setup's `Summary.txt` (plus that run's `Detail.txt` and rule report) is copied to `%ProgramData%\SqlPatchGui\logs\setup\<server>_<time>\`.
- **On failure:** the summary opens in Notepad automatically.
- **Setup summary button:** fetches the latest summary from the destination at any time.
- **Check:** shows the result of the last setup run.

The tool explains exit codes. For example, `0x84BE0BC2` means *a Windows restart is pending: reboot, then run again*.

## Files
- **Tool log:** `%ProgramData%\SqlPatchGui\logs\SqlPatchGui_<timestamp>.log`
- **Last-used values:** `%APPDATA%\SqlPatchGui\settings.json`

## Multi-server pipeline
For patching many servers on a schedule (AG- and FCI-aware, one replica per run, driven by a `DBAPatch` control database), see [`pipeline/`](pipeline/README.md):
- `Invoke-SqlPatch.ps1`: entry point (Discover / Download / Plan / Patch)
- `module/SqlPatch.psm1`: pipeline logic (dbatools)
- `sql/01_PatchControl.sql`, `sql/02_SampleInventory.sql`: control database and example inventory

# PITR-RestoreBuilder

> **WARNING: do not use this in production until you have tested it in a non-production environment.**
> Run it end to end in a lab or test environment first: build the chain, run preflight, generate the scripts, run the restore, and check the restored data. Read every generated script before you run it, especially the `MOVE` paths, `STOPAT`, `REPLACE` and `RECOVERY`. The tool is provided as is, with no warranty. You are responsible for what you run against your servers.

A WinForms GUI that builds a **point-in-time restore** of a SQL Server database from msdb backup history. It finds the right FULL, DIFF and LOG backups for a `STOPAT` time, checks the destination server for anything that would stop the restore, and generates the restore script and a SQL Agent job.

**Read-only and generate-only by default.** Nothing is restored by the tool itself.

## Start
- **Normal use:** double-click **`Start-PITR-RestoreBuilder.cmd`**. No admin rights are needed.
- **TDE certificate transfer:** double-click **`Start-PITR-RestoreBuilder-Admin.cmd`** (asks for elevation). Elevation is only needed to create the temporary SMB share used to move the certificate.
- Downloaded files are flagged by Windows. Unblock them once:
  ```powershell
  Get-ChildItem C:\Apps\PITR-RestoreBuilder | Unblock-File
  ```

## How it works
1. **Build chain** (source, reads msdb only). Picks the latest FULL that finished before STOPAT, then the DIFF based on that FULL, then the log backups up to and including the first log that ends after STOPAT, all linked by LSN. Stops with a clear message on a broken log chain, a recovery fork, damaged or snapshot backups, or a STOPAT that no log reaches. Modes: Auto, FULL only, FULL + DIFF, FULL + DIFF + LOG, plus three **Continue** modes for a restore already in progress (see below).
2. **Preflight** (destination, read-only). Runs `RESTORE HEADERONLY` on every backup file and checks that the LSNs match msdb, so overwritten or missing files are caught. Also checks `FILELISTONLY`, the SQL Server version, permissions, whether the target database exists or is in an AG, blob credentials, TDE and backup-encryption certificates, target folders, disk space, restore buffer memory and SQL Agent. Each failure says why and how to fix it.
3. **Generate scripts.** Writes these files to the output folder:

| File | Purpose |
|---|---|
| `01_restore.sql` | Restore script for SSMS: one `GO` batch per restore, plus a wrong-server / existing-database guard |
| `02_create_sql_agent_job.sql` | SQL Agent job, owner `sa`, one-time schedule. **One step per backup type:** FULL, DIFF, all LOGs, then RECOVERY if chosen |
| `03_tde_certificate_transfer.sql` | Manual certificate backup and restore steps (only when needed) |
| `04_blob_credential.sql` | SAS credential for the blob container (only when restoring from URL) |
| `05_azcopy_to_blob.ps1` | Copies the backups to blob storage when the destination can't reach the share (DMZ to Azure, for example) |
| `00_preflight_report.txt`, `chain.csv` | What was checked and which backups were picked |

## Restore defaults
- Every step: `WITH FILE = n, NORECOVERY, STATS = 10, BUFFERCOUNT = 450, MAXTRANSFERSIZE = 4194304`. All of these can be changed in the GUI.
- `STOPAT` is added only on the last LOG step, in **source server local time**, which is the time msdb and the transaction log use.
- `RECOVERY` is used only when you tick it; otherwise the database is left RESTORING. `REPLACE` is used only when you tick it.
- `CHECKSUM` is not added, on the assumption that backups are already taken WITH CHECKSUM.
- BUFFERCOUNT 450 × 4 MB is about 1.8 GB of memory outside the buffer pool. Preflight warns when that is more than 25% of the destination's free memory.

## WITH MOVE
Strategies:
- **Auto** picks destination folders in this order:
  1. a destination database whose name contains your pattern
  2. the folder most user databases use
  3. the instance default paths
- **Largest free disk:** the drive or mount point with the most free space, using a folder on it that user databases already use. If no such folder exists, it uses the drive root and the comment asks you to pick a subfolder.
- **Manual:** the folders you enter.
- **Instance default:** the instance's default data and log paths.

**Browse the destination disks.** The `...` button next to the Data (MDF), NDF and Log (LDF) boxes opens the destination's drives and mount points with free and total space. Expand them to choose a folder. The listing shows what the **destination SQL Server service account** can see (read-only: `sys.dm_os_enumerate_fixed_drives` / `dm_os_volume_stats`, and `xp_dirtree` for folders, which needs sysadmin). **Set folder for selected rows** applies a chosen folder to the selected files in the MOVE grid. The folder must already exist, because RESTORE does not create folders.

Folders you type in always win. Each `MOVE` line carries a `-- REVIEW` comment saying where the folder came from. You can edit any target path in the grid on tab 2.

## Continue a restore in progress (DIFF and/or LOG only)
For example, the FULL was restored yesterday WITH NORECOVERY and you now want to apply the DIFF and logs. Pick one of these modes on tab 1:

| Mode | Restores |
|---|---|
| Continue: DIFF only | the newest DIFF on the FULL that is already restored |
| Continue: DIFF + LOGs | that DIFF (if one is newer than where the database is), then the logs up to STOPAT |
| Continue: LOGs only | the logs from where the database is now, up to STOPAT |

- **Restoring DBs...** lists the destination's databases that are in RESTORING state, with their last restore. Pick the one to continue.
- The tool reads where that database is now from the destination, without changing it: `redo_start_lsn` and `differential_base_lsn` in `sys.master_files`, plus `restorehistory`. It then builds the chain from that point using the source msdb.
- No FULL and no `MOVE` are included, because the files already exist.
- The script and the job refuse to run if the database is no longer RESTORING.
- **FULL still restoring?** You don't have to wait. SQL Server only records the next-restore point when a RESTORE finishes. The tool therefore finds the running RESTORE on the target (percent done, time left), identifies which FULL/DIFF it is using (from this tool's step comments, or by file name), and plans the DIFF and logs from that backup. The continue script and job **wait** (up to 48 hours, checking every minute) until the running restore has finished, then check that the database is RESTORING and carry on. Seeing another session's restore needs VIEW SERVER STATE. The job runs as the SQL Agent service account, which has it.

## Backup paths and shares
msdb stores backup paths as the source server saw them, for example `C:\Apps\BKUP\...`. From another server that path points at the destination's own C: drive. When **Convert local paths** is ticked (the default):
- The tool **reads** the shares on the source server (WinRM, then DCOM). It never creates or changes a share.
- It uses the deepest regular share that covers every backup folder in the chain, for example `\\SRC\bkup\...` rather than `\\SRC\C\...`, or a share that only covers the FULL folder.
- An admin share such as `C$` is used only as a last resort, with a warning.
- **Remap from / to** overrides everything, and also handles moving to a different share or to blob storage.

The destination's **SQL Server service account** must be able to read the share, at both the share level **and** the NTFS level. With a virtual account (`NT SERVICE\MSSQLSERVER`) on a domain, grant the destination computer account `DOMAIN\HOST$`. Being able to open the share yourself is not enough, because SQL Server reads the files as its own service account.

## Optional execution
These only run when **Allow execution** is ticked, and each one lists what it will change and asks first:
- azcopy upload of the backup files to blob storage
- `CREATE CREDENTIAL` (SAS) on the destination
- TDE / backup-encryption certificate transfer:
  - `BACKUP CERTIFICATE` on the source to a temporary share on this PC
  - copy to `C:\TDE` on the destination
  - `CREATE MASTER KEY` (if missing) and `CREATE CERTIFICATE` on the destination
  - the share is removed afterwards. The certificate **and private key** files stay on disk; move them somewhere safe and delete them.
- Create the SQL Agent job, and optionally start it. The tool refuses if options changed after the scripts were generated.

## Job notifications (tab 3)
- **E-mail to** (one or more addresses, separated by `;`): the job gets a **Notify: succeeded** and/or **Notify: failed** step that sends through Database Mail. The **profile is chosen when the job runs**: the default profile, else the first one. The failure e-mail names the step that failed and its error message. The job still ends as FAILED, and a mail problem never turns a successful restore into a failed job.
- **Operator:** **Load operators** lists the destination's enabled SQL Agent operators. The job's built-in notification (`@notify_level_email` on success and/or failure) uses it. This needs a mail profile set in SQL Agent > Properties > Alert System.
- **On success / On failure** choose which events notify. All of these are remembered on this PC.
- Preflight checks that Database Mail XPs is on, that a profile exists and which profile would be used, that the operator exists and has an e-mail address, and that SQL Agent has a mail profile.

## Server lists and cache (tab 1)
- **Server list from** defaults to the machine the tool runs on. Load the list from your own inventory table (first column = server name) or from a Central Management Server.
- The query is read-only: it is checked for blocked keywords and runs inside a transaction that is always rolled back.
- Servers you connect to are remembered with their auth type and SQL login name. **Passwords are never saved.**
- Settings are stored per user in `%APPDATA%\PITR-RestoreBuilder\settings.json`.

## Requirements
- Windows PowerShell 5.1 (no modules needed: System.Data.SqlClient + WinForms).
- **Source:** read access to msdb and `VIEW SERVER STATE` (for the TDE lookup).
- **Destination:** `CREATE ANY DATABASE` (dbcreator) for `RESTORE HEADERONLY` and the restore itself, and `VIEW SERVER STATE`. Creating the job needs SQL Agent rights in msdb (sysadmin, or SQLAgentOperatorRole plus job ownership by `sa`).
- `RESTORE FROM URL` with SAS needs SQL Server 2016+ on the destination.
- azcopy v10 in PATH for the optional blob upload.

## Known limits
- Backups taken on another AG replica are not in this server's msdb. Run the tool against the replica that took them.
- Snapshot/VDI backups (third-party tools) are skipped. Restore those with the tool that created them.
- Files added to the database after the FULL are recreated at their original path during the log restore. Preflight warns about this.
- TDE certificate rotation: backups taken before a rotation need the old certificate too.
- The window is single-threaded, so a large azcopy upload from inside the tool freezes it until the upload finishes. For big copies, run `05_azcopy_to_blob.ps1` separately.

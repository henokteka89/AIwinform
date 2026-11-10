# TdeCertMigration

> **WARNING: test this in a non-production environment before you use it in production.**
> Run What-If, read every step, then try Execute between two test instances first. The tool is provided as is, with no warranty.

A WinForms GUI that moves **TDE certificates with their private keys** from a source SQL Server to a destination SQL Server, so encrypted databases (or their backups) can be restored or attached on the destination.

**What-If by default.** It shows every step, the exact T-SQL or file operation, the server it runs on and what it does. Nothing changes until you pick **Execute** and type the destination server name to confirm.

## Start
Double-click **`Start-TdeCertMigration.cmd`**. You don't need to run it as administrator on this PC.

Files downloaded from the internet are flagged by Windows. Unblock them once:
```powershell
Get-ChildItem C:\Apps\TdeCertMigration | Unblock-File
```

## Tabs
| Tab | What you do there |
|---|---|
| **1. Source** | Connect (Windows or SQL login, optional encryption). Lists the certificates in `master` with thumbprint, expiry, private-key protection and the databases that use each one for TDE. Certificates used by TDE are ticked for you. |
| **2. Destination** | Connect. Shows whether `master` has a **database master key (DMK)** and whether the service master key encrypts it. If there's no DMK, tick **Create the database master key** and enter a password. Lists the certificates already on the destination. Choose what happens when a *different* certificate has the same name: **Stop** (the default) or create it with a suffix. |
| **3. Transfer and password** | Export folder (as the source SQL Server sees it) and import folder (as the destination sees it), plus the same folders as this PC reaches them. By default it uses the admin share `\\host\X$\...`; untick the auto option to type your own share. Also the **private-key password**, used for `ENCRYPTION BY PASSWORD` on the source and `DECRYPTION BY PASSWORD` on the destination. |
| **4. Plan and run** | **What-If** fills the step grid. Click a step to see what it does, where it runs, which login it runs as, and its script. The full plan appears as one SQLCMD-mode script with the passwords masked, and you can save it. **Execute** runs the same plan with a progress bar and colours each step's status. It stops at the first failure and writes everything to the log. |

## Steps (in this order)
| # | Runs on | Step | Changes? |
|---|---|---|---|
| 1 | Source | Check that each certificate still exists with the same thumbprint and a private key protected by the master key | no |
| 2 | Destination | Check the DMK. If it isn't encrypted by the service master key, check that its password opens it | no |
| 3 | Source | `xp_create_subdir` for the export folder, only if it's missing | folder |
| 4 | Source | `BACKUP CERTIFICATE ... TO FILE ... WITH PRIVATE KEY (FILE ..., ENCRYPTION BY PASSWORD ...)` to **new time-stamped files** | files |
| 5 | Destination | `xp_create_subdir` for the import folder, only if it's missing | folder |
| 6 | This PC | Copy the `.cer` / `.pvk` files and compare SHA-256 hashes. It never overwrites a file. | files |
| 7 | Destination | `CREATE MASTER KEY`, only if there is none **and** you ticked the option | DMK |
| 8 | Destination | `CREATE CERTIFICATE ... FROM FILE ... WITH PRIVATE KEY (FILE ..., DECRYPTION BY PASSWORD ...)` | certificate |
| 9 | Destination | Verify that the thumbprints match the source and that the private key is protected by the master key | no |
| 10 | This PC | Optional: delete the copies in the import folder after the verify step passes. The source export is kept. | files |

## Production safety
- **Nothing is dropped, altered, replaced or overwritten.** The tool runs only the statements listed above.
- **No database is touched.** It does not run `ALTER DATABASE`, `CREATE DATABASE ENCRYPTION KEY` or `SET ENCRYPTION ON`. It only puts the certificate on the destination.
- A certificate that's already on the destination (same thumbprint) is **skipped**. If a different certificate has the same name, that certificate **stops** (or gets a suffix, if you chose that).
- An existing DMK is never changed.
- **Wrong-server guard:** every SQL step begins with `IF @@SERVERNAME <> '<planned server>' RAISERROR ... RETURN`.
- **Order:** read-only checks run first, then the export, then the copy with its hash check. The destination is changed only after all of these succeed. A failure stops the run, and the steps after it don't run.
- **The plan is rebuilt from fresh SELECTs** every time you click the button, so it matches what's on the servers right now.
- **Passwords** are never shown in the plan, the log or the saved script, which use `<<CERT_PASSWORD>>` / `<<DMK_PASSWORD>>` placeholders.
- What-If runs only SELECTs, plus a `Test-Path` on the admin-share roots.

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+ on Windows. No modules needed.
- **sysadmin** on both instances. `xp_create_subdir` needs sysadmin; BACKUP and CREATE CERTIFICATE need CONTROL on the certificate or database.
- This PC must be able to read the source folder and write to the destination folder. With the default admin share, that means you're a local admin on both hosts and SMB is open. Otherwise, type a share you can reach.
- The SQL Server service creates the folders itself, under the instance's default backup folder (`<backup folder>\TDE_Migration`), so it already has the rights it needs to write and read the files.

## After a run
- The export folder holds the **certificate AND its private key**. Move the files to a safe place and keep the password in your vault.
- If a DMK was created, back it up: `BACKUP MASTER KEY TO FILE = '...' ENCRYPTION BY PASSWORD = '...'`.
- **Availability groups:** every replica needs the certificate. Run the tool once for each replica.
- Log: `%LOCALAPPDATA%\TdeCertMigration\TdeCertMigration_<date>.log`

## Running the plan by hand
Save the script from tab 4 and open it in SSMS with **Query > SQLCMD Mode**. Replace the password placeholders and run it one step at a time. The "this PC" steps (copy, delete) are written as comments, so do those yourself.

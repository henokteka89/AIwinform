# TdeCertMigration

> **WARNING: test this in a non-production environment before you use it in production.**
> Run What-If, read every step, then try Execute between two test instances first. The tool is provided as is, with no warranty.

A WinForms GUI that moves **TDE certificates with their private keys** from a source SQL Server to a destination SQL Server, so encrypted databases (or their backups) can be restored or attached on the destination.

**Two transfer methods**
- **Direct (default, SQL Server 2012+):** the certificate and its private key go over the two SQL connections, with the key encrypted by your password: `CERTENCODED` / `CERTPRIVATEKEY` on the source, then `CREATE CERTIFICATE ... FROM BINARY` on the destination. **No folders, files, shares or copy**, so Windows "access denied" errors on folders such as `C:\TDE\Backup` can't happen. The bytes stay only in the tool's memory and are cleared when the run ends. Before anything is created, the tool checks that the SHA-1 of the certificate bytes equals the thumbprint.
- **Files:** `BACKUP CERTIFICATE` to `.cer` / `.pvk` files, copy them, then `CREATE CERTIFICATE ... FROM FILE`. Use this for SQL Server 2008 R2 or older, or when you also want key files to keep in a vault.

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
| **3. Transfer and password** | **Transfer method** (Direct or Files) and the **private-key password**. File method only: export folder (as the source SQL Server sees it) and import folder (as the destination sees it). **How this PC reaches them:** **temporary shares** (the default, see below), admin shares `\\host\X$\...`, or paths you type. Also the **private-key password**, used for `ENCRYPTION BY PASSWORD` on the source and `DECRYPTION BY PASSWORD` on the destination. |
| **4. Plan and run** | **What-If** fills the step grid. Click a step to see what it does, where it runs, which login it runs as, and its script. The full plan appears as one SQLCMD-mode script with the passwords masked, and you can save it. **Execute** runs the same plan with a progress bar and colours each step's status. It stops at the first failure and writes everything to the log. |

## Steps, direct method
| # | Runs on | Step | Changes? |
|---|---|---|---|
| 1 | Source | Check that each certificate exists with the same thumbprint and a private key protected by the master key | no |
| 2 | Destination | Check the DMK (and its password if needed) | no |
| 3 | Source | `SELECT CERTENCODED(...), CERTPRIVATEKEY(..., <password>)` into memory, then check that the SHA-1 equals the thumbprint | no |
| 4 | Destination | `CREATE MASTER KEY`, only if there is none **and** you ticked the option | DMK |
| 5 | Destination | `CREATE CERTIFICATE ... FROM BINARY = 0x... WITH PRIVATE KEY (BINARY = 0x..., DECRYPTION BY PASSWORD ...)` | certificate |
| 6 | Destination | Verify that the thumbprints match and that the private key is protected by the master key | no |

In What-If and the saved script, the bytes appear as `<<CERTIFICATE_BYTES_FROM_SOURCE>>` / `<<ENCRYPTED_PRIVATE_KEY_BYTES_FROM_SOURCE>>`.

## Steps, file method (in this order)
| # | Runs on | Step | Changes? |
|---|---|---|---|
| 1 | Source | Check that each certificate still exists with the same thumbprint and a private key protected by the master key | no |
| 2 | Destination | Check the DMK. If it isn't encrypted by the service master key, check that its password opens it | no |
| 3 | Source | `xp_create_subdir` for the export folder, only if it's missing | folder |
| 4 | Source | `BACKUP CERTIFICATE ... TO FILE ... WITH PRIVATE KEY (FILE ..., ENCRYPTION BY PASSWORD ...)` to **new time-stamped files** | files |
| 5 | Destination | `xp_create_subdir` for the import folder, only if it's missing | folder |
| 6a | Source host | **Temporary share** on the export folder: hidden (`TDEMIG_SRC_<stamp>$`), `-Temporary`, READ for you only | share |
| 6b | Destination host | **Temporary share** on the import folder: hidden (`TDEMIG_DST_<stamp>$`), `-Temporary`, CHANGE for you only | share |
| 6c | This PC | Copy the `.cer` / `.pvk` files and compare SHA-256 hashes. It never overwrites a file. | files |
| 6d | This PC | **Remove the temporary shares.** The folders and files are kept. This step runs even if an earlier step failed. | share removed |
| 7 | Destination | `CREATE MASTER KEY`, only if there is none **and** you ticked the option | DMK |
| 8 | Destination | `CREATE CERTIFICATE ... FROM FILE ... WITH PRIVATE KEY (FILE ..., DECRYPTION BY PASSWORD ...)` | certificate |
| 9 | Destination | Verify that the thumbprints match the source and that the private key is protected by the master key | no |
| 10 | This PC | Optional: delete the copies in the import folder after the verify step passes. The source export is kept. When this is ticked, the destination share stays until this step, then it's removed. | files |

## Temporary shares (default)
- The shares are created from this PC over WinRM, or DCOM if WinRM doesn't answer, with `New-SmbShare -CimSession ... -Temporary`. Windows also drops a temporary share on reboot.
- They're hidden (`$`), named `TDEMIG_SRC_<stamp>$` / `TDEMIG_DST_<stamp>$`, and grant only your account (read on the source, change on the destination). Existing shares, the folders and their NTFS permissions aren't changed.
- The tool removes **only the shares it created in this run**, by exact name, and only if each one still points at the same folder. **The folder and the files stay.**
- Removal always runs: after the copy, after a failure, and if you close the window. If it can't remove a share, the log names it, along with the `Remove-SmbShare` command to run by hand.
- No share is needed when a folder is already a UNC path or the SQL Server runs on this PC.
- Needs: local admin on both hosts, and WinRM (5985) or DCOM/RPC reachable from this PC.
- **Failover cluster instances:** a share on a clustered disk may need a cluster scope. If the share step fails there, choose **Admin shares** or **Paths I type**.

## Production safety
- **Nothing is dropped, altered, replaced or overwritten.** The tool runs only the statements listed above.
- **No database is touched.** It does not run `ALTER DATABASE`, `CREATE DATABASE ENCRYPTION KEY` or `SET ENCRYPTION ON`. It only puts the certificate on the destination.
- A certificate that's already on the destination (same thumbprint) is **skipped**. If a different certificate has the same name, that certificate **stops** (or gets a suffix, if you chose that).
- An existing DMK is never changed.
- **Wrong-server guard:** every SQL step begins with `IF @@SERVERNAME <> '<planned server>' RAISERROR ... RETURN`.
- **Order:** read-only checks run first, then the export, then the copy with its hash check. The destination is changed only after all of these succeed. A failure stops the run, and the steps after it don't run.
- **The plan is rebuilt from fresh SELECTs** every time you click the button, so it matches what's on the servers right now.
- **Passwords** are never shown in the plan, the log or the saved script, which use `<<CERT_PASSWORD>>` / `<<DMK_PASSWORD>>` placeholders.
- What-If runs only SELECTs, plus `Test-WSMan` (temporary shares) or a `Test-Path` on the admin-share roots.

## Requirements
- Windows PowerShell 5.1 or PowerShell 7+ on Windows. No modules needed.
- **sysadmin** on both instances. `xp_create_subdir` needs sysadmin; BACKUP and CREATE CERTIFICATE need CONTROL on the certificate or database.
- This PC must be able to read the source folder and write to the destination folder through SMB. With temporary shares or admin shares, that means you're a local admin on both hosts. Otherwise, type a share you can reach.
- The SQL Server service creates the folders itself, under the instance's default backup folder (`<backup folder>\TDE_Migration`), so it already has the rights it needs to write and read the files.

## After a run
- The export folder holds the **certificate AND its private key**. Move the files to a safe place and keep the password in your vault.
- If a DMK was created, back it up: `BACKUP MASTER KEY TO FILE = '...' ENCRYPTION BY PASSWORD = '...'`.
- **Availability groups:** every replica needs the certificate. Run the tool once for each replica.
- Log: `%LOCALAPPDATA%\TdeCertMigration\TdeCertMigration_<date>.log`

## Running the plan by hand
Save the script from tab 4 and open it in SSMS with **Query > SQLCMD Mode**. Replace the password placeholders and run it one step at a time. The "this PC" steps (copy, delete) are written as comments, so do those yourself.

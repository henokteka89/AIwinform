#Requires -Version 5.1
<#
.SYNOPSIS
    TDE Certificate Migration - WinForms GUI that moves TDE certificates (with their private keys)
    from a source SQL Server to a destination SQL Server.

.DESCRIPTION
    WHAT-IF BY DEFAULT
      What-If builds the full plan and shows, for every step, the exact T-SQL or file operation, which
      server it runs on, and what it does. To build the plan it only runs SELECTs against both servers.
      Nothing is changed.

    EXECUTE (only when "Execute" is selected AND you type the destination server name to confirm)
      Source      : EXEC xp_create_subdir (export folder), BACKUP CERTIFICATE ... WITH PRIVATE KEY
                    (ENCRYPTION BY PASSWORD) to NEW, time-stamped files.
      This PC     : creates a hidden, temporary SMB share on the export folder (source host) and on the import
                    folder (destination host), copies the .cer/.pvk files, verifies SHA-256, never overwrites, and
                    removes the shares again when done (also after a failure). Folders and files are kept.
      Destination : EXEC xp_create_subdir (import folder), CREATE MASTER KEY (only if master has none AND you
                    ticked it), CREATE CERTIFICATE ... FROM FILE ... WITH PRIVATE KEY (DECRYPTION BY PASSWORD).

    SAFETY RULES (production)
      - No DROP, no ALTER, no overwrite of any certificate, key, file or database.
      - No database is touched: TDE is NOT enabled on any database (no ALTER DATABASE / CREATE DATABASE ENCRYPTION KEY).
      - A certificate that is already on the destination (same thumbprint) is skipped.
        A different certificate with the same name stops that certificate (unless you choose the suffix option).
      - Every SQL step starts with a guard: it stops if it is connected to a different server than planned.
      - Order: all read-only checks -> source export -> copy and hash check -> only then any change on the destination.
        A failure stops the run; the remaining steps are not run.
      - Passwords are never written to the screen, the log or the saved script.

.NOTES
    Windows PowerShell 5.1, no modules needed (System.Data.SqlClient + WinForms).
    Needs sysadmin (or CONTROL SERVER) on both instances, and this PC must be able to reach both folders
    (admin share \\host\X$\... by default, or any share you type in).
    Run: powershell.exe -ExecutionPolicy Bypass -STA -File .\TdeCertMigration.ps1
#>
[CmdletBinding()] param()

$ErrorActionPreference = 'Stop'
$script:AppName = 'TDE Certificate Migration'
$script:Version = '1.0.0'
$script:TokCert = '<<#CERT_PASSWORD#>>'
$script:TokDmk  = '<<#DMK_PASSWORD#>>'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:State = [pscustomobject]@{ Src = $null; Dst = $null; SrcConn = $null; DstConn = $null; SrcCerts = $null; DstCerts = $null }
$script:Plan = $null
$script:Running = $false
$script:SqlMessages = New-Object System.Collections.Generic.List[string]
$script:CreatedShares = @{}   # shares this run created: key HOST|NAME -> spec. Only these are ever removed.
$script:CimSessions = @{}

$script:LogDir = Join-Path $env:LOCALAPPDATA 'TdeCertMigration'
try { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null } catch { $script:LogDir = $env:TEMP }
$script:LogFile = Join-Path $script:LogDir ("TdeCertMigration_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))

#region ============================ general helpers ============================

function DbVal  { param($v) if ($null -eq $v -or $v -is [System.DBNull]) { return $null }; return $v }
function DbVal0 { param($v) if ($null -eq $v -or $v -is [System.DBNull]) { return 0 };     return $v }
function QN { param([string]$s) return '[' + $s.Replace(']', ']]') + ']' }
function QS { param([string]$s) if ($null -eq $s) { $s = '' }; return "N'" + $s.Replace("'", "''") + "'" }
function ConvertTo-SafeName { param([string]$s) return ($s -replace '[\\/:*?"<>|\s\[\]'']', '_') }

function Join-WinPath {
    param([string]$Dir, [string]$Leaf)
    return $Dir.TrimEnd('\') + '\' + $Leaf
}

# C:\X on HOST -> \\HOST\C$\X (or C:\X when HOST is this PC). UNC paths are returned unchanged.
function Get-UncPath {
    param([string]$Path, [string]$HostName)
    if (-not $Path) { return '' }
    $Path = $Path.Trim().TrimEnd('\')
    if ($Path -match '^\\\\') { return $Path }
    if ($Path -match '^([A-Za-z]):(\\.*)?$') {
        if (-not $HostName -or $HostName -ieq $env:COMPUTERNAME) { return $Path }
        $rest = ''
        if ($Matches[2]) { $rest = $Matches[2] }
        return "\\$HostName\$($Matches[1].ToUpper())`$$rest"
    }
    return $Path
}

function Get-SqlErrorText {
    param($ErrorRecord)
    $e = $ErrorRecord.Exception
    while ($e) {
        if ($e -is [System.Data.SqlClient.SqlException]) {
            return (($e.Errors | ForEach-Object { "Msg $($_.Number): $($_.Message)" }) -join ' | ')
        }
        $e = $e.InnerException
    }
    return $ErrorRecord.Exception.Message
}

function Get-Me { return "$env:USERDOMAIN\$env:USERNAME" }

#endregion

#region ============================ SQL access ============================

function New-ConnString {
    param([hashtable]$Conn)
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source']      = $Conn.Server
    $b['Initial Catalog']  = 'master'
    $b['Application Name'] = 'TdeCertMigration'
    $b['Connect Timeout']  = 15
    if ($Conn.Encrypt) { $b['Encrypt'] = $true; $b['TrustServerCertificate'] = $true }
    if ($Conn.Auth -eq 'SQL login') { $b['User ID'] = $Conn.User; $b['Password'] = $Conn.Pass }
    else { $b['Integrated Security'] = $true }
    return $b.ConnectionString
}

function Invoke-Sql {
    param([hashtable]$Conn, [string]$Query, [int]$Timeout = 120)
    $script:SqlMessages.Clear()
    $cn = New-Object System.Data.SqlClient.SqlConnection (New-ConnString $Conn)
    $cn.add_InfoMessage([System.Data.SqlClient.SqlInfoMessageEventHandler] {
        param($sender, $e)
        foreach ($er in $e.Errors) { $script:SqlMessages.Add([string]$er.Message) }
    })
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $ds = New-Object System.Data.DataSet
        [void]$da.Fill($ds)
        if ($ds.Tables.Count -gt 0) { return , $ds.Tables[0] }
        return , (New-Object System.Data.DataTable)
    }
    finally { $cn.Dispose() }
}

# Read-only
function Get-ServerInfo {
    param([hashtable]$Conn)
    $r = (Invoke-Sql $Conn @"
SELECT COALESCE(@@SERVERNAME, CONVERT(nvarchar(128), SERVERPROPERTY('ServerName'))) AS ServerName,
       CONVERT(nvarchar(128), SERVERPROPERTY('MachineName'))               AS MachineName,
       CONVERT(nvarchar(128), SERVERPROPERTY('ProductVersion'))            AS ProductVersion,
       CONVERT(nvarchar(128), SERVERPROPERTY('ProductLevel'))              AS ProductLevel,
       CONVERT(nvarchar(128), SERVERPROPERTY('Edition'))                   AS Edition,
       CONVERT(nvarchar(512), SERVERPROPERTY('InstanceDefaultBackupPath')) AS BackupPath,
       ISNULL(IS_SRVROLEMEMBER('sysadmin'), 0)                             AS IsSysadmin,
       (SELECT COUNT(*) FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##') AS HasDmk,
       (SELECT CONVERT(int, is_master_key_encrypted_by_server) FROM sys.databases WHERE database_id = 1) AS DmkBySmk,
       (SELECT COUNT(*) FROM sys.fn_my_permissions(NULL, 'SERVER') WHERE permission_name = N'CONTROL SERVER') AS HasControlServer;
"@).Rows[0]
    $ver = [string]$r.ProductVersion
    $o = [pscustomobject]@{
        ServerName     = [string]$r.ServerName
        HostName       = [string]$r.MachineName
        Version        = $ver
        Major          = [int]($ver.Split('.')[0])
        Level          = [string]$r.ProductLevel
        Edition        = [string]$r.Edition
        BackupPath     = [string](DbVal $r.BackupPath)
        IsSysadmin     = ([int](DbVal0 $r.IsSysadmin) -eq 1)
        HasControl     = ([int](DbVal0 $r.HasControlServer) -gt 0)
        HasDmk         = ([int](DbVal0 $r.HasDmk) -gt 0)
        DmkBySmk       = ([int](DbVal0 $r.DmkBySmk) -eq 1)
        ServiceAccount = ''
    }
    if (-not $o.BackupPath) {
        try {
            $t = Invoke-Sql $Conn "DECLARE @p nvarchar(512); EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'BackupDirectory', @p OUTPUT; SELECT @p AS p;"
            $o.BackupPath = [string](DbVal $t.Rows[0].p)
        } catch { }
    }
    try {
        $s = Invoke-Sql $Conn "SELECT TOP (1) service_account FROM sys.dm_server_services WHERE servicename LIKE N'SQL Server (%';"
        if ($s.Rows.Count) { $o.ServiceAccount = [string]$s.Rows[0].service_account }
    } catch { }
    return $o
}

# Read-only
function Get-Certificates {
    param([hashtable]$Conn)
    $t = Invoke-Sql $Conn @"
SELECT CAST(0 AS bit) AS Migrate,
       c.name AS Name,
       CONVERT(varchar(130), c.thumbprint, 1) AS Thumbprint,
       c.subject AS Subject,
       CONVERT(varchar(19), c.start_date, 120)  AS StartDate,
       CONVERT(varchar(19), c.expiry_date, 120) AS ExpiryDate,
       c.pvt_key_encryption_type_desc AS PrivateKey,
       ISNULL(STUFF((SELECT N', ' + DB_NAME(k.database_id)
                     FROM sys.dm_database_encryption_keys AS k
                     WHERE k.encryptor_thumbprint = c.thumbprint AND k.database_id <> 2
                     FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, N''), N'') AS UsedByTde
FROM master.sys.certificates AS c
WHERE c.name NOT LIKE N'##%'
ORDER BY CASE WHEN EXISTS (SELECT 1 FROM sys.dm_database_encryption_keys AS k
                           WHERE k.encryptor_thumbprint = c.thumbprint AND k.database_id <> 2) THEN 0 ELSE 1 END,
         c.name;
"@
    foreach ($col in $t.Columns) { $col.ReadOnly = $false }
    return , $t
}

#endregion

#region ============================ plan ============================

function Add-Step {
    param($Plan, [string]$Where, [string]$Title, [string]$What, [string]$Kind = 'Sql', [string]$Sql = '',
          [string]$Status = 'Planned', [hashtable]$Data = @{}, [switch]$Changes, [switch]$AlwaysRun)
    switch ($Where) {
        'SOURCE'      { $server = $Plan.Src.ServerName }
        'DESTINATION' { $server = $Plan.Dst.ServerName }
        default       { $server = $env:COMPUTERNAME }
    }
    $s = [pscustomobject]@{
        No = $Plan.Steps.Count + 1; Where = $Where; Server = $server; Title = $Title; What = $What
        Kind = $Kind; Sql = $Sql; Status = $Status; Data = $Data; Changes = [bool]$Changes; AlwaysRun = [bool]$AlwaysRun; Result = ''
    }
    [void]$Plan.Steps.Add($s)
}

# Every SQL step starts with this: stops the batch when connected to the wrong server
function Get-Guard {
    param([string]$ServerName, [string]$Role)
    $msg = "STOP: this step is for the $Role server $ServerName, but the connection is to a different server. Nothing was changed."
    return "IF @@SERVERNAME <> $(QS $ServerName)`r`nBEGIN RAISERROR($(QS $msg.Replace('%', '%%')), 16, 1); RETURN; END;`r`n"
}

function Get-StepCode {
    param($Step, [switch]$Mask)
    switch ($Step.Kind) {
        'Sql' {
            if ($Mask) { return $Step.Sql.Replace($script:TokCert, "N'<<CERT_PASSWORD>>'").Replace($script:TokDmk, "N'<<DMK_PASSWORD>>'") }
            return $Step.Sql.Replace($script:TokCert, (QS $txtCertPwd.Text)).Replace($script:TokDmk, (QS $txtDmkPwd.Text))
        }
        'Copy' {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("# PowerShell, run by the tool on this PC ($env:COMPUTERNAME) as $(Get-Me)")
            [void]$sb.AppendLine("# Never overwrites: stops if a file with the same name is already in the destination folder.")
            foreach ($f in $Step.Data.Files) {
                $s = Join-WinPath $Step.Data.SrcDir $f; $d = Join-WinPath $Step.Data.DstDir $f
                [void]$sb.AppendLine("Copy-Item -LiteralPath '$s' -Destination '$d'")
                [void]$sb.AppendLine("(Get-FileHash -LiteralPath '$s' -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath '$d' -Algorithm SHA256).Hash   # must be True")
            }
            return $sb.ToString().TrimEnd()
        }
        'Delete' {
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine("# PowerShell, run by the tool on this PC ($env:COMPUTERNAME) as $(Get-Me)")
            [void]$sb.AppendLine("# Deletes ONLY the files this run copied, and only after the certificates were verified on the destination.")
            foreach ($f in $Step.Data.Files) { [void]$sb.AppendLine("Remove-Item -LiteralPath '$(Join-WinPath $Step.Data.Dir $f)'") }
            return $sb.ToString().TrimEnd()
        }
        'ShareCreate' {
            $d = $Step.Data
            $acc = '-ChangeAccess'
            if ($d.Access -eq 'Read') { $acc = '-ReadAccess' }
            $lines = @(
                "# PowerShell, run by the tool on this PC ($env:COMPUTERNAME) as $(Get-Me), against $($d.Host) (WinRM, DCOM if WinRM fails)"
                '# Adds a NEW hidden, temporary share on the existing folder (gone after a reboot anyway). Stops if the name is taken.'
                ('$cim = New-CimSession -ComputerName ''{0}''' -f $d.Host)
                ('New-SmbShare -CimSession $cim -Name ''{0}'' -Path ''{1}'' {2} ''{3}'' -Temporary -CachingMode None' -f $d.Name, $d.Path, $acc, $d.Account)
                ('# this PC then uses \\{0}\{1}' -f $d.Host, $d.Name)
            )
            return ($lines -join "`r`n")
        }
        'ShareRemove' {
            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add("# PowerShell, run by the tool on this PC ($env:COMPUTERNAME) as $(Get-Me). Runs even if an earlier step failed.")
            $lines.Add('# Removes ONLY the share this run created (and only if it still points at the same folder). The folder and files are kept.')
            foreach ($d in $Step.Data.Specs) {
                $lines.Add(('Remove-SmbShare -CimSession (New-CimSession -ComputerName ''{0}'') -Name ''{1}'' -Force   # folder {2} is kept' -f $d.Host, $d.Name, $d.Path))
            }
            return ($lines -join "`r`n")
        }
        default { return '-- (no action: information only)' }
    }
}

# Temporary share for a folder on a SQL Server host. Returns $null when no share is needed
# (the path is already UNC, or the host is this PC).
function New-ShareSpec {
    param([string]$HostName, [string]$Path, [string]$Name, [string]$Access)
    if ($Path -match '^\\') { return $null }
    if ($Path -notmatch '^[A-Za-z]:\') { throw "Folder '$Path' must be a local drive path (D:\...) or a UNC path." }
    if ($HostName -ieq $env:COMPUTERNAME) { return $null }
    return @{ Host = $HostName; Path = $Path; Name = $Name; Access = $Access; Account = (Get-Me) }
}

function New-Plan {
    $src = $script:State.Src; $dst = $script:State.Dst
    if (-not $src) { throw 'Connect to the source first (Source tab).' }
    if (-not $dst) { throw 'Connect to the destination first (Destination tab).' }

    # fresh read-only look at both servers, so the plan matches what is there now
    $src = Get-ServerInfo $script:State.SrcConn; $script:State.Src = $src
    $dst = Get-ServerInfo $script:State.DstConn; $script:State.Dst = $dst
    $script:State.DstCerts = Get-Certificates $script:State.DstConn
    $script:State.DstCerts.Columns.Remove('Migrate')
    $gridDst.DataSource = $script:State.DstCerts
    Set-DestInfo

    if ($src.ServerName -ieq $dst.ServerName) { throw "Source and destination are the same server ($($src.ServerName))." }
    if (-not $script:State.SrcCerts) { throw 'No certificates loaded from the source.' }
    [void]$gridSrc.EndEdit()
    $sel = @($script:State.SrcCerts.Rows | Where-Object { $_.Migrate -eq $true })
    if (-not $sel.Count) { throw 'Tick at least one certificate on the Source tab.' }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reach = $cboReach.SelectedIndex   # 0 temporary shares, 1 admin shares, 2 typed paths
    $srcSql = $txtSrcFolder.Text.Trim().TrimEnd('\'); $dstSql = $txtDstFolder.Text.Trim().TrimEnd('\')
    if (-not $srcSql -or -not $dstSql) { throw 'Fill in the export and import folders on the Transfer tab.' }
    $srcShare = $null; $dstShare = $null
    if ($reach -eq 0) {
        $srcShare = New-ShareSpec $src.HostName $srcSql "TDEMIG_SRC_$stamp`$" 'Read'
        $dstShare = New-ShareSpec $dst.HostName $dstSql "TDEMIG_DST_$stamp`$" 'Change'
        $srcUnc = Get-UncPath $srcSql $src.HostName; $dstUnc = Get-UncPath $dstSql $dst.HostName
        if ($srcShare) { $srcUnc = "\\$($srcShare.Host)\$($srcShare.Name)" }
        if ($dstShare) { $dstUnc = "\\$($dstShare.Host)\$($dstShare.Name)" }
    } else {
        $srcUnc = $txtSrcUnc.Text.Trim().TrimEnd('\'); $dstUnc = $txtDstUnc.Text.Trim().TrimEnd('\')
        if (-not $srcUnc -or -not $dstUnc) { throw "Fill in both 'as THIS PC reaches it' paths on the Transfer tab." }
    }
    $plan = [pscustomobject]@{
        Steps = New-Object System.Collections.ArrayList; Warnings = New-Object System.Collections.Generic.List[string]
        Certs = @(); Stamp = $stamp; Src = $src; Dst = $dst; Created = Get-Date
        SrcSql = $srcSql; SrcUnc = ''; DstSql = $dstSql; DstUnc = ''; Reach = $reach
        NeedsCertPwd = $false; NeedsDmkPwd = $false; Blocked = $false
    }
    $suffix = ($cboConflict.SelectedIndex -eq 1)

    $certs = foreach ($r in $sel) {
        $name = [string]$r.Name; $thumb = [string]$r.Thumbprint
        if ($thumb -notmatch '^0x[0-9A-Fa-f]+$') { throw "Unexpected thumbprint format for [$name]: $thumb" }
        $c = [pscustomobject]@{ Name = $name; Thumb = $thumb; DestName = $name; File = (ConvertTo-SafeName "$($name)_$stamp"); State = 'Move'; Note = ''; PrivateKey = [string]$r.PrivateKey }
        $same = @($script:State.DstCerts.Rows | Where-Object { [string]$_.Thumbprint -eq $thumb })
        $byName = @($script:State.DstCerts.Rows | Where-Object { [string]$_.Name -eq $name -and [string]$_.Thumbprint -ne $thumb })
        if ($same.Count) {
            $c.State = 'Exists'; $c.Note = "already on the destination as [$($same[0].Name)] (same thumbprint)"
        } elseif ($c.PrivateKey -eq 'NO_PRIVATE_KEY') {
            $c.State = 'Blocked'; $c.Note = 'has no private key on the source, so it cannot be used for TDE on the destination'
        } elseif ($c.PrivateKey -eq 'ENCRYPTED_BY_PASSWORD') {
            $c.State = 'Blocked'; $c.Note = 'private key is protected by a password, not by the master key; this tool only exports master-key-protected keys'
        } elseif ($byName.Count) {
            if ($suffix) {
                $c.DestName = "$($name)_migrated_$($stamp.Substring(0, 8))"
                if (@($script:State.DstCerts.Rows | Where-Object { [string]$_.Name -eq $c.DestName }).Count) {
                    $c.State = 'Blocked'; $c.Note = "a different certificate named [$name] and one named [$($c.DestName)] already exist on the destination"
                } else { $c.Note = "a different certificate named [$name] exists on the destination; this one will be created as [$($c.DestName)]" }
            } else {
                $c.State = 'Blocked'; $c.Note = "a DIFFERENT certificate named [$name] already exists on the destination (nothing will be replaced). Choose the suffix option on the Destination tab to create it under another name"
            }
        }
        $c
    }
    $plan.Certs = @($certs)
    $plan.SrcUnc = $srcUnc; $plan.DstUnc = $dstUnc
    $move = @($plan.Certs | Where-Object { $_.State -eq 'Move' })

    # warnings (do not block)
    if ($dst.Major -lt $src.Major) { $plan.Warnings.Add("Destination is SQL Server major version $($dst.Major), source is $($src.Major). Certificates backed up on a newer version may not restore on an older one: test first.") }
    foreach ($x in @(@{ S = $src; R = 'source' }, @{ S = $dst; R = 'destination' })) {
        if (-not $x.S.IsSysadmin -and -not $x.S.HasControl) { $plan.Warnings.Add("Your login is not sysadmin / CONTROL SERVER on the $($x.R) $($x.S.ServerName); BACKUP / CREATE CERTIFICATE and xp_create_subdir will probably fail.") }
        elseif (-not $x.S.IsSysadmin) { $plan.Warnings.Add("Your login is not sysadmin on the $($x.R) $($x.S.ServerName); xp_create_subdir needs sysadmin. Create the folder by hand if that step fails.") }
    }
    foreach ($c in $plan.Certs) {
        $row = @($script:State.SrcCerts.Rows | Where-Object { [string]$_.Thumbprint -eq $c.Thumb })[0]
        if ($row -and [string]$row.ExpiryDate -and ([datetime][string]$row.ExpiryDate) -lt (Get-Date)) {
            $plan.Warnings.Add("Certificate [$($c.Name)] expired on $($row.ExpiryDate). TDE still works with an expired certificate, but plan a rotation.")
        }
    }
    $sameFolder = ($srcSql -ieq $dstSql -and ($src.HostName -ieq $dst.HostName -or $srcSql -match '^\\\\'))
    if ($reach -ne 0 -and $srcUnc -ieq $dstUnc) { $sameFolder = $true }
    if ($sameFolder) { $srcShare = $null; $dstShare = $null }
    if ($move.Count -and -not $sameFolder -and $reach -eq 0) {
        foreach ($sh in @($srcShare, $dstShare) | Where-Object { $_ }) {
            try { Test-WSMan -ComputerName $sh.Host -ErrorAction Stop | Out-Null }
            catch { $plan.Warnings.Add("WinRM to $($sh.Host) did not answer; the share step will try DCOM. If both fail, choose 'Admin share' or 'Paths I type' on the Transfer tab.") }
        }
    }
    if ($move.Count -and -not $sameFolder -and $reach -ne 0) {
        foreach ($u in @($srcUnc, $dstUnc)) {
            if ($u -match '^(\\\\[^\\]+\\[^\\]+)') {
                $root = $Matches[1]
                if (-not (Test-Path -LiteralPath $root)) { $plan.Warnings.Add("This PC cannot open $root. The copy step will fail: check the admin share / firewall, or type a share you can reach on the Transfer tab.") }
            }
        }
    }

    # ---------- steps ----------
    $sg = Get-Guard $src.ServerName 'SOURCE'
    $dg = Get-Guard $dst.ServerName 'DESTINATION'

    # 1. read-only check on the source
    if ($move.Count) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($sg)
        foreach ($c in $move) {
            $n = QS $c.Name
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM master.sys.certificates WHERE name = $n AND thumbprint = $($c.Thumb))")
            [void]$sb.AppendLine("    BEGIN RAISERROR(N'Certificate %s with thumbprint $($c.Thumb) is not in master on the source.', 16, 1, $n); RETURN; END;")
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM master.sys.certificates WHERE name = $n AND pvt_key_encryption_type = 'MK')")
            [void]$sb.AppendLine("    BEGIN RAISERROR(N'The private key of certificate %s is not protected by the master key; it cannot be exported here.', 16, 1, $n); RETURN; END;")
        }
        [void]$sb.AppendLine("PRINT N'Source check OK: $($move.Count) certificate(s) found with a private key.';")
        Add-Step $plan 'SOURCE' 'Check the certificates and private keys (read-only)' `
            "Read-only. Confirms that each selected certificate is still in master on $($src.ServerName) with the same thumbprint and a private key protected by the database master key. Stops the run if not." `
            -Sql $sb.ToString()
    }

    # 2. read-only check on the destination: master key state
    $openDmk = ''; $closeDmk = ''; $createDmk = $false
    if ($move.Count) {
        if ($dst.HasDmk -and $dst.DmkBySmk) {
            Add-Step $plan 'DESTINATION' 'Check the database master key (read-only)' `
                "Read-only. master on $($dst.ServerName) already has a database master key, encrypted by the service master key, so SQL Server opens it automatically for CREATE CERTIFICATE. Nothing is created." `
                -Sql ($dg + "IF NOT EXISTS (SELECT 1 FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##')`r`n    BEGIN RAISERROR(N'master has no database master key on the destination any more. Build the plan again.', 16, 1); RETURN; END;`r`nPRINT N'Destination database master key: present.';")
        } elseif ($dst.HasDmk) {
            $plan.NeedsDmkPwd = $true
            $openDmk = "OPEN MASTER KEY DECRYPTION BY PASSWORD = $script:TokDmk;`r`n"; $closeDmk = "CLOSE MASTER KEY;`r`n"
            Add-Step $plan 'DESTINATION' 'Check that the DMK password opens the master key' `
                "master on $($dst.ServerName) has a database master key that is NOT encrypted by the service master key, so it must be opened with its password before CREATE CERTIFICATE. This step only opens and closes it (no change). Enter the DMK password on the Destination tab." `
                -Sql ($dg + $openDmk + $closeDmk + "PRINT N'Destination database master key: password OK.';")
        } elseif ($chkCreateDmk.Checked) {
            $plan.NeedsDmkPwd = $true; $createDmk = $true
            Add-Step $plan 'DESTINATION' 'Confirm master has no master key yet (read-only)' `
                "Read-only. master on $($dst.ServerName) has no database master key. It will be created later (after the export and the copy succeed)." `
                -Sql ($dg + "IF EXISTS (SELECT 1 FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##')`r`n    PRINT N'A database master key now exists; the create step will skip it.';`r`nELSE PRINT N'No database master key yet: it will be created.';")
        } else {
            Add-Step $plan 'DESTINATION' 'BLOCKED: master has no database master key' `
                "master on $($dst.ServerName) has no database master key, and CREATE CERTIFICATE with a private key needs one. Tick 'Create the database master key' on the Destination tab and enter a password, or create it yourself." `
                -Kind 'Info' -Status 'Blocked'
            $plan.Blocked = $true
        }
    }

    # information rows for certificates that will not be moved
    foreach ($c in @($plan.Certs | Where-Object { $_.State -eq 'Exists' })) {
        Add-Step $plan 'DESTINATION' "Skip [$($c.Name)]: already on the destination" "Certificate [$($c.Name)] $($c.Note). Nothing to do for it." -Kind 'Info' -Status 'Skip'
    }
    foreach ($c in @($plan.Certs | Where-Object { $_.State -eq 'Blocked' })) {
        Add-Step $plan 'SOURCE' "BLOCKED: [$($c.Name)]" "Certificate [$($c.Name)] $($c.Note)." -Kind 'Info' -Status 'Blocked'
        $plan.Blocked = $true
    }

    if ($move.Count) {
        $plan.NeedsCertPwd = $true
        $files = New-Object System.Collections.Generic.List[string]
        foreach ($c in $move) { $files.Add("$($c.File).cer"); $files.Add("$($c.File).pvk") }

        # 3. export folder on the source
        Add-Step $plan 'SOURCE' 'Create the export folder (if missing)' `
            "Creates $srcSql on $($src.HostName) if it does not exist (xp_create_subdir does nothing when it already exists). The folder is created by the SQL Server service account, so it can write the files there. Nothing else is touched." `
            -Sql ($sg + "EXEC master.dbo.xp_create_subdir $(QS $srcSql);`r`nPRINT N'Export folder ready: $($srcSql.Replace("'", "''"))';") -Changes

        # 4. backup on the source
        foreach ($c in $move) {
            $cer = Join-WinPath $srcSql "$($c.File).cer"; $pvk = Join-WinPath $srcSql "$($c.File).pvk"
            Add-Step $plan 'SOURCE' "Back up certificate [$($c.Name)] with its private key" `
                "Writes two NEW files on $($src.HostName): $cer (public certificate) and $pvk (private key, encrypted with your certificate password). The certificate itself is not changed. The file names carry a time stamp, so no existing file is overwritten (BACKUP CERTIFICATE fails rather than overwrite)." `
                -Sql ($sg + "BACKUP CERTIFICATE $(QN $c.Name)`r`n    TO FILE = $(QS $cer)`r`n    WITH PRIVATE KEY (FILE = $(QS $pvk),`r`n                      ENCRYPTION BY PASSWORD = $script:TokCert);`r`nPRINT N'Backed up certificate $($c.Name.Replace("'", "''")).';") -Changes
        }

        # 5. import folder on the destination + copy
        if (-not $sameFolder) {
            Add-Step $plan 'DESTINATION' 'Create the import folder (if missing)' `
                "Creates $dstSql on $($dst.HostName) if it does not exist. The folder is created by the destination SQL Server service account, so it can read the files. Nothing else is touched." `
                -Sql ($dg + "EXEC master.dbo.xp_create_subdir $(QS $dstSql);`r`nPRINT N'Import folder ready: $($dstSql.Replace("'", "''"))';") -Changes
            if ($srcShare) {
                Add-Step $plan 'SOURCE' "Create temporary share $($srcShare.Name) on $($srcShare.Host)" `
                    "Shares the existing export folder $srcSql on $($srcShare.Host) as \\$($srcShare.Host)\$($srcShare.Name): hidden (`$), temporary (also gone after a reboot), READ access for $($srcShare.Account) only. Run from this PC over WinRM (DCOM if WinRM fails). The folder and other shares are not touched. Removed again right after the copy." `
                    -Kind 'ShareCreate' -Data $srcShare -Changes
            }
            if ($dstShare) {
                Add-Step $plan 'DESTINATION' "Create temporary share $($dstShare.Name) on $($dstShare.Host)" `
                    "Shares the existing import folder $dstSql on $($dstShare.Host) as \\$($dstShare.Host)\$($dstShare.Name): hidden (`$), temporary (also gone after a reboot), CHANGE access for $($dstShare.Account) only. The folder and other shares are not touched. Removed again when it is no longer needed." `
                    -Kind 'ShareCreate' -Data $dstShare -Changes
            }
            Add-Step $plan 'THIS PC' "Copy $($files.Count) file(s) to the destination and verify SHA-256" `
                "This PC copies the .cer and .pvk files from $srcUnc to $dstUnc, then compares the SHA-256 hash of every copy with the original. It never overwrites: it stops if a file with the same name is already there." `
                -Kind 'Copy' -Data @{ SrcDir = $srcUnc; DstDir = $dstUnc; Files = @($files) } -Changes
            $keepDst = ($dstShare -and $chkDeleteDest.Checked)   # the delete step needs the destination share
            $now = @(@($srcShare, $dstShare) | Where-Object { $_ -and -not ($keepDst -and $_ -eq $dstShare) })
            if ($now.Count) {
                Add-Step $plan 'THIS PC' "Remove the temporary share(s): $(($now | ForEach-Object { $_.Name }) -join ', ')" `
                    "Removes only the share(s) this run created, and only if they still point at the same folder. The folders and the certificate files stay where they are. This step runs even if an earlier step failed." `
                    -Kind 'ShareRemove' -Data @{ Specs = $now } -Changes -AlwaysRun
            }
        } else {
            Add-Step $plan 'THIS PC' 'No copy needed' "The source export folder and the destination import folder are the same folder ($srcUnc), so nothing is copied." -Kind 'Info' -Status 'Skip'
        }

        # 6. master key (only now, after export and copy succeeded)
        if ($createDmk) {
            Add-Step $plan 'DESTINATION' 'Create the database master key in master' `
                "Creates the database master key in master on $($dst.ServerName) (only if there still is none). It protects the certificate private key; SQL Server also encrypts it by the service master key so it opens automatically. BACK IT UP afterwards (BACKUP MASTER KEY) and keep the password in your vault." `
                -Sql ($dg + "IF NOT EXISTS (SELECT 1 FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##')`r`nBEGIN`r`n    CREATE MASTER KEY ENCRYPTION BY PASSWORD = $script:TokDmk;`r`n    PRINT N'Database master key created.';`r`nEND`r`nELSE PRINT N'A database master key already exists; not created.';") -Changes
        }

        # 7. create certificates on the destination
        foreach ($c in $move) {
            $cer = Join-WinPath $dstSql "$($c.File).cer"; $pvk = Join-WinPath $dstSql "$($c.File).pvk"
            $msgExists = "A certificate with thumbprint $($c.Thumb) already exists on the destination; not created."
            $msgName = "A different certificate named %s already exists on the destination; nothing was created or replaced."
            $q = $dg +
                "IF EXISTS (SELECT 1 FROM master.sys.certificates WHERE thumbprint = $($c.Thumb))`r`n    BEGIN PRINT $(QS $msgExists); RETURN; END;`r`n" +
                "IF EXISTS (SELECT 1 FROM master.sys.certificates WHERE name = $(QS $c.DestName))`r`n    BEGIN RAISERROR($(QS $msgName), 16, 1, $(QS $c.DestName)); RETURN; END;`r`n" +
                $openDmk +
                "CREATE CERTIFICATE $(QN $c.DestName)`r`n    FROM FILE = $(QS $cer)`r`n    WITH PRIVATE KEY (FILE = $(QS $pvk),`r`n                      DECRYPTION BY PASSWORD = $script:TokCert);`r`n" +
                $closeDmk +
                "PRINT N'Created certificate $($c.DestName.Replace("'", "''")).';"
            $what = "Creates certificate [$($c.DestName)] in master on $($dst.ServerName) from the copied files, decrypting the private key with your certificate password and protecting it with the destination master key. Skipped if the same certificate is already there; stops (changes nothing) if another certificate has that name."
            if ($c.Note) { $what += " Note: $($c.Note)." }
            Add-Step $plan 'DESTINATION' "Create certificate [$($c.DestName)] from the files" $what -Sql $q -Changes
        }

        # 8. verify (read-only)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($dg)
        foreach ($c in $move) {
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM master.sys.certificates WHERE thumbprint = $($c.Thumb) AND pvt_key_encryption_type = 'MK')")
            [void]$sb.AppendLine("    BEGIN RAISERROR(N'Certificate $($c.Thumb) is not on the destination with a master-key-protected private key.', 16, 1); RETURN; END;")
        }
        [void]$sb.AppendLine("SELECT name, CONVERT(varchar(130), thumbprint, 1) AS thumbprint, pvt_key_encryption_type_desc, expiry_date")
        [void]$sb.AppendLine("FROM master.sys.certificates WHERE thumbprint IN ($((@($move | ForEach-Object { $_.Thumb })) -join ', '));")
        [void]$sb.AppendLine("PRINT N'Destination check OK: thumbprints match the source and private keys are protected by the master key.';")
        Add-Step $plan 'DESTINATION' 'Verify the certificates (read-only)' `
            "Read-only. Confirms every migrated certificate is in master on $($dst.ServerName) with the SAME thumbprint as on the source and a private key protected by the master key. This is what TDE needs to restore or attach the encrypted databases." `
            -Sql $sb.ToString()

        # 9. optional clean-up of the copies
        if ($chkDeleteDest.Checked) {
            if ($sameFolder) {
                Add-Step $plan 'THIS PC' 'Delete copies: skipped' 'The import folder is the same as the export folder; the files are kept (they are the only backup).' -Kind 'Info' -Status 'Skip'
            } else {
                Add-Step $plan 'THIS PC' "Delete the $($files.Count) copied file(s) from the destination folder" `
                    "Deletes only the files this run copied to $dstUnc, and only after the verification step passed. The export on the source ($srcSql on $($src.HostName)) is kept: move it to a safe place." `
                    -Kind 'Delete' -Data @{ Dir = $dstUnc; Files = @($files) } -Changes
            }
        }
        if (-not $sameFolder -and $dstShare -and $chkDeleteDest.Checked) {
            Add-Step $plan 'THIS PC' "Remove the temporary share: $($dstShare.Name)" `
                "Removes only the share this run created on $($dstShare.Host), and only if it still points at $dstSql. The folder is kept. This step runs even if an earlier step failed." `
                -Kind 'ShareRemove' -Data @{ Specs = @($dstShare) } -Changes -AlwaysRun
        }
    } elseif (-not $plan.Blocked) {
        Add-Step $plan 'THIS PC' 'Nothing to do' 'All selected certificates are already on the destination.' -Kind 'Info' -Status 'Skip'
    }
    return $plan
}

function Get-ConnDesc {
    param([string]$Where)
    switch ($Where) {
        'SOURCE'      { $c = $script:State.SrcConn }
        'DESTINATION' { $c = $script:State.DstConn }
        default       { return "this PC ($env:COMPUTERNAME) as $(Get-Me)" }
    }
    if ($c.Auth -eq 'SQL login') { return "SQL login $($c.User)" }
    return "Windows login $(Get-Me)"
}

function Get-PlanScript {
    param($Plan)
    $sb = New-Object System.Text.StringBuilder
    $line = '-' * 100
    [void]$sb.AppendLine('/*' + ('=' * 98))
    [void]$sb.AppendLine("  $script:AppName $script:Version - plan generated $($Plan.Created.ToString('yyyy-MM-dd HH:mm:ss')) by $(Get-Me) on $env:COMPUTERNAME")
    [void]$sb.AppendLine("  SOURCE      : $($Plan.Src.ServerName)  (host $($Plan.Src.HostName), $($Plan.Src.Version) $($Plan.Src.Level), $($Plan.Src.Edition))")
    [void]$sb.AppendLine("  DESTINATION : $($Plan.Dst.ServerName)  (host $($Plan.Dst.HostName), $($Plan.Dst.Version) $($Plan.Dst.Level), $($Plan.Dst.Edition))")
    [void]$sb.AppendLine("  Export folder (source SQL Server)      : $($Plan.SrcSql)     from this PC: $($Plan.SrcUnc)")
    [void]$sb.AppendLine("  Import folder (destination SQL Server) : $($Plan.DstSql)     from this PC: $($Plan.DstUnc)")
    $how = @('temporary hidden shares, created by the run and removed when done (folders kept)', 'admin shares (\\host\X$)', 'paths typed by you')[$Plan.Reach]
    [void]$sb.AppendLine("  This PC reaches the folders through: $how")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('  Passwords are shown as <<CERT_PASSWORD>> / <<DMK_PASSWORD>>; they are never saved.')
    [void]$sb.AppendLine('  To run it by hand: SSMS > Query > SQLCMD Mode, replace the placeholders, run one step at a time,')
    [void]$sb.AppendLine('  and do the THIS PC steps yourself (they are shown as comments).')
    [void]$sb.AppendLine('  What this plan never does: DROP / ALTER anything, overwrite files or certificates, or touch any database.')
    if ($Plan.Warnings.Count) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('  WARNINGS')
        foreach ($w in $Plan.Warnings) { [void]$sb.AppendLine("   - $w") }
    }
    if ($Plan.Blocked) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('  BLOCKED: fix the BLOCKED steps below before Execute is allowed.')
    }
    [void]$sb.AppendLine(('=' * 98) + '*/')
    $n = $Plan.Steps.Count
    foreach ($s in $Plan.Steps) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("-- $line")
        [void]$sb.AppendLine("-- STEP $($s.No) of $n  |  $($s.Where)  |  $($s.Server)  |  $($s.Status.ToUpper())")
        [void]$sb.AppendLine("-- $($s.Title)")
        [void]$sb.AppendLine("-- Runs as : $(Get-ConnDesc $s.Where)")
        [void]$sb.AppendLine("-- Changes : $(if ($s.Changes) { 'YES' } else { 'no (read-only / information)' })")
        foreach ($w in (Split-Words $s.What 96)) { [void]$sb.AppendLine("-- $w") }
        [void]$sb.AppendLine("-- $line")
        if ($s.Kind -eq 'Sql') {
            [void]$sb.AppendLine(":CONNECT $($s.Server)")
            [void]$sb.AppendLine((Get-StepCode $s -Mask).TrimEnd())
            [void]$sb.AppendLine('GO')
        } elseif ($s.Kind -in 'Copy', 'Delete', 'ShareCreate', 'ShareRemove') {
            foreach ($l in ((Get-StepCode $s -Mask) -split "`r?`n")) { [void]$sb.AppendLine("-- $l") }
        }
    }
    return $sb.ToString()
}

function Split-Words {
    param([string]$Text, [int]$Width)
    $out = New-Object System.Collections.Generic.List[string]
    $cur = ''
    foreach ($w in ($Text -split '\s+')) {
        if ($cur.Length -and ($cur.Length + 1 + $w.Length) -gt $Width) { $out.Add($cur); $cur = $w }
        elseif ($cur.Length) { $cur += " $w" } else { $cur = $w }
    }
    if ($cur) { $out.Add($cur) }
    return $out
}

#endregion

#region ============================ execution ============================

function Invoke-CopyStep {
    param($Step)
    $d = $Step.Data
    if (-not (Test-Path -LiteralPath $d.DstDir)) { throw "This PC cannot see $($d.DstDir). Check the path / admin share on the Transfer tab." }
    foreach ($f in $d.Files) {
        $s = Join-WinPath $d.SrcDir $f; $t = Join-WinPath $d.DstDir $f
        if (-not (Test-Path -LiteralPath $s)) { throw "File not found from this PC: $s. Check the 'from this PC' source path on the Transfer tab." }
        if (Test-Path -LiteralPath $t) { throw "$t already exists. The tool never overwrites; nothing was copied over it." }
        Copy-Item -LiteralPath $s -Destination $t
        $h1 = (Get-FileHash -LiteralPath $s -Algorithm SHA256).Hash
        $h2 = (Get-FileHash -LiteralPath $t -Algorithm SHA256).Hash
        if ($h1 -ne $h2) { throw "SHA-256 mismatch after copying $f. Do not use the copy; check the network/share." }
        Write-Log ("Copied {0} ({1:N0} bytes, SHA-256 {2}...)" -f $f, (Get-Item -LiteralPath $t).Length, $h1.Substring(0, 16)) 'OK'
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Invoke-DeleteStep {
    param($Step)
    foreach ($f in $Step.Data.Files) {
        $p = Join-WinPath $Step.Data.Dir $f
        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p; Write-Log "Deleted copy $p" 'OK' }
    }
}

function Get-HostCim {
    param([string]$HostName)
    if ($script:CimSessions.ContainsKey($HostName)) { return $script:CimSessions[$HostName] }
    try { $s = New-CimSession -ComputerName $HostName -OperationTimeoutSec 60 -ErrorAction Stop }
    catch {
        Write-Log "WinRM to $HostName failed ($($_.Exception.Message)); trying DCOM." 'WARN'
        $s = New-CimSession -ComputerName $HostName -SessionOption (New-CimSessionOption -Protocol Dcom) -OperationTimeoutSec 60 -ErrorAction Stop
    }
    $script:CimSessions[$HostName] = $s
    return $s
}

function Close-HostCims {
    foreach ($k in @($script:CimSessions.Keys)) { try { Remove-CimSession -CimSession $script:CimSessions[$k] } catch { } }
    $script:CimSessions = @{}
}

function Invoke-ShareCreateStep {
    param($Step)
    $d = $Step.Data
    $cim = Get-HostCim $d.Host
    if (Get-SmbShare -CimSession $cim -Name $d.Name -ErrorAction SilentlyContinue) { throw "A share named $($d.Name) already exists on $($d.Host). It was not touched." }
    $p = @{ CimSession = $cim; Name = $d.Name; Path = $d.Path; Temporary = $true; CachingMode = 'None'; ErrorAction = 'Stop'
            Description = "TdeCertMigration temporary share ($(Get-Me), $(Get-Date -Format 'yyyy-MM-dd HH:mm')). Safe to remove." }
    if ($d.Access -eq 'Read') { $p.ReadAccess = $d.Account } else { $p.ChangeAccess = $d.Account }
    New-SmbShare @p | Out-Null
    $script:CreatedShares["$($d.Host)|$($d.Name)"] = $d
    Write-Log "Created temporary share \\$($d.Host)\$($d.Name) -> $($d.Path) ($($d.Access) for $($d.Account))." 'OK'
    if (-not (Test-Path -LiteralPath "\\$($d.Host)\$($d.Name)")) { throw "The share was created but this PC cannot open \\$($d.Host)\$($d.Name) (NTFS rights on $($d.Path)?). It will be removed." }
}

function Remove-CreatedShare {
    param($d)
    $key = "$($d.Host)|$($d.Name)"
    if (-not $script:CreatedShares.ContainsKey($key)) { Write-Log "Share $($d.Name) on $($d.Host) was not created by this run; nothing to remove." 'INFO'; return }
    $cim = Get-HostCim $d.Host
    $sh = Get-SmbShare -CimSession $cim -Name $d.Name -ErrorAction SilentlyContinue
    if (-not $sh) { Write-Log "Share $($d.Name) on $($d.Host) is already gone." 'INFO' }
    elseif ([string]$sh.Path.TrimEnd('\') -ine $d.Path.TrimEnd('\')) { throw "Share $($d.Name) on $($d.Host) now points at $($sh.Path), not $($d.Path); it was NOT removed. Check it by hand." }
    else {
        Remove-SmbShare -CimSession $cim -Name $d.Name -Force -ErrorAction Stop
        Write-Log "Removed temporary share \\$($d.Host)\$($d.Name). The folder $($d.Path) and its files are kept." 'OK'
    }
    [void]$script:CreatedShares.Remove($key)
}

function Invoke-ShareRemoveStep {
    param($Step)
    $err = @()
    foreach ($d in $Step.Data.Specs) {
        try { Remove-CreatedShare $d } catch { $err += $_.Exception.Message; Write-Log $_.Exception.Message 'FAIL' }
    }
    if ($err.Count) { throw ($err -join ' | ') }
}

# Safety net: removes any share this run created that is still there (failure, closing the window)
function Remove-LeftoverShares {
    foreach ($k in @($script:CreatedShares.Keys)) {
        $d = $script:CreatedShares[$k]
        try { Remove-CreatedShare $d } catch { Write-Log "Could not remove share $($d.Name) on $($d.Host): $($_.Exception.Message). Remove it by hand: Remove-SmbShare -Name '$($d.Name)' on $($d.Host)." 'FAIL' }
    }
}

function Test-ExecuteInputs {
    param($Plan)
    if ($Plan.Blocked) { throw 'The plan has BLOCKED steps (red rows). Fix them and run What-If again.' }
    if (-not @($Plan.Steps | Where-Object { $_.Changes }).Count) { throw 'Nothing to execute: the plan makes no changes.' }
    if ($Plan.NeedsCertPwd) {
        if (-not $txtCertPwd.Text) { $tabs.SelectedTab = $tabXfer; throw 'Enter the private key (certificate) password on the Transfer tab.' }
        if ($txtCertPwd.Text -cne $txtCertPwd2.Text) { $tabs.SelectedTab = $tabXfer; throw 'The certificate password and its confirmation do not match.' }
        if ($txtCertPwd.Text.Length -lt 8) { $tabs.SelectedTab = $tabXfer; throw 'Use a certificate password of at least 8 characters (SQL Server also applies the Windows password policy).' }
    }
    if ($Plan.NeedsDmkPwd) {
        if (-not $txtDmkPwd.Text) { $tabs.SelectedTab = $tabDst; throw 'Enter the database master key password on the Destination tab.' }
        if (-not $script:State.Dst.HasDmk -and $txtDmkPwd.Text -cne $txtDmkPwd2.Text) { $tabs.SelectedTab = $tabDst; throw 'The DMK password and its confirmation do not match.' }
    }
}

function Confirm-Execute {
    param($Plan)
    $chg = @($Plan.Steps | Where-Object { $_.Changes -and $_.Status -eq 'Planned' })
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "$script:AppName - confirm Execute"; $f.Size = New-Object System.Drawing.Size(760, 470)
    $f.StartPosition = 'CenterParent'; $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $t = New-Object System.Windows.Forms.TextBox; $t.Multiline = $true; $t.ReadOnly = $true; $t.ScrollBars = 'Vertical'
    $t.Font = New-Object System.Drawing.Font('Consolas', 9); $t.Location = New-Object System.Drawing.Point(10, 10); $t.Size = New-Object System.Drawing.Size(725, 300)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("These steps make changes:`r`n")
    foreach ($s in $chg) { [void]$sb.AppendLine("  $($s.No). [$($s.Where)] $($s.Server): $($s.Title)") }
    [void]$sb.AppendLine("`r`nNothing is dropped, altered or overwritten, and no database is touched.")
    [void]$sb.AppendLine("The run stops at the first failure.")
    if ($Plan.Warnings.Count) { [void]$sb.AppendLine("`r`nWARNINGS:"); foreach ($w in $Plan.Warnings) { [void]$sb.AppendLine("  - $w") } }
    $t.Text = $sb.ToString()
    $l = New-Object System.Windows.Forms.Label; $l.Text = "Type the DESTINATION server name ($($Plan.Dst.ServerName)) to confirm:"
    $l.Location = New-Object System.Drawing.Point(10, 320); $l.Size = New-Object System.Drawing.Size(725, 20)
    $i = New-Object System.Windows.Forms.TextBox; $i.Location = New-Object System.Drawing.Point(10, 344); $i.Size = New-Object System.Drawing.Size(400, 23)
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = 'Execute'; $ok.Enabled = $false; $ok.DialogResult = 'OK'
    $ok.Location = New-Object System.Drawing.Point(530, 385); $ok.Size = New-Object System.Drawing.Size(100, 30)
    $cn = New-Object System.Windows.Forms.Button; $cn.Text = 'Cancel'; $cn.DialogResult = 'Cancel'
    $cn.Location = New-Object System.Drawing.Point(635, 385); $cn.Size = New-Object System.Drawing.Size(100, 30)
    $expected = $Plan.Dst.ServerName
    $i.Add_TextChanged({ $ok.Enabled = ($i.Text.Trim() -ieq $expected) }.GetNewClosure())
    $f.Controls.AddRange(@($t, $l, $i, $ok, $cn)); $f.AcceptButton = $ok; $f.CancelButton = $cn
    $r = $f.ShowDialog($form)
    $f.Dispose()
    return ($r -eq [System.Windows.Forms.DialogResult]::OK)
}

function Invoke-Plan {
    param($Plan)
    $script:Running = $true
    Set-Busy $true
    $todo = @($Plan.Steps | Where-Object { $_.Status -eq 'Planned' })
    $progress.Minimum = 0; $progress.Maximum = [Math]::Max(1, $todo.Count); $progress.Value = 0
    $tabsBottom.SelectedTab = $tabLog
    Write-Log "===== EXECUTE started: $($Plan.Src.ServerName) -> $($Plan.Dst.ServerName), $($todo.Count) step(s) =====" 'INFO'
    $failed = $null; $done = 0
    try {
        foreach ($s in $Plan.Steps) {
            if ($failed -and -not ($s.AlwaysRun -and $s.Status -eq 'Planned')) { if ($s.Status -eq 'Planned') { Set-StepStatus $s 'Not run' } ; continue }
            if ($s.Status -eq 'Skip') { Set-StepStatus $s 'Skipped'; continue }
            if ($s.Status -ne 'Planned') { continue }
            Set-StepStatus $s 'Running'
            $lblProg.Text = "Step $($s.No) of $($Plan.Steps.Count): $($s.Title)"
            Write-Log "Step $($s.No) [$($s.Where) $($s.Server)] $($s.Title)" 'INFO'
            [System.Windows.Forms.Application]::DoEvents()
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                switch ($s.Kind) {
                    'Sql' {
                        $conn = $script:State.DstConn
                        if ($s.Where -eq 'SOURCE') { $conn = $script:State.SrcConn }
                        $t = Invoke-Sql $conn (Get-StepCode $s) -Timeout 300
                        foreach ($m in $script:SqlMessages) { Write-Log "  $m" 'INFO' }
                        if ($t -and $t.Rows.Count) {
                            foreach ($row in $t.Rows) { Write-Log ('  ' + (($t.Columns | ForEach-Object { "$($_.ColumnName)=$($row[$_.ColumnName])" }) -join '  ')) 'INFO' }
                        }
                    }
                    'Copy'   { Invoke-CopyStep $s }
                    'Delete' { Invoke-DeleteStep $s }
                    'ShareCreate' { Invoke-ShareCreateStep $s }
                    'ShareRemove' { Invoke-ShareRemoveStep $s }
                }
                $sw.Stop()
                Set-StepStatus $s 'Done' ("{0:N1}s" -f $sw.Elapsed.TotalSeconds)
                Write-Log "Step $($s.No) done ($("{0:N1}" -f $sw.Elapsed.TotalSeconds)s)." 'OK'
            } catch {
                $sw.Stop()
                $msg = Get-SqlErrorText $_
                foreach ($m in $script:SqlMessages) { Write-Log "  $m" 'INFO' }
                Set-StepStatus $s 'Failed' $msg
                Write-Log "Step $($s.No) FAILED: $msg" 'FAIL'
                if (-not $failed) { $failed = $s }
            }
            $done++; $progress.Value = [Math]::Min($progress.Maximum, $done)
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    finally {
        Remove-LeftoverShares
        Close-HostCims
        $script:Running = $false; Set-Busy $false
    }

    if ($failed) {
        $lblProg.Text = "FAILED at step $($failed.No). Remaining steps were not run."
        Write-Log "===== EXECUTE stopped at step $($failed.No). Steps after it were NOT run. =====" 'FAIL'
        $hint = ''
        if (@($Plan.Steps | Where-Object { $_.Status -eq 'Done' -and ($_.Kind -eq 'Copy' -or $_.Title -like 'Back up*') }).Count) {
            $hint = "`r`n`r`nCertificate files were already written. They hold the private key: secure or delete them ($($Plan.SrcSql) on $($Plan.Src.HostName), $($Plan.DstSql) on $($Plan.Dst.HostName))."
        }
        [void][System.Windows.Forms.MessageBox]::Show("Step $($failed.No) failed:`r`n$($failed.Title)`r`n`r`n$($failed.Result)`r`n`r`nThe steps after it were not run. See the Log tab.$hint", $script:AppName, 'OK', 'Error')
    } else {
        $progress.Value = $progress.Maximum
        $lblProg.Text = 'Completed successfully.'
        Write-Log '===== EXECUTE completed successfully =====' 'OK'
        $moved = @($Plan.Certs | Where-Object { $_.State -eq 'Move' })
        $next = New-Object System.Text.StringBuilder
        [void]$next.AppendLine("$($moved.Count) certificate(s) are now on $($Plan.Dst.ServerName) with matching thumbprints.")
        [void]$next.AppendLine('')
        [void]$next.AppendLine('Next steps:')
        [void]$next.AppendLine(" - The export ($($Plan.SrcSql) on $($Plan.Src.HostName)) holds the certificate AND private key: move it to a safe place.")
        [void]$next.AppendLine(' - Keep the certificate password in your vault: it is needed to restore these files anywhere else.')
        if (@($Plan.Steps | Where-Object { $_.Title -eq 'Create the database master key in master' -and $_.Status -eq 'Done' }).Count) {
            [void]$next.AppendLine(' - A database master key was created: back it up (BACKUP MASTER KEY TO FILE = ... ENCRYPTION BY PASSWORD = ...).')
        }
        [void]$next.AppendLine(' - Availability group: every replica needs the certificate; run the tool again for each replica.')
        Write-Log ($next.ToString().TrimEnd()) 'INFO'
        [void][System.Windows.Forms.MessageBox]::Show($next.ToString(), $script:AppName, 'OK', 'Information')
    }
    try { Connect-Dest -Quiet } catch { }
}

#endregion

#region ============================ UI helpers ============================

function Add-Ctl {
    param($Parent, $Ctl, [int]$X, [int]$Y, [int]$W, [int]$H = 23)
    $Ctl.Location = New-Object System.Drawing.Point($X, $Y)
    $Ctl.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($Ctl)
    return $Ctl
}
function New-Label {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 150, [int]$H = 23)
    $l = New-Object System.Windows.Forms.Label; $l.Text = $Text; $l.TextAlign = 'MiddleLeft'
    return (Add-Ctl $Parent $l $X $Y $W $H)
}
function New-TextBox {
    param($Parent, [int]$X, [int]$Y, [int]$W, [string]$Text = '', [switch]$Password)
    $t = New-Object System.Windows.Forms.TextBox; $t.Text = $Text
    if ($Password) { $t.UseSystemPasswordChar = $true }
    return (Add-Ctl $Parent $t $X $Y $W 23)
}
function New-Combo {
    param($Parent, [int]$X, [int]$Y, [int]$W, [string[]]$Items, [int]$Selected = 0)
    $c = New-Object System.Windows.Forms.ComboBox; $c.DropDownStyle = 'DropDownList'
    [void]$c.Items.AddRange([object[]]$Items); $c.SelectedIndex = $Selected
    return (Add-Ctl $Parent $c $X $Y $W 23)
}
function New-Check {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [bool]$Checked = $false)
    $c = New-Object System.Windows.Forms.CheckBox; $c.Text = $Text; $c.Checked = $Checked
    return (Add-Ctl $Parent $c $X $Y $W 23)
}
function New-Button {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 150, [int]$H = 27)
    $b = New-Object System.Windows.Forms.Button; $b.Text = $Text
    return (Add-Ctl $Parent $b $X $Y $W $H)
}
function New-Group {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox; $g.Text = $Text
    return (Add-Ctl $Parent $g $X $Y $W $H)
}
function New-Grid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false; $g.RowHeadersVisible = $false
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false; $g.AutoSizeColumnsMode = 'DisplayedCells'
    $g.BackgroundColor = [System.Drawing.Color]::White
    return $g
}
function New-Mono {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true; $t.ReadOnly = $true; $t.ScrollBars = 'Both'; $t.WordWrap = $false
    $t.Font = New-Object System.Drawing.Font('Consolas', 9); $t.BackColor = [System.Drawing.Color]::White
    return $t
}

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "{0} [{1,-4}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    if ($txtLog) {
        $txtLog.AppendText($line.Replace("`n", "`r`n" + (' ' * 27)) + "`r`n")
    }
}

function Set-Status { param([string]$t) $lblStatus.Text = $t; [System.Windows.Forms.Application]::DoEvents() }

function Set-Busy {
    param([bool]$Busy)
    foreach ($c in @($btnRun, $btnSrcConnect, $btnDstConnect, $rbWhatIf, $rbExec, $btnSave)) { $c.Enabled = -not $Busy }
    if ($Busy) { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor } else { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Invoke-Ui {
    param([scriptblock]$Action)
    if ($script:Running) { return }
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $Action }
    catch {
        $m = Get-SqlErrorText $_
        Write-Log $m 'FAIL'
        [void][System.Windows.Forms.MessageBox]::Show($m, $script:AppName, 'OK', 'Error')
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; Set-Status 'Ready' }
}

$script:StatusColors = @{
    'Planned' = [System.Drawing.Color]::White
    'Skip'    = [System.Drawing.Color]::Gainsboro
    'Skipped' = [System.Drawing.Color]::Gainsboro
    'Blocked' = [System.Drawing.Color]::MistyRose
    'Running' = [System.Drawing.Color]::LightYellow
    'Done'    = [System.Drawing.Color]::Honeydew
    'Failed'  = [System.Drawing.Color]::LightCoral
    'Not run' = [System.Drawing.Color]::LightGray
}

function Set-StepStatus {
    param($Step, [string]$Status, [string]$Result = '')
    $Step.Status = $Status; $Step.Result = $Result
    $i = $Step.No - 1
    if ($script:StepTable -and $i -lt $script:StepTable.Rows.Count) {
        $script:StepTable.Rows[$i]['Status'] = $Status
        $script:StepTable.Rows[$i]['Result'] = $Result
        if ($i -lt $gridSteps.Rows.Count) {
            $gridSteps.Rows[$i].DefaultCellStyle.BackColor = $script:StatusColors[$Status]
            if ($Status -eq 'Running') { $gridSteps.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $i - 3) }
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-Plan {
    param($Plan)
    $dt = New-Object System.Data.DataTable
    foreach ($c in 'No', 'Where', 'Server', 'Step', 'Changes', 'Status', 'Result') { [void]$dt.Columns.Add($c) }
    foreach ($s in $Plan.Steps) {
        $chg = 'no'
        if ($s.Changes) { $chg = 'YES' }
        [void]$dt.Rows.Add($s.No, $s.Where, $s.Server, $s.Title, $chg, $s.Status, '')
    }
    $script:StepTable = $dt
    $gridSteps.DataSource = $dt
    foreach ($col in $gridSteps.Columns) { $col.SortMode = 'NotSortable' }
    for ($i = 0; $i -lt $gridSteps.Rows.Count; $i++) { $gridSteps.Rows[$i].DefaultCellStyle.BackColor = $script:StatusColors[$Plan.Steps[$i].Status] }
    $txtScript.Text = (Get-PlanScript $Plan)
    $progress.Value = 0
    $nChg = @($Plan.Steps | Where-Object { $_.Changes -and $_.Status -eq 'Planned' }).Count
    $lblProg.Text = "$($Plan.Steps.Count) step(s), $nChg make changes." + $(if ($Plan.Blocked) { '  BLOCKED - see red rows.' } else { '' })
    if ($gridSteps.Rows.Count) { $gridSteps.Rows[0].Selected = $true; Show-StepDetail 0 }
    foreach ($w in $Plan.Warnings) { Write-Log $w 'WARN' }
}

function Show-StepDetail {
    param([int]$Index)
    if (-not $script:Plan -or $Index -lt 0 -or $Index -ge $script:Plan.Steps.Count) { return }
    $s = $script:Plan.Steps[$Index]
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("STEP $($s.No) of $($script:Plan.Steps.Count)   [$($s.Status)]")
    [void]$sb.AppendLine("Runs on : $($s.Where)  $($s.Server)")
    [void]$sb.AppendLine("Runs as : $(Get-ConnDesc $s.Where)")
    [void]$sb.AppendLine("Changes : $(if ($s.Changes) { 'YES' } else { 'no (read-only / information)' })")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine($s.Title)
    [void]$sb.AppendLine('')
    foreach ($w in (Split-Words $s.What 70)) { [void]$sb.AppendLine($w) }
    if ($s.Result) { [void]$sb.AppendLine(''); [void]$sb.AppendLine("Result: $($s.Result)") }
    if ($s.Kind -ne 'Info') {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('---------------- script ----------------')
        [void]$sb.AppendLine((Get-StepCode $s -Mask))
    }
    $txtDetail.Text = $sb.ToString()
}

#endregion

#region ============================ form ============================

$form = New-Object System.Windows.Forms.Form
$form.Text = "$script:AppName $script:Version"
$form.Size = New-Object System.Drawing.Size(1210, 820)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(8, 8); $tabs.Size = New-Object System.Drawing.Size(1178, 742)
$tabs.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($tabs)
$tabSrc  = New-Object System.Windows.Forms.TabPage; $tabSrc.Text  = '1. Source'
$tabDst  = New-Object System.Windows.Forms.TabPage; $tabDst.Text  = '2. Destination'
$tabXfer = New-Object System.Windows.Forms.TabPage; $tabXfer.Text = '3. Transfer and password'
$tabRun  = New-Object System.Windows.Forms.TabPage; $tabRun.Text  = '4. Plan and run'
$tabs.TabPages.AddRange(@($tabSrc, $tabDst, $tabXfer, $tabRun))

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 755); $lblStatus.Size = New-Object System.Drawing.Size(1170, 20)
$lblStatus.Anchor = 'Bottom,Left,Right'; $lblStatus.Text = "Ready.  What-If is the default: nothing is changed until you choose Execute and confirm.  Log: $script:LogFile"
$form.Controls.Add($lblStatus)

function New-ConnGroup {
    param($Parent, [string]$Title)
    $g = New-Group $Parent $Title 10 8 1150 112
    [void](New-Label $g 'Server\instance' 10 22 110)
    $srv = New-TextBox $g 125 22 330
    [void](New-Label $g 'Authentication' 470 22 95)
    $auth = New-Combo $g 570 22 120 @('Windows', 'SQL login')
    $enc = New-Check $g 'Encrypt connection (trust server certificate)' 710 22 320
    [void](New-Label $g 'Login' 10 52 110)
    $usr = New-TextBox $g 125 52 200
    [void](New-Label $g 'Password' 340 52 65)
    $pw = New-TextBox $g 410 52 200 -Password
    $btn = New-Button $g 'Connect and load certificates' 630 50 230
    $info = New-Label $g '' 10 80 1130 24
    $info.ForeColor = [System.Drawing.Color]::DarkSlateBlue
    $usr.Enabled = $false; $pw.Enabled = $false
    $auth.Add_SelectedIndexChanged({ $on = ($auth.SelectedIndex -eq 1); $usr.Enabled = $on; $pw.Enabled = $on }.GetNewClosure())
    return [pscustomobject]@{ Server = $srv; Auth = $auth; Encrypt = $enc; User = $usr; Pass = $pw; Button = $btn; Info = $info }
}

function Get-Conn {
    param($G)
    $s = $G.Server.Text.Trim()
    if (-not $s) { throw 'Enter the server name.' }
    return @{ Server = $s; Auth = $G.Auth.Text; User = $G.User.Text; Pass = $G.Pass.Text; Encrypt = $G.Encrypt.Checked }
}

# ---------- tab 1: source ----------
$cSrc = New-ConnGroup $tabSrc 'Source SQL Server (certificates are backed up here)'
$btnSrcConnect = $cSrc.Button
[void](New-Label $tabSrc 'Certificates in master on the source. Tick the ones to migrate (certificates used by TDE are ticked for you).' 10 128 700)
$btnSelTde  = New-Button $tabSrc 'Tick TDE ones' 760 125 130
$btnSelAll  = New-Button $tabSrc 'Tick all' 895 125 130
$btnSelNone = New-Button $tabSrc 'Untick all' 1030 125 130
$gridSrc = New-Grid
[void](Add-Ctl $tabSrc $gridSrc 10 158 1150 545)
$gridSrc.Anchor = 'Top,Bottom,Left,Right'

# ---------- tab 2: destination ----------
$cDst = New-ConnGroup $tabDst 'Destination SQL Server (certificates are created here)'
$btnDstConnect = $cDst.Button
$gDmk = New-Group $tabDst 'Database master key (DMK) in master on the destination' 10 126 1150 118
$lblDmk = New-Label $gDmk 'Connect to the destination to check the master key.' 10 20 1130 22
$lblDmk.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$chkCreateDmk = New-Check $gDmk 'Create the database master key if master has none' 10 46 380
[void](New-Label $gDmk 'DMK password' 400 46 95)
$txtDmkPwd = New-TextBox $gDmk 500 46 200 -Password
[void](New-Label $gDmk 'Confirm' 715 46 60)
$txtDmkPwd2 = New-TextBox $gDmk 780 46 200 -Password
[void](New-Label $gDmk 'The DMK password is only needed when master has no DMK (to create it), or when its DMK is not encrypted by the service master key (to OPEN it). An existing DMK is never changed.' 10 76 1130 34)
[void](New-Label $tabDst 'If a DIFFERENT certificate with the same name exists' 10 252 330)
$cboConflict = New-Combo $tabDst 345 252 520 @('Stop: do not migrate that certificate (nothing is replaced)', 'Create it with a suffix: <name>_migrated_<yyyyMMdd>')
[void](New-Label $tabDst 'Certificates already in master on the destination (read-only view):' 10 282 600)
$gridDst = New-Grid
$gridDst.ReadOnly = $true
[void](Add-Ctl $tabDst $gridDst 10 308 1150 395)
$gridDst.Anchor = 'Top,Bottom,Left,Right'

# ---------- tab 3: transfer ----------
$gFold = New-Group $tabXfer 'Folders' 10 8 1150 200
[void](New-Label $gFold 'Export folder on the SOURCE (as the source SQL Server sees it)' 10 24 400)
$txtSrcFolder = New-TextBox $gFold 415 24 540
[void](New-Label $gFold '   same folder, as THIS PC reaches it' 10 54 400)
$txtSrcUnc = New-TextBox $gFold 415 54 540
[void](New-Label $gFold 'Import folder on the DESTINATION (as the destination SQL Server sees it)' 10 92 400)
$txtDstFolder = New-TextBox $gFold 415 92 540
[void](New-Label $gFold '   same folder, as THIS PC reaches it' 10 122 400)
$txtDstUnc = New-TextBox $gFold 415 122 540
[void](New-Label $gFold 'How THIS PC reaches the folders' 10 154 200)
$cboReach = New-Combo $gFold 215 154 740 @('Temporary shares: created on each folder for the copy, removed when done (folders kept)', 'Admin shares (\\host\X$\...): I am a local admin on both hosts', 'Paths I type in the boxes above')
$btnResetFold = New-Button $gFold 'Use instance backup folders' 965 22 175
[void](New-Label $gFold 'Default: <instance default backup folder>\TDE_Migration. The SQL Server service creates the folder itself, so it has the rights it needs.' 10 178 1130 20)

$gPwd = New-Group $tabXfer 'Private key password' 10 216 1150 120
[void](New-Label $gPwd 'Used for ENCRYPTION BY PASSWORD when the private key is backed up on the source, and DECRYPTION BY PASSWORD when it is created on the destination.' 10 20 1130 22)
[void](New-Label $gPwd 'Password' 10 50 70)
$txtCertPwd = New-TextBox $gPwd 85 50 260 -Password
[void](New-Label $gPwd 'Confirm' 360 50 60)
$txtCertPwd2 = New-TextBox $gPwd 425 50 260 -Password
$chkShowPwd = New-Check $gPwd 'Show passwords' 700 50 150
[void](New-Label $gPwd 'Keep it in your password vault together with the .cer/.pvk files: without it the private key cannot be restored. At least 8 characters; SQL Server also applies the Windows password policy. Not needed for What-If.' 10 80 1130 34)

$gOpt = New-Group $tabXfer 'After the run' 10 344 1150 84
$chkDeleteDest = New-Check $gOpt 'After the verify step passes, delete the copies this run put in the destination import folder (the source export is always kept)' 10 22 1100
[void](New-Label $gOpt 'File names: <certificate>_<yyyyMMdd_HHmmss>.cer / .pvk - a new name every run, so no file is ever overwritten.' 10 50 1100 22)

# ---------- tab 4: plan and run ----------
$rbWhatIf = New-Object System.Windows.Forms.RadioButton; $rbWhatIf.Text = 'What-If: show each step, its script and the server it runs on (changes nothing)'; $rbWhatIf.Checked = $true
[void](Add-Ctl $tabRun $rbWhatIf 10 8 520 24)
$rbExec = New-Object System.Windows.Forms.RadioButton; $rbExec.Text = 'Execute: run the steps with progress (asks to confirm)'
[void](Add-Ctl $tabRun $rbExec 540 8 360 24)
$btnRun = New-Button $tabRun 'Show plan (What-If)' 905 4 160 30
$btnSave = New-Button $tabRun 'Save script...' 1070 4 90 30
$progress = New-Object System.Windows.Forms.ProgressBar
[void](Add-Ctl $tabRun $progress 10 40 600 20)
$lblProg = New-Label $tabRun 'Build the plan with What-If first.' 620 38 540 24
$progress.Anchor = 'Top,Left'; $lblProg.Anchor = 'Top,Left,Right'

$split = New-Object System.Windows.Forms.SplitContainer
$split.Orientation = 'Horizontal'
[void](Add-Ctl $tabRun $split 10 68 1150 635)
$split.Anchor = 'Top,Bottom,Left,Right'
try { $split.SplitterDistance = 330 } catch { }

$splitTop = New-Object System.Windows.Forms.SplitContainer
$splitTop.Dock = 'Fill'
$split.Panel1.Controls.Add($splitTop)
try { $splitTop.SplitterDistance = 640 } catch { }
$gridSteps = New-Grid
$gridSteps.ReadOnly = $true; $gridSteps.Dock = 'Fill'
$splitTop.Panel1.Controls.Add($gridSteps)
$txtDetail = New-Mono
$txtDetail.Dock = 'Fill'; $txtDetail.Text = 'Select a step to see what it does, where it runs and its script.'
$splitTop.Panel2.Controls.Add($txtDetail)

$tabsBottom = New-Object System.Windows.Forms.TabControl
$tabsBottom.Dock = 'Fill'
$split.Panel2.Controls.Add($tabsBottom)
$tabScript = New-Object System.Windows.Forms.TabPage; $tabScript.Text = 'Full script (SQLCMD mode, passwords masked)'
$tabLog = New-Object System.Windows.Forms.TabPage; $tabLog.Text = 'Log'
$tabsBottom.TabPages.AddRange(@($tabScript, $tabLog))
$txtScript = New-Mono; $txtScript.Dock = 'Fill'; $tabScript.Controls.Add($txtScript)
$txtLog = New-Mono; $txtLog.Dock = 'Fill'; $tabLog.Controls.Add($txtLog)

#endregion

#region ============================ UI logic ============================

function Get-SharePreview {
    param([string]$Path, [string]$HostName, [string]$Tag)
    if (-not $Path) { return '' }
    if ($Path.Trim() -match '^\\\\' -or -not $HostName -or $HostName -ieq $env:COMPUTERNAME) { return (Get-UncPath $Path $HostName) }
    return "\\$HostName\TDEMIG_$($Tag)_<time stamp>`$     (temporary share: created by the run, removed when done)"
}

function Update-Unc {
    $sh = $null; $dh = $null
    if ($script:State.Src) { $sh = $script:State.Src.HostName }
    if ($script:State.Dst) { $dh = $script:State.Dst.HostName }
    switch ($cboReach.SelectedIndex) {
        0 { $txtSrcUnc.Text = Get-SharePreview $txtSrcFolder.Text $sh 'SRC'; $txtDstUnc.Text = Get-SharePreview $txtDstFolder.Text $dh 'DST' }
        1 { $txtSrcUnc.Text = Get-UncPath $txtSrcFolder.Text $sh; $txtDstUnc.Text = Get-UncPath $txtDstFolder.Text $dh }
        default { }
    }
}

function Get-DefaultFolder {
    param($Info)
    $b = $Info.BackupPath
    if (-not $b) { $b = 'C:\' }
    if ($b -match '^https?://') { $b = 'C:\' }   # backups to URL: fall back to a local folder
    return (Join-WinPath $b 'TDE_Migration')
}

function Connect-Source {
    $c = Get-Conn $cSrc
    Set-Status "Connecting to $($c.Server)..."
    $info = Get-ServerInfo $c
    $certs = Get-Certificates $c
    foreach ($r in $certs.Rows) { if ([string]$r.UsedByTde) { $r.Migrate = $true } }
    $script:State.Src = $info; $script:State.SrcConn = $c; $script:State.SrcCerts = $certs
    $gridSrc.DataSource = $certs
    foreach ($col in $gridSrc.Columns) { $col.ReadOnly = ($col.Name -ne 'Migrate') }
    $nTde = @($certs.Rows | Where-Object { [string]$_.UsedByTde }).Count
    $cSrc.Info.Text = "Connected: $($info.ServerName) on host $($info.HostName) | $($info.Version) $($info.Level) $($info.Edition) | service: $($info.ServiceAccount) | sysadmin: $($info.IsSysadmin) | $($certs.Rows.Count) certificate(s), $nTde used by TDE"
    if (-not $txtSrcFolder.Text.Trim()) { $txtSrcFolder.Text = Get-DefaultFolder $info }
    Update-Unc
    Write-Log "Source connected: $($info.ServerName) (host $($info.HostName), $($info.Version)), $($certs.Rows.Count) certificate(s), $nTde used by TDE." 'OK'
    $script:Plan = $null
}

function Set-DestInfo {
    $info = $script:State.Dst
    if (-not $info) { return }
    if ($info.HasDmk -and $info.DmkBySmk) {
        $lblDmk.Text = 'Master key: PRESENT, encrypted by the service master key. Nothing to do; no password needed.'
        $lblDmk.ForeColor = [System.Drawing.Color]::DarkGreen
        $chkCreateDmk.Enabled = $false; $txtDmkPwd.Enabled = $false; $txtDmkPwd2.Enabled = $false
    } elseif ($info.HasDmk) {
        $lblDmk.Text = 'Master key: PRESENT but NOT encrypted by the service master key. Enter its password: it is used to OPEN it for CREATE CERTIFICATE (it is not changed).'
        $lblDmk.ForeColor = [System.Drawing.Color]::DarkOrange
        $chkCreateDmk.Enabled = $false; $txtDmkPwd.Enabled = $true; $txtDmkPwd2.Enabled = $false
    } else {
        $lblDmk.Text = 'Master key: MISSING. CREATE CERTIFICATE needs one. Tick the box below and enter a password to have it created (after the export and copy succeed).'
        $lblDmk.ForeColor = [System.Drawing.Color]::Firebrick
        $chkCreateDmk.Enabled = $true; $txtDmkPwd.Enabled = $true; $txtDmkPwd2.Enabled = $true
    }
}

function Connect-Dest {
    param([switch]$Quiet)
    $c = $script:State.DstConn
    if (-not $Quiet) { $c = Get-Conn $cDst; Set-Status "Connecting to $($c.Server)..." }
    $info = Get-ServerInfo $c
    $certs = Get-Certificates $c
    $certs.Columns.Remove('Migrate')
    $script:State.Dst = $info; $script:State.DstConn = $c; $script:State.DstCerts = $certs
    $gridDst.DataSource = $certs
    $cDst.Info.Text = "Connected: $($info.ServerName) on host $($info.HostName) | $($info.Version) $($info.Level) $($info.Edition) | service: $($info.ServiceAccount) | sysadmin: $($info.IsSysadmin) | $($certs.Rows.Count) certificate(s)"
    Set-DestInfo
    if (-not $Quiet) {
        if (-not $txtDstFolder.Text.Trim()) { $txtDstFolder.Text = Get-DefaultFolder $info }
        Update-Unc
        $dmk = 'missing'
        if ($info.HasDmk) { $dmk = "present (encrypted by service master key: $($info.DmkBySmk))" }
        Write-Log "Destination connected: $($info.ServerName) (host $($info.HostName), $($info.Version)), DMK $dmk, $($certs.Rows.Count) certificate(s)." 'OK'
        $script:Plan = $null
    }
}

$btnSrcConnect.Add_Click({ Invoke-Ui { Connect-Source } })
$btnDstConnect.Add_Click({ Invoke-Ui { Connect-Dest } })

$btnSelTde.Add_Click({ if ($script:State.SrcCerts) { foreach ($r in $script:State.SrcCerts.Rows) { $r.Migrate = [bool]([string]$r.UsedByTde) } } })
$btnSelAll.Add_Click({ if ($script:State.SrcCerts) { foreach ($r in $script:State.SrcCerts.Rows) { $r.Migrate = $true } } })
$btnSelNone.Add_Click({ if ($script:State.SrcCerts) { foreach ($r in $script:State.SrcCerts.Rows) { $r.Migrate = $false } } })
$gridSrc.Add_CurrentCellDirtyStateChanged({ if ($gridSrc.IsCurrentCellDirty) { [void]$gridSrc.CommitEdit('CurrentCellChange') } })

$txtSrcFolder.Add_TextChanged({ Update-Unc })
$txtDstFolder.Add_TextChanged({ Update-Unc })
$cboReach.Add_SelectedIndexChanged({ $ro = ($cboReach.SelectedIndex -ne 2); $txtSrcUnc.ReadOnly = $ro; $txtDstUnc.ReadOnly = $ro; Update-Unc })
$txtSrcUnc.ReadOnly = $true; $txtDstUnc.ReadOnly = $true
$btnResetFold.Add_Click({
    if ($script:State.Src) { $txtSrcFolder.Text = Get-DefaultFolder $script:State.Src }
    if ($script:State.Dst) { $txtDstFolder.Text = Get-DefaultFolder $script:State.Dst }
    Update-Unc
})
$chkShowPwd.Add_CheckedChanged({
    foreach ($t in @($txtCertPwd, $txtCertPwd2, $txtDmkPwd, $txtDmkPwd2)) { $t.UseSystemPasswordChar = -not $chkShowPwd.Checked }
})

$rbWhatIf.Add_CheckedChanged({ if ($rbWhatIf.Checked) { $btnRun.Text = 'Show plan (What-If)' } else { $btnRun.Text = 'Execute...' } })
$gridSteps.Add_SelectionChanged({ if ($gridSteps.CurrentRow) { Show-StepDetail $gridSteps.CurrentRow.Index } })

$btnRun.Add_Click({
    Invoke-Ui {
        Set-Status 'Building the plan (read-only)...'
        $script:Plan = New-Plan
        Show-Plan $script:Plan
        Write-Log "Plan built (What-If): $($script:Plan.Steps.Count) step(s), $(@($script:Plan.Steps | Where-Object { $_.Changes -and $_.Status -eq 'Planned' }).Count) with changes$(if ($script:Plan.Blocked) { ', BLOCKED' })." 'INFO'
        if ($rbExec.Checked) {
            Test-ExecuteInputs $script:Plan
            if (Confirm-Execute $script:Plan) { Invoke-Plan $script:Plan }
            else { Write-Log 'Execute cancelled; nothing was changed.' 'INFO' }
        } else {
            $tabsBottom.SelectedTab = $tabScript
        }
    }
})

$btnSave.Add_Click({
    Invoke-Ui {
        if (-not $script:Plan) { throw 'Build the plan first (Show plan).' }
        $d = New-Object System.Windows.Forms.SaveFileDialog
        $d.Filter = 'SQL script (*.sql)|*.sql|All files (*.*)|*.*'
        $d.FileName = "TdeCertMigration_$(ConvertTo-SafeName $script:Plan.Src.ServerName)_to_$(ConvertTo-SafeName $script:Plan.Dst.ServerName)_$($script:Plan.Stamp).sql"
        if ($d.ShowDialog($form) -eq 'OK') {
            [System.IO.File]::WriteAllText($d.FileName, (Get-PlanScript $script:Plan), (New-Object System.Text.UTF8Encoding($true)))
            Write-Log "Saved plan script (passwords masked) to $($d.FileName)" 'OK'
        }
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    if ($script:Running) {
        [void][System.Windows.Forms.MessageBox]::Show('A run is in progress. Wait for it to finish.', $script:AppName, 'OK', 'Warning')
        $e.Cancel = $true
        return
    }
    if ($script:CreatedShares.Count) { Remove-LeftoverShares; Close-HostCims }
})

$form.Add_Shown({
    try { $split.SplitterDistance = 330; $splitTop.SplitterDistance = 640 } catch { }
    Write-Log "$script:AppName $script:Version started by $(Get-Me) on $env:COMPUTERNAME. Mode: What-If (nothing is changed until you choose Execute and confirm)." 'INFO'
})

#endregion

[void]$form.ShowDialog()
$form.Dispose()

#Requires -Version 5.1
<#
.SYNOPSIS
    PITR Restore Builder - WinForms GUI that builds a FULL / FULL+DIFF / FULL+DIFF+LOG (STOPAT)
    restore script from msdb backup history, and preflights it on the destination server.

.DESCRIPTION
    READ-ONLY / GENERATE-ONLY BY DEFAULT
      Source      : SELECTs only (msdb backup history, sys.databases, TDE certificate info).
      Destination : SELECTs, RESTORE HEADERONLY, RESTORE FILELISTONLY, xp_fileexist, xp_fixeddrives.
      Nothing is restored by this tool. It writes .sql / .ps1 files for you to review and run.

    RESTORE DEFAULTS
      Every step: WITH FILE = n, NORECOVERY, STATS = 10, BUFFERCOUNT = 450, MAXTRANSFERSIZE = 4194304
      STOPAT only on the last LOG step. RECOVERY only when "WITH RECOVERY" is ticked. REPLACE only when ticked.

    OPTIONAL EXECUTION (only when "Allow execution" is ticked AND each action is confirmed)
      - azcopy upload of backup files to blob storage
      - CREATE CREDENTIAL (SAS) for the container on the destination
      - TDE / backup-encryption certificate transfer: BACKUP CERTIFICATE on the source to a temporary
        SMB share on this machine, copy to the destination folder (default C:\TDE), CREATE CERTIFICATE on
        the destination. The share is removed afterwards (the folder is kept).
      - Create (and optionally start) the SQL Agent job that runs the restore. Owner = sa (SUSER_SNAME(0x01)).

.NOTES
    Windows PowerShell 5.1, no modules needed (System.Data.SqlClient + WinForms).
    STOPAT is in the SOURCE server's local time: msdb and the transaction log record source local time.
    Run   : powershell.exe -ExecutionPolicy Bypass -File .\PITR-RestoreBuilder.ps1
    Run elevated (Run as administrator) if you want the tool to create the temporary TDE share.
#>
[CmdletBinding()] param()

$ErrorActionPreference = 'Stop'
$script:AppName = 'PITR Restore Builder'
$script:Version = '1.7'

#region ============================ general helpers ============================

function DbVal  { param($v) if ($null -eq $v -or $v -is [System.DBNull]) { return $null }; return $v }
function DbVal0 { param($v) if ($null -eq $v -or $v -is [System.DBNull]) { return 0 };     return $v }
function QN  { param([string]$s) return '[' + $s.Replace(']', ']]') + ']' }
function QS  { param([string]$s) if ($null -eq $s) { $s = '' }; return "N'" + $s.Replace("'", "''") + "'" }
function ConvertTo-SafeName { param([string]$s) return ($s -replace '[\\/:*?"<>|\s\[\]]', '_') }

function Get-Leaf {
    param([string]$p)
    $i = [Math]::Max($p.LastIndexOf('\'), $p.LastIndexOf('/'))
    if ($i -ge 0) { return $p.Substring($i + 1) }
    return $p
}

function Get-DirOf {
    param([string]$p)
    if (-not $p) { return $null }
    $i = [Math]::Max($p.LastIndexOf('\'), $p.LastIndexOf('/'))
    if ($i -le 0) { return $null }
    $d = $p.Substring(0, $i)
    if ($d -match '^[A-Za-z]:$') { $d += '\' }
    return $d
}

function Join-WinPath {
    param([string]$Dir, [string]$Leaf)
    return $Dir.TrimEnd('\') + '\' + $Leaf
}

function Get-ContainerUrl {
    param([string]$Url)
    $u = [uri]$Url
    $seg = $u.AbsolutePath.Trim('/').Split('/')[0]
    return "$($u.Scheme)://$($u.Authority)/$seg"
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

function Test-IsAdmin {
    $p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Account that a SQL Server service presents on the network (virtual/local accounts -> DOMAIN\HOST$)
function Get-NetworkIdentity {
    param([string]$Account, [string]$Domain, [string]$HostName)
    if (-not $Account) { return $null }
    if ($Account -match '^(NT SERVICE|NT AUTHORITY)\\' -or $Account -eq 'LocalSystem' -or $Account -match '^\.\\') {
        if ($Domain -and $Domain -ne $HostName -and $Domain -ne 'WORKGROUP') { return "$Domain\$HostName`$" }
        return $null
    }
    return $Account
}

function ConvertTo-DataTable {
    param($Objects, [string[]]$Props)
    $dt = New-Object System.Data.DataTable
    foreach ($p in $Props) { [void]$dt.Columns.Add($p) }
    foreach ($o in $Objects) {
        $r = $dt.NewRow()
        foreach ($p in $Props) {
            $v = $o.$p
            if ($null -eq $v) { $r[$p] = [System.DBNull]::Value }
            elseif ($v -is [datetime]) { $r[$p] = $v.ToString('yyyy-MM-dd HH:mm:ss') }
            else { $r[$p] = [string]$v }
        }
        $dt.Rows.Add($r)
    }
    return , $dt
}

#endregion

#region ============================ SQL access (read-only) ============================

function New-ConnString {
    param([hashtable]$Conn, [string]$Database = 'master')
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source']      = $Conn.Server
    $b['Initial Catalog']  = $Database
    $b['Application Name'] = 'PITR-RestoreBuilder'
    $b['Connect Timeout']  = 15
    if ($Conn.Auth -eq 'SQL login') { $b['User ID'] = $Conn.User; $b['Password'] = $Conn.Pass }
    else { $b['Integrated Security'] = $true }
    return $b.ConnectionString
}

function Invoke-Sql {
    param([hashtable]$Conn, [string]$Query, [hashtable]$Params = @{}, [int]$Timeout = 120, [string]$Database = 'master')
    $cn = New-Object System.Data.SqlClient.SqlConnection (New-ConnString -Conn $Conn -Database $Database)
    try {
        $cn.Open()
        $cmd = $cn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        foreach ($k in $Params.Keys) {
            $v = $Params[$k]
            if ($null -eq $v) { $v = [System.DBNull]::Value }
            [void]$cmd.Parameters.AddWithValue("@$k", $v)
        }
        $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $ds = New-Object System.Data.DataSet
        [void]$da.Fill($ds)
        if ($ds.Tables.Count -gt 0) { return , $ds.Tables[0] }
        return , (New-Object System.Data.DataTable)
    }
    finally { $cn.Dispose() }
}

# Runs a script that contains GO separators (as SSMS would): one batch at a time
function Invoke-SqlBatches {
    param([hashtable]$Conn, [string]$Script, [string]$Database = 'master', [int]$Timeout = 300)
    $batches = [regex]::Split($Script, '(?im)^[ \t]*GO[ \t]*(?:--[^\r\n]*)?\r?$')
    foreach ($b in $batches) {
        $body = ($b -split "`r?`n" | Where-Object { $_.Trim() -and -not $_.Trim().StartsWith('--') }) -join ''
        if (-not $body.Trim()) { continue }
        Invoke-Sql $Conn $b -Database $Database -Timeout $Timeout | Out-Null
    }
}

#endregion

#region ============================ source: msdb history ============================

function Get-SourceInfo {
    param([hashtable]$Conn, [string]$Db)
    $r = (Invoke-Sql $Conn @"
SELECT @@SERVERNAME AS ServerName, SYSDATETIME() AS ServerTime,
       CAST(PARSENAME(CONVERT(varchar(32), SERVERPROPERTY('ProductVersion')), 4) AS int) AS Major,
       CONVERT(nvarchar(128), SERVERPROPERTY('ComputerNamePhysicalNetBIOS')) AS HostName,
       DEFAULT_DOMAIN() AS DomainName;
"@).Rows[0]
    $info = [pscustomobject]@{
        ServerName = [string]$r.ServerName; ServerTime = [datetime]$r.ServerTime; Major = [int]$r.Major
        HostName = [string]$r.HostName; DomainName = [string](DbVal $r.DomainName); ServiceAccount = $null
        DbExists = $false; RecoveryModel = $null; IsEncrypted = $false; TdeCertName = $null; TdeCertThumb = $null
        ServerFilter = 'AND bs.server_name = @@SERVERNAME'; FilterNote = $null
    }
    try {
        $s = Invoke-Sql $Conn "SELECT TOP (1) service_account FROM sys.dm_server_services WHERE servicename LIKE N'SQL Server (%';"
        if ($s.Rows.Count) { $info.ServiceAccount = [string]$s.Rows[0][0] }
    } catch { }
    if ($Db) {
        $d = Invoke-Sql $Conn @"
SELECT d.recovery_model_desc, d.is_encrypted, c.name AS cert_name,
       CONVERT(varchar(64), dek.encryptor_thumbprint, 1) AS dek_thumb
FROM sys.databases d
LEFT JOIN sys.dm_database_encryption_keys dek ON dek.database_id = d.database_id
LEFT JOIN master.sys.certificates c ON c.thumbprint = dek.encryptor_thumbprint
WHERE d.name = @db;
"@ @{ db = $Db }
        if ($d.Rows.Count) {
            $x = $d.Rows[0]
            $info.DbExists = $true
            $info.RecoveryModel = [string]$x.recovery_model_desc
            $info.IsEncrypted = [bool](DbVal0 $x.is_encrypted)
            $info.TdeCertName = [string](DbVal $x.cert_name)
            $info.TdeCertThumb = [string](DbVal $x.dek_thumb)
            if (-not $info.TdeCertName) { $info.TdeCertThumb = $null }
        }
        $c = (Invoke-Sql $Conn @"
SELECT SUM(CASE WHEN server_name = @@SERVERNAME THEN 1 ELSE 0 END) AS local_rows, COUNT(*) AS all_rows
FROM msdb.dbo.backupset WHERE database_name = @db;
"@ @{ db = $Db }).Rows[0]
        if ([int](DbVal0 $c.local_rows) -eq 0 -and [int](DbVal0 $c.all_rows) -gt 0) {
            $info.ServerFilter = ''
            $info.FilterNote = "msdb has $($c.all_rows) backup rows for [$Db] but none with server_name = @@SERVERNAME (server renamed, or rows imported by RESTORE). Using all rows - check the file paths carefully."
        }
    }
    return $info
}

function Get-SourceCertByThumb {
    param([hashtable]$Conn, [string]$Thumb)
    $t = Invoke-Sql $Conn "SELECT name FROM master.sys.certificates WHERE thumbprint = CONVERT(varbinary(64), @t, 1);" @{ t = $Thumb }
    if ($t.Rows.Count) { return [string]$t.Rows[0][0] }
    return $null
}

function Get-BsSelect {
    param([int]$Major, [string]$Top = '', [string]$Filter = '')
    $enc = if ($Major -ge 12) { 'CONVERT(varchar(64), bs.encryptor_thumbprint, 1)' } else { 'CAST(NULL AS varchar(64))' }
@"
SELECT $Top bs.backup_set_id, bs.media_set_id, bs.position, bs.type, bs.is_copy_only, bs.is_snapshot, bs.is_damaged,
       bs.has_bulk_logged_data, bs.backup_start_date, bs.backup_finish_date, bs.first_lsn, bs.last_lsn,
       bs.checkpoint_lsn, bs.database_backup_lsn, bs.first_recovery_fork_guid, bs.last_recovery_fork_guid,
       bs.backup_size, bs.compressed_backup_size, bs.software_major_version, bs.recovery_model,
       $enc AS encryptor_thumb,
       STUFF((SELECT N'|' + CAST(ISNULL(mf.device_type, 2) AS nvarchar(5)) + N'>' + mf.physical_device_name
              FROM msdb.dbo.backupmediafamily mf
              WHERE mf.media_set_id = bs.media_set_id AND mf.mirror = 0
              ORDER BY mf.family_sequence_number
              FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 1, N'') AS devices
FROM msdb.dbo.backupset bs
WHERE bs.database_name = @db $Filter
"@
}

function ConvertTo-BackupSet {
    param($r)
    $devs = @()
    $raw = [string](DbVal $r.devices)
    foreach ($d in $raw.Split('|')) {
        if ($d.Length -gt 0) {
            $i = $d.IndexOf('>')
            $devs += [pscustomobject]@{ DeviceType = [int]$d.Substring(0, $i); Path = $d.Substring($i + 1) }
        }
    }
    $typeName = switch ([string]$r.type) { 'D' { 'FULL' } 'I' { 'DIFF' } 'L' { 'LOG' } default { [string]$r.type } }
    return [pscustomobject]@{
        Step = 0; Type = $typeName; Id = [int]$r.backup_set_id; MediaSetId = [int]$r.media_set_id; Position = [int]$r.position
        IsCopyOnly = [bool](DbVal0 $r.is_copy_only); IsSnapshot = [bool](DbVal0 $r.is_snapshot); IsDamaged = [bool](DbVal0 $r.is_damaged)
        HasBulk = [bool](DbVal0 $r.has_bulk_logged_data)
        Start = [datetime]$r.backup_start_date; Finish = [datetime]$r.backup_finish_date
        FirstLsn = [decimal]$r.first_lsn; LastLsn = [decimal]$r.last_lsn
        CheckpointLsn = (DbVal $r.checkpoint_lsn); DatabaseBackupLsn = (DbVal $r.database_backup_lsn)
        FirstFork = (DbVal $r.first_recovery_fork_guid); LastFork = (DbVal $r.last_recovery_fork_guid)
        SizeMB = [math]::Round([double](DbVal0 $r.backup_size) / 1MB, 1)
        CompressedMB = [math]::Round([double](DbVal0 $r.compressed_backup_size) / 1MB, 1)
        SwMajor = [int](DbVal0 $r.software_major_version); RecoveryModel = [string](DbVal $r.recovery_model)
        EncryptorThumb = [string](DbVal $r.encryptor_thumb)
        Devices = $devs; Files = (($devs | ForEach-Object { $_.Path }) -join ' ; ')
    }
}

function Get-BackupChain {
    param([hashtable]$Conn, [string]$Db, [datetime]$StopAt, [string]$Mode, [int]$PinnedFullId = 0, [int]$SrcMajor = 13, [string]$ServerFilter = '',
          $ContinueLsn = $null, $DiffBaseLsn = $null)
    $warn = New-Object System.Collections.Generic.List[string]
    $fmt = 'yyyy-MM-dd hh:mm:ss tt'
    $full = $null
    $isContinue = ($Mode -like 'Continue*')
    if ($isContinue -and -not $ContinueLsn) { throw 'Continue mode needs the redo point (redo_start_lsn) of the RESTORING target database on the destination.' }

    # ---- FULL (skipped when continuing a restore that already has its FULL) ----
    if ($isContinue) { $eff = $Mode }
    else {
    if ($PinnedFullId -gt 0) {
        $q = (Get-BsSelect $SrcMajor '' $ServerFilter) + " AND bs.type = 'D' AND bs.backup_set_id = @id;"
        $rows = Invoke-Sql $Conn $q @{ db = $Db; id = $PinnedFullId }
    } else {
        $q = (Get-BsSelect $SrcMajor 'TOP (20)' $ServerFilter) + " AND bs.type = 'D' AND bs.backup_finish_date <= @stop ORDER BY bs.backup_finish_date DESC;"
        $rows = Invoke-Sql $Conn $q @{ db = $Db; stop = $StopAt }
    }
    $full = $null
    foreach ($f in @($rows.Rows | ForEach-Object { ConvertTo-BackupSet $_ } | Sort-Object Finish -Descending)) {
        if ($f.IsDamaged) { $warn.Add("Skipped damaged FULL id $($f.Id) ($($f.Finish))."); continue }
        if ($f.IsSnapshot -or @($f.Devices | Where-Object { $_.DeviceType -eq 7 }).Count) {
            $warn.Add("Skipped snapshot/VDI FULL id $($f.Id) ($($f.Finish)): restore it with the tool that created it."); continue
        }
        $full = $f; break
    }
    if (-not $full) { throw "No usable FULL backup of [$Db] that finished on or before $($StopAt.ToString('yyyy-MM-dd HH:mm:ss')) was found in msdb." }
    if ($full.Finish -gt $StopAt) { throw "The pinned FULL (id $($full.Id)) finished at $($full.Finish), after STOPAT. Pick an earlier FULL or a later STOPAT." }
    $fmt = 'yyyy-MM-dd hh:mm:ss tt'
    if ($PinnedFullId -gt 0) { $warn.Add("FULL is PINNED to id $($full.Id) (finished $($full.Finish.ToString($fmt))) by the checkbox on tab 1; newer FULLs are ignored.") }
    $q = (Get-BsSelect $SrcMajor 'TOP (5)' $ServerFilter) + " AND bs.type = 'D' AND bs.backup_finish_date > @stop ORDER BY bs.backup_finish_date ASC;"
    $newer = @((Invoke-Sql $Conn $q @{ db = $Db; stop = $StopAt }).Rows | ForEach-Object { ConvertTo-BackupSet $_ } | Where-Object { $_.Finish -gt $StopAt } | Sort-Object Finish)
    if ($newer.Count) {
        $warn.Add(("CHECK STOPAT: a newer FULL (id {0}, finished {1}) exists AFTER your STOPAT {2}, so it is not used. If you meant a time after it, fix STOPAT (watch AM/PM) and build again." -f $newer[0].Id, $newer[0].Finish.ToString($fmt), $StopAt.ToString($fmt)))
    }

    $eff = $Mode
    if ($Mode -eq 'Auto') {
        if ($full.RecoveryModel -eq 'SIMPLE') { $eff = 'FullDiff'; $warn.Add('Database was in SIMPLE recovery at FULL time: no log backups, using FULL + DIFF.') }
        else { $eff = 'FullDiffLog' }
    }
    }   # end of FULL section

    # ---- DIFF when continuing: newest DIFF on the restored FULL, newer than the current restore point ----
    $diff = $null
    if ($eff -eq 'ContinueDiff' -or $eff -eq 'ContinueDiffLog') {
        if (-not $DiffBaseLsn) { $warn.Add('The differential base of the RESTORING database is unknown; DIFF skipped.') }
        else {
            $q = (Get-BsSelect $SrcMajor 'TOP (20)' $ServerFilter) + " AND bs.type = 'I' AND bs.backup_finish_date <= @stop AND bs.database_backup_lsn = @base AND bs.last_lsn > @redo ORDER BY bs.backup_finish_date DESC;"
            $rows = Invoke-Sql $Conn $q @{ db = $Db; stop = $StopAt; base = [decimal]$DiffBaseLsn; redo = [decimal]$ContinueLsn }
            foreach ($d in @($rows.Rows | ForEach-Object { ConvertTo-BackupSet $_ } | Sort-Object Finish -Descending)) {
                if ($d.IsDamaged -or $d.IsSnapshot) { $warn.Add("Skipped DIFF id $($d.Id) (damaged or snapshot)."); continue }
                $diff = $d; break
            }
        }
        if (-not $diff) {
            if ($eff -eq 'ContinueDiff') { throw ("No DIFF found that is based on the FULL already restored (differential base LSN $DiffBaseLsn), is newer than the current restore point (LSN $ContinueLsn) and finished before STOPAT " + $StopAt.ToString($fmt) + '.') }
            $warn.Add('No newer DIFF on the restored FULL before STOPAT; continuing with LOGs only.')
        }
    }

    # ---- DIFF ----
    if ($eff -eq 'FullDiff' -or $eff -eq 'FullDiffLog') {
        if ($full.IsCopyOnly) { $warn.Add('Selected FULL is COPY_ONLY: it cannot be a differential base, DIFFs skipped.') }
        else {
            $q = (Get-BsSelect $SrcMajor 'TOP (20)' $ServerFilter) + " AND bs.type = 'I' AND bs.backup_finish_date <= @stop AND bs.database_backup_lsn = @ckpt ORDER BY bs.backup_finish_date DESC;"
            $rows = Invoke-Sql $Conn $q @{ db = $Db; stop = $StopAt; ckpt = $full.CheckpointLsn }
            foreach ($d in @($rows.Rows | ForEach-Object { ConvertTo-BackupSet $_ } | Sort-Object Finish -Descending)) {
                if ($d.IsDamaged -or $d.IsSnapshot) { $warn.Add("Skipped DIFF id $($d.Id) (damaged or snapshot)."); continue }
                $diff = $d; break
            }
            if (-not $diff) { $warn.Add('No DIFF based on this FULL before STOPAT; continuing from the FULL.') }
        }
    }

    # ---- LOGS ----
    $logs = @(); $usesStop = $false
    if ($eff -eq 'FullDiffLog' -or $eff -eq 'ContinueDiffLog' -or $eff -eq 'ContinueLog') {
        if ($diff) { $base = $diff }
        elseif ($full) { $base = $full }
        else { $base = [pscustomobject]@{ Type = 'current restore point of the RESTORING database'; Id = '-'; Finish = $null; LastLsn = [decimal]$ContinueLsn; LastFork = $null } }
        $f2 = $ServerFilter -replace 'bs\.', 'b2.'
        $q = (Get-BsSelect $SrcMajor '' $ServerFilter) + @"
 AND bs.type = 'L' AND bs.last_lsn > @startLsn
 AND (bs.backup_finish_date <= @stop
      OR bs.backup_set_id IN (SELECT TOP (5) b2.backup_set_id FROM msdb.dbo.backupset b2
                              WHERE b2.database_name = @db AND b2.type = 'L' AND b2.backup_finish_date > @stop $f2
                              ORDER BY b2.backup_finish_date))
ORDER BY bs.first_lsn, bs.backup_finish_date;
"@
        $rows = Invoke-Sql $Conn $q @{ db = $Db; stop = $StopAt; startLsn = $base.LastLsn }
        $cands = @($rows.Rows | ForEach-Object { ConvertTo-BackupSet $_ } | Where-Object { -not $_.IsDamaged })
        $cur = $base.LastLsn; $prev = $base
        while ($true) {
            $next = $cands | Where-Object { $_.FirstLsn -le $cur -and $_.LastLsn -gt $cur } |
                    Sort-Object @{ Expression = { $_.IsCopyOnly } }, Finish | Select-Object -First 1
            if (-not $next) {
                $later = $cands | Where-Object { $_.FirstLsn -gt $cur } | Sort-Object FirstLsn | Select-Object -First 1
                if ($later) {
                    throw ("Log chain is broken: no log backup covers LSN $cur (after $($prev.Type) id $($prev.Id), finished $($prev.Finish)). " +
                           "The next log starts at LSN $($later.FirstLsn) ($($later.Finish)). Choose an earlier STOPAT or a newer FULL.")
                }
                $lastFinish = if ($logs.Count) { $logs[-1].Finish } else { $base.Finish }
                throw ("No log backup reaches STOPAT $($StopAt.ToString('yyyy-MM-dd HH:mm:ss')). The latest restorable point is about $lastFinish. " +
                       "Choose an earlier STOPAT, take/copy a newer log backup, or use mode FULL + DIFF.")
            }
            if ($prev.Type -eq 'LOG' -and $prev.LastFork -and $next.FirstFork -and ($prev.LastFork -ne $next.FirstFork)) {
                $warn.Add("Recovery fork changes between log $($prev.Id) and log $($next.Id): make sure this is the branch you want.")
            }
            $logs += $next; $cur = $next.LastLsn; $prev = $next
            if ($next.Finish -ge $StopAt) { break }
        }
        $usesStop = $true
        if ($logs[-1].HasBulk) { $warn.Add("The last log (id $($logs[-1].Id)) contains bulk-logged operations: STOPAT inside it can fail. Pick a STOPAT outside that log, or restore it fully.") }
        foreach ($l in $logs) { if ($l.IsCopyOnly) { $warn.Add("Log id $($l.Id) is COPY_ONLY; it is used because it continues the LSN chain.") } }
    }

    $sets = New-Object System.Collections.Generic.List[object]
    if ($full) { $sets.Add($full) }
    if ($diff) { $sets.Add($diff) }
    foreach ($l in $logs) { $sets.Add($l) }
    if ($sets.Count -eq 0) { throw 'Nothing to restore for this mode and STOPAT.' }
    $n = 0; foreach ($s in $sets) { $n++; $s.Step = $n }
    $last = $sets[$sets.Count - 1]
    if ($eff -eq 'FullOnly' -or $eff -eq 'FullDiff' -or $eff -eq 'ContinueDiff') { $warn.Add("Mode $eff restores to the end of the last backup ($($last.Finish)); STOPAT is only used to choose the backups.") }
    return [pscustomobject]@{
        Sets = $sets; Warnings = $warn; Mode = $eff; UsesStopAt = $usesStop; StopAt = $StopAt
        RestorePoint = $(if ($usesStop) { $StopAt } else { $last.Finish })
    }
}

function Get-ChainSummary {
    param($Chain)
    $parts = foreach ($t in 'FULL', 'DIFF', 'LOG') { $k = @($Chain.Sets | Where-Object { $_.Type -eq $t }).Count; if ($k) { "$k $t" } }
    return (@($parts) -join ', ')
}

function Get-BackupFiles {
    param([hashtable]$Conn, $Chain)
    if ($Chain.Sets[0].Type -ne 'FULL') { return [pscustomobject]@{ Files = @(); AddedLater = @() } }
    $fullId = $Chain.Sets[0].Id
    $t = Invoke-Sql $Conn @"
SELECT bf.logical_name, bf.physical_name, bf.file_type, bf.file_number, bf.file_size, bf.filegroup_name
FROM msdb.dbo.backupfile bf
WHERE bf.backup_set_id = @id AND bf.is_present = 1
ORDER BY bf.file_number;
"@ @{ id = $fullId }
    $files = foreach ($r in $t.Rows) {
        [pscustomobject]@{ LogicalName = [string]$r.logical_name; PhysicalName = [string]$r.physical_name; FileType = [string]$r.file_type
                           FileNumber = [int](DbVal0 $r.file_number); SizeBytes = [double](DbVal0 $r.file_size); FileGroup = [string](DbVal $r.filegroup_name) }
    }
    $files = @($files)
    $added = @()
    if ($Chain.Sets.Count -gt 1) {
        $ids = ($Chain.Sets | Select-Object -Skip 1 | ForEach-Object { [int]$_.Id }) -join ','
        $t2 = Invoke-Sql $Conn "SELECT DISTINCT logical_name FROM msdb.dbo.backupfile WHERE backup_set_id IN ($ids);"
        $known = @($files | ForEach-Object { $_.LogicalName })
        foreach ($r in $t2.Rows) { if ($known -notcontains [string]$r.logical_name) { $added += [string]$r.logical_name } }
    }
    return [pscustomobject]@{ Files = $files; AddedLater = $added }
}

#endregion

#region ============================ destination (read-only) ============================

function Get-DestInfo {
    param([hashtable]$Conn, [string]$Target)
    $r = (Invoke-Sql $Conn @"
SELECT @@SERVERNAME AS ServerName, SYSDATETIME() AS ServerTime,
       CAST(PARSENAME(CONVERT(varchar(32), SERVERPROPERTY('ProductVersion')), 4) AS int) AS Major,
       CONVERT(nvarchar(128), SERVERPROPERTY('Edition')) AS Edition,
       CONVERT(nvarchar(128), SERVERPROPERTY('ComputerNamePhysicalNetBIOS')) AS HostName,
       DEFAULT_DOMAIN() AS DomainName,
       IS_SRVROLEMEMBER('sysadmin') AS IsSysadmin,
       HAS_PERMS_BY_NAME(NULL, NULL, 'CREATE ANY DATABASE') AS CanCreateDb,
       CONVERT(nvarchar(512), SERVERPROPERTY('InstanceDefaultDataPath')) AS DefaultData,
       CONVERT(nvarchar(512), SERVERPROPERTY('InstanceDefaultLogPath')) AS DefaultLog,
       CONVERT(nvarchar(512), SERVERPROPERTY('ErrorLogFileName')) AS ErrorLogFile;
"@).Rows[0]
    $d = [pscustomobject]@{
        ServerName = [string]$r.ServerName; ServerTime = [datetime]$r.ServerTime; Major = [int]$r.Major; Edition = [string]$r.Edition
        HostName = [string]$r.HostName; DomainName = [string](DbVal $r.DomainName)
        IsSysadmin = ([int](DbVal0 $r.IsSysadmin) -eq 1); CanCreateDb = ([int](DbVal0 $r.CanCreateDb) -eq 1)
        DefaultData = [string](DbVal $r.DefaultData); DefaultLog = [string](DbVal $r.DefaultLog); ErrorLogFile = [string](DbVal $r.ErrorLogFile)
        ServiceAccount = $null; AgentStatus = $null; AvailMB = $null; TotalMB = $null
        TargetExists = $false; TargetState = $null; TargetInAg = $false
        Certs = @(); HasDmk = $false; Credentials = @(); UserFiles = @(); Volumes = @(); FixedDrives = @(); Drives = @()
        TargetRedoLsn = $null; TargetDiffBaseLsn = $null; TargetLastRestore = $null
    }
    try {
        foreach ($s in (Invoke-Sql $Conn 'SELECT servicename, service_account, status_desc FROM sys.dm_server_services;').Rows) {
            if ([string]$s.servicename -like 'SQL Server (*') { $d.ServiceAccount = [string]$s.service_account }
            if ([string]$s.servicename -like 'SQL Server Agent*') { $d.AgentStatus = [string]$s.status_desc }
        }
    } catch { }
    try {
        $m = (Invoke-Sql $Conn 'SELECT available_physical_memory_kb / 1024 AS AvailMB, total_physical_memory_kb / 1024 AS TotalMB FROM sys.dm_os_sys_memory;').Rows[0]
        $d.AvailMB = [double]$m.AvailMB; $d.TotalMB = [double]$m.TotalMB
    } catch { }
    if ($Target) {
        $t = Invoke-Sql $Conn 'SELECT name, state_desc FROM sys.databases WHERE name = @t;' @{ t = $Target }
        if ($t.Rows.Count) {
            $d.TargetExists = $true; $d.TargetState = [string]$t.Rows[0].state_desc
            try {
                $a = Invoke-Sql $Conn 'SELECT CASE WHEN replica_id IS NULL THEN 0 ELSE 1 END AS in_ag FROM sys.databases WHERE name = @t;' @{ t = $Target }
                $d.TargetInAg = ([int]$a.Rows[0].in_ag -eq 1)
            } catch { }
            try {
                $rl = Invoke-Sql $Conn 'SELECT redo_start_lsn, differential_base_lsn FROM sys.master_files WHERE database_id = DB_ID(@t) AND file_id = 1;' @{ t = $Target }
                if ($rl.Rows.Count) { $d.TargetRedoLsn = DbVal $rl.Rows[0].redo_start_lsn; $d.TargetDiffBaseLsn = DbVal $rl.Rows[0].differential_base_lsn }
            } catch { }
            try {
                $lr = Invoke-Sql $Conn @"
SELECT TOP (1) rh.restore_date, rh.restore_type, bs.backup_finish_date
FROM msdb.dbo.restorehistory rh LEFT JOIN msdb.dbo.backupset bs ON bs.backup_set_id = rh.backup_set_id
WHERE rh.destination_database_name = @t ORDER BY rh.restore_history_id DESC;
"@ @{ t = $Target }
                if ($lr.Rows.Count) {
                    $x = $lr.Rows[0]
                    $rt = switch ([string]$x.restore_type) { 'D' { 'FULL' } 'I' { 'DIFF' } 'L' { 'LOG' } default { [string]$x.restore_type } }
                    $d.TargetLastRestore = '{0} restored {1:yyyy-MM-dd hh:mm:ss tt} (backup finished {2})' -f $rt, ([datetime]$x.restore_date), $(if ($x.backup_finish_date -is [datetime]) { $x.backup_finish_date.ToString('yyyy-MM-dd hh:mm:ss tt') } else { '?' })
                }
            } catch { }
        }
    }
    $d.Certs = @((Invoke-Sql $Conn 'SELECT name, CONVERT(varchar(64), thumbprint, 1) AS thumb FROM master.sys.certificates;').Rows |
                 ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Thumb = [string]$_.thumb } })
    $d.HasDmk = ([int](Invoke-Sql $Conn "SELECT COUNT(*) AS n FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##';").Rows[0].n -gt 0)
    $d.Credentials = @((Invoke-Sql $Conn 'SELECT name, credential_identity FROM sys.credentials;').Rows |
                       ForEach-Object { [pscustomobject]@{ Name = [string]$_.name; Identity = [string]$_.credential_identity } })
    $d.UserFiles = @((Invoke-Sql $Conn @"
SELECT d.name AS db, mf.type, mf.file_id, mf.physical_name
FROM sys.master_files mf JOIN sys.databases d ON d.database_id = mf.database_id
WHERE mf.database_id > 4;
"@).Rows | ForEach-Object { [pscustomobject]@{ Db = [string]$_.db; Type = [int]$_.type; FileId = [int]$_.file_id; Path = [string]$_.physical_name } })
    if (-not $d.DefaultData -or -not $d.DefaultLog) {
        foreach ($mf in (Invoke-Sql $Conn 'SELECT type, physical_name FROM sys.master_files WHERE database_id = 3;').Rows) {
            if ([int]$mf.type -eq 0 -and -not $d.DefaultData) { $d.DefaultData = Get-DirOf ([string]$mf.physical_name) }
            if ([int]$mf.type -eq 1 -and -not $d.DefaultLog)  { $d.DefaultLog  = Get-DirOf ([string]$mf.physical_name) }
        }
    }
    try {
        $d.Volumes = @((Invoke-Sql $Conn @"
SELECT DISTINCT vs.volume_mount_point, vs.available_bytes, vs.total_bytes
FROM sys.master_files mf CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs;
"@).Rows | ForEach-Object { [pscustomobject]@{ Mount = [string]$_.volume_mount_point; AvailBytes = [double]$_.available_bytes } })
    } catch { }
    try {
        $d.FixedDrives = @((Invoke-Sql $Conn 'EXEC master.dbo.xp_fixeddrives;').Rows |
                           ForEach-Object { [pscustomobject]@{ Drive = [string]$_[0]; FreeMB = [double]$_[1] } })
    } catch { }
    try { $d.Drives = @(Get-DestDrives $Conn) } catch { }
    return $d
}

# Drives and mount points on the destination with free/total space (read-only DMVs; xp_fixeddrives fallback)
function Get-DestDrives {
    param([hashtable]$Conn)
    $list = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    try {
        foreach ($r in (Invoke-Sql $Conn 'SELECT fixed_drive_path, free_space_in_bytes FROM sys.dm_os_enumerate_fixed_drives;').Rows) {
            $root = [string]$r.fixed_drive_path; if (-not $root.EndsWith('\')) { $root += '\' }
            if ($seen.Add($root)) { $list.Add([pscustomobject]@{ Root = $root; FreeBytes = [double]$r.free_space_in_bytes; TotalBytes = $null }) }
        }
    } catch {
        try {
            foreach ($r in (Invoke-Sql $Conn 'EXEC master.dbo.xp_fixeddrives;').Rows) {
                $root = ([string]$r[0]) + ':\'
                if ($seen.Add($root)) { $list.Add([pscustomobject]@{ Root = $root; FreeBytes = [double]$r[1] * 1MB; TotalBytes = $null }) }
            }
        } catch { }
    }
    try {
        $q = 'SELECT DISTINCT vs.volume_mount_point, vs.available_bytes, vs.total_bytes FROM sys.master_files mf CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs;'
        foreach ($r in (Invoke-Sql $Conn $q).Rows) {
            $root = [string]$r.volume_mount_point
            $ex = $null
            foreach ($e in $list) { if ($e.Root -eq $root) { $ex = $e } }
            if ($ex) { $ex.FreeBytes = [double]$r.available_bytes; $ex.TotalBytes = [double]$r.total_bytes }
            elseif ($seen.Add($root)) { $list.Add([pscustomobject]@{ Root = $root; FreeBytes = [double]$r.available_bytes; TotalBytes = [double]$r.total_bytes }) }
        }
    } catch { }
    return @($list | Sort-Object Root)
}

# Subfolders of a destination folder, as the SQL Server service account sees them
function Get-DestSubfolders {
    param([hashtable]$Conn, [string]$Path)
    $names = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($r in (Invoke-Sql $Conn 'EXEC master.sys.xp_dirtree @p, 1, 0;' @{ p = $Path }).Rows) { $names.Add([string]$r.subdirectory) }
    } catch {
        foreach ($r in (Invoke-Sql $Conn "SELECT file_or_directory_name FROM sys.dm_os_enumerate_filesystem(@p, N'*') WHERE is_directory = 1;" @{ p = $Path }).Rows) { $names.Add([string]$r.file_or_directory_name) }
    }
    return @($names | Sort-Object)
}

function Get-VolumeFor {
    param([string]$Path, $Dest)
    $best = $null
    foreach ($v in $Dest.Volumes) {
        if ($Path.StartsWith($v.Mount, [StringComparison]::OrdinalIgnoreCase) -and (-not $best -or $v.Mount.Length -gt $best.Mount.Length)) { $best = $v }
    }
    if ($best) { return $best }
    if ($Path -match '^([A-Za-z]):') {
        $drv = $Matches[1]
        $fd = $Dest.FixedDrives | Where-Object { $_.Drive -eq $drv } | Select-Object -First 1
        if ($fd) { return [pscustomobject]@{ Mount = "$($drv):\"; AvailBytes = $fd.FreeMB * 1MB } }
    }
    return $null
}

#endregion

#region ============================ WITH MOVE plan ============================

function Get-MovePlan {
    param($Files, $Dest, [string]$TargetDb, [string]$Strategy, [string]$DataDir, [string]$NdfDir, [string]$LogDir, [string]$Pattern)
    $dData = $null; $dLog = $null; $dNdf = $null; $whyD = ''; $whyL = ''; $whyN = ''
    $userFiles = @($Dest.UserFiles | Where-Object { $_.Db -ne $TargetDb })

    if ($Strategy -like 'Manual*' -and (-not $DataDir -or -not $LogDir)) { throw 'MOVE strategy "Manual": enter at least the data (mdf) folder and the log (ldf) folder.' }

    if ($Strategy -like 'Largest*') {
        $big = @($Dest.Drives | Sort-Object FreeBytes -Descending)
        if ($big.Count) {
            $dv = $big[0]
            $why = 'largest free disk {0} ({1:N0} GB free)' -f $dv.Root, ($dv.FreeBytes / 1GB)
            $onIt = @($userFiles | Where-Object { $_.Path.StartsWith($dv.Root, [StringComparison]::OrdinalIgnoreCase) })
            $g = $onIt | Where-Object { $_.Type -eq 0 } | Group-Object { Get-DirOf $_.Path } | Sort-Object Count -Descending | Select-Object -First 1
            if ($g) { $dData = $g.Name; $whyD = "$why, folder already used for data files" } else { $dData = $dv.Root; $whyD = "$why, drive ROOT - pick a subfolder with the ... button" }
            $g = $onIt | Where-Object { $_.Type -eq 1 } | Group-Object { Get-DirOf $_.Path } | Sort-Object Count -Descending | Select-Object -First 1
            if ($g) { $dLog = $g.Name; $whyL = "$why, folder already used for log files" } else { $dLog = $dData; $whyL = $whyD }
        }
    }
    if ($Strategy -like 'Auto*') {
        if ($Pattern) {
            $sim = @($userFiles | Where-Object { $_.Db -like "*$Pattern*" })
            if ($sim.Count) {
                $simDb = ($sim | Sort-Object { $_.Db.Length }, Db | Select-Object -First 1).Db
                $p = $sim | Where-Object { $_.Db -eq $simDb -and $_.Type -eq 0 -and $_.FileId -eq 1 } | Select-Object -First 1
                $l = $sim | Where-Object { $_.Db -eq $simDb -and $_.Type -eq 1 } | Select-Object -First 1
                $n = $sim | Where-Object { $_.Db -eq $simDb -and $_.Type -eq 0 -and $_.FileId -ne 1 } | Select-Object -First 1
                if ($p) { $dData = Get-DirOf $p.Path; $whyD = "similar database [$simDb] on the destination" }
                if ($l) { $dLog  = Get-DirOf $l.Path; $whyL = "similar database [$simDb] on the destination" }
                if ($n) { $dNdf  = Get-DirOf $n.Path; $whyN = "similar database [$simDb] on the destination" }
            }
        }
        if (-not $dData) {
            $g = $userFiles | Where-Object { $_.Type -eq 0 } | Group-Object { Get-DirOf $_.Path } | Sort-Object Count -Descending | Select-Object -First 1
            if ($g) { $dData = $g.Name; $whyD = "folder used by most user-database data files ($($g.Count)) on the destination" }
        }
        if (-not $dLog) {
            $g = $userFiles | Where-Object { $_.Type -eq 1 } | Group-Object { Get-DirOf $_.Path } | Sort-Object Count -Descending | Select-Object -First 1
            if ($g) { $dLog = $g.Name; $whyL = "folder used by most user-database log files ($($g.Count)) on the destination" }
        }
    }
    if (-not $dData) { $dData = $Dest.DefaultData; $whyD = 'instance default data path' }
    if (-not $dLog)  { $dLog  = $Dest.DefaultLog;  $whyL = 'instance default log path' }
    if ($DataDir) { $dData = $DataDir; $whyD = 'entered manually' }
    if ($LogDir)  { $dLog  = $LogDir;  $whyL = 'entered manually' }
    if ($NdfDir)  { $dNdf  = $NdfDir;  $whyN = 'entered manually' }
    if (-not $dNdf) { $dNdf = $dData; $whyN = $whyD }
    if (-not $dData -or -not $dLog) { throw 'Could not work out destination folders. Enter the data and log folders manually.' }

    $used = @{}
    foreach ($u in $Dest.UserFiles) { $used[$u.Path.ToLower()] = $u.Db }
    $stamp = Get-Date -Format 'yyyyMMddHHmm'

    $plan = foreach ($f in $Files) {
        $leaf = Get-Leaf $f.PhysicalName
        $ext = ''; $dot = $leaf.LastIndexOf('.'); if ($dot -ge 0) { $ext = $leaf.Substring($dot) }
        switch ($f.FileType) {
            'L'     { $dir = $dLog;  $why = $whyL; if (-not $ext) { $ext = '.ldf' }; $kind = 'LDF' }
            'S'     { $dir = $dData; $why = $whyD; $ext = ''; $kind = 'FILESTREAM/IMOLTP folder' }
            'F'     { $dir = $dData; $why = $whyD; $kind = 'FULLTEXT' }
            default {
                if ($f.FileNumber -eq 1) { $dir = $dData; $why = $whyD; if (-not $ext) { $ext = '.mdf' }; $kind = 'MDF' }
                else                     { $dir = $dNdf;  $why = $whyN; if (-not $ext) { $ext = '.ndf' }; $kind = 'NDF' }
            }
        }
        $base = ConvertTo-SafeName ($TargetDb + '_' + $f.LogicalName)
        $target = Join-WinPath $dir ($base + $ext)
        $comment = "REVIEW - folder from $why"
        $owner = $used[$target.ToLower()]
        if ($owner -and $owner -ne $TargetDb) {
            $target = Join-WinPath $dir ($base + '_' + $stamp + $ext)
            $comment += "; renamed, original name is used by [$owner]"
        }
        [pscustomobject]@{ LogicalName = $f.LogicalName; Kind = $kind; SizeMB = [math]::Round($f.SizeBytes / 1MB, 0)
                           SourcePath = $f.PhysicalName; TargetPath = $target; Computed = $target; Comment = $comment }
    }
    return @($plan)
}

#endregion

#region ============================ path resolution & checks ============================

function Resolve-BackupPath {
    param($Device, $Opt)
    $path = $Device.Path
    $isUrl = ($Device.DeviceType -eq 9) -or ($path -like 'http*://*')
    if ($Opt.BlobMode -and -not $isUrl) {
        $ct = if ($Opt.ContainerUrl) { $Opt.ContainerUrl } else { 'https://<account>.blob.core.windows.net/<container>' }
        return [pscustomobject]@{ Kind = 'URL'; Path = "$ct/$($Opt.Db)/$(Get-Leaf $path)"; Original = $path }
    }
    if ($Opt.RemapFrom -and $path.StartsWith($Opt.RemapFrom, [StringComparison]::OrdinalIgnoreCase)) {
        $new = $Opt.RemapTo + $path.Substring($Opt.RemapFrom.Length)
        if ($new -like 'http*') { return [pscustomobject]@{ Kind = 'URL'; Path = $new.Replace('\', '/'); Original = $path } }
        return [pscustomobject]@{ Kind = 'DISK'; Path = $new; Original = $path }
    }
    if ($Opt.AutoUnc -and -not $isUrl -and $path -match '^[A-Za-z]:\\') {
        $dir = Get-DirOf $path
        $m = $script:State.ShareMap | Where-Object { $_.Local -eq $dir } | Select-Object -First 1
        if ($m) { return [pscustomobject]@{ Kind = 'DISK'; Path = (Join-WinPath $m.Unc (Get-Leaf $path)); Original = $path } }
    }
    return [pscustomobject]@{ Kind = $(if ($isUrl) { 'URL' } else { 'DISK' }); Path = $path; Original = $path }
}

# Lists disk shares on the source host (read-only WMI/CIM query: WSMan first, then DCOM)
function Get-SourceShares {
    param([string]$HostName)
    $err = $null
    foreach ($proto in 'Wsman', 'Dcom') {
        try {
            $so = New-CimSessionOption -Protocol $proto
            $cs = New-CimSession -ComputerName $HostName -SessionOption $so -OperationTimeoutSec 20 -ErrorAction Stop
            try {
                $sh = @(Get-CimInstance -CimSession $cs -ClassName Win32_Share -ErrorAction Stop |
                        Where-Object { $_.Path -and ([uint32]$_.Type -eq 0 -or [uint32]$_.Type -eq 2147483648) } |
                        ForEach-Object { [pscustomobject]@{ Name = [string]$_.Name; Path = ([string]$_.Path).TrimEnd('\'); IsAdmin = ([uint32]$_.Type -ne 0) } })
            } finally { Remove-CimSession $cs -ErrorAction SilentlyContinue }
            if ($sh.Count) { return $sh }
        } catch { $err = "$proto : $($_.Exception.Message)" }
    }
    throw "Could not list the shares on $HostName ($err)."
}

# For every LOCAL backup folder in the chain, find the share on the source server that exposes it
function Update-ShareMap {
    $script:State.ShareMap = @()
    if (-not $script:Opt.AutoUnc) { return }
    $src = $script:State.Source
    $local = @(foreach ($s in $script:State.Chain.Sets) { foreach ($d in $s.Devices) { if ($d.DeviceType -ne 9 -and $d.Path -match '^[A-Za-z]:\\') { $d.Path } } })
    if (-not $local.Count) { return }
    $dirs = @($local | ForEach-Object { Get-DirOf $_ } | Sort-Object -Unique)
    Write-Log "$($local.Count) backup file(s) are recorded in msdb with a LOCAL path on $($src.HostName) (e.g. $($local[0])). Checking whether that folder is already shared on $($src.HostName) (read-only lookup; no share is created or changed)..." 'INFO'
    try { $shares = Get-SourceShares $src.HostName }
    catch {
        Write-Log "$($_.Exception.Message) Fill in Remap from/to on tab 2 instead, e.g. from $($dirs[0])\ to \\$($src.HostName)\<share>\" 'WARN'
        return
    }
    # 1) Prefer ONE regular share that covers every backup folder in the chain: the deepest such share
    #    (e.g. [bkup] on C:\Apps\BKUP rather than [C] on C:\ or [FULL] on ...\FULL, which misses DIFF/LOG).
    $regular = @($shares | Where-Object { -not $_.IsAdmin })
    $common = $null
    foreach ($sh in ($regular | Sort-Object { $_.Path.Length } -Descending)) {
        $missing = 0
        foreach ($dir in $dirs) { if (-not ($dir + '\').StartsWith($sh.Path + '\', [StringComparison]::OrdinalIgnoreCase)) { $missing++ } }
        if ($missing -eq 0) { $common = $sh; break }
    }
    if ($common) { Write-Log "Existing share [$($common.Name)] ($($common.Path)) covers all $($dirs.Count) backup folder(s) in this chain; using it for every file." 'OK' }
    else { Write-Log 'No single regular share covers every backup folder; choosing the most specific share per folder.' 'WARN' }

    $map = @()
    foreach ($dir in $dirs) {
        $best = $common
        if (-not $best) {
            # 2) per folder: the most specific regular share, then (last resort) an admin share
            $best = $shares | Where-Object { ($dir + '\').StartsWith($_.Path + '\', [StringComparison]::OrdinalIgnoreCase) } |
                    Sort-Object @{ Expression = { $_.IsAdmin }; Ascending = $true }, @{ Expression = { $_.Path.Length }; Descending = $true } |
                    Select-Object -First 1
        }
        if (-not $best) { Write-Log "No share on $($src.HostName) covers $dir. Share that folder (read for the destination SQL service account) or use Remap from/to." 'WARN'; continue }
        $unc = "\\$($src.HostName)\$($best.Name)" + $dir.Substring($best.Path.Length)
        $map += [pscustomobject]@{ Local = $dir; Unc = $unc; Share = $best.Name; IsAdmin = $best.IsAdmin }
        if ($best.IsAdmin) { Write-Log "No regular share covers $dir; only the admin share $($best.Name) does -> $unc. The destination SQL service account would need admin rights on $($src.HostName). Better: use Remap from/to with the share you want." 'WARN' }
        else { Write-Log "  $dir -> $unc" 'OK' }
    }
    $script:State.ShareMap = $map
}

function Get-FromClause {
    param($Set, $Opt)
    return (@($Set.Devices | ForEach-Object { $r = Resolve-BackupPath $_ $Opt; "$($r.Kind) = $(QS $r.Path)" }) -join ",`r`n     ")
}

function Get-ErrorHint {
    param([string]$Msg, [string]$SvcAcct)
    $h = [pscustomobject]@{ Kind = 'Other'; Thumb = $null; Fix = 'Read the full message; try the same RESTORE HEADERONLY manually on the destination.' }
    if ($Msg -match "thumbprint\s*'?(0x[0-9A-Fa-f]+)") {
        $h.Kind = 'Cert'; $h.Thumb = $Matches[1]
        $h.Fix = 'The destination is missing the certificate that encrypts this backup (TDE or backup encryption). Use Execute fixes or 03_tde_certificate_transfer.sql.'
    }
    elseif ($Msg -match 'error 5\(|Access is denied') {
        $h.Kind = 'Access'
        $h.Fix = "Grant the destination SQL Server service account [$SvcAcct] read on the share AND NTFS, or copy the files to blob (05_azcopy_to_blob.ps1) and tick 'Files were copied to blob'."
    }
    elseif ($Msg -match 'error (53|64|67|1231|1311|1326)\(|network path|network name') {
        $h.Kind = 'Network'
        $h.Fix = "The destination cannot reach this path (DMZ / Azure / firewall). Copy the files to blob with 05_azcopy_to_blob.ps1, tick 'Files were copied to blob', and run preflight again."
    }
    elseif ($Msg -match 'error (2|3)\(|cannot find the (file|path)') {
        $h.Kind = 'Missing'
        $h.Fix = 'File not found from the destination. If the path is local to the source (C:\...), the destination looks on its own drive: tick the share conversion option or use Remap from/to. Otherwise the file was removed by retention or moved.'
    }
    elseif ($Msg -match '(?i)credential|shared access|error 86\(|\b403\b|authenticat|not authorized') {
        $h.Kind = 'Credential'
        $h.Fix = 'Create or refresh the SAS credential on the destination (04_blob_credential.sql). The SAS needs Read + List and must not be expired.'
    }
    elseif ($Msg -match '(?i)version.*(incompatible|not supported)|cannot be restored on this server') {
        $h.Kind = 'Version'; $h.Fix = 'The backup is from a newer SQL Server version than the destination.'
    }
    elseif ($Msg -match '(?i)media family|media set') {
        $h.Kind = 'Stripe'; $h.Fix = 'A stripe of this media set is missing or wrong: every file of the backup set is needed.'
    }
    elseif ($Msg -match '(?i)\b404\b|blob') {
        $h.Kind = 'Blob'; $h.Fix = 'Check the blob URL (container and blob name) and that the blob exists.'
    }
    return $h
}

function Add-Check {
    param([string]$Status, [string]$Name, [string]$Detail, [string]$Fix = '')
    $script:State.Checks.Add([pscustomobject]@{ Status = $Status; Check = $Name; Detail = $Detail; Fix = $Fix })
    $lvl = @{ PASS = 'OK'; WARN = 'WARN'; FAIL = 'FAIL'; INFO = 'INFO' }[$Status]
    $txt = "[$Status] $Name - $Detail"
    if ($Fix) { $txt += "  >> FIX: $Fix" }
    Write-Log $txt $lvl
}

function Get-RequiredCerts {
    $o = $script:Opt; $src = $script:State.Source; $chain = $script:State.Chain
    $need = New-Object System.Collections.Generic.List[object]
    if ($src -and $src.TdeCertThumb) { $need.Add([pscustomobject]@{ Name = $src.TdeCertName; Thumb = $src.TdeCertThumb; Reason = 'TDE' }) }
    $thumbs = @($chain.Sets | Where-Object { $_.EncryptorThumb } | ForEach-Object { $_.EncryptorThumb } | Sort-Object -Unique)
    foreach ($t in $script:State.ErrorThumbs) { if ($thumbs -notcontains $t) { $thumbs += $t } }
    foreach ($t in $thumbs) {
        if (@($need | Where-Object { $_.Thumb -eq $t }).Count) { continue }
        $name = $null; try { $name = Get-SourceCertByThumb $o.Src $t } catch { }
        $need.Add([pscustomobject]@{ Name = $name; Thumb = $t; Reason = 'Backup encryption / found in restore error' })
    }
    return $need
}

function Get-MissingContainers {
    $o = $script:Opt; $d = $script:State.Dest
    $urls = @(foreach ($s in $script:State.Chain.Sets) { foreach ($dev in $s.Devices) { $r = Resolve-BackupPath $dev $o; if ($r.Kind -eq 'URL') { $r.Path } } })
    $containers = @($urls | ForEach-Object { Get-ContainerUrl $_ } | Sort-Object -Unique)
    $missing = @()
    foreach ($ct in $containers) {
        $m = @($d.Credentials | Where-Object { $ct.StartsWith($_.Name.TrimEnd('/'), [StringComparison]::OrdinalIgnoreCase) -or
                                               $_.Name.TrimEnd('/').StartsWith($ct, [StringComparison]::OrdinalIgnoreCase) })
        if (-not $m.Count) { $missing += $ct }
    }
    return [pscustomobject]@{ All = $containers; Missing = $missing }
}

function Invoke-Preflight {
    $o = $script:Opt; $c = $script:State.Chain; $d = $script:State.Dest
    $script:State.Checks = New-Object System.Collections.Generic.List[object]
    $script:State.AccessFailures = New-Object System.Collections.Generic.List[string]
    $script:State.ErrorThumbs = @()
    $n = $c.Sets.Count

    Add-Check 'INFO' 'Restore chain' ("Mode {0}: {1} set(s): {2}. Restore point: {3:yyyy-MM-dd HH:mm:ss}" -f $c.Mode, $n, (Get-ChainSummary $c), $c.RestorePoint)
    foreach ($w in $c.Warnings) { Add-Check 'WARN' 'Chain note' $w }
    Add-Check 'PASS' 'Destination connection' "$($d.ServerName) (host $($d.HostName), major v$($d.Major), $($d.Edition)); SQL service account: $($d.ServiceAccount)"

    $srcMajor = ($c.Sets | Measure-Object SwMajor -Maximum).Maximum
    if ($d.Major -lt $srcMajor) { Add-Check 'FAIL' 'Version' "Backups are from major v$srcMajor, destination is v$($d.Major). A backup can't be restored to an older version." 'Use a destination at the same or a newer version.' }
    else { Add-Check 'PASS' 'Version' "backup v$srcMajor -> destination v$($d.Major)" }

    if ($d.IsSysadmin -or $d.CanCreateDb) { Add-Check 'PASS' 'Permissions' 'CREATE ANY DATABASE / sysadmin present (needed for HEADERONLY and RESTORE).' }
    else { Add-Check 'FAIL' 'Permissions' 'Your login lacks CREATE ANY DATABASE on the destination.' 'Use a login in dbcreator or sysadmin.' }

    if ($c.Mode -like 'Continue*') {
        if (-not $d.TargetExists) { Add-Check 'FAIL' 'Target database' "[$($o.Target)] does not exist on the destination. Continue modes need a database that is already RESTORING." 'Use the Restoring DBs button on tab 1, or restore a FULL first.' }
        elseif ($d.TargetState -ne 'RESTORING') { Add-Check 'FAIL' 'Target database' "[$($o.Target)] is $($d.TargetState), not RESTORING, so no DIFF or LOG can be applied." 'It was recovered already. Start again with a FULL (WITH REPLACE or a new name).' }
        else { Add-Check 'PASS' 'Target database' "[$($o.Target)] is RESTORING; next restore starts at LSN $($d.TargetRedoLsn). Last restore: $($d.TargetLastRestore)" }
    }
    elseif ($d.TargetExists) {
        if ($d.TargetInAg) { Add-Check 'FAIL' 'Target database' "[$($o.Target)] exists and is in an availability group." 'Remove it from the AG first, or restore under another name.' }
        elseif ($o.Replace) { Add-Check 'WARN' 'Target database' "[$($o.Target)] exists ($($d.TargetState)) and WITH REPLACE will overwrite it." }
        else { Add-Check 'FAIL' 'Target database' "[$($o.Target)] already exists ($($d.TargetState))." 'Pick another target name, or tick WITH REPLACE.' }
    } else { Add-Check 'PASS' 'Target database' "[$($o.Target)] does not exist yet." }

    $needMB = [math]::Round($o.BufferCount * [double]$o.MaxTransfer / 1MB)
    if ($d.AvailMB) {
        if ($needMB -gt 0.25 * $d.AvailMB) { Add-Check 'WARN' 'Restore buffer memory' "BUFFERCOUNT $($o.BufferCount) x MAXTRANSFERSIZE $($o.MaxTransfer) = ~$needMB MB; destination has $([math]::Round($d.AvailMB)) MB free of $([math]::Round($d.TotalMB)) MB." 'Lower BUFFERCOUNT, or restore when the box is quiet.' }
        else { Add-Check 'PASS' 'Restore buffer memory' "~$needMB MB needed; $([math]::Round($d.AvailMB)) MB free." }
    }

    # blob / credentials
    $ctInfo = Get-MissingContainers
    if ($ctInfo.All.Count) {
        if ($d.Major -lt 13) { Add-Check 'FAIL' 'Restore from URL' 'RESTORE FROM URL with a SAS credential needs SQL Server 2016+ on the destination.' }
        foreach ($ct in $ctInfo.All) {
            if ($ctInfo.Missing -contains $ct) { Add-Check 'FAIL' 'Blob credential' "No credential named after $ct on the destination." 'Run 04_blob_credential.sql, or Execute fixes with a SAS token.' }
            else {
                $cr = $d.Credentials | Where-Object { $ct.StartsWith($_.Name.TrimEnd('/'), [StringComparison]::OrdinalIgnoreCase) -or $_.Name.TrimEnd('/').StartsWith($ct, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
                if ($cr.Identity -ne 'SHARED ACCESS SIGNATURE') { Add-Check 'WARN' 'Blob credential' "[$($cr.Name)] uses identity '$($cr.Identity)' (storage key). The script uses SAS-style URL restores." 'Create a SAS credential for the container.' }
                else { Add-Check 'PASS' 'Blob credential' "[$($cr.Name)] (SAS) exists; the header read below proves whether the SAS still works." }
            }
        }
    }

    # local paths that were not converted to a share
    if ($script:State.Source.HostName -ne $d.HostName) {
        $stillLocal = @(foreach ($s in $c.Sets) { foreach ($dev in $s.Devices) { $r = Resolve-BackupPath $dev $o; if ($r.Kind -eq 'DISK' -and $r.Path -match '^[A-Za-z]:\\') { $r.Path } } })
        $stillLocal = @($stillLocal | Sort-Object -Unique)
        if ($stillLocal.Count) {
            Add-Check 'FAIL' 'Local backup paths' "$($stillLocal.Count) file(s) have a path local to $($script:State.Source.HostName) (e.g. $($stillLocal[0])); $($d.HostName) would look on its OWN drive." "Tick 'Convert local paths to the source share' and Build chain again, or set Remap from '$(Get-DirOf $stillLocal[0])\' to '\\$($script:State.Source.HostName)\<share>\'."
        } elseif ($script:State.ShareMap.Count) {
            Add-Check 'PASS' 'Local backup paths' ("Converted to shares: " + (($script:State.ShareMap | ForEach-Object { "$($_.Local) -> $($_.Unc)" }) -join '; '))
        }
    }

    # headers
    $i = 0
    foreach ($s in $c.Sets) {
        $i++
        Set-Status "Reading backup header $i of $n ..."
        $label = "Header $i/$n $($s.Type) id $($s.Id) ($($s.Finish.ToString('yyyy-MM-dd HH:mm')))"
        $from = Get-FromClause $s $o
        try {
            $h = Invoke-Sql $o.Dest "RESTORE HEADERONLY FROM $from;" -Timeout 600
            $row = $null
            foreach ($r in $h.Rows) { if ([int]$r.Position -eq $s.Position) { $row = $r; break } }
            if (-not $row) { Add-Check 'FAIL' $label "Position $($s.Position) is not in the file(s): the file was replaced or reused." 'Find the original file.' }
            elseif ([decimal]$row.FirstLSN -ne $s.FirstLsn -or [decimal]$row.LastLSN -ne $s.LastLsn) {
                Add-Check 'FAIL' $label "LSNs in the file ($($row.FirstLSN) - $($row.LastLSN)) differ from msdb ($($s.FirstLsn) - $($s.LastLsn)): the file was overwritten by another backup." 'Find the original file.'
            }
            else { Add-Check 'PASS' $label 'readable from the destination; LSNs match msdb.' }
        } catch {
            $msg = Get-SqlErrorText $_
            $hint = Get-ErrorHint $msg $d.ServiceAccount
            Add-Check 'FAIL' $label $msg $hint.Fix
            if ($hint.Kind -eq 'Access' -or $hint.Kind -eq 'Network') {
                foreach ($dev in $s.Devices) { if ($dev.DeviceType -ne 9 -and -not $script:State.AccessFailures.Contains($dev.Path)) { $script:State.AccessFailures.Add($dev.Path) } }
            }
            if ($hint.Thumb -and ($script:State.ErrorThumbs -notcontains $hint.Thumb)) { $script:State.ErrorThumbs += $hint.Thumb }
        }
    }

    # FILELISTONLY on the FULL (also detects a missing TDE certificate)
    $full = $c.Sets[0]
    try {
        $fl = Invoke-Sql $o.Dest "RESTORE FILELISTONLY FROM $(Get-FromClause $full $o) WITH FILE = $($full.Position);" -Timeout 600
        $names = @($fl.Rows | ForEach-Object { [string]$_.LogicalName })
        $miss = @($script:State.Files | Where-Object { $names -notcontains $_.LogicalName } | ForEach-Object { $_.LogicalName })
        if ($miss.Count) { Add-Check 'WARN' 'File list' "msdb lists files not found in FILELISTONLY: $($miss -join ', ')" }
        else { Add-Check 'PASS' 'File list' "$($names.Count) logical file(s) match msdb." }
    } catch {
        $msg = Get-SqlErrorText $_
        $hint = Get-ErrorHint $msg $d.ServiceAccount
        if ($hint.Thumb -and ($script:State.ErrorThumbs -notcontains $hint.Thumb)) { $script:State.ErrorThumbs += $hint.Thumb }
        Add-Check $(if ($hint.Kind -eq 'Cert') { 'FAIL' } else { 'WARN' }) 'File list (FILELISTONLY)' $msg $hint.Fix
    }

    # certificates
    $script:State.MissingCerts = New-Object System.Collections.Generic.List[object]
    if ($script:State.Source.IsEncrypted -and -not $script:State.Source.TdeCertThumb) {
        Add-Check 'WARN' 'TDE' 'Source database is encrypted by an asymmetric key / EKM provider. This tool cannot transfer that; set it up manually.'
    }
    foreach ($cert in (Get-RequiredCerts)) {
        if (@($d.Certs | Where-Object { $_.Thumb -eq $cert.Thumb }).Count) { Add-Check 'PASS' "Certificate ($($cert.Reason))" "[$($cert.Name)] $($cert.Thumb) exists on the destination." }
        else {
            $script:State.MissingCerts.Add($cert)
            Add-Check 'FAIL' "Certificate ($($cert.Reason))" "[$($cert.Name)] $($cert.Thumb) is NOT on the destination; the restore will fail." 'Execute fixes (needs a password), or run 03_tde_certificate_transfer.sql by hand.'
        }
    }
    if ($script:State.MissingCerts.Count -and -not $d.HasDmk) { Add-Check 'WARN' 'Database master key' 'Destination master has no database master key; one will be created (enter a DMK password).' }
    if ($script:State.Source.TdeCertThumb) { Add-Check 'INFO' 'TDE rotation' 'Backups taken before a TDE certificate rotation need the OLD certificate too. FILELISTONLY above only proves the FULL.' }

    # folders & space
    $moves = Get-MoveRows
    foreach ($dir in @($moves | ForEach-Object { Get-DirOf $_.TargetPath } | Sort-Object -Unique)) {
        try {
            $x = Invoke-Sql $o.Dest 'EXEC master.dbo.xp_fileexist @p;' @{ p = $dir }
            if ([int]$x.Rows[0][1] -eq 1) { Add-Check 'PASS' 'Target folder' "$dir exists." }
            else { Add-Check 'FAIL' 'Target folder' "$dir does not exist on the destination (RESTORE will not create it)." "Create it and give the SQL service account [$($d.ServiceAccount)] full control." }
        } catch { Add-Check 'WARN' 'Target folder' "Could not check ${dir}: $(Get-SqlErrorText $_)" }
    }
    $byVol = @{}
    foreach ($m in $moves) {
        $v = Get-VolumeFor $m.TargetPath $d
        $key = if ($v) { $v.Mount } else { '(unknown)' }
        if (-not $byVol.ContainsKey($key)) { $byVol[$key] = [pscustomobject]@{ Vol = $v; Need = 0.0 } }
        $byVol[$key].Need += [double]$m.SizeMB * 1MB
    }
    foreach ($k in $byVol.Keys) {
        $e = $byVol[$k]; $needGB = [math]::Round($e.Need / 1GB, 1)
        if (-not $e.Vol) { Add-Check 'WARN' 'Disk space' "Could not find free space for $k (needs ~$needGB GB)." 'Check it by hand.'; continue }
        $freeGB = [math]::Round($e.Vol.AvailBytes / 1GB, 1)
        if ($e.Need -gt $e.Vol.AvailBytes) { Add-Check 'FAIL' 'Disk space' "$k needs ~$needGB GB (files are created at full size, including the log), only $freeGB GB free." 'Pick another folder or free space.' }
        elseif ($e.Need -gt 0.9 * $e.Vol.AvailBytes) { Add-Check 'WARN' 'Disk space' "$k needs ~$needGB GB of $freeGB GB free (over 90%)." }
        else { Add-Check 'PASS' 'Disk space' "$k needs ~$needGB GB, $freeGB GB free." }
    }
    if ($script:State.Files.Count -and $script:State.AddedLater.Count) {
        Add-Check 'WARN' 'Files added after the FULL' "Files created after the FULL: $($script:State.AddedLater -join ', '). The log restore recreates them at their ORIGINAL path; add MOVE for them on that step if the path is invalid on the destination."
    }
    if ($d.AgentStatus -and $d.AgentStatus -ne 'Running') { Add-Check 'WARN' 'SQL Agent' "SQL Server Agent is $($d.AgentStatus) on the destination: a scheduled job would not run." }

    $f = @($script:State.Checks | Where-Object Status -eq 'FAIL').Count
    $w = @($script:State.Checks | Where-Object Status -eq 'WARN').Count
    return [pscustomobject]@{ Fail = $f; Warn = $w }
}

#endregion

#region ============================ script generation ============================

function Get-MoveRows {
    $t = $script:State.MoveTable
    if (-not $t) { return @() }
    return @(foreach ($r in $t.Rows) {
        $c = [string]$r.Comment
        if ([string]$r.TargetPath -ne [string]$r.Computed) { $c = 'edited by user' }
        [pscustomobject]@{ LogicalName = [string]$r.LogicalName; Kind = [string]$r.Kind; SizeMB = [double]$r.SizeMB; TargetPath = [string]$r.TargetPath; Comment = $c }
    })
}

function Join-WithItems {
    param($Items)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $line = '     ' + $Items[$i].Text
        if ($i -lt $Items.Count - 1) { $line += ',' }
        if ($Items[$i].Comment) { $line += '   -- ' + $Items[$i].Comment }
        [void]$sb.AppendLine($line)
    }
    return $sb.ToString().TrimEnd()
}

function New-RestoreStatements {
    param($Chain, $Moves, $Opt)
    $list = New-Object System.Collections.Generic.List[object]
    $n = $Chain.Sets.Count; $i = 0
    foreach ($s in $Chain.Sets) {
        $i++
        $kw = if ($s.Type -eq 'LOG') { 'LOG' } else { 'DATABASE' }
        $items = New-Object System.Collections.Generic.List[object]
        $items.Add([pscustomobject]@{ Text = "FILE = $($s.Position)"; Comment = '' })
        if ($s.Type -eq 'FULL') {
            foreach ($m in $Moves) { $items.Add([pscustomobject]@{ Text = "MOVE $(QS $m.LogicalName) TO $(QS $m.TargetPath)"; Comment = "$($m.Kind): $($m.Comment)" }) }
            if ($Opt.Replace) { $items.Add([pscustomobject]@{ Text = 'REPLACE'; Comment = 'overwrites the existing database' }) }
        }
        $items.Add([pscustomobject]@{ Text = 'NORECOVERY'; Comment = '' })
        if ($s.Type -eq 'LOG' -and $i -eq $n -and $Chain.UsesStopAt) {
            $items.Add([pscustomobject]@{ Text = "STOPAT = N'$($Chain.StopAt.ToString('yyyy-MM-ddTHH:mm:ss.fff'))'"; Comment = 'source server local time' })
        }
        $items.Add([pscustomobject]@{ Text = "STATS = $($Opt.Stats)"; Comment = '' })
        $items.Add([pscustomobject]@{ Text = "BUFFERCOUNT = $($Opt.BufferCount)"; Comment = '' })
        $items.Add([pscustomobject]@{ Text = "MAXTRANSFERSIZE = $($Opt.MaxTransfer)"; Comment = '' })

        $copy = if ($s.IsCopyOnly) { ' COPY_ONLY' } else { '' }
        $header = "-- Step $i/$n : $($s.Type)$copy  backup_set_id=$($s.Id)  finished $($s.Finish.ToString('yyyy-MM-dd HH:mm:ss'))  LSN $($s.FirstLsn) -> $($s.LastLsn)  ($($s.SizeMB) MB)"
        foreach ($dev in $s.Devices) { $r = Resolve-BackupPath $dev $Opt; if ($r.Path -ne $r.Original) { $header += "`r`n--   msdb path: $($r.Original)  ->  read from: $($r.Path)" } }
        $sql = "RESTORE $kw $(QN $Opt.Target)`r`nFROM $(Get-FromClause $s $Opt)`r`nWITH`r`n$(Join-WithItems $items);"
        $list.Add([pscustomobject]@{
            Header = $header; Sql = $sql; Type = $s.Type
            Print = "Step $i/$n $($s.Type) id $($s.Id) ($($s.Finish.ToString('yyyy-MM-dd HH:mm:ss')))"
            StepName = ('{0:00} {1} {2:yyyy-MM-dd HH:mm} id {3}' -f $i, $s.Type, $s.Finish, $s.Id)
        })
    }
    return $list
}

function New-Banner {
    param([string]$Title)
    $o = $script:Opt; $c = $script:State.Chain; $d = $script:State.Dest
    $counts = Get-ChainSummary $c
    $pf = 'NOT RUN'
    if ($script:State.Checks) {
        $nf = @($script:State.Checks | Where-Object { $_.Status -eq 'FAIL' }).Count
        $nw = @($script:State.Checks | Where-Object { $_.Status -eq 'WARN' }).Count
        $pf = "$nf FAIL, $nw WARN (see 00_preflight_report.txt)"
    }
    $final = if ($o.WithRecovery) { 'RECOVERY (database comes ONLINE after the last step)' } else { 'NORECOVERY (database stays RESTORING; recover manually)' }
    $stop = if ($c.UsesStopAt) { $c.StopAt.ToString('yyyy-MM-dd HH:mm:ss') + ' (source server local time)' } else { "not used (mode $($c.Mode)); restore point ~ $($c.RestorePoint)" }
    $mem = [math]::Round($o.BufferCount * [double]$o.MaxTransfer / 1MB)
@"
/* ==========================================================================================
   $Title
   Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME ($script:AppName $script:Version)
   Source      : $($script:State.Source.ServerName)  database [$($o.Db)]   (from msdb history)
   Destination : $($d.ServerName)  database [$($o.Target)]
   Mode        : $($c.Mode)   chain: $counts
   STOPAT      : $stop
   Final state : $final
   Options     : STATS = $($o.Stats), BUFFERCOUNT = $($o.BufferCount), MAXTRANSFERSIZE = $($o.MaxTransfer) (~$mem MB buffers)$(if ($o.Replace) { ', REPLACE' })
   Preflight   : $pf
   REVIEW every MOVE target path before running.
   ========================================================================================== */
"@
}

function New-SsmsScript {
    param($Statements)
    $o = $script:Opt; $d = $script:State.Dest
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((New-Banner 'POINT-IN-TIME RESTORE SCRIPT (run in SSMS on the DESTINATION)'))
    [void]$sb.AppendLine('SET NOCOUNT ON;')
    $srvMsg = ("STOP: connected to %s, but this script is for {0}." -f $d.ServerName).Replace('%s', '#S#').Replace('%', '%%').Replace('#S#', '%s')
    [void]$sb.AppendLine("IF @@SERVERNAME <> $(QS $d.ServerName) BEGIN RAISERROR($(QS $srvMsg), 16, 1, @@SERVERNAME) WITH NOWAIT; SET NOEXEC ON; END")
    [void]$sb.AppendLine('GO')
    if ($script:State.Chain.Mode -like 'Continue*') {
        $m = ("STOP: database [{0}] is not in RESTORING state, so DIFF/LOG restores cannot continue." -f $o.Target).Replace('%', '%%')
        [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.databases WHERE name = $(QS $o.Target) AND state_desc = N'RESTORING') BEGIN RAISERROR($(QS $m), 16, 1) WITH NOWAIT; SET NOEXEC ON; END")
        [void]$sb.AppendLine('GO')
    } elseif (-not $o.Replace) {
        $m = ("STOP: database [{0}] already exists. Use another name or regenerate with REPLACE." -f $o.Target).Replace('%', '%%')
        [void]$sb.AppendLine("IF DB_ID($(QS $o.Target)) IS NOT NULL BEGIN RAISERROR($(QS $m), 16, 1) WITH NOWAIT; SET NOEXEC ON; END")
        [void]$sb.AppendLine('GO')
    } else {
        [void]$sb.AppendLine("-- REPLACE chosen. If [$($o.Target)] is online, disconnect users first:")
        [void]$sb.AppendLine("-- ALTER DATABASE $(QN $o.Target) SET SINGLE_USER WITH ROLLBACK IMMEDIATE;")
        [void]$sb.AppendLine('GO')
    }
    foreach ($s in $Statements) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine($s.Header)
        [void]$sb.AppendLine("PRINT $(QS $s.Print);")
        [void]$sb.AppendLine($s.Sql)
        [void]$sb.AppendLine('GO')
    }
    [void]$sb.AppendLine('')
    if ($o.WithRecovery) {
        [void]$sb.AppendLine('-- Final step: bring the database online (WITH RECOVERY was requested)')
        [void]$sb.AppendLine("RESTORE DATABASE $(QN $o.Target) WITH RECOVERY;")
    } else {
        [void]$sb.AppendLine('-- Database is left RESTORING (NORECOVERY). When the data is verified and no more logs are needed, run:')
        [void]$sb.AppendLine("-- RESTORE DATABASE $(QN $o.Target) WITH RECOVERY;")
    }
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine('SET NOEXEC OFF;')
    [void]$sb.AppendLine('GO')
    return $sb.ToString()
}

function New-JobScript {
    param($Statements, [string]$JobName)
    $o = $script:Opt; $d = $script:State.Dest
    $outFile = $null
    if ($d.ErrorLogFile) {
        $outFile = Join-WinPath (Get-DirOf $d.ErrorLogFile) ("PITR_{0}_{1}.log" -f (ConvertTo-SafeName $o.Target), (Get-Date -Format 'yyyyMMdd_HHmm'))
        if ($outFile.Length -gt 200) { $outFile = $null }
    }
    $guard = New-Object System.Text.StringBuilder
    [void]$guard.AppendLine('SET NOCOUNT ON;')
    $srvMsg = ("Wrong server %s - this job was generated for {0}." -f $d.ServerName).Replace('%s', '#S#').Replace('%', '%%').Replace('#S#', '%s')
    [void]$guard.AppendLine("IF @@SERVERNAME <> $(QS $d.ServerName) BEGIN RAISERROR($(QS $srvMsg), 16, 1, @@SERVERNAME); RETURN; END")
    if ($script:State.Chain.Mode -like 'Continue*') {
        $m = ("Database [{0}] is not RESTORING - job stopped (continue mode)." -f $o.Target).Replace('%', '%%')
        [void]$guard.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.databases WHERE name = $(QS $o.Target) AND state_desc = N'RESTORING') BEGIN RAISERROR($(QS $m), 16, 1); RETURN; END")
    } elseif (-not $o.Replace) {
        $m = ("Database [{0}] already exists - job stopped (REPLACE not chosen)." -f $o.Target).Replace('%', '%%')
        [void]$guard.AppendLine("IF DB_ID($(QS $o.Target)) IS NOT NULL BEGIN RAISERROR($(QS $m), 16, 1); RETURN; END")
    }
    [void]$guard.AppendLine("PRINT N'Safety checks passed on ' + @@SERVERNAME + N' at ' + CONVERT(nvarchar(30), SYSDATETIME(), 121);")

    $steps = New-Object System.Collections.Generic.List[object]
    $steps.Add([pscustomobject]@{ Name = '00 Safety checks'; Cmd = $guard.ToString() })
    $stepNo = 0
    foreach ($t in 'FULL', 'DIFF', 'LOG') {
        $grp = New-Object System.Collections.Generic.List[object]
        foreach ($st in $Statements) { if ($st.Type -eq $t) { $grp.Add($st) } }
        if ($grp.Count -eq 0) { continue }
        $stepNo++
        $cmd = New-Object System.Text.StringBuilder
        [void]$cmd.AppendLine('SET NOCOUNT ON;')
        [void]$cmd.AppendLine("DECLARE @step nvarchar(400) = N'';")
        [void]$cmd.AppendLine('BEGIN TRY')
        foreach ($st in $grp) {
            $lbl = $st.Print.Replace('%', '%%')
            [void]$cmd.AppendLine('')
            [void]$cmd.AppendLine($st.Header)
            [void]$cmd.AppendLine("    SET @step = $(QS $lbl); RAISERROR(@step, 0, 1) WITH NOWAIT;")
            [void]$cmd.AppendLine($st.Sql)
        }
        [void]$cmd.AppendLine('END TRY')
        [void]$cmd.AppendLine('BEGIN CATCH')
        [void]$cmd.AppendLine("    DECLARE @err nvarchar(4000) = N'Failed at ' + @step + N': Msg ' + CAST(ERROR_NUMBER() AS nvarchar(10)) + N' - ' + ERROR_MESSAGE() + N' (full detail, e.g. OS error, is in the SQL Server error log)';")
        [void]$cmd.AppendLine('    RAISERROR(@err, 16, 1);')
        [void]$cmd.AppendLine('END CATCH')
        $first = $grp[0]; $last = $grp[$grp.Count - 1]
        if ($t -eq 'LOG') { $nm = '{0:00} LOG x {1} ({2} -> {3})' -f $stepNo, $grp.Count, ($first.Print -replace '^.*\((.*)\)$', '$1'), ($last.Print -replace '^.*\((.*)\)$', '$1') }
        else { $nm = '{0:00} {1} - {2}' -f $stepNo, $t, ($first.Print -replace '^Step \d+/\d+ ', '') }
        $steps.Add([pscustomobject]@{ Name = $nm; Cmd = $cmd.ToString() })
    }
    if ($o.WithRecovery) { $stepNo++; $steps.Add([pscustomobject]@{ Name = ('{0:00} WITH RECOVERY' -f $stepNo); Cmd = "RESTORE DATABASE $(QN $o.Target) WITH RECOVERY;" }) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((New-Banner "SQL AGENT JOB '$JobName' (run on the DESTINATION)"))
    if ($o.NoSchedule) { [void]$sb.AppendLine('-- No schedule: start it with sp_start_job (bottom of this script).') }
    else { [void]$sb.AppendLine("-- One-time schedule: $($o.JobRunAt.ToString('yyyy-MM-dd HH:mm:ss')) DESTINATION server local time.") }
    [void]$sb.AppendLine('-- One job step per backup type (FULL / DIFF / all LOGs). Inside a step the restores run in order and stop at the first failure,')
    [void]$sb.AppendLine('-- reporting which backup failed. A failed step quits the job.')
    [void]$sb.AppendLine('SET NOCOUNT ON;')
    [void]$sb.AppendLine("DECLARE @jobName sysname = $(QS $JobName), @jobId uniqueidentifier, @owner sysname = SUSER_SNAME(0x01);  -- sa, even if renamed")
    [void]$sb.AppendLine("IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = @jobName) BEGIN RAISERROR(N'Job %s already exists - nothing created.', 16, 1, @jobName); RETURN; END")
    [void]$sb.AppendLine('BEGIN TRY')
    [void]$sb.AppendLine('    BEGIN TRANSACTION;')
    $desc = "PITR restore of [$($o.Db)] from $($script:State.Source.ServerName) as [$($o.Target)], mode $($script:State.Chain.Mode), STOPAT $($script:State.Chain.StopAt.ToString('yyyy-MM-dd HH:mm:ss')). Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERDOMAIN\$env:USERNAME."
    [void]$sb.AppendLine("    EXEC msdb.dbo.sp_add_job @job_name = @jobName, @enabled = 1, @owner_login_name = @owner, @description = $(QS $desc), @job_id = @jobId OUTPUT;")
    $k = 0
    foreach ($st in $steps) {
        $k++
        $succ = if ($k -lt $steps.Count) { 3 } else { 1 }
        $flags = if ($k -eq 1) { 0 } else { 2 }
        $of = if ($outFile) { ", @output_file_name = $(QS $outFile), @flags = $flags" } else { '' }
        $nm = $st.Name; if ($nm.Length -gt 128) { $nm = $nm.Substring(0, 128) }
        [void]$sb.AppendLine("    EXEC msdb.dbo.sp_add_jobstep @job_id = @jobId, @step_id = $k, @step_name = $(QS $nm), @subsystem = N'TSQL', @database_name = N'master',")
        [void]$sb.AppendLine("         @on_success_action = $succ, @on_fail_action = 2, @retry_attempts = 0$of,")
        [void]$sb.AppendLine("         @command = $(QS $st.Cmd);")
    }
    if (-not $o.NoSchedule) {
        [void]$sb.AppendLine("    EXEC msdb.dbo.sp_add_jobschedule @job_id = @jobId, @name = $(QS ('One time ' + $o.JobRunAt.ToString('yyyy-MM-dd HH:mm'))), @enabled = 1, @freq_type = 1,")
        [void]$sb.AppendLine("         @active_start_date = $($o.JobRunAt.ToString('yyyyMMdd')), @active_start_time = $([int]$o.JobRunAt.ToString('HHmmss'));")
    }
    [void]$sb.AppendLine("    EXEC msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)';")
    [void]$sb.AppendLine('    COMMIT TRANSACTION;')
    [void]$sb.AppendLine("    PRINT N'Job created: ' + @jobName;")
    [void]$sb.AppendLine('END TRY')
    [void]$sb.AppendLine('BEGIN CATCH')
    [void]$sb.AppendLine('    IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;')
    [void]$sb.AppendLine('    DECLARE @err nvarchar(4000) = ERROR_MESSAGE();')
    [void]$sb.AppendLine('    RAISERROR(@err, 16, 1);')
    [void]$sb.AppendLine('END CATCH')
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine("-- To start it now instead of waiting for the schedule:")
    [void]$sb.AppendLine("-- EXEC msdb.dbo.sp_start_job @job_name = $(QS $JobName);")
    if ($outFile) { [void]$sb.AppendLine("-- Step output is appended to: $outFile") }
    return $sb.ToString()
}

function New-TdeManualScript {
    param($Certs)
    $o = $script:Opt; $src = $script:State.Source; $d = $script:State.Dest
    $srcId = Get-NetworkIdentity $src.ServiceAccount $src.DomainName $src.HostName
    $me = "$env:USERDOMAIN\$env:USERNAME"
    $grant = @($srcId, (Get-NetworkIdentity $d.ServiceAccount $d.DomainName $d.HostName), $me) | Where-Object { $_ } | Sort-Object -Unique
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((New-Banner 'TDE / BACKUP-ENCRYPTION CERTIFICATE TRANSFER (MANUAL STEPS)'))
    [void]$sb.AppendLine('/* Certificates the destination needs:')
    foreach ($c in $Certs) { [void]$sb.AppendLine("     [$($c.Name)]  $($c.Thumb)  ($($c.Reason))") }
    [void]$sb.AppendLine(@"

   STEP 0  Choose a folder the SOURCE SQL Server service account [$($src.ServiceAccount)] can WRITE to.
           Option A - a temporary share on this machine ($env:COMPUTERNAME). In an elevated PowerShell:
               New-Item -ItemType Directory -Path C:\PITR_TDE -Force
               New-SmbShare -Name PITR_TDE -Path C:\PITR_TDE -FullAccess $(($grant | ForEach-Object { "'$_'" }) -join ',')
               icacls C:\PITR_TDE $(($grant | ForEach-Object { "/grant '$($_):(OI)(CI)M'" }) -join ' ')
           Option B - no rights to create shares: use a folder local to the SOURCE server, e.g. C:\TDE on $($src.HostName),
               and replace <<SOURCE_FOLDER>> below with it.
   STEP 1  Run the SOURCE section on $($src.ServerName). Replace <<CERT_PASSWORD>> with a strong password (keep it in your vault).
   STEP 2  Copy the .cer and .pvk files to $($o.DestCertDir) on the destination host $($d.HostName), for example:
               robocopy "<<SOURCE_FOLDER>>" "\\$($d.HostName)\$($o.DestCertDir.Substring(0,1))`$$($o.DestCertDir.Substring(2))" *.cer *.pvk
   STEP 3  Run the DESTINATION section on $($d.ServerName).
   STEP 4  Remove the share (Remove-SmbShare -Name PITR_TDE -Force). Keep the key files in a safe place, then delete the copies.
*/
"@)
    [void]$sb.AppendLine("-- ===================== SOURCE: $($src.ServerName) =====================")
    [void]$sb.AppendLine('USE master;')
    [void]$sb.AppendLine("IF @@SERVERNAME <> $(QS $src.ServerName) BEGIN RAISERROR(N'Run this section on the SOURCE server.', 16, 1); SET NOEXEC ON; END")
    [void]$sb.AppendLine('GO')
    foreach ($c in $Certs) {
        $fn = ConvertTo-SafeName $c.Name
        [void]$sb.AppendLine("BACKUP CERTIFICATE $(QN $c.Name) TO FILE = N'<<SOURCE_FOLDER>>\$fn.cer'")
        [void]$sb.AppendLine("    WITH PRIVATE KEY (FILE = N'<<SOURCE_FOLDER>>\$fn.pvk', ENCRYPTION BY PASSWORD = N'<<CERT_PASSWORD>>');")
    }
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine('SET NOEXEC OFF;')
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("-- ===================== DESTINATION: $($d.ServerName) =====================")
    [void]$sb.AppendLine('USE master;')
    [void]$sb.AppendLine("IF @@SERVERNAME <> $(QS $d.ServerName) BEGIN RAISERROR(N'Run this section on the DESTINATION server.', 16, 1); SET NOEXEC ON; END")
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##')")
    [void]$sb.AppendLine("    CREATE MASTER KEY ENCRYPTION BY PASSWORD = N'<<DMK_PASSWORD>>';   -- only runs if master has no DMK; back it up afterwards")
    [void]$sb.AppendLine('GO')
    foreach ($c in $Certs) {
        $fn = ConvertTo-SafeName $c.Name
        [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.certificates WHERE thumbprint = $($c.Thumb))")
        [void]$sb.AppendLine("    CREATE CERTIFICATE $(QN $c.Name) FROM FILE = N'$(Join-WinPath $o.DestCertDir "$fn.cer")'")
        [void]$sb.AppendLine("        WITH PRIVATE KEY (FILE = N'$(Join-WinPath $o.DestCertDir "$fn.pvk")', DECRYPTION BY PASSWORD = N'<<CERT_PASSWORD>>');")
        [void]$sb.AppendLine("SELECT name, thumbprint FROM sys.certificates WHERE thumbprint = $($c.Thumb);   -- must return 1 row")
    }
    [void]$sb.AppendLine('GO')
    [void]$sb.AppendLine('SET NOEXEC OFF;')
    [void]$sb.AppendLine('GO')
    return $sb.ToString()
}

function New-CredentialScript {
    param([string[]]$Containers)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine((New-Banner 'BLOB SAS CREDENTIAL (run on the DESTINATION)'))
    [void]$sb.AppendLine("-- Replace <<SAS_TOKEN>> with the SAS token WITHOUT the leading '?'. Needs Read + List; check the expiry date.")
    [void]$sb.AppendLine('USE master;')
    foreach ($ct in $Containers) {
        [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.credentials WHERE name = $(QS $ct))")
        [void]$sb.AppendLine("    CREATE CREDENTIAL $(QN $ct) WITH IDENTITY = N'SHARED ACCESS SIGNATURE', SECRET = N'<<SAS_TOKEN>>';")
        [void]$sb.AppendLine("-- To refresh an expired SAS instead:")
        [void]$sb.AppendLine("-- ALTER CREDENTIAL $(QN $ct) WITH IDENTITY = N'SHARED ACCESS SIGNATURE', SECRET = N'<<SAS_TOKEN>>';")
    }
    [void]$sb.AppendLine('GO')
    return $sb.ToString()
}

function Get-AzCopyList {
    $o = $script:Opt
    $o2 = $o.PSObject.Copy(); $o2.BlobMode = $false
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($s in $script:State.Chain.Sets) {
        foreach ($dev in $s.Devices) {
            $want = ($dev.DeviceType -ne 9) -and ($o.BlobMode -or $script:State.AccessFailures.Contains([string]$dev.Path))
            if ($want) {
                $r = Resolve-BackupPath $dev $o2
                if ($r.Kind -eq 'DISK' -and -not $paths.Contains($r.Path)) { $paths.Add($r.Path) }
            }
        }
    }
    $ct = if ($o.ContainerUrl) { $o.ContainerUrl } else { 'https://<account>.blob.core.windows.net/<container>' }
    return @($paths | Sort-Object -Unique | ForEach-Object { [pscustomobject]@{ Src = $_; Dst = "$ct/$($o.Db)/$(Get-Leaf $_)" } })
}

function New-AzCopyScript {
    param($List)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<#')
    [void]$sb.AppendLine("  05_azcopy_to_blob.ps1 - generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $script:AppName")
    [void]$sb.AppendLine('  Run on a machine that can READ the source backup share AND reach the storage account (azcopy v10 in PATH).')
    [void]$sb.AppendLine("  Then tick 'Files were copied to blob', make sure the destination has the SAS credential (04_blob_credential.sql),")
    [void]$sb.AppendLine('  and run preflight again.')
    [void]$sb.AppendLine('#>')
    [void]$sb.AppendLine("`$sas = '<<SAS_TOKEN>>'   # SAS with Read, Write, Create, List - no leading '?'")
    [void]$sb.AppendLine('$files = @(')
    $i = 0
    foreach ($f in $List) {
        $i++; $comma = if ($i -lt $List.Count) { ',' } else { '' }
        [void]$sb.AppendLine("    @{ Src = '$($f.Src.Replace("'", "''"))'; Dst = '$($f.Dst.Replace("'", "''"))' }$comma")
    }
    [void]$sb.AppendLine(')')
    [void]$sb.AppendLine('foreach ($f in $files) {')
    [void]$sb.AppendLine('    Write-Host "Copying $($f.Src)"')
    [void]$sb.AppendLine('    azcopy copy $f.Src ($f.Dst + ''?'' + $sas) --blob-type BlockBlob --overwrite=false')
    [void]$sb.AppendLine('    if ($LASTEXITCODE -ne 0) { throw "azcopy failed for $($f.Src) (exit $LASTEXITCODE)" }')
    [void]$sb.AppendLine('}')
    return $sb.ToString()
}

function Write-Outputs {
    $o = $script:Opt
    $moves = Get-MoveRows
    $stmts = New-RestoreStatements $script:State.Chain $moves $o
    $files = [ordered]@{}
    $files['01_restore.sql'] = New-SsmsScript $stmts
    $files['02_create_sql_agent_job.sql'] = New-JobScript $stmts $o.JobName
    $script:State.JobScript = $files['02_create_sql_agent_job.sql']; $script:State.JobName = $o.JobName
    $script:State.GenOpt = $o.PSObject.Copy()

    $certs = New-Object System.Collections.Generic.List[object]
    if ($script:State.Checks) { foreach ($x in $script:State.MissingCerts) { $certs.Add($x) } }
    else { foreach ($x in (Get-RequiredCerts)) { $certs.Add($x) } }
    if ($certs.Count) { $files['03_tde_certificate_transfer.sql'] = New-TdeManualScript $certs }
    $ct = Get-MissingContainers
    if ($ct.All.Count) { $files['04_blob_credential.sql'] = New-CredentialScript $ct.All }
    $az = Get-AzCopyList
    if ($az.Count) { $files['05_azcopy_to_blob.ps1'] = New-AzCopyScript $az }

    $rep = New-Object System.Text.StringBuilder
    [void]$rep.AppendLine("$script:AppName $script:Version - report generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$rep.AppendLine("Source $($script:State.Source.ServerName) [$($o.Db)] -> Destination $($script:State.Dest.ServerName) [$($o.Target)]")
    [void]$rep.AppendLine('')
    if ($script:State.Checks) {
        foreach ($c in $script:State.Checks) {
            [void]$rep.AppendLine(("{0,-5} {1}: {2}" -f $c.Status, $c.Check, $c.Detail))
            if ($c.Fix) { [void]$rep.AppendLine("      FIX: $($c.Fix)") }
        }
    } else { [void]$rep.AppendLine('Preflight was NOT run.') }
    [void]$rep.AppendLine(''); [void]$rep.AppendLine('---- chain ----')
    foreach ($s in $script:State.Chain.Sets) { [void]$rep.AppendLine(("{0,2} {1,-4} id {2,-8} {3:yyyy-MM-dd HH:mm:ss}  LSN {4} -> {5}  {6}" -f $s.Step, $s.Type, $s.Id, $s.Finish, $s.FirstLsn, $s.LastLsn, $s.Files)) }
    [void]$rep.AppendLine(''); [void]$rep.AppendLine('---- log ----')
    foreach ($l in $script:Report) { [void]$rep.AppendLine($l) }
    $files['00_preflight_report.txt'] = $rep.ToString()

    $dir = Join-Path $o.OutRoot ('{0}_to_{1}_{2}' -f (ConvertTo-SafeName $o.Db), (ConvertTo-SafeName $o.Target), (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    foreach ($k in $files.Keys) { [System.IO.File]::WriteAllText((Join-Path $dir $k), $files[$k], (New-Object System.Text.UTF8Encoding($true))) }
    $script:State.Chain.Sets | Select-Object Step, Type, Id, IsCopyOnly, Start, Finish, FirstLsn, LastLsn, SizeMB, Files |
        Export-Csv -Path (Join-Path $dir 'chain.csv') -NoTypeInformation
    $script:State.Outputs = $files; $script:State.OutDir = $dir
    return $dir
}

#endregion

#region ============================ optional execution ============================

function Invoke-AzCopyUpload {
    param($List)
    $o = $script:Opt
    if (-not (Get-Command azcopy -ErrorAction SilentlyContinue)) { throw 'azcopy was not found in PATH on this machine. Install azcopy v10 or run 05_azcopy_to_blob.ps1 elsewhere.' }
    foreach ($f in $List) {
        Write-Log "azcopy: $($f.Src) -> $($f.Dst)"
        $out = & azcopy copy $f.Src ($f.Dst + '?' + $o.Sas) --blob-type BlockBlob --overwrite=false 2>&1
        $tail = @($out | Select-Object -Last 3) -join ' | '
        if ($LASTEXITCODE -ne 0) { throw "azcopy failed for $($f.Src) (exit $LASTEXITCODE): $tail" }
        Write-Log "  $tail" 'OK'
    }
}

function Invoke-CreateCredential {
    param([string[]]$Containers)
    foreach ($ct in $Containers) {
        Invoke-Sql $script:Opt.Dest "CREATE CREDENTIAL $(QN $ct) WITH IDENTITY = N'SHARED ACCESS SIGNATURE', SECRET = $(QS $script:Opt.Sas);" | Out-Null
        Write-Log "Created credential [$ct] on $($script:State.Dest.ServerName)." 'OK'
    }
}

function Invoke-TdeTransfer {
    param($Certs)
    $o = $script:Opt; $src = $script:State.Source; $d = $script:State.Dest
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $localDir = Join-Path $o.StageRoot "PITR_TDE_$stamp"
    New-Item -ItemType Directory -Path $localDir -Force | Out-Null
    $shareName = "PITR_TDE_$stamp"; $shareCreated = $false
    $srcLocal = ($src.HostName -eq $env:COMPUTERNAME); $destLocal = ($d.HostName -eq $env:COMPUTERNAME)
    $srcId = Get-NetworkIdentity $src.ServiceAccount $src.DomainName $src.HostName
    $dstId = Get-NetworkIdentity $d.ServiceAccount $d.DomainName $d.HostName
    $grantLocal = {
        param([string]$Acct)
        if (-not $Acct) { return }
        & icacls.exe $localDir /grant "$($Acct):(OI)(CI)M" | Out-Null
        if ($LASTEXITCODE -ne 0) { Write-Log "icacls grant for $Acct failed (exit $LASTEXITCODE)." 'WARN' }
    }
    $newShare = {
        if ($shareCreated) { return }
        if (-not (Test-IsAdmin)) { throw 'Creating the temporary share needs PowerShell started with "Run as administrator". Use 03_tde_certificate_transfer.sql instead.' }
        New-SmbShare -Name $shareName -Path $localDir -FullAccess "$env:USERDOMAIN\$env:USERNAME" | Out-Null
        $shareCreated = $true   # dot-sourced below, so this sets the function's variable
        Write-Log "Created temporary share \\$env:COMPUTERNAME\$shareName -> $localDir" 'OK'
        foreach ($id in @($srcId, $dstId) | Where-Object { $_ }) {
            try { Grant-SmbShareAccess -Name $shareName -AccountName $id -AccessRight Full -Force | Out-Null; & $grantLocal $id; Write-Log "Granted $id on the share." 'OK' }
            catch { Write-Log "Could not grant $id on the share: $($_.Exception.Message)" 'WARN' }
        }
    }
    try {
        if ($srcLocal) { & $grantLocal $src.ServiceAccount; $srcDir = $localDir }
        else { . $newShare; $srcDir = "\\$env:COMPUTERNAME\$shareName" }

        foreach ($c in $Certs) {
            if (-not $c.Name) { throw "Certificate $($c.Thumb) was not found on the source; transfer it manually." }
            $fn = ConvertTo-SafeName $c.Name
            Invoke-Sql $o.Src "BACKUP CERTIFICATE $(QN $c.Name) TO FILE = $(QS (Join-WinPath $srcDir "$fn.cer")) WITH PRIVATE KEY (FILE = $(QS (Join-WinPath $srcDir "$fn.pvk")), ENCRYPTION BY PASSWORD = $(QS $o.CertPwd));" | Out-Null
            Write-Log "Backed up certificate [$($c.Name)] on the source to $srcDir." 'OK'
        }
        $files = @(Get-ChildItem -Path $localDir -File | Where-Object { $_.Extension -in '.cer', '.pvk' })
        if ($files.Count -lt 2 * $Certs.Count) { throw "Expected $(2 * $Certs.Count) certificate files in $localDir, found $($files.Count)." }

        $destDir = $null
        if ($destLocal) {
            New-Item -ItemType Directory -Path $o.DestCertDir -Force | Out-Null
            Copy-Item -Path $files.FullName -Destination $o.DestCertDir -Force
            $destDir = $o.DestCertDir
        } else {
            $admin = "\\$($d.HostName)\$($o.DestCertDir.Substring(0,1))`$$($o.DestCertDir.Substring(2))"
            try {
                New-Item -ItemType Directory -Path $admin -Force | Out-Null
                Copy-Item -Path $files.FullName -Destination $admin -Force
                $destDir = $o.DestCertDir
                Write-Log "Copied certificate files to $admin" 'OK'
            } catch {
                Write-Log "Could not copy to $admin ($($_.Exception.Message)). Trying to let the destination SQL Server read from the share." 'WARN'
                . $newShare
                if ($dstId) { try { Grant-SmbShareAccess -Name $shareName -AccountName $dstId -AccessRight Full -Force | Out-Null; & $grantLocal $dstId } catch { } }
                $destDir = "\\$env:COMPUTERNAME\$shareName"
            }
        }

        if (-not $d.HasDmk) {
            Invoke-Sql $o.Dest "CREATE MASTER KEY ENCRYPTION BY PASSWORD = $(QS $o.DmkPwd);" | Out-Null
            Write-Log 'Created the database master key in master on the destination. Back it up (BACKUP MASTER KEY).' 'OK'
        }
        foreach ($c in $Certs) {
            $fn = ConvertTo-SafeName $c.Name
            $name = $c.Name
            if (@($d.Certs | Where-Object { $_.Name -eq $name }).Count) { $name = "$($c.Name)_PITR_$stamp"; Write-Log "A different certificate named [$($c.Name)] exists on the destination; creating it as [$name]." 'WARN' }
            Invoke-Sql $o.Dest "CREATE CERTIFICATE $(QN $name) FROM FILE = $(QS (Join-WinPath $destDir "$fn.cer")) WITH PRIVATE KEY (FILE = $(QS (Join-WinPath $destDir "$fn.pvk")), DECRYPTION BY PASSWORD = $(QS $o.CertPwd));" | Out-Null
            $chk = Invoke-Sql $o.Dest 'SELECT COUNT(*) AS n FROM master.sys.certificates WHERE thumbprint = CONVERT(varbinary(64), @t, 1);' @{ t = $c.Thumb }
            if ([int]$chk.Rows[0].n -eq 1) { Write-Log "Certificate [$name] $($c.Thumb) is now on the destination." 'OK' }
            else { throw "Certificate $($c.Thumb) not found on the destination after CREATE CERTIFICATE." }
        }
    }
    finally {
        if ($shareCreated) {
            try { Remove-SmbShare -Name $shareName -Force -Confirm:$false; Write-Log "Removed share $shareName (folder kept)." 'OK' }
            catch { Write-Log "Could not remove share ${shareName}: $($_.Exception.Message). Remove it by hand." 'FAIL' }
        }
        Write-Log "SECURITY: $localDir and $($o.DestCertDir) on $($d.HostName) hold the certificate AND private key. Move them to a safe place, then delete them." 'WARN'
    }
}

#endregion

#region ============================ GUI ============================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:State = @{
    Source = $null; Dest = $null; Chain = $null; Files = @(); AddedLater = @(); Checks = $null; MoveTable = $null
    Outputs = $null; OutDir = $null; JobScript = $null; JobName = $null; ShareMap = @()
    AccessFailures = (New-Object System.Collections.Generic.List[string]); MissingCerts = (New-Object System.Collections.Generic.List[object]); ErrorThumbs = @()
}
$script:Report = New-Object System.Collections.Generic.List[string]
$script:TargetAuto = $true; $script:JobAuto = $true; $script:PatternAuto = $true; $script:Setting = $false

function Add-Ctl {
    param($Parent, $Ctl, [int]$X, [int]$Y, [int]$W, [int]$H = 23)
    $Ctl.Location = New-Object System.Drawing.Point($X, $Y)
    $Ctl.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($Ctl)
    return $Ctl
}
function New-Label {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 150, [int]$H = 23, [switch]$PassThru)
    $l = New-Object System.Windows.Forms.Label; $l.Text = $Text; $l.TextAlign = 'MiddleLeft'
    $c = Add-Ctl $Parent $l $X $Y $W $H
    if ($PassThru) { return $c }
}
function New-TextBox {
    param($Parent, [int]$X, [int]$Y, [int]$W, [string]$Text = '', [switch]$Password)
    $t = New-Object System.Windows.Forms.TextBox; $t.Text = $Text
    if ($Password) { $t.UseSystemPasswordChar = $true }
    return (Add-Ctl $Parent $t $X $Y $W 23)
}
function New-Combo {
    param($Parent, [int]$X, [int]$Y, [int]$W, [string[]]$Items, [int]$Selected = 0, [switch]$Editable)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = $(if ($Editable) { 'DropDown' } else { 'DropDownList' })
    if ($Items) { [void]$c.Items.AddRange([object[]]$Items); if (-not $Editable) { $c.SelectedIndex = $Selected } }
    return (Add-Ctl $Parent $c $X $Y $W 23)
}
function New-Check {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [bool]$Checked = $false)
    $c = New-Object System.Windows.Forms.CheckBox; $c.Text = $Text; $c.Checked = $Checked
    return (Add-Ctl $Parent $c $X $Y $W 23)
}
function New-Num {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$Min, [int]$Max, [int]$Value)
    $n = New-Object System.Windows.Forms.NumericUpDown; $n.Minimum = $Min; $n.Maximum = $Max; $n.Value = $Value
    return (Add-Ctl $Parent $n $X $Y $W 23)
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
    param([bool]$ReadOnly = $true)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.ReadOnly = $ReadOnly; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false
    $g.SelectionMode = 'FullRowSelect'; $g.MultiSelect = $false; $g.RowHeadersVisible = $false
    $g.AutoSizeColumnsMode = 'DisplayedCells'; $g.BackgroundColor = [System.Drawing.Color]::White
    return $g
}
function New-DateTime {
    param($Parent, [int]$X, [int]$Y, [int]$W, [datetime]$Value)
    $p = New-Object System.Windows.Forms.DateTimePicker
    $p.Format = 'Custom'; $p.CustomFormat = 'yyyy-MM-dd  hh:mm:ss tt'; $p.Value = $Value
    return (Add-Ctl $Parent $p $X $Y $W 23)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "$script:AppName $script:Version  -  read-only / generate-only unless 'Allow execution' is ticked"
$form.Size = New-Object System.Drawing.Size(1200, 840)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$bottom = New-Object System.Windows.Forms.Panel; $bottom.Dock = 'Bottom'; $bottom.Height = 50
$tabs = New-Object System.Windows.Forms.TabControl; $tabs.Dock = 'Fill'
$form.Controls.Add($bottom); $form.Controls.Add($tabs); $tabs.BringToFront()

function New-Tab { param([string]$Text) $p = New-Object System.Windows.Forms.TabPage; $p.Text = $Text; $p.AutoScroll = $true; $tabs.TabPages.Add($p); return $p }
$tab1 = New-Tab '1. Servers & backups'
$tab2 = New-Tab '2. Restore options & WITH MOVE'
$tab3 = New-Tab '3. Fixes & SQL Agent job'
$tab4 = New-Tab 'Chain & preflight'
$tab5 = New-Tab 'Log'
$tab6 = New-Tab 'Script preview'

# ---------- tab 1 ----------
$gSrc = New-Group $tab1 'Source (msdb history is read here)' 10 8 560 150
New-Label $gSrc 'Server\instance' 10 24 110
$txtSrc = New-Combo $gSrc 125 24 420 @() -Editable
$txtSrc.AutoCompleteMode = 'SuggestAppend'; $txtSrc.AutoCompleteSource = 'ListItems'
New-Label $gSrc 'Authentication' 10 54 110
$cboSrcAuth = New-Combo $gSrc 125 54 150 @('Windows', 'SQL login')
New-Label $gSrc 'Login' 10 84 110
$txtSrcUser = New-TextBox $gSrc 125 84 150
New-Label $gSrc 'Password' 290 84 65
$txtSrcPass = New-TextBox $gSrc 360 84 185 -Password
$btnLoadDbs = New-Button $gSrc 'Connect && load databases' 125 114 200

$gDst = New-Group $tab1 'Destination (preflight runs here, nothing is restored)' 580 8 580 150
New-Label $gDst 'Server\instance' 10 24 110
$txtDst = New-Combo $gDst 125 24 440 @() -Editable
$txtDst.AutoCompleteMode = 'SuggestAppend'; $txtDst.AutoCompleteSource = 'ListItems'
New-Label $gDst 'Authentication' 10 54 110
$cboDstAuth = New-Combo $gDst 125 54 150 @('Windows', 'SQL login')
New-Label $gDst 'Login' 10 84 110
$txtDstUser = New-TextBox $gDst 125 84 150
New-Label $gDst 'Password' 290 84 65
$txtDstPass = New-TextBox $gDst 360 84 205 -Password
$lblDstTime = New-Label $gDst 'Destination server time: (run preflight)' 10 116 560 -PassThru

New-Label $tab1 'Source database' 10 168 120
$cboDb = New-Combo $tab1 135 168 300 @() -Editable
New-Label $tab1 'Restore as (target)' 450 168 130
$txtTarget = New-TextBox $tab1 585 168 320
$btnDefName = New-Button $tab1 'Default name' 912 166 100
$btnRestoring = New-Button $tab1 'Restoring DBs...' 1017 166 140

New-Label $tab1 'Restore mode' 10 198 120
$cboMode = New-Combo $tab1 135 198 300 @('Auto (point in time when logs exist)', 'FULL only', 'FULL + DIFF', 'FULL + DIFF + LOG (point in time)', 'Continue: DIFF only (target already RESTORING)', 'Continue: DIFF + LOGs (point in time)', 'Continue: LOGs only (point in time)')
New-Label $tab1 'STOPAT (source time)' 450 198 130
$dtpStop = New-DateTime $tab1 585 198 190 (Get-Date)
$btnStopNow = New-Button $tab1 'Now' 780 196 50
$lblSrcTime = New-Label $tab1 'Source server time: (connect first)' 835 198 330 -PassThru

New-Label $tab1 'History (days)' 10 228 120
$numDays = New-Num $tab1 135 228 70 1 3650 14
$btnHistory = New-Button $tab1 'Load backup history' 215 226 160
$btnUseFinish = New-Button $tab1 'STOPAT = selected row Finish' 385 226 210
$chkPin = New-Check $tab1 'Start from the selected FULL row (pin it)' 610 228 330
$gridHist = New-Grid
[void](Add-Ctl $tab1 $gridHist 10 260 1150 450)
$gridHist.Anchor = 'Top,Bottom,Left,Right'

# server list row (shift everything on tab 1 down to make room)
foreach ($c in @($tab1.Controls)) { $c.Top += 72 }
$gridHist.Height -= 72
$gList = New-Group $tab1 'Server list (fills the Source and Destination dropdowns; cached on this PC; read-only SELECT)' 10 4 1150 70
New-Label $gList 'List from server' 10 18 105
$txtInvServer = New-TextBox $gList 118 18 220 $env:COMPUTERNAME
New-Label $gList 'Database' 345 18 60
$txtInvDb = New-TextBox $gList 405 18 100 'DBA'
New-Label $gList 'Source' 512 18 48
$cboInvPreset = New-Combo $gList 560 18 230 @('Custom table / query', 'Central Management Server')
$btnInvLoad = New-Button $gList 'Load list' 798 16 85
$btnInvQuery = New-Button $gList 'Query...' 888 16 70
$btnRecentClear = New-Button $gList 'Clear recent' 963 16 95
$lblInvInfo = New-Label $gList '' 10 42 500 18 -PassThru
New-Label $gList 'List auth' 512 42 55 18
$cboInvAuth = New-Combo $gList 568 40 110 @('Windows', 'SQL login')
New-Label $gList 'Login' 684 42 38 18
$txtInvUser = New-TextBox $gList 722 40 130
New-Label $gList 'Password' 858 42 58 18
$txtInvPass = New-TextBox $gList 918 40 140 -Password

# ---------- tab 2 ----------
$gPerf = New-Group $tab2 'WITH options' 10 8 560 200
New-Label $gPerf 'BUFFERCOUNT' 10 24 130
$numBuf = New-Num $gPerf 145 24 90 1 10000 450
New-Label $gPerf 'MAXTRANSFERSIZE' 10 54 130
$cboMts = New-Combo $gPerf 145 54 120 @('65536', '131072', '262144', '524288', '1048576', '2097152', '4194304') 6
New-Label $gPerf 'STATS' 10 84 130
$numStats = New-Num $gPerf 145 84 90 1 100 10
$lblMem = New-Label $gPerf '' 250 24 300 -PassThru
$chkRecovery = New-Check $gPerf 'WITH RECOVERY at the end (default: NORECOVERY, database stays RESTORING)' 10 116 540
$chkReplace = New-Check $gPerf 'WITH REPLACE (overwrite an existing database with the target name)' 10 142 540
New-Label $gPerf 'CHECKSUM is not added: backups are already taken WITH CHECKSUM.' 10 168 540

$gLoc = New-Group $tab2 'Backup file location' 580 8 580 200
New-Label $gLoc 'Remap from prefix' 10 24 120
$txtRemapFrom = New-TextBox $gLoc 135 24 430
New-Label $gLoc 'Remap to prefix' 10 54 120
$txtRemapTo = New-TextBox $gLoc 135 54 430
$chkBlob = New-Check $gLoc 'Files were copied to blob: restore FROM URL <container>/<database>/<file>' 10 84 560
New-Label $gLoc 'Container URL' 10 112 120
$txtContainer = New-TextBox $gLoc 135 112 430 'https://<account>.blob.core.windows.net/<container>'
New-Label $gLoc 'SAS token' 10 142 120
$txtSas = New-TextBox $gLoc 135 142 430 -Password
$chkAutoUnc = New-Check $gLoc 'Convert local paths (C:\...) to the source server share that holds them (UNC)' 10 168 560 $true

$gMove = New-Group $tab2 'WITH MOVE - edit the TargetPath column if needed' 10 214 1150 120
New-Label $gMove 'Strategy' 10 22 70
$cboMove = New-Combo $gMove 85 22 470 @('Auto: similar database -> most-used user-database folder -> instance default', 'Manual: use the folders below', 'Instance default data/log paths', 'Largest free disk on the destination')
New-Label $gMove 'Similar DB name contains' 570 22 160
$txtPattern = New-TextBox $gMove 735 22 250
$btnRecalc = New-Button $gMove 'Recalculate MOVE' 995 20 140
New-Label $gMove 'Data (mdf)' 10 52 70
$txtDataDir = New-TextBox $gMove 85 52 262
$btnBrData = New-Button $gMove '...' 350 51 32 25
New-Label $gMove 'NDF' 392 52 32
$txtNdfDir = New-TextBox $gMove 425 52 262
$btnBrNdf = New-Button $gMove '...' 690 51 32 25
New-Label $gMove 'Log (ldf)' 732 52 60
$txtLogDir = New-TextBox $gMove 795 52 300
$btnBrLog = New-Button $gMove '...' 1098 51 32 25
$btnSetRows = New-Button $gMove 'Set folder for selected rows...' 10 82 220 27
New-Label $gMove '... = browse the destination disks (free space, folders). Typed folders win over the strategy. Empty NDF = with the MDF.' 240 84 900
$gridMove = New-Grid -ReadOnly $false
[void](Add-Ctl $tab2 $gridMove 10 340 1150 370)
$gridMove.Anchor = 'Top,Bottom,Left,Right'
$gridMove.MultiSelect = $true

# ---------- tab 3 ----------
$gTde = New-Group $tab3 'TDE / backup-encryption certificate (only used by Execute fixes)' 10 8 570 210
New-Label $gTde 'Certificate file password' 10 24 170
$txtCertPwd = New-TextBox $gTde 185 24 370 -Password
New-Label $gTde 'Confirm password' 10 54 170
$txtCertPwd2 = New-TextBox $gTde 185 54 370 -Password
New-Label $gTde 'DMK password (if dest has none)' 10 84 170
$txtDmkPwd = New-TextBox $gTde 185 84 370 -Password
New-Label $gTde 'Staging folder on this PC' 10 114 170
$txtStage = New-TextBox $gTde 185 114 370 'C:\PITR_TDE'
New-Label $gTde 'Certificate folder on dest' 10 144 170
$txtDestCert = New-TextBox $gTde 185 144 370 'C:\TDE'
New-Label $gTde 'A temporary share is created on this PC and removed afterwards.' 10 174 550

$gJob = New-Group $tab3 'SQL Agent job on the destination (owner sa)' 590 8 570 210
New-Label $gJob 'Job name' 10 24 110
$txtJob = New-TextBox $gJob 125 24 430
New-Label $gJob 'Run once at' 10 54 110
$dtpJob = New-DateTime $gJob 125 54 180 ((Get-Date).AddMinutes(30))
New-Label $gJob '(destination server local time)' 315 54 240
$chkNoSched = New-Check $gJob 'No schedule - create the job only, start it manually' 10 84 540
$chkStartNow = New-Check $gJob 'Start the job right after creating it' 10 110 540
New-Label $gJob 'One job step per backup type: FULL, DIFF, all LOGs (+ RECOVERY).' 10 140 540

$gOut = New-Group $tab3 'Output' 10 224 1150 60
New-Label $gOut 'Output folder' 10 24 100
$txtOut = New-TextBox $gOut 115 24 900 (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PITR_Restore')
$btnBrowse = New-Button $gOut 'Browse...' 1025 22 110

$gExec = New-Group $tab3 'Execution (everything above is generate-only)' 10 290 1150 130
$chkAllow = New-Check $gExec 'Allow execution (OFF = read-only)' 10 24 400
$chkAllow.ForeColor = [System.Drawing.Color]::DarkRed
$btnExecFixes = New-Button $gExec 'Execute fixes (azcopy / credential / TDE)...' 10 56 330 30
$btnCreateJob = New-Button $gExec 'Create SQL Agent job on destination...' 350 56 300 30
New-Label $gExec 'Each action lists exactly what it will change and asks first. The restore itself only runs through the job you create.' 10 94 1100

# ---------- tab 4 ----------
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'; $split.SplitterDistance = 260
$tab4.Controls.Add($split)
$gridChain = New-Grid; $gridChain.Dock = 'Fill'; $split.Panel1.Controls.Add($gridChain)
$gridChecks = New-Grid; $gridChecks.Dock = 'Fill'; $split.Panel2.Controls.Add($gridChecks)

# ---------- tab 5 / 6 ----------
$rtbLog = New-Object System.Windows.Forms.RichTextBox
$rtbLog.Dock = 'Fill'; $rtbLog.ReadOnly = $true; $rtbLog.Font = New-Object System.Drawing.Font('Consolas', 9); $rtbLog.WordWrap = $false
$tab5.Controls.Add($rtbLog)
$txtPreview = New-Object System.Windows.Forms.RichTextBox
$txtPreview.Dock = 'Fill'; $txtPreview.ReadOnly = $true; $txtPreview.WordWrap = $false; $txtPreview.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtPreview.MaxLength = [int]::MaxValue
$cboPreview = New-Object System.Windows.Forms.ComboBox; $cboPreview.Dock = 'Top'; $cboPreview.DropDownStyle = 'DropDownList'
$tab6.Controls.Add($txtPreview); $tab6.Controls.Add($cboPreview)

# ---------- bottom ----------
$btnBuild = New-Button $bottom '1. Build chain' 10 10 140 30
$btnPreflight = New-Button $bottom '2. Preflight on destination' 160 10 200 30
$btnGenerate = New-Button $bottom '3. Generate scripts' 370 10 160 30
$btnOpenOut = New-Button $bottom 'Open output folder' 540 10 150 30
$lblStatus = New-Label $bottom 'Ready' 705 13 470 23 -PassThru

#endregion

#region ============================ GUI logic ============================

function Write-Log {
    param([string]$Msg, [ValidateSet('INFO', 'OK', 'WARN', 'FAIL')][string]$Level = 'INFO')
    $line = '{0:HH:mm:ss} {1,-4} {2}' -f (Get-Date), $Level, $Msg
    $script:Report.Add($line)
    $color = switch ($Level) { 'OK' { 'DarkGreen' } 'WARN' { 'DarkOrange' } 'FAIL' { 'Red' } default { 'Black' } }
    $rtbLog.SelectionStart = $rtbLog.TextLength
    $rtbLog.SelectionColor = [System.Drawing.Color]::$color
    $rtbLog.AppendText($line + "`r`n")
    $rtbLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Status { param([string]$Text) $lblStatus.Text = $Text; [System.Windows.Forms.Application]::DoEvents() }

function Confirm-Action {
    param([string]$Text)
    $r = [System.Windows.Forms.MessageBox]::Show($Text, "$script:AppName - confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Show-Info { param([string]$Text) [void][System.Windows.Forms.MessageBox]::Show($Text, $script:AppName, 'OK', 'Information') }

function Invoke-Ui {
    param([scriptblock]$Action)
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $Action }
    catch {
        $m = Get-SqlErrorText $_
        $ii = $_.InvocationInfo
        if ($ii -and $ii.ScriptLineNumber) { $m += "  (script line $($ii.ScriptLineNumber): $($ii.Line.Trim()))" }
        Write-Log $m 'FAIL'
        [void][System.Windows.Forms.MessageBox]::Show($m, $script:AppName, 'OK', 'Error')
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; Set-Status 'Ready' }
}

function Read-Options {
    $modes = @('Auto', 'FullOnly', 'FullDiff', 'FullDiffLog', 'ContinueDiff', 'ContinueDiffLog', 'ContinueLog')
    $pin = 0
    if ($chkPin.Checked) {
        $row = $gridHist.CurrentRow
        if (-not $row -or [string]$row.Cells['Type'].Value -ne 'FULL') { throw 'Pin is ticked: select a FULL row in the backup history grid.' }
        $pin = [int]$row.Cells['Id'].Value
    }
    $ct = $txtContainer.Text.Trim().TrimEnd('/')
    if ($ct -like '*<account>*') { $ct = '' }
    $script:Opt = [pscustomobject]@{
        Src  = @{ Server = $txtSrc.Text.Trim(); Auth = $cboSrcAuth.Text; User = $txtSrcUser.Text; Pass = $txtSrcPass.Text }
        Dest = @{ Server = $txtDst.Text.Trim(); Auth = $cboDstAuth.Text; User = $txtDstUser.Text; Pass = $txtDstPass.Text }
        Db = $cboDb.Text.Trim(); Target = $txtTarget.Text.Trim(); Mode = $modes[$cboMode.SelectedIndex]; StopAt = $dtpStop.Value
        PinnedFullId = $pin; BufferCount = [int]$numBuf.Value; MaxTransfer = [int]$cboMts.Text; Stats = [int]$numStats.Value
        WithRecovery = $chkRecovery.Checked; Replace = $chkReplace.Checked
        RemapFrom = $txtRemapFrom.Text.Trim(); RemapTo = $txtRemapTo.Text.Trim()
        BlobMode = $chkBlob.Checked; ContainerUrl = $ct; Sas = $txtSas.Text.Trim().TrimStart('?')
        MoveStrategy = $cboMove.Text; Pattern = $txtPattern.Text.Trim()
        DataDir = $txtDataDir.Text.Trim(); NdfDir = $txtNdfDir.Text.Trim(); LogDir = $txtLogDir.Text.Trim()
        CertPwd = $txtCertPwd.Text; CertPwd2 = $txtCertPwd2.Text; DmkPwd = $txtDmkPwd.Text
        StageRoot = $txtStage.Text.Trim(); DestCertDir = $txtDestCert.Text.Trim().TrimEnd('\')
        JobName = $txtJob.Text.Trim(); JobRunAt = $dtpJob.Value; NoSchedule = $chkNoSched.Checked; StartNow = $chkStartNow.Checked
        OutRoot = $txtOut.Text.Trim(); AllowExec = $chkAllow.Checked; AutoUnc = $chkAutoUnc.Checked
    }
    if ($script:Opt.BlobMode -and -not $script:Opt.ContainerUrl) { Write-Log "Blob option is ticked but the container URL is a placeholder; scripts will contain <account>/<container>." 'WARN' }
}

function Assert-Source { if (-not $script:Opt.Src.Server) { throw 'Enter the source server.' }; if (-not $script:Opt.Db) { throw 'Pick the source database.' } }
function Assert-Dest {
    if (-not $script:Opt.Dest.Server) { throw 'Enter the destination server.' }
    if (-not $script:Opt.Target) { throw 'Enter the target database name.' }
}

function Update-DefaultNames {
    $script:Setting = $true
    try {
        $db = $cboDb.Text.Trim()
        if ($db -and $script:TargetAuto) { $txtTarget.Text = '{0}_PITR_{1:yyyyMMdd_HHmm}' -f $db, $dtpStop.Value }
        if ($db -and $script:PatternAuto) { $txtPattern.Text = $db }
        if ($script:JobAuto -and $txtTarget.Text) { $txtJob.Text = 'PITR restore - {0} - {1:yyyyMMdd_HHmm}' -f $txtTarget.Text, (Get-Date) }
    } finally { $script:Setting = $false }
}

function Update-MemLabel {
    $mb = [math]::Round([double]$numBuf.Value * [double]$cboMts.Text / 1MB)
    $lblMem.Text = "Restore buffers: ~$mb MB (BUFFERCOUNT x MAXTRANSFERSIZE), outside the buffer pool"
}

function Update-MoveGrid {
    $o = $script:Opt
    if ($script:State.Chain -and $script:State.Chain.Sets[0].Type -ne 'FULL') {
        $dt = New-Object System.Data.DataTable
        foreach ($c in 'LogicalName', 'Kind', 'SizeMB', 'SourcePath', 'TargetPath', 'Comment', 'Computed') { [void]$dt.Columns.Add($c) }
        $script:State.MoveTable = $dt; $gridMove.DataSource = $dt
        Write-Log 'No FULL in this chain (continue mode): the database files already exist, so no WITH MOVE is needed.' 'INFO'
        return
    }
    if (-not $script:State.Dest) { throw 'Destination information is needed for MOVE paths: fill in the destination and run preflight.' }
    $plan = Get-MovePlan -Files $script:State.Files -Dest $script:State.Dest -TargetDb $o.Target -Strategy $o.MoveStrategy `
                         -DataDir $o.DataDir -NdfDir $o.NdfDir -LogDir $o.LogDir -Pattern $o.Pattern
    $dt = New-Object System.Data.DataTable
    foreach ($c in 'LogicalName', 'Kind', 'SizeMB', 'SourcePath', 'TargetPath', 'Comment', 'Computed') { [void]$dt.Columns.Add($c) }
    foreach ($p in $plan) {
        $r = $dt.NewRow()
        foreach ($c in 'LogicalName', 'Kind', 'SizeMB', 'SourcePath', 'TargetPath', 'Comment', 'Computed') { $r[$c] = [string]$p.$c }
        $dt.Rows.Add($r)
    }
    foreach ($c in 'LogicalName', 'Kind', 'SizeMB', 'SourcePath', 'Comment') { $dt.Columns[$c].ReadOnly = $true }
    $script:State.MoveTable = $dt
    $gridMove.DataSource = $dt
    if ($gridMove.Columns['Computed']) { $gridMove.Columns['Computed'].Visible = $false }
    if ($gridMove.Columns['TargetPath']) { $gridMove.Columns['TargetPath'].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow }
    Write-Log "MOVE plan: $($plan.Count) file(s). Review the TargetPath column on tab 2." 'OK'
}

function Show-Preview {
    $cboPreview.Items.Clear()
    foreach ($k in $script:State.Outputs.Keys) { [void]$cboPreview.Items.Add($k) }
    if ($cboPreview.Items.Count) { $cboPreview.SelectedIndex = [Math]::Min(1, $cboPreview.Items.Count - 1) }
}

# ---------- events ----------
$cboDb.Add_TextChanged({ Update-DefaultNames })
$dtpStop.Add_ValueChanged({ Update-DefaultNames })
$txtTarget.Add_TextChanged({ if (-not $script:Setting) { $script:TargetAuto = $false; Update-DefaultNames } })
$txtJob.Add_TextChanged({ if (-not $script:Setting) { $script:JobAuto = $false } })
$txtPattern.Add_TextChanged({ if (-not $script:Setting) { $script:PatternAuto = $false } })
$btnDefName.Add_Click({ $script:TargetAuto = $true; $script:JobAuto = $true; Update-DefaultNames })
$numBuf.Add_ValueChanged({ Update-MemLabel })
$cboMts.Add_SelectedIndexChanged({ Update-MemLabel })
$cboPreview.Add_SelectedIndexChanged({ if ($cboPreview.SelectedItem) { $txtPreview.Text = $script:State.Outputs[[string]$cboPreview.SelectedItem] } })
$gridChecks.Add_DataBindingComplete({
    foreach ($row in $gridChecks.Rows) {
        $st = [string]$row.Cells['Status'].Value
        $bc = [System.Drawing.Color]::White
        if ($st -eq 'FAIL') { $bc = [System.Drawing.Color]::MistyRose }
        elseif ($st -eq 'WARN') { $bc = [System.Drawing.Color]::LightYellow }
        elseif ($st -eq 'PASS') { $bc = [System.Drawing.Color]::Honeydew }
        $row.DefaultCellStyle.BackColor = $bc
    }
})
$btnBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = $fb.SelectedPath }
})
$btnOpenOut.Add_Click({ if ($script:State.OutDir -and (Test-Path $script:State.OutDir)) { Start-Process explorer.exe $script:State.OutDir } else { Show-Info 'Nothing generated yet.' } })

$btnLoadDbs.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:Opt.Src.Server) { throw 'Enter the source server.' }
    Set-Status 'Connecting to source...'
    $t = Invoke-Sql $script:Opt.Src @"
SELECT name FROM sys.databases WHERE database_id > 4
UNION
SELECT DISTINCT database_name FROM msdb.dbo.backupset WHERE database_name NOT IN (N'master', N'model', N'msdb', N'tempdb')
ORDER BY name;
"@
    $keep = $cboDb.Text
    $cboDb.Items.Clear()
    foreach ($r in $t.Rows) { [void]$cboDb.Items.Add([string]$r.name) }
    if ($keep) { $cboDb.Text = $keep }
    $si = Get-SourceInfo $script:Opt.Src $null
    $script:State.Source = $si
    Add-RecentServer $script:Opt.Src.Server $script:Opt.Src.Auth $script:Opt.Src.User
    $lblSrcTime.Text = "Source: {0:yyyy-MM-dd hh:mm:ss tt}  (this PC: {1:hh:mm:ss tt})" -f $si.ServerTime, (Get-Date)
    $dtpStop.Value = $si.ServerTime
    Write-Log "Connected to source $($si.ServerName) (major v$($si.Major), host $($si.HostName)). $($t.Rows.Count) database(s). STOPAT set to the source server's current time." 'OK'
} })

$btnHistory.Add_Click({ Invoke-Ui {
    Read-Options; Assert-Source
    $si = Get-SourceInfo $script:Opt.Src $script:Opt.Db
    $script:State.Source = $si
    if ($si.FilterNote) { Write-Log $si.FilterNote 'WARN' }
    $q = @"
SELECT TOP (1000) bs.backup_set_id AS [Id],
       CASE bs.type WHEN 'D' THEN 'FULL' WHEN 'I' THEN 'DIFF' WHEN 'L' THEN 'LOG' ELSE bs.type END AS [Type],
       bs.is_copy_only AS [CopyOnly], bs.backup_start_date AS [Start], bs.backup_finish_date AS [Finish],
       CAST(bs.backup_size / 1048576.0 AS decimal(18,1)) AS [SizeMB],
       CAST(bs.compressed_backup_size / 1048576.0 AS decimal(18,1)) AS [CompressedMB],
       bs.recovery_model AS [RecoveryModel],
       (SELECT COUNT(*) FROM msdb.dbo.backupmediafamily mf WHERE mf.media_set_id = bs.media_set_id AND mf.mirror = 0) AS [Stripes],
       (SELECT TOP (1) mf.physical_device_name FROM msdb.dbo.backupmediafamily mf WHERE mf.media_set_id = bs.media_set_id AND mf.mirror = 0 ORDER BY mf.family_sequence_number) AS [FirstFile]
FROM msdb.dbo.backupset bs
WHERE bs.database_name = @db AND bs.type IN ('D', 'I', 'L') AND bs.backup_finish_date >= DATEADD(DAY, -@days, GETDATE()) $($si.ServerFilter)
ORDER BY bs.backup_finish_date DESC;
"@
    $t = Invoke-Sql $script:Opt.Src $q @{ db = $script:Opt.Db; days = [int]$numDays.Value }
    $gridHist.DataSource = $t
    if ($gridHist.Columns['Start'])  { $gridHist.Columns['Start'].DefaultCellStyle.Format = 'yyyy-MM-dd hh:mm:ss tt' }
    if ($gridHist.Columns['Finish']) { $gridHist.Columns['Finish'].DefaultCellStyle.Format = 'yyyy-MM-dd hh:mm:ss tt' }
    Write-Log "Loaded $($t.Rows.Count) backup(s) of [$($script:Opt.Db)] from the last $([int]$numDays.Value) day(s). Recovery model now: $($si.RecoveryModel). TDE: $(if ($si.TdeCertName) { $si.TdeCertName } else { 'no' })" 'OK'
} })

$btnStopNow.Add_Click({ Invoke-Ui {
    Read-Options
    $now = Get-Date
    if ($script:Opt.Src.Server) {
        try { $now = [datetime](Invoke-Sql $script:Opt.Src 'SELECT SYSDATETIME() AS t;').Rows[0].t; $lblSrcTime.Text = 'Source server time: {0:yyyy-MM-dd hh:mm:ss tt}' -f $now }
        catch { Write-Log "Could not read the source server time; using this PC's time. $(Get-SqlErrorText $_)" 'WARN' }
    }
    $dtpStop.Value = $now
    Write-Log ('STOPAT set to now: {0:yyyy-MM-dd hh:mm:ss tt} (source server time).' -f $now) 'OK'
} })

function Show-DestFolderPicker {
    param([string]$Title, [string]$Start)
    $conn = $script:Opt.Dest
    if (-not $conn.Server) { throw 'Enter the destination server first (tab 1).' }
    Set-Status "Reading disks on $($conn.Server)..."
    $drives = @(Get-DestDrives $conn)
    if (-not $drives.Count) { throw "Could not list the disks on $($conn.Server). Needs VIEW SERVER STATE (or sysadmin for xp_fixeddrives)." }
    $script:PickConn = $conn
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "$Title - disks on $($conn.Server)"; $dlg.Size = New-Object System.Drawing.Size(760, 620); $dlg.StartPosition = 'CenterParent'; $dlg.Font = $form.Font
    New-Label $dlg 'Disks and folders as the destination SQL Server service account sees them. Expand a disk to see its folders.' 10 8 720
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.HideSelection = $false
    [void](Add-Ctl $dlg $tree 10 34 725 450)
    $maxFree = ($drives | Measure-Object FreeBytes -Maximum).Maximum
    $bold = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
    foreach ($dv in $drives) {
        $txt = '{0}    {1:N1} GB free' -f $dv.Root, ($dv.FreeBytes / 1GB)
        if ($dv.TotalBytes) { $txt += ' of {0:N0} GB ({1:N0}% free)' -f ($dv.TotalBytes / 1GB), (100 * $dv.FreeBytes / $dv.TotalBytes) }
        if ($dv.FreeBytes -eq $maxFree) { $txt += '    <- most free space' }
        $node = New-Object System.Windows.Forms.TreeNode -ArgumentList $txt
        $node.Tag = $dv.Root
        if ($dv.FreeBytes -eq $maxFree) { $node.NodeFont = $bold }
        [void]$node.Nodes.Add('(loading...)')
        [void]$tree.Nodes.Add($node)
    }
    New-Label $dlg 'Selected folder (you can also type a path; it must already exist on the destination):' 10 492 720
    $script:PickPath = New-TextBox $dlg 10 516 725 $Start
    $ok = New-Button $dlg 'Use this folder' 495 548 130
    $cancel = New-Button $dlg 'Cancel' 635 548 100
    $ok.DialogResult = 'OK'; $cancel.DialogResult = 'Cancel'; $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    $tree.Add_BeforeExpand({
        $n = $_.Node
        if ($n.Nodes.Count -eq 1 -and $null -eq $n.Nodes[0].Tag -and $n.Nodes[0].Text -eq '(loading...)') {
            $n.Nodes.Clear()
            try {
                $subs = @(Get-DestSubfolders $script:PickConn ([string]$n.Tag))
                foreach ($sub in $subs) {
                    $c = New-Object System.Windows.Forms.TreeNode -ArgumentList $sub
                    $c.Tag = Join-WinPath ([string]$n.Tag) $sub
                    [void]$c.Nodes.Add('(loading...)')
                    [void]$n.Nodes.Add($c)
                }
                if (-not $subs.Count) { [void]$n.Nodes.Add('(no subfolders, or no access for the SQL service account)') }
            } catch { [void]$n.Nodes.Add('(cannot list: ' + (Get-SqlErrorText $_) + ')') }
        }
    })
    $tree.Add_AfterSelect({ if ($_.Node.Tag) { $script:PickPath.Text = [string]$_.Node.Tag } })
    Set-Status 'Ready'
    $r = $dlg.ShowDialog($form)
    $path = $script:PickPath.Text.Trim()
    $dlg.Dispose()
    if ($r -eq [System.Windows.Forms.DialogResult]::OK -and $path) { return $path }
    return $null
}

$btnBrData.Add_Click({ Invoke-Ui { Read-Options; $p = Show-DestFolderPicker 'Folder for the MDF' $txtDataDir.Text; if ($p) { $txtDataDir.Text = $p } } })
$btnBrNdf.Add_Click({ Invoke-Ui { Read-Options; $p = Show-DestFolderPicker 'Folder for the NDF files' $txtNdfDir.Text; if ($p) { $txtNdfDir.Text = $p } } })
$btnBrLog.Add_Click({ Invoke-Ui { Read-Options; $p = Show-DestFolderPicker 'Folder for the LDF' $txtLogDir.Text; if ($p) { $txtLogDir.Text = $p } } })
$btnSetRows.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:State.MoveTable -or $script:State.MoveTable.Rows.Count -eq 0) { throw 'No MOVE rows yet: build the chain (with the destination filled in) first.' }
    $rows = @($gridMove.SelectedRows)
    if (-not $rows.Count -and $gridMove.CurrentRow) { $rows = @($gridMove.CurrentRow) }
    if (-not $rows.Count) { throw 'Select one or more rows in the MOVE grid first (click the row; Ctrl/Shift for several).' }
    $p = Show-DestFolderPicker ("Folder for {0} selected file(s)" -f $rows.Count) ''
    if (-not $p) { return }
    foreach ($gr in $rows) {
        $dr = $gr.DataBoundItem.Row
        $dr['TargetPath'] = Join-WinPath $p (Get-Leaf ([string]$dr['TargetPath']))
    }
    Write-Log "Set folder $p for $($rows.Count) file(s). Run preflight to re-check folders and space." 'OK'
} })

$cboMode.Add_SelectedIndexChanged({
    if ($cboMode.SelectedIndex -ge 4) {
        $script:TargetAuto = $false
        Write-Log 'Continue mode: set "Restore as" to the database that is already RESTORING on the destination (Restoring DBs... button). The FULL is not restored again.' 'INFO'
    }
})

$btnRestoring.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:Opt.Dest.Server) { throw 'Enter the destination server first.' }
    $t = Invoke-Sql $script:Opt.Dest @"
SELECT d.name,
       (SELECT TOP (1) CONVERT(varchar(19), rh.restore_date, 120) + ' ' +
               CASE rh.restore_type WHEN 'D' THEN 'FULL' WHEN 'I' THEN 'DIFF' WHEN 'L' THEN 'LOG' ELSE rh.restore_type END
        FROM msdb.dbo.restorehistory rh WHERE rh.destination_database_name = d.name ORDER BY rh.restore_history_id DESC) AS last_restore
FROM sys.databases d WHERE d.state_desc = N'RESTORING' ORDER BY d.name;
"@
    if (-not $t.Rows.Count) { Show-Info "No database is in RESTORING state on $($script:Opt.Dest.Server)."; return }
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    foreach ($r in $t.Rows) {
        $label = [string]$r.name
        if ($r.last_restore -isnot [System.DBNull]) { $label += '    (last restore: ' + [string]$r.last_restore + ')' }
        $it = $menu.Items.Add($label)
        $it.Tag = [string]$r.name
        $it.Add_Click({ $txtTarget.Text = [string]$this.Tag; $script:TargetAuto = $false; Write-Log "Target set to RESTORING database [$($this.Tag)]." 'OK' })
    }
    $menu.Show($btnRestoring, (New-Object System.Drawing.Point(0, $btnRestoring.Height)))
} })

$btnUseFinish.Add_Click({
    $row = $gridHist.CurrentRow
    if ($row) { $dtpStop.Value = [datetime]$row.Cells['Finish'].Value } else { Show-Info 'Select a row in the history grid first.' }
})

$btnRecalc.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:State.Chain) { throw 'Build the chain first.' }
    if (-not $script:State.Dest) { Assert-Dest; $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target }
    Update-MoveGrid
} })

$btnBuild.Add_Click({ Invoke-Ui {
    Read-Options; Assert-Source
    Set-Status 'Building the restore chain from msdb...'
    $si = Get-SourceInfo $script:Opt.Src $script:Opt.Db
    $script:State.Source = $si
    if ($si.FilterNote) { Write-Log $si.FilterNote 'WARN' }
    Write-Log ('STOPAT used: {0:yyyy-MM-dd hh:mm:ss tt} ({0:yyyy-MM-dd HH:mm:ss} 24h), source server time now {1:yyyy-MM-dd hh:mm:ss tt}.' -f $script:Opt.StopAt, $si.ServerTime) 'INFO'
    if ($script:Opt.StopAt -gt $si.ServerTime) { Write-Log 'STOPAT is in the future on the source server: the chain will need a log backup that does not exist yet. Use the Now button or an earlier time.' 'WARN' }
    $redo = $null; $dbase = $null
    if ($script:Opt.Mode -like 'Continue*') {
        Assert-Dest
        $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target
        $dd = $script:State.Dest
        if (-not $dd.TargetExists -or $dd.TargetState -ne 'RESTORING') {
            $names = @((Invoke-Sql $script:Opt.Dest "SELECT name FROM sys.databases WHERE state_desc = N'RESTORING' ORDER BY name;").Rows | ForEach-Object { [string]$_.name })
            $st = 'does not exist'; if ($dd.TargetExists) { $st = "is $($dd.TargetState)" }
            $lst = 'none'; if ($names.Count) { $lst = $names -join ', ' }
            throw "Continue mode: [$($script:Opt.Target)] $st on $($dd.ServerName); it must be RESTORING. RESTORING databases there: $lst."
        }
        if (-not $dd.TargetRedoLsn) { throw "Could not read redo_start_lsn for [$($script:Opt.Target)] on $($dd.ServerName) (needs VIEW SERVER STATE / VIEW ANY DEFINITION)." }
        $redo = $dd.TargetRedoLsn; $dbase = $dd.TargetDiffBaseLsn
        Write-Log "Continuing [$($script:Opt.Target)] on $($dd.ServerName): RESTORING, next restore starts at LSN $redo (differential base LSN $dbase). Last restore: $($dd.TargetLastRestore)" 'INFO'
    }
    $chain = Get-BackupChain -Conn $script:Opt.Src -Db $script:Opt.Db -StopAt $script:Opt.StopAt -Mode $script:Opt.Mode `
                             -PinnedFullId $script:Opt.PinnedFullId -SrcMajor $si.Major -ServerFilter $si.ServerFilter -ContinueLsn $redo -DiffBaseLsn $dbase
    $script:State.Chain = $chain
    $script:State.Checks = $null; $script:State.MoveTable = $null; $gridChecks.DataSource = $null
    foreach ($w in $chain.Warnings) { Write-Log $w 'WARN' }
    $bf = Get-BackupFiles $script:Opt.Src $chain
    $script:State.Files = $bf.Files; $script:State.AddedLater = $bf.AddedLater
    if ($bf.AddedLater.Count) { Write-Log "Files added after the FULL: $($bf.AddedLater -join ', ')" 'WARN' }
    Set-Status 'Checking whether local backup folders are shared...'
    Update-ShareMap
    $gridChain.DataSource = ConvertTo-DataTable $chain.Sets @('Step', 'Type', 'Id', 'IsCopyOnly', 'Start', 'Finish', 'FirstLsn', 'LastLsn', 'SizeMB', 'Files')
    Write-Log ("Chain built (mode {0}): {1}. First: {2} id {3} finished {4:yyyy-MM-dd hh:mm:ss tt}. Restore point {5:yyyy-MM-dd hh:mm:ss tt}." -f $chain.Mode, (Get-ChainSummary $chain), $chain.Sets[0].Type, $chain.Sets[0].Id, $chain.Sets[0].Finish, $chain.RestorePoint) 'OK'
    if ($script:Opt.Dest.Server -and $script:Opt.Target) {
        try { $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target; Add-RecentServer $script:Opt.Dest.Server $script:Opt.Dest.Auth $script:Opt.Dest.User; Update-MoveGrid }
        catch { Write-Log "Destination not read yet: $(Get-SqlErrorText $_)" 'WARN' }
    }
    $tabs.SelectedTab = $tab4
} })

$btnPreflight.Add_Click({ Invoke-Ui {
    Read-Options; Assert-Dest
    if (-not $script:State.Chain) { throw 'Build the chain first (button 1).' }
    Set-Status 'Reading destination...'
    $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target
    Add-RecentServer $script:Opt.Dest.Server $script:Opt.Dest.Auth $script:Opt.Dest.User
    $lblDstTime.Text = "Destination server time: {0:yyyy-MM-dd HH:mm:ss}  (host {1}, service account {2})" -f $script:State.Dest.ServerTime, $script:State.Dest.HostName, $script:State.Dest.ServiceAccount
    if (-not $script:State.MoveTable) { Update-MoveGrid }
    $r = Invoke-Preflight
    $gridChecks.DataSource = ConvertTo-DataTable $script:State.Checks @('Status', 'Check', 'Detail', 'Fix')
    $tabs.SelectedTab = $tab4
    $msg = "Preflight finished: $($r.Fail) FAIL, $($r.Warn) WARN."
    if ($script:State.AccessFailures.Count) { $msg += "`r`n`r`nThe destination can't read $($script:State.AccessFailures.Count) file(s). Generate scripts writes 05_azcopy_to_blob.ps1 to copy them to blob storage." }
    if ($script:State.MissingCerts.Count) { $msg += "`r`n`r`n$($script:State.MissingCerts.Count) certificate(s) missing on the destination. See 03_tde_certificate_transfer.sql, or Execute fixes." }
    Write-Log $msg $(if ($r.Fail) { 'FAIL' } elseif ($r.Warn) { 'WARN' } else { 'OK' })
    Show-Info $msg
} })

$btnGenerate.Add_Click({ Invoke-Ui {
    Read-Options; Assert-Dest
    if (-not $script:State.Chain) { throw 'Build the chain first (button 1).' }
    if (-not $script:State.Dest) { $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target }
    if (-not $script:State.MoveTable) { Update-MoveGrid }
    if (-not $script:State.Checks) { Write-Log 'Preflight was not run: the scripts are generated, but nothing was checked on the destination.' 'WARN' }
    if (-not $script:Opt.NoSchedule -and $script:Opt.JobRunAt -le $script:State.Dest.ServerTime) { Write-Log "Job time $($script:Opt.JobRunAt) is in the past on the destination ($($script:State.Dest.ServerTime)). Fix it before creating the job." 'WARN' }
    $dir = Write-Outputs
    Show-Preview
    $tabs.SelectedTab = $tab6
    Write-Log "Scripts written to $dir : $($script:State.Outputs.Keys -join ', ')" 'OK'
} })

$btnExecFixes.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:Opt.AllowExec) { throw "Tick 'Allow execution' on tab 3 first." }
    if (-not $script:State.Checks) { throw 'Run preflight first so the tool knows what is missing.' }
    $script:State.Dest = Get-DestInfo $script:Opt.Dest $script:Opt.Target
    $acts = @(); $doAz = $false; $doCred = @(); $doTde = $false
    $az = Get-AzCopyList
    if ($script:Opt.BlobMode -and $az.Count) {
        if (-not $script:Opt.ContainerUrl -or -not $script:Opt.Sas) { throw 'Blob upload: enter the container URL and a SAS token (tab 2).' }
        $gb = [math]::Round((($script:State.Chain.Sets | Measure-Object CompressedMB -Sum).Sum) / 1024, 1)
        $acts += "azcopy: upload $($az.Count) file(s) (~$gb GB) from this PC to $($script:Opt.ContainerUrl)/$($script:Opt.Db)/"
        $doAz = $true
    }
    $ct = Get-MissingContainers
    if ($ct.Missing.Count) {
        if (-not $script:Opt.Sas) { Write-Log 'Credential missing on destination but no SAS token entered; skipped.' 'WARN' }
        else { $doCred = $ct.Missing; $acts += "DESTINATION $($script:State.Dest.ServerName): CREATE CREDENTIAL for $($ct.Missing -join ', ')" }
    }
    $script:State.MissingCerts = New-Object System.Collections.Generic.List[object]
    foreach ($c in (Get-RequiredCerts)) { if (-not @($script:State.Dest.Certs | Where-Object { $_.Thumb -eq $c.Thumb }).Count) { $script:State.MissingCerts.Add($c) } }
    if ($script:State.MissingCerts.Count) {
        if (-not $script:Opt.CertPwd) { Write-Log 'Certificate(s) missing but no certificate password entered (tab 3); TDE transfer skipped.' 'WARN' }
        elseif ($script:Opt.CertPwd -ne $script:Opt.CertPwd2) { throw 'The two certificate passwords do not match.' }
        elseif (-not $script:State.Dest.HasDmk -and -not $script:Opt.DmkPwd) { throw 'The destination has no database master key: enter a DMK password (tab 3).' }
        else {
            $names = ($script:State.MissingCerts | ForEach-Object { "[$($_.Name)]" }) -join ', '
            $acts += "SOURCE $($script:State.Source.ServerName): BACKUP CERTIFICATE $names (+ private key) to a temporary share on $env:COMPUTERNAME"
            $acts += "Copy the .cer/.pvk to $($script:Opt.DestCertDir) on $($script:State.Dest.HostName)$(if (-not $script:State.Dest.HasDmk) { '; CREATE MASTER KEY in master' }); CREATE CERTIFICATE on the DESTINATION; remove the share"
            $doTde = $true
        }
    }
    if (-not $acts.Count) { Show-Info 'Nothing to execute (or the needed passwords/SAS were not entered).'; return }
    if (-not (Confirm-Action ("These changes will be made:`r`n`r`n- " + ($acts -join "`r`n- ") + "`r`n`r`nContinue?"))) { Write-Log 'Execution cancelled.' 'WARN'; return }
    if ($doAz) { Set-Status 'azcopy upload running (the window may not respond until it finishes)...'; Invoke-AzCopyUpload $az }
    if ($doCred.Count) { Invoke-CreateCredential $doCred }
    if ($doTde) { Invoke-TdeTransfer $script:State.MissingCerts }
    Write-Log 'Execution finished. Run preflight again to confirm everything passes.' 'OK'
    Show-Info 'Done. Run preflight again to confirm.'
} })

$btnCreateJob.Add_Click({ Invoke-Ui {
    Read-Options
    if (-not $script:Opt.AllowExec) { throw "Tick 'Allow execution' on tab 3 first." }
    if (-not $script:State.JobScript) { throw 'Generate the scripts first (button 3).' }
    $g = $script:State.GenOpt
    $changed = @()
    if ($script:Opt.Target -ne $g.Target) { $changed += "target name ($($g.Target) -> $($script:Opt.Target))" }
    if ($script:Opt.Db -ne $g.Db) { $changed += "source database ($($g.Db) -> $($script:Opt.Db))" }
    if ($script:Opt.StopAt -ne $g.StopAt) { $changed += "STOPAT ($($g.StopAt.ToString('yyyy-MM-dd hh:mm:ss tt')) -> $($script:Opt.StopAt.ToString('yyyy-MM-dd hh:mm:ss tt')))" }
    if ($script:Opt.JobName -ne $g.JobName) { $changed += "job name" }
    if ($script:Opt.JobRunAt -ne $g.JobRunAt -or $script:Opt.NoSchedule -ne $g.NoSchedule) { $changed += 'job schedule' }
    if ($script:Opt.WithRecovery -ne $g.WithRecovery -or $script:Opt.Replace -ne $g.Replace) { $changed += 'RECOVERY/REPLACE option' }
    if ($changed.Count) { throw ("Options changed after the scripts were generated: " + ($changed -join '; ') + ". Click 'Generate scripts' again (and Build chain first if STOPAT or the database changed), then create the job.") }
    $d = Get-DestInfo $script:Opt.Dest $script:Opt.Target
    if ($d.ServerName -ne $script:State.Dest.ServerName) { throw "Connected to $($d.ServerName), but the job was generated for $($script:State.Dest.ServerName)." }
    if (-not $g.NoSchedule -and $g.JobRunAt -le $d.ServerTime) { throw "The job time ($($g.JobRunAt.ToString('yyyy-MM-dd hh:mm:ss tt'))) is in the past on the destination (now $($d.ServerTime.ToString('yyyy-MM-dd hh:mm:ss tt'))). Change it and generate again." }
    $fails = -1
    if ($script:State.Checks) { $fails = @($script:State.Checks | Where-Object { $_.Status -eq 'FAIL' }).Count }
    $when = 'no schedule'
    if (-not $g.NoSchedule) { $when = "one time at $($g.JobRunAt.ToString('yyyy-MM-dd hh:mm:ss tt')) (destination time)" }
    $fin = 'NORECOVERY'; if ($g.WithRecovery) { $fin = 'RECOVERY' }
    $txt = "Create SQL Agent job on $($d.ServerName)?`r`n`r`nJob: $($script:State.JobName)`r`nOwner: sa`r`nSchedule: $when`r`nRestores [$($g.Db)] as [$($g.Target)], STOPAT $($script:State.Chain.StopAt.ToString('yyyy-MM-dd hh:mm:ss tt')), final state $fin"
    if ($script:Opt.StartNow) { $txt += "`r`n`r`nThe job will be STARTED immediately." }
    if ($fails -gt 0) { $txt += "`r`n`r`nWARNING: preflight has $fails FAIL(s)." } elseif ($fails -lt 0) { $txt += "`r`n`r`nWARNING: preflight was not run." }
    $txt += "`r`n`r`nThis creates exactly the job in 02_create_sql_agent_job.sql (folder: $($script:State.OutDir))."
    if (-not (Confirm-Action $txt)) { Write-Log 'Job creation cancelled.' 'WARN'; return }
    Invoke-SqlBatches $script:Opt.Dest $script:State.JobScript -Database 'msdb'
    Write-Log "Job '$($script:State.JobName)' created on $($d.ServerName) ($when)." 'OK'
    if ($script:Opt.StartNow) {
        Invoke-Sql $script:Opt.Dest 'EXEC msdb.dbo.sp_start_job @job_name = @n;' @{ n = $script:State.JobName } -Database 'msdb' | Out-Null
        Write-Log "Job '$($script:State.JobName)' started. Watch it in Job Activity Monitor." 'OK'
    }
    Show-Info "Job created$(if ($script:Opt.StartNow) { ' and started' })."
} })

# ---------- settings / server cache ----------
$script:SettingsDir  = Join-Path $env:APPDATA 'PITR-RestoreBuilder'
$script:SettingsFile = Join-Path $script:SettingsDir 'settings.json'
$script:DefaultInvQuery = "SELECT ServerName, Environment, Description`r`nFROM dbo.ServerInventory`r`nWHERE IsActive = 1`r`nORDER BY ServerName;"
$script:CmsQuery = "SELECT s.server_name AS ServerName, g.name AS GroupName, s.description AS Description`r`nFROM msdb.dbo.sysmanagement_shared_registered_servers s`r`nJOIN msdb.dbo.sysmanagement_shared_server_groups g ON g.server_group_id = s.server_group_id`r`nORDER BY g.name, s.server_name;"

function Get-Prop {
    param($Obj, [string]$Name, $Default)
    if ($null -ne $Obj -and ($Obj.PSObject.Properties.Name -contains $Name) -and $null -ne $Obj.$Name) { return $Obj.$Name }
    return $Default
}

function Read-Settings {
    $s = [pscustomobject]@{
        Recent = @(); InvPreset = 0; InvServer = ''; InvAuth = 'Windows'; InvUser = ''; InvDatabase = 'DBA'; InvQuery = $script:DefaultInvQuery
        InvServers = @(); InvLoadedAt = ''; LastSource = ''; LastSourceAuth = 'Windows'; LastSourceUser = ''
        LastDest = ''; LastDestAuth = 'Windows'; LastDestUser = ''; OutRoot = ''; StageRoot = ''; DestCertDir = ''
    }
    if (Test-Path -LiteralPath $script:SettingsFile) {
        try {
            $j = Get-Content -LiteralPath $script:SettingsFile -Raw | ConvertFrom-Json
            foreach ($p in @($s.PSObject.Properties.Name)) { $v = Get-Prop $j $p $null; if ($null -ne $v) { $s.$p = $v } }
        } catch { $script:SettingsError = $_.Exception.Message }
    }
    $rec = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($s.Recent)) {
        if ($r -and $r.Name) { $rec.Add([pscustomobject]@{ Name = [string]$r.Name; Auth = [string](Get-Prop $r 'Auth' 'Windows'); User = [string](Get-Prop $r 'User' ''); LastUsed = [string](Get-Prop $r 'LastUsed' '') }) }
    }
    $s.Recent = $rec.ToArray()
    $inv = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($s.InvServers)) {
        if ($r -and $r.Name) { $inv.Add([pscustomobject]@{ Name = [string]$r.Name; Info = [string](Get-Prop $r 'Info' '') }) }
    }
    $s.InvServers = $inv.ToArray()
    return $s
}

function Save-Settings {
    try {
        if (-not (Test-Path -LiteralPath $script:SettingsDir)) { New-Item -ItemType Directory -Path $script:SettingsDir -Force | Out-Null }
        $script:Settings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:SettingsFile -Encoding UTF8
    } catch { Write-Log "Could not save settings to $($script:SettingsFile): $($_.Exception.Message)" 'WARN' }
}

function Update-ServerLists {
    $names = New-Object System.Collections.Generic.List[string]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in @($script:Settings.Recent)) { if ($seen.Add($r.Name)) { $names.Add($r.Name) } }
    foreach ($r in @($script:Settings.InvServers)) { if ($seen.Add($r.Name)) { $names.Add($r.Name) } }
    foreach ($cb in @($txtSrc, $txtDst)) {
        $keep = $cb.Text
        $cb.BeginUpdate(); $cb.Items.Clear()
        foreach ($n in $names) { [void]$cb.Items.Add($n) }
        $cb.EndUpdate(); $cb.Text = $keep
    }
    $cnt = @($script:Settings.InvServers).Count
    $rc = @($script:Settings.Recent).Count
    if ($cnt) { $lblInvInfo.Text = "Cached: $cnt server(s), loaded $($script:Settings.InvLoadedAt); $rc recent." } else { $lblInvInfo.Text = "No server list cached yet; $rc recent server(s)." }
}

function Show-QueryDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Server list query'; $dlg.Size = New-Object System.Drawing.Size(820, 380); $dlg.StartPosition = 'CenterParent'
    $dlg.Font = $form.Font; $dlg.MinimizeBox = $false; $dlg.MaximizeBox = $false
    New-Label $dlg 'First column = server\instance name. Other columns (environment, group, ...) are optional. SELECT only.' 10 8 780
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true; $tb.ScrollBars = 'Both'; $tb.WordWrap = $false; $tb.Font = New-Object System.Drawing.Font('Consolas', 9)
    $tb.Text = $script:InvQueryText
    [void](Add-Ctl $dlg $tb 10 34 785 250)
    $ok = New-Button $dlg 'OK' 600 295 90
    $cancel = New-Button $dlg 'Cancel' 700 295 90
    $reset = New-Button $dlg 'Reset to example' 10 295 140
    $ok.DialogResult = 'OK'; $cancel.DialogResult = 'Cancel'; $dlg.AcceptButton = $null; $dlg.CancelButton = $cancel
    $reset.Add_Click({ if ($cboInvPreset.SelectedIndex -eq 1) { $tb.Text = $script:CmsQuery } else { $tb.Text = $script:DefaultInvQuery } })
    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $script:InvQueryText = $tb.Text; Write-Log 'Server list query updated.' 'OK' }
    $dlg.Dispose()
}

function Add-RecentServer {
    param([string]$Name, [string]$Auth, [string]$User)
    if (-not $Name) { return }
    $u = ''; if ($Auth -eq 'SQL login') { $u = $User }
    $list = New-Object System.Collections.Generic.List[object]
    $list.Add([pscustomobject]@{ Name = $Name; Auth = $Auth; User = $u; LastUsed = (Get-Date -Format 'yyyy-MM-dd HH:mm') })
    foreach ($r in @($script:Settings.Recent)) { if ($r.Name -ne $Name -and $list.Count -lt 30) { $list.Add($r) } }
    $script:Settings.Recent = $list.ToArray()
    Save-Settings
    Update-ServerLists
}

function Set-AuthFromRecent {
    param($Combo, $AuthCombo, $UserBox)
    $n = [string]$Combo.SelectedItem
    if (-not $n) { return }
    $r = @($script:Settings.Recent) | Where-Object { $_.Name -eq $n } | Select-Object -First 1
    if ($r) { $AuthCombo.SelectedItem = $r.Auth; if ($r.User) { $UserBox.Text = $r.User } }
}

# Read-only query for the server list: keyword guard + transaction that is always rolled back
function Invoke-SqlReadOnly {
    param([hashtable]$Conn, [string]$Query, [string]$Database = 'master')
    $blocked = '(?i)\b(insert|update|delete|merge|drop|alter|create|truncate|exec|execute|grant|revoke|deny|backup|restore|dbcc|shutdown|kill|into|openrowset|opendatasource|xp_\w+)\b'
    if ($Query -match $blocked) { throw "The server-list query must be a plain SELECT. Blocked keyword: '$($Matches[1])'." }
    $cn = New-Object System.Data.SqlClient.SqlConnection (New-ConnString -Conn $Conn -Database $Database)
    try {
        $cn.Open()
        $tx = $cn.BeginTransaction()
        try {
            $cmd = $cn.CreateCommand(); $cmd.Transaction = $tx; $cmd.CommandText = $Query; $cmd.CommandTimeout = 60
            $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
            $dt = New-Object System.Data.DataTable
            [void]$da.Fill($dt)
            return , $dt
        } finally { $tx.Rollback() }
    } finally { $cn.Dispose() }
}

$txtSrc.Add_SelectedIndexChanged({ Set-AuthFromRecent $txtSrc $cboSrcAuth $txtSrcUser })
$txtDst.Add_SelectedIndexChanged({ Set-AuthFromRecent $txtDst $cboDstAuth $txtDstUser })

$cboInvPreset.Add_SelectedIndexChanged({
    if ($script:LoadingSettings) { return }
    if ($cboInvPreset.SelectedIndex -eq 1) { $txtInvDb.Text = 'msdb'; $script:InvQueryText = $script:CmsQuery }
    elseif ($script:InvQueryText -eq $script:CmsQuery) { $txtInvDb.Text = 'DBA'; $script:InvQueryText = $script:DefaultInvQuery }
})

$btnInvLoad.Add_Click({ Invoke-Ui {
    $c = @{ Server = $txtInvServer.Text.Trim(); Auth = $cboInvAuth.Text; User = $txtInvUser.Text; Pass = $txtInvPass.Text }
    if (-not $c.Server) { throw 'Enter the server that holds the server list.' }
    $db = $txtInvDb.Text.Trim(); if (-not $db) { $db = 'master' }
    Set-Status "Reading server list from $($c.Server)..."
    $t = Invoke-SqlReadOnly $c $script:InvQueryText $db
    if ($t.Columns.Count -lt 1) { throw 'The query returned no columns.' }
    $list = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $t.Rows) {
        $n = ([string]$r[0]).Trim()
        if ($n -and $seen.Add($n)) {
            $parts = New-Object System.Collections.Generic.List[string]
            for ($i = 1; $i -lt $t.Columns.Count; $i++) { $v = [string]$r[$i]; if ($v) { $parts.Add($v) } }
            $list.Add([pscustomobject]@{ Name = $n; Info = ($parts -join ' | ') })
        }
    }
    $s = $script:Settings
    $s.InvServers = $list.ToArray(); $s.InvLoadedAt = ('{0:yyyy-MM-dd HH:mm} from {1}.{2}' -f (Get-Date), $c.Server, $db)
    $s.InvPreset = $cboInvPreset.SelectedIndex; $s.InvServer = $c.Server; $s.InvAuth = $c.Auth; $s.InvUser = $c.User
    $s.InvDatabase = $db; $s.InvQuery = $script:InvQueryText
    Save-Settings
    Update-ServerLists
    Write-Log "Loaded $($list.Count) server name(s) from $($c.Server).$db and cached them in $($script:SettingsFile)." 'OK'
} })

$btnInvQuery.Add_Click({ Show-QueryDialog })
$btnRecentClear.Add_Click({
    if (Confirm-Action 'Remove all recent servers from the cache?') { $script:Settings.Recent = @(); Save-Settings; Update-ServerLists }
})

$form.Add_FormClosing({
    $s = $script:Settings
    $s.LastSource = $txtSrc.Text.Trim(); $s.LastSourceAuth = $cboSrcAuth.Text; $s.LastSourceUser = $txtSrcUser.Text
    $s.LastDest = $txtDst.Text.Trim(); $s.LastDestAuth = $cboDstAuth.Text; $s.LastDestUser = $txtDstUser.Text
    $s.OutRoot = $txtOut.Text.Trim(); $s.StageRoot = $txtStage.Text.Trim(); $s.DestCertDir = $txtDestCert.Text.Trim()
    $s.InvPreset = $cboInvPreset.SelectedIndex; $s.InvServer = $txtInvServer.Text.Trim(); $s.InvAuth = $cboInvAuth.Text
    $s.InvUser = $txtInvUser.Text; $s.InvDatabase = $txtInvDb.Text.Trim(); $s.InvQuery = $script:InvQueryText
    Save-Settings
})

# ---------- startup: restore settings ----------
$script:Settings = Read-Settings
$st = $script:Settings
$script:LoadingSettings = $true
$cboInvPreset.SelectedIndex = [Math]::Max(0, [Math]::Min(1, [int]$st.InvPreset))
$script:InvQueryText = [string]$st.InvQuery
if (-not $script:InvQueryText) { $script:InvQueryText = $script:DefaultInvQuery }
if ($st.InvServer) { $txtInvServer.Text = $st.InvServer } else { $txtInvServer.Text = $env:COMPUTERNAME }; $cboInvAuth.SelectedItem = $st.InvAuth; $txtInvUser.Text = $st.InvUser; $txtInvDb.Text = $st.InvDatabase
$txtSrc.Text = $st.LastSource; $cboSrcAuth.SelectedItem = $st.LastSourceAuth; $txtSrcUser.Text = $st.LastSourceUser
$txtDst.Text = $st.LastDest; $cboDstAuth.SelectedItem = $st.LastDestAuth; $txtDstUser.Text = $st.LastDestUser
if ($st.OutRoot) { $txtOut.Text = $st.OutRoot }
if ($st.StageRoot) { $txtStage.Text = $st.StageRoot }
if ($st.DestCertDir) { $txtDestCert.Text = $st.DestCertDir }
$script:LoadingSettings = $false
Update-ServerLists

Update-MemLabel
if ($script:SettingsError) { Write-Log "Settings file could not be read ($script:SettingsError); starting with defaults." 'WARN' }
Write-Log "Settings: $script:SettingsFile" 'INFO'
Write-Log "$script:AppName $script:Version started on $env:COMPUTERNAME as $env:USERDOMAIN\$env:USERNAME. Read-only mode." 'INFO'
[void]$form.ShowDialog()
$form.Dispose()

#endregion

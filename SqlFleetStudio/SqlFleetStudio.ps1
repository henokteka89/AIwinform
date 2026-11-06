<#
.SYNOPSIS
    SQL Fleet Studio - an SSMS-style query tool that runs T-SQL against many SQL Servers at once.

.DESCRIPTION
    * Server list: pulled from a central server (CMS or your own inventory table) and cached locally.
      Filter it, check the servers you want, run. No more reconnecting window by window.
    * Query library: save the scripts you use all the time; search, copy, or send them to the editor.
    * Execution: runs in parallel on every checked server; results are merged into one grid with a
      ServerName column. Messages, errors and per-server status are shown separately.

    PRODUCTION SAFETY (on by default)
    * READ-ONLY mode: write statements are blocked before anything is sent, AND every run is wrapped
      in a transaction that is always rolled back - so even something that slips past the check is undone.
      Temp tables (#t) and table variables (@t) are allowed, so normal DBA diagnostic scripts work.
    * Writes must be armed per run ("Allow writes (this run only)") and confirmed by typing EXECUTE.
      The switch turns itself off after the run.
    * Every session runs with SET DEADLOCK_PRIORITY LOW and a LOCK_TIMEOUT, so production work wins.
    * Query timeout, max rows per result set, parallelism limit, cancel (Esc).
    * Every execution (who, when, where, what, outcome) is appended to %APPDATA%\SqlFleetStudio\execution.log.

.NOTES
    Requires Windows PowerShell 5.1 (System.Data.SqlClient ships with .NET Framework).
    Run:  powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\SqlFleetStudio.ps1
    Keys: F5 or Ctrl+E = execute (selection only, if text is selected), Esc = cancel, Ctrl+A = select all.
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Any error anywhere in the script (including startup) is reported with its exact line and logged.
trap {
    $msg = "{0}`r`n`r`n{1}`r`n`r`n{2}" -f $_, $_.InvocationInfo.PositionMessage, $_.ScriptStackTrace
    try {
        $logDir = Join-Path $env:APPDATA 'SqlFleetStudio'
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
        Add-Content -Path (Join-Path $logDir 'error.log') -Value ("==== {0}  script error  PS {1}`r`n{2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $PSVersionTable.PSVersion, $msg)
    } catch { }
    Write-Host $msg -ForegroundColor Red
    try { [void][System.Windows.Forms.MessageBox]::Show("SQL Fleet Studio hit an error:`r`n`r`n$msg", 'SQL Fleet Studio') } catch { }
    break
}

$AppName    = 'SQL Fleet Studio'
$AppVersion = '1.0'

# ------------------------------------------------------------------ prerequisites
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName Microsoft.VisualBasic

if (-not ('System.Data.SqlClient.SqlConnection' -as [type])) {
    [void][System.Windows.Forms.MessageBox]::Show('System.Data.SqlClient is not available in this PowerShell. Run the script with Windows PowerShell 5.1 (powershell.exe).', $AppName)
    return
}
[System.Windows.Forms.Application]::EnableVisualStyles()
# Anything that escapes an event handler is logged and reported instead of closing the app.
try { [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException) } catch { }

# ------------------------------------------------------------------ storage
$DataDir = Join-Path $env:APPDATA 'SqlFleetStudio'
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Path $DataDir | Out-Null }
$Paths = @{
    Settings = Join-Path $DataDir 'settings.json'
    Servers  = Join-Path $DataDir 'servers.json'
    Library  = Join-Path $DataDir 'library.json'
    Draft    = Join-Path $DataDir 'draft.sql'
    Log      = Join-Path $DataDir 'execution.log'
    ErrorLog = Join-Path $DataDir 'error.log'
}

$DefaultServerQuery = @'
-- First column = server name. Any extra columns (Group, Environment, ...) become filterable columns.
-- Default below reads a Central Management Server. Replace with your own inventory table if you have one, e.g.
-- SELECT ServerName, Environment, Application FROM dbo.ServerInventory WHERE IsActive = 1 ORDER BY ServerName;
SELECT s.server_name AS ServerName,
       g.name        AS [Group],
       s.description AS Description
FROM msdb.dbo.sysmanagement_shared_registered_servers AS s
JOIN msdb.dbo.sysmanagement_shared_server_groups AS g
     ON g.server_group_id = s.server_group_id
ORDER BY g.name, s.server_name;
'@

$DefaultSettings = [ordered]@{
    CentralServer          = ''
    CentralDatabase        = 'msdb'
    ServerListQuery        = $DefaultServerQuery
    AuthMode               = 'Windows'
    SqlUser                = ''
    Encrypt                = $false
    TrustServerCertificate = $true
    ReadOnlyIntent         = $false
    MultiSubnetFailover    = $false
    ConnectTimeout         = 15
    QueryTimeout           = 300
    MaxRows                = 50000
    MaxParallel            = 8
    LockTimeoutMs          = 10000
    ProdPattern            = '(?i)prod|prd'
    ConfirmServerCount     = 10
    DefaultDatabase        = 'master'
}

# ------------------------------------------------------------------ state
$State = @{
    Settings      = $null
    Columns       = @('Server')
    Servers       = New-Object System.Collections.Generic.List[object]
    Manual        = New-Object System.Collections.Generic.List[string]
    RefreshedAt   = ''
    Source        = ''
    Checked       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    ProdSet       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    VisibleSet    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    Library       = New-Object System.Collections.Generic.List[object]
    LibSelectedId = $null
    Rebuilding    = $false
    Running       = $false
    Jobs          = $null
    Pool          = $null
    Shared        = $null
    ResultMap     = $null
    StatusItems   = @{}
    RunStart      = $null
    RunInfo       = $null
    SqlCredential = $null
}

# ------------------------------------------------------------------ read-only guard
# Comments, string literals, [bracketed] and "quoted" identifiers are stripped before checking.
$StripPattern = @'
/\*[\s\S]*?\*/|--[^\r\n]*|N?'(?:[^']|'')*'|\[(?:[^\]]|\]\])*\]|"(?:[^"]|"")*"
'@

$BlockRules = @(
    @{ N = 'INSERT';                    R = '\bINSERT(?>\s+(?:TOP\s*\([^)]*\)\s*)?(?:INTO\s+)?)(?![#@])' }
    @{ N = 'UPDATE';                    R = '\bUPDATE(?>\s+(?:TOP\s*\([^)]*\)\s*)?)(?![#@])' }
    @{ N = 'DELETE';                    R = '\bDELETE(?>\s+(?:TOP\s*\([^)]*\)\s*)?(?:FROM\s+)?)(?![#@])' }
    @{ N = 'MERGE';                     R = '\bMERGE(?>\s+(?:TOP\s*\([^)]*\)\s*)?(?:INTO\s+)?)(?![#@])' }
    @{ N = 'TRUNCATE';                  R = '\bTRUNCATE\s+TABLE(?>\s+)(?![#@])' }
    @{ N = 'SELECT INTO';               R = '\bINTO(?>\s+)(?![#@])' }
    @{ N = 'CREATE';                    R = '\bCREATE(?>\s+)(?!TABLE\s+#)(?!(?:UNIQUE\s+)?(?:(?:NON)?CLUSTERED\s+)?INDEX\s+\S+\s+ON\s+#)' }
    @{ N = 'ALTER';                     R = '\bALTER(?>\s+)(?!TABLE\s+#)' }
    @{ N = 'DROP';                      R = '\bDROP(?>\s+)(?!TABLE\s+(?:IF\s+EXISTS\s+)?#)' }
    @{ N = 'GRANT/REVOKE/DENY';         R = '\b(?:GRANT|REVOKE|DENY)\b' }
    @{ N = 'BACKUP';                    R = '\bBACKUP\s+(?:DATABASE|LOG|CERTIFICATE|MASTER|SERVICE|SYMMETRIC)\b' }
    @{ N = 'RESTORE';                   R = '\bRESTORE\s+(?:DATABASE|LOG)\b' }
    @{ N = 'KILL';                      R = '\bKILL\s+\d' }
    @{ N = 'SHUTDOWN';                  R = '\bSHUTDOWN\b' }
    @{ N = 'RECONFIGURE';               R = '\bRECONFIGURE\b' }
    @{ N = 'BULK INSERT';               R = '\bBULK\s+INSERT\b' }
    @{ N = 'transaction control';       R = '\b(?:COMMIT|ROLLBACK|SAVE\s+TRAN(?:SACTION)?|BEGIN\s+(?:DISTRIBUTED\s+)?TRAN(?:SACTION)?)\b' }
    @{ N = 'ENABLE/DISABLE TRIGGER';    R = '\b(?:ENABLE|DISABLE)\s+TRIGGER\b' }
    @{ N = 'xp_cmdshell';               R = '\bxp_cmdshell\b' }
    @{ N = 'sp_configure';              R = '\bsp_configure\b' }
    @{ N = 'OPENROWSET/OPENDATASOURCE'; R = '\b(?:OPENROWSET|OPENDATASOURCE)\b' }
)
$WarnRules = @(
    @{ N = 'EXEC / EXECUTE';            R = '\bEXEC(?:UTE)?\b' }
    @{ N = 'sp_executesql';             R = '\bsp_executesql\b' }
    @{ N = 'DBCC';                      R = '\bDBCC\b' }
    @{ N = 'OPENQUERY (linked server)'; R = '\bOPENQUERY\b' }
)

function Test-ReadOnlySql([string]$Sql) {
    $clean = [regex]::Replace($Sql, $StripPattern, ' ')
    $opt = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $blocked = @(foreach ($r in $BlockRules) { if ([regex]::IsMatch($clean, $r.R, $opt)) { $r.N } })
    $warn    = @(foreach ($r in $WarnRules)  { if ([regex]::IsMatch($clean, $r.R, $opt)) { $r.N } })
    [pscustomobject]@{ Blocked = $blocked; Warnings = $warn }
}

function Split-SqlBatches([string]$Sql) {
    # Splits on SSMS-style GO lines (GO, GO 5).
    $out = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    foreach ($line in ($Sql -split "\r?\n")) {
        $m = [regex]::Match($line, '^\s*GO(?:\s+(\d+))?\s*(?:--.*)?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            $b = $sb.ToString()
            if ($b.Trim()) {
                $times = 1
                if ($m.Groups[1].Success) { $times = [Math]::Max(1, [Math]::Min([int]$m.Groups[1].Value, 1000)) }
                for ($i = 0; $i -lt $times; $i++) { $out.Add($b) }
            }
            [void]$sb.Clear()
        } else {
            [void]$sb.AppendLine($line)
        }
    }
    $b = $sb.ToString()
    if ($b.Trim()) { $out.Add($b) }
    $out.ToArray()
}

# ------------------------------------------------------------------ worker (runs in a background runspace, one per server)
$WorkerScript = {
    param($Server, $ConnStr, $Cred, [string[]]$Batches, [bool]$ReadOnly, [int]$Timeout, [int]$MaxRows, [int]$LockTimeoutMs, $Shared)

    $ErrorActionPreference = 'Stop'
    $msgs   = New-Object System.Collections.ArrayList
    $tables = New-Object System.Collections.ArrayList
    $res = [pscustomobject]@{
        Server = $Server; Tables = $tables; Messages = $msgs; Error = $null
        Cancelled = $false; Truncated = $false; Rows = 0; RowsAffected = 0; Ms = 0
    }
    $hex = {
        param([byte[]]$b)
        $max = [Math]::Min($b.Length, 64)
        $s = '0x' + ([BitConverter]::ToString($b, 0, $max)).Replace('-', '')
        if ($b.Length -gt 64) { $s += ('... ({0} bytes)' -f $b.Length) }
        $s
    }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $conn = $null
    $tx = $null
    $tmp = $null
    try {
        if ($Shared.ContainsKey('__cancel')) {
            $res.Cancelled = $true
            [void]$msgs.Add('Cancelled before start.')
            return $res
        }
        $Shared["state:$Server"] = 'Connecting'
        $conn = New-Object System.Data.SqlClient.SqlConnection $ConnStr
        if ($Cred) { $conn.Credential = $Cred }
        $conn.FireInfoMessageEventOnUserErrors = $false
        $onInfo = { param($s, $e) foreach ($er in $e.Errors) { [void]$msgs.Add($er.Message) } }.GetNewClosure()
        $conn.add_InfoMessage($onInfo)
        $conn.Open()
        $Shared["state:$Server"] = 'Running'

        # Be a polite guest on production: lose deadlocks, don't wait forever on locks.
        $pre = $conn.CreateCommand()
        $pre.CommandText = "SET DEADLOCK_PRIORITY LOW; SET LOCK_TIMEOUT $LockTimeoutMs;"
        [void]$pre.ExecuteNonQuery()

        if ($ReadOnly) { $tx = $conn.BeginTransaction() }

        $onDone = { param($s, $e) [void]$msgs.Add(('({0} row(s) affected)' -f $e.RecordCount)) }.GetNewClosure()
        $stop = $false
        $resultIndex = 0

        foreach ($batch in $Batches) {
            if ($stop) { break }
            if ($Shared.ContainsKey('__cancel')) { $res.Cancelled = $true; break }

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $batch
            $cmd.CommandTimeout = $Timeout
            if ($tx) { $cmd.Transaction = $tx }
            $cmd.add_StatementCompleted($onDone)
            $Shared["cmd:$Server"] = $cmd

            $reader = $cmd.ExecuteReader()
            try {
                do {
                    if ($reader.FieldCount -gt 0) {
                        $n = $reader.FieldCount
                        $dt = New-Object System.Data.DataTable ("R$resultIndex")
                        $used = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                        $binCols = New-Object System.Collections.Generic.List[int]
                        for ($i = 0; $i -lt $n; $i++) {
                            $name = $reader.GetName($i)
                            if ([string]::IsNullOrWhiteSpace($name)) { $name = '(No column name)' }
                            $base = $name; $k = 2
                            while (-not $used.Add($name)) { $name = "$base ($k)"; $k++ }
                            $t = $null
                            try { $t = $reader.GetFieldType($i) } catch { $t = $null }
                            if ($t -eq [byte[]]) { $binCols.Add($i); $t = [string] }
                            if ($null -eq $t) { $t = [string] }
                            [void]$dt.Columns.Add($name, $t)
                        }
                        $srvName = 'ServerName'; $k = 2
                        while ($used.Contains($srvName)) { $srvName = "ServerName$k"; $k++ }
                        $srvCol = $dt.Columns.Add($srvName, [string])
                        $srvCol.DefaultValue = $Server

                        $vals = New-Object object[] $n
                        $count = 0
                        $dt.BeginLoadData()
                        while ($reader.Read()) {
                            if ($count -ge $MaxRows) { $res.Truncated = $true; $stop = $true; break }
                            try { [void]$reader.GetValues($vals) }
                            catch {
                                for ($i = 0; $i -lt $n; $i++) {
                                    try { $vals[$i] = $reader.GetValue($i) } catch { $vals[$i] = '<unsupported type>' }
                                }
                            }
                            foreach ($bi in $binCols) { $v = $vals[$bi]; if ($v -is [byte[]]) { $vals[$bi] = & $hex $v } }
                            [void]$dt.Rows.Add($vals)
                            $count++
                        }
                        $dt.EndLoadData()
                        $srvCol.SetOrdinal(0)
                        $dt.ExtendedProperties['Index'] = $resultIndex
                        [void]$tables.Add($dt)
                        $res.Rows += $count
                        $resultIndex++
                        if ($stop) { break }
                    }
                } while ($reader.NextResult())
            } finally {
                if ($stop) { try { $cmd.Cancel() } catch { } }
                try { $reader.Close() } catch { }
                try { if ($reader.RecordsAffected -gt 0) { $res.RowsAffected += $reader.RecordsAffected } } catch { }
                $Shared["cmd:$Server"] = $null
            }
        }
        if ($res.Truncated) { [void]$msgs.Add(('Stopped at {0} rows (Max rows). Remaining rows/results were cancelled.' -f $MaxRows)) }
    } catch {
        $ex = $_.Exception
        while ($ex -and -not ($ex -is [System.Data.SqlClient.SqlException]) -and $ex.InnerException) { $ex = $ex.InnerException }
        if ($Shared.ContainsKey('__cancel')) { $res.Cancelled = $true }
        if ($ex -is [System.Data.SqlClient.SqlException]) {
            $lines = foreach ($er in $ex.Errors) {
                'Msg {0}, Level {1}, State {2}, Line {3}' -f $er.Number, $er.Class, $er.State, $er.LineNumber
                $er.Message
            }
            $res.Error = ($lines -join "`r`n")
        } else {
            $res.Error = $ex.Message
        }
    } finally {
        if ($tx) { try { $tx.Rollback() } catch { } }
        if ($conn) { try { $conn.Dispose() } catch { } }
        $sw.Stop()
        $res.Ms = $sw.ElapsedMilliseconds
    }
    $res
}

# ------------------------------------------------------------------ helpers
function Read-JsonFile([string]$Path) {
    if (Test-Path $Path) {
        try {
            $raw = Get-Content -Path $Path -Raw -Encoding UTF8
            if ($raw -and $raw.Trim()) { return ($raw | ConvertFrom-Json) }
        } catch { }
    }
    return $null
}

function Write-JsonFile([string]$Path, $Object) {
    $json = ConvertTo-Json -InputObject $Object -Depth 6
    $tmp = "$Path.tmp"
    Set-Content -Path $tmp -Value $json -Encoding UTF8
    Move-Item -Path $tmp -Destination $Path -Force
}

# Windows PowerShell 5.1 can throw "Argument types do not match" on @() around .NET generic
# collections (List[T], HashSet[T], dictionary key/value collections). Copy them explicitly instead.
function ConvertTo-PlainArray($Collection) {
    $list = New-Object System.Collections.ArrayList
    if ($null -ne $Collection) { foreach ($x in $Collection) { [void]$list.Add($x) } }
    , $list.ToArray()
}

function ConvertTo-CrLf([string]$s) {
    if ($null -eq $s) { return '' }
    ($s -replace "`r`n", "`n") -replace "`n", "`r`n"
}

function Get-InnerMessage($err) {
    $ex = if ($err -is [System.Management.Automation.ErrorRecord]) { $err.Exception } else { $err }
    while ($ex.InnerException -and ($ex -is [System.Management.Automation.MethodInvocationException] -or $ex -is [System.Management.Automation.RuntimeException])) { $ex = $ex.InnerException }
    $ex.Message
}

function Show-Info([string]$m) { [void][System.Windows.Forms.MessageBox]::Show($form, $m, $AppName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) }
function Show-Warn([string]$m) { [void][System.Windows.Forms.MessageBox]::Show($form, $m, $AppName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) }
function Show-Err([string]$m)  { [void][System.Windows.Forms.MessageBox]::Show($form, $m, $AppName, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
function Ask-YesNo([string]$m) {
    [System.Windows.Forms.MessageBox]::Show($form, $m, $AppName, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question) -eq [System.Windows.Forms.DialogResult]::Yes
}

function Write-ErrorLog($Err, [string]$Where = '') {
    try {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add(('==== {0}  {1}  PS {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Where, $PSVersionTable.PSVersion))
        if ($Err -is [System.Management.Automation.ErrorRecord]) {
            $lines.Add([string]$Err)
            if ($Err.InvocationInfo) { $lines.Add([string]$Err.InvocationInfo.PositionMessage) }
            if ($Err.ScriptStackTrace) { $lines.Add('Script stack:'); $lines.Add([string]$Err.ScriptStackTrace) }
            $ex = $Err.Exception
        } else { $ex = $Err }
        while ($ex) { $lines.Add(('{0}: {1}' -f $ex.GetType().FullName, $ex.Message)); if ($ex.StackTrace) { $lines.Add([string]$ex.StackTrace) }; $ex = $ex.InnerException }
        Add-Content -Path $Paths.ErrorLog -Value $lines -Encoding UTF8
    } catch { }
}

function Invoke-Safe([scriptblock]$Action, [string]$Where = 'handler') {
    try { & $Action }
    catch {
        Write-ErrorLog $_ $Where
        try { $form.Cursor = [System.Windows.Forms.Cursors]::Default } catch { }
        try { Show-Err ((Get-InnerMessage $_) + "`r`n`r`n(Details written to $($Paths.ErrorLog))") } catch { }
    }
}

function Write-ExecLog([System.Collections.IDictionary]$Entry) {
    try {
        $o = [ordered]@{ Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); User = "$env:USERDOMAIN\$env:USERNAME"; Machine = $env:COMPUTERNAME }
        foreach ($k in $Entry.Keys) { $o[$k] = $Entry[$k] }
        Add-Content -Path $Paths.Log -Value (ConvertTo-Json -InputObject $o -Compress -Depth 4) -Encoding UTF8
    } catch { }
}

# ------------------------------------------------------------------ settings / cache / library persistence
function Load-Settings {
    $s = [ordered]@{}
    foreach ($k in $DefaultSettings.Keys) { $s[$k] = $DefaultSettings[$k] }
    $f = Read-JsonFile $Paths.Settings
    if ($f) { foreach ($p in $f.PSObject.Properties) { if ($s.Contains($p.Name)) { $s[$p.Name] = $p.Value } } }
    $s
}

function Save-Settings { Write-JsonFile $Paths.Settings $State.Settings }

function Load-ServerCache {
    $c = Read-JsonFile $Paths.Servers
    $State.Servers.Clear(); $State.Manual.Clear()
    if (-not $c) { return }
    if ($c.Columns) { $State.Columns = @($c.Columns | ForEach-Object { [string]$_ }) }
    foreach ($s in @($c.Servers)) { if ($s) { $State.Servers.Add($s) } }
    foreach ($m in @($c.Manual))  { if ($m) { $State.Manual.Add([string]$m) } }
    foreach ($x in @($c.Checked)) { if ($x) { [void]$State.Checked.Add([string]$x) } }
    $State.RefreshedAt = [string]$c.RefreshedAt
    $State.Source = [string]$c.Source
}

function Save-ServerCache {
    $obj = [ordered]@{
        RefreshedAt = $State.RefreshedAt
        Source      = $State.Source
        Columns     = @($State.Columns)
        Servers     = (ConvertTo-PlainArray $State.Servers)
        Manual      = (ConvertTo-PlainArray $State.Manual)
        Checked     = (ConvertTo-PlainArray $State.Checked)
    }
    Write-JsonFile $Paths.Servers $obj
}

function New-LibEntry([string]$Name, [string]$Category, [string]$Sql, [string]$Id = $null) {
    if (-not $Id) { $Id = [guid]::NewGuid().ToString() }
    [pscustomobject]@{ Id = $Id; Name = $Name; Category = $Category; Sql = $Sql; Updated = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
}

function Get-SeedLibrary {
    $seed = New-Object System.Collections.Generic.List[object]
    $seed.Add((New-LibEntry 'Server / version info' 'Health' @'
SELECT @@SERVERNAME                      AS server_name,
       SERVERPROPERTY('ProductVersion')  AS product_version,
       SERVERPROPERTY('ProductLevel')    AS product_level,
       SERVERPROPERTY('Edition')         AS edition,
       SERVERPROPERTY('IsHadrEnabled')   AS hadr_enabled,
       sqlserver_start_time
FROM sys.dm_os_sys_info;
'@))
    $seed.Add((New-LibEntry 'Database sizes' 'Storage' @'
SELECT d.name AS database_name, d.state_desc, d.recovery_model_desc,
       CAST(SUM(CASE WHEN mf.type = 0 THEN mf.size END) * 8 / 1024.0 AS decimal(18,1)) AS data_mb,
       CAST(SUM(CASE WHEN mf.type = 1 THEN mf.size END) * 8 / 1024.0 AS decimal(18,1)) AS log_mb
FROM sys.databases AS d
JOIN sys.master_files AS mf ON mf.database_id = d.database_id
GROUP BY d.name, d.state_desc, d.recovery_model_desc
ORDER BY d.name;
'@))
    $seed.Add((New-LibEntry 'Disk free space' 'Storage' @'
SELECT DISTINCT vs.volume_mount_point, vs.logical_volume_name,
       CAST(vs.total_bytes     / 1073741824.0 AS decimal(18,1)) AS total_gb,
       CAST(vs.available_bytes / 1073741824.0 AS decimal(18,1)) AS free_gb,
       CAST(vs.available_bytes * 100.0 / vs.total_bytes AS decimal(5,1)) AS free_pct
FROM sys.master_files AS mf
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) AS vs
ORDER BY vs.volume_mount_point;
'@))
    $seed.Add((New-LibEntry 'Last backups per database' 'Backups' @'
SELECT d.name AS database_name, d.recovery_model_desc,
       MAX(CASE WHEN b.type = 'D' THEN b.backup_finish_date END) AS last_full,
       MAX(CASE WHEN b.type = 'I' THEN b.backup_finish_date END) AS last_diff,
       MAX(CASE WHEN b.type = 'L' THEN b.backup_finish_date END) AS last_log
FROM sys.databases AS d
LEFT JOIN msdb.dbo.backupset AS b ON b.database_name = d.name
WHERE d.name <> 'tempdb'
GROUP BY d.name, d.recovery_model_desc
ORDER BY d.name;
'@))
    $seed.Add((New-LibEntry 'Active requests' 'Activity' @'
SELECT r.session_id, r.status, r.command, DB_NAME(r.database_id) AS database_name,
       r.blocking_session_id, r.wait_type, r.wait_time, r.cpu_time, r.total_elapsed_time,
       s.login_name, s.host_name, s.program_name,
       SUBSTRING(t.text, r.statement_start_offset / 2 + 1,
                 (CASE r.statement_end_offset WHEN -1 THEN DATALENGTH(t.text) ELSE r.statement_end_offset END
                  - r.statement_start_offset) / 2 + 1) AS statement_text
FROM sys.dm_exec_requests AS r
JOIN sys.dm_exec_sessions AS s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) AS t
WHERE r.session_id <> @@SPID AND s.is_user_process = 1
ORDER BY r.total_elapsed_time DESC;
'@))
    $seed.Add((New-LibEntry 'Blocking' 'Activity' @'
SELECT r.session_id, r.blocking_session_id, r.wait_type, r.wait_time, r.wait_resource,
       DB_NAME(r.database_id) AS database_name, t.text AS sql_text
FROM sys.dm_exec_requests AS r
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) AS t
WHERE r.blocking_session_id <> 0
ORDER BY r.wait_time DESC;
'@))
    $seed.Add((New-LibEntry 'Failed Agent jobs (last 24h)' 'Agent' @'
SELECT j.name AS job_name, h.step_id, h.step_name,
       msdb.dbo.agent_datetime(h.run_date, h.run_time) AS run_datetime,
       h.message
FROM msdb.dbo.sysjobhistory AS h
JOIN msdb.dbo.sysjobs AS j ON j.job_id = h.job_id
WHERE h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -24, GETDATE())
ORDER BY run_datetime DESC;
'@))
    $seed.Add((New-LibEntry 'AG health' 'HADR' @'
SELECT ag.name AS ag_name, ar.replica_server_name, ars.role_desc, ars.synchronization_health_desc,
       DB_NAME(drs.database_id) AS database_name, drs.synchronization_state_desc,
       drs.log_send_queue_size, drs.redo_queue_size
FROM sys.availability_groups AS ag
JOIN sys.availability_replicas AS ar ON ar.group_id = ag.group_id
JOIN sys.dm_hadr_availability_replica_states AS ars ON ars.replica_id = ar.replica_id
LEFT JOIN sys.dm_hadr_database_replica_states AS drs ON drs.replica_id = ar.replica_id
ORDER BY ag.name, ar.replica_server_name;
'@))
    $seed
}

function Load-Library {
    $State.Library.Clear()
    if (-not (Test-Path $Paths.Library)) {
        foreach ($q in (Get-SeedLibrary)) { $State.Library.Add($q) }
        Save-Library
        return
    }
    $raw = Read-JsonFile $Paths.Library
    foreach ($q in @($raw)) {
        if (-not $q) { continue }
        $e = New-LibEntry ([string]$q.Name) ([string]$q.Category) (ConvertTo-CrLf ([string]$q.Sql)) ([string]$q.Id)
        if ($q.Updated) { $e.Updated = [string]$q.Updated }
        $State.Library.Add($e)
    }
}

function Save-Library {
    $items = ConvertTo-PlainArray $State.Library
    Write-JsonFile $Paths.Library $items
}

function Find-LibEntry([string]$Id) {
    foreach ($q in $State.Library) { if ($q.Id -eq $Id) { return $q } }
    $null
}

# ------------------------------------------------------------------ SQL helpers
function New-SqlConnectionString([string]$Server, [string]$Database, [switch]$NoIntent) {
    $S = $State.Settings
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    # PowerShell treats the builder as a dictionary, so $b.DataSource would be taken as a keyword
    # named "DataSource". Set the real connection-string keywords through the indexer instead.
    $b['Data Source'] = $Server
    $b['Initial Catalog'] = $Database
    $b['Application Name'] = 'SqlFleetStudio'
    $b['Connect Timeout'] = [int]$S.ConnectTimeout
    $b['Encrypt'] = [bool]$S.Encrypt
    $b['TrustServerCertificate'] = [bool]$S.TrustServerCertificate
    $b['MultiSubnetFailover'] = [bool]$S.MultiSubnetFailover
    if ($S.AuthMode -ne 'SQL') { $b['Integrated Security'] = $true }
    if ($S.ReadOnlyIntent -and -not $NoIntent) { $b['ApplicationIntent'] = 'ReadOnly' }
    $b.ToString()
}

function Get-SqlCred {
    if ($State.Settings.AuthMode -ne 'SQL') { return $null }
    if (-not $State.SqlCredential) {
        $gc = @{ Message = 'SQL login for SQL Fleet Studio (kept in memory only, never saved)' }
        if ($State.Settings.SqlUser) { $gc['UserName'] = [string]$State.Settings.SqlUser }
        $c = $null
        try { $c = Get-Credential @gc } catch { $c = $null }
        if (-not $c) { return $null }
        $pw = $c.Password.Copy()
        $pw.MakeReadOnly()
        $State.SqlCredential = [System.Data.SqlClient.SqlCredential]::new($c.UserName.TrimStart('\'), $pw)
    }
    $State.SqlCredential
}

function Invoke-SqlSync([string]$Server, [string]$Database, [string]$Query, [int]$Timeout = 60) {
    $cred = $null
    if ($State.Settings.AuthMode -eq 'SQL') {
        $cred = Get-SqlCred
        if (-not $cred) { throw 'No SQL credential provided.' }
    }
    $conn = New-Object System.Data.SqlClient.SqlConnection (New-SqlConnectionString $Server $Database -NoIntent)
    if ($cred) { $conn.Credential = $cred }
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt = New-Object System.Data.DataTable
        [void]$da.Fill($dt)
        , $dt
    } finally {
        $conn.Dispose()
    }
}

# ------------------------------------------------------------------ UI construction helpers
$FontUi     = New-Object System.Drawing.Font('Segoe UI', 9)
$FontBold   = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$FontEditor = New-Object System.Drawing.Font('Consolas', 10.5)
$FontGrid   = New-Object System.Drawing.Font('Consolas', 9)

function New-Button([string]$Text, [int]$Width = 90) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Width = $Width; $b.Height = 27
    $b.Margin = New-Object System.Windows.Forms.Padding(2, 3, 2, 3)
    $b
}
function New-Label([string]$Text) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(8, 8, 2, 2)
    $l
}
function New-Num([int]$Min, [int]$Max, [int]$Value, [int]$Width = 80) {
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Minimum = $Min; $n.Maximum = $Max
    $n.Value = [Math]::Min([Math]::Max($Value, $Min), $Max)
    $n.Width = $Width; $n.ThousandsSeparator = $true
    $n.Margin = New-Object System.Windows.Forms.Padding(2, 5, 2, 2)
    $n
}
function New-Flow {
    $f = New-Object System.Windows.Forms.FlowLayoutPanel
    $f.Dock = 'Top'; $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'; $f.WrapContents = $true
    $f.Padding = New-Object System.Windows.Forms.Padding(3)
    $f
}
function New-CodeBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true; $t.ScrollBars = 'Both'; $t.WordWrap = $false
    $t.AcceptsTab = $true; $t.AcceptsReturn = $true; $t.HideSelection = $false
    $t.MaxLength = 0; $t.Font = $FontEditor; $t.Dock = 'Fill'
    $t.add_KeyDown({
        param($s, $e)
        Invoke-Safe {
            if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true }
        } 'editor KeyDown'
    })
    $t
}

$ToolTip = New-Object System.Windows.Forms.ToolTip

# ------------------------------------------------------------------ main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "$AppName $AppVersion"
$form.Size = New-Object System.Drawing.Size(1500, 920)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 600)
$form.StartPosition = 'CenterScreen'
$form.Font = $FontUi
$form.KeyPreview = $true

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Spring = $true
$lblStatus.TextAlign = 'MiddleLeft'
$lblElapsed = New-Object System.Windows.Forms.ToolStripStatusLabel
[void]$statusStrip.Items.Add($lblStatus)
[void]$statusStrip.Items.Add($lblElapsed)

$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = 'Fill'
$splitMain.FixedPanel = 'Panel1'

# ---------------- left: servers
$pnlSrvTop = New-Flow
$lblSrvTitle = New-Label 'Servers'
$lblSrvTitle.Font = $FontBold
$btnRefresh = New-Button 'Refresh' 70
$ToolTip.SetToolTip($btnRefresh, 'Re-run the server list query on the central server and update the cache')
$btnAddSrv = New-Button 'Add...' 60
$ToolTip.SetToolTip($btnAddSrv, 'Add servers manually (not from the central list)')
$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Width = 170
$txtFilter.Margin = New-Object System.Windows.Forms.Padding(8, 5, 2, 2)
$ToolTip.SetToolTip($txtFilter, 'Filter: space-separated words, all must match (any column)')
$btnAll = New-Button 'All' 45
$ToolTip.SetToolTip($btnAll, 'Check all servers visible under the current filter')
$btnNone = New-Button 'None' 50
$ToolTip.SetToolTip($btnNone, 'Uncheck every server (including hidden ones)')
$pnlSrvTop.Controls.AddRange(@($lblSrvTitle, $btnRefresh, $btnAddSrv))
$pnlSrvTop.SetFlowBreak($btnAddSrv, $true)
$pnlSrvTop.Controls.AddRange(@((New-Label 'Filter:'), $txtFilter, $btnAll, $btnNone))

$lvServers = New-Object System.Windows.Forms.ListView
$lvServers.Dock = 'Fill'
$lvServers.View = 'Details'
$lvServers.CheckBoxes = $true
$lvServers.FullRowSelect = $true
$lvServers.HideSelection = $false
$lvServers.ShowItemToolTips = $true

$lblSrvSummary = New-Object System.Windows.Forms.Label
$lblSrvSummary.Dock = 'Bottom'
$lblSrvSummary.Height = 22
$lblSrvSummary.TextAlign = 'MiddleLeft'

$cmsSrv = New-Object System.Windows.Forms.ContextMenuStrip
$miCheckOnly = $cmsSrv.Items.Add('Check only the selected server(s)')
$miCheckSel  = $cmsSrv.Items.Add('Check selected')
$miUncheck   = $cmsSrv.Items.Add('Uncheck selected')
$miCopyName  = $cmsSrv.Items.Add('Copy name(s)')
$miRemove    = $cmsSrv.Items.Add('Remove manually added server(s)')
$lvServers.ContextMenuStrip = $cmsSrv

$splitMain.Panel1.Controls.Add($lvServers)
$splitMain.Panel1.Controls.Add($lblSrvSummary)
$splitMain.Panel1.Controls.Add($pnlSrvTop)
$lvServers.BringToFront()

# ---------------- right: tabs
$tabsMain = New-Object System.Windows.Forms.TabControl
$tabsMain.Dock = 'Fill'
$tabExec     = New-Object System.Windows.Forms.TabPage 'Query'
$tabLib      = New-Object System.Windows.Forms.TabPage 'Query Library'
$tabSettings = New-Object System.Windows.Forms.TabPage 'Settings'
$tabsMain.TabPages.AddRange(@($tabExec, $tabLib, $tabSettings))
$splitMain.Panel2.Controls.Add($tabsMain)

# ---------------- Query tab
$tbExec = New-Flow
$btnExecute = New-Button 'Execute (F5)' 100
$btnExecute.Font = $FontBold
$btnCancel = New-Button 'Cancel (Esc)' 90
$btnCancel.Enabled = $false
$cboDatabase = New-Object System.Windows.Forms.ComboBox
$cboDatabase.DropDownStyle = 'DropDown'
$cboDatabase.Width = 160
$cboDatabase.Margin = New-Object System.Windows.Forms.Padding(2, 5, 2, 2)
$btnDbList = New-Button 'List DBs' 65
$ToolTip.SetToolTip($btnDbList, 'Load database names from the first checked server')
$chkWrites = New-Object System.Windows.Forms.CheckBox
$chkWrites.Text = 'Allow writes (this run only)'
$chkWrites.AutoSize = $true
$chkWrites.Margin = New-Object System.Windows.Forms.Padding(12, 7, 2, 2)
$chkWrites.ForeColor = [System.Drawing.Color]::Firebrick
$numTimeout = New-Num 0 86400 300 65
$ToolTip.SetToolTip($numTimeout, 'Query timeout in seconds per batch (0 = no timeout)')
$numMaxRows = New-Num 1 10000000 50000 85
$ToolTip.SetToolTip($numMaxRows, 'Maximum rows fetched per result set per server')
$btnSaveToLib = New-Button 'Save to library' 105
$btnExport = New-Button 'Export CSV' 80
$btnClear = New-Button 'Clear' 55
$tbExec.Controls.AddRange(@($btnExecute, $btnCancel, (New-Label 'Database:'), $cboDatabase, $btnDbList, $chkWrites,
        (New-Label 'Timeout (s):'), $numTimeout, (New-Label 'Max rows:'), $numMaxRows, $btnSaveToLib, $btnExport, $btnClear))

$lblBanner = New-Object System.Windows.Forms.Label
$lblBanner.Dock = 'Top'
$lblBanner.Height = 26
$lblBanner.TextAlign = 'MiddleLeft'
$lblBanner.Font = $FontBold
$lblBanner.ForeColor = [System.Drawing.Color]::White

$splitExec = New-Object System.Windows.Forms.SplitContainer
$splitExec.Dock = 'Fill'
$splitExec.Orientation = 'Horizontal'

$txtEditor = New-CodeBox
$splitExec.Panel1.Controls.Add($txtEditor)

$tabResults = New-Object System.Windows.Forms.TabControl
$tabResults.Dock = 'Fill'
$tabMessages = New-Object System.Windows.Forms.TabPage 'Messages'
$tabStatus = New-Object System.Windows.Forms.TabPage 'Servers'
$txtMessages = New-Object System.Windows.Forms.TextBox
$txtMessages.Multiline = $true; $txtMessages.ReadOnly = $true; $txtMessages.ScrollBars = 'Both'
$txtMessages.WordWrap = $false; $txtMessages.Font = $FontGrid; $txtMessages.Dock = 'Fill'
$txtMessages.MaxLength = 0; $txtMessages.BackColor = [System.Drawing.Color]::White
$tabMessages.Controls.Add($txtMessages)
$lvStatus = New-Object System.Windows.Forms.ListView
$lvStatus.Dock = 'Fill'; $lvStatus.View = 'Details'; $lvStatus.FullRowSelect = $true; $lvStatus.Font = $FontGrid
[void]$lvStatus.Columns.Add('Server', 220)
[void]$lvStatus.Columns.Add('Status', 90)
[void]$lvStatus.Columns.Add('Rows', 80)
[void]$lvStatus.Columns.Add('Seconds', 70)
[void]$lvStatus.Columns.Add('Info', 700)
$tabStatus.Controls.Add($lvStatus)
$tabResults.TabPages.AddRange(@($tabMessages, $tabStatus))
$splitExec.Panel2.Controls.Add($tabResults)

$tabExec.Controls.Add($splitExec)
$tabExec.Controls.Add($lblBanner)
$tabExec.Controls.Add($tbExec)
$splitExec.BringToFront()

# ---------------- Library tab
$splitLib = New-Object System.Windows.Forms.SplitContainer
$splitLib.Dock = 'Fill'
$pnlLibSearch = New-Flow
$txtLibSearch = New-Object System.Windows.Forms.TextBox
$txtLibSearch.Width = 230
$txtLibSearch.Margin = New-Object System.Windows.Forms.Padding(2, 5, 2, 2)
$pnlLibSearch.Controls.AddRange(@((New-Label 'Search:'), $txtLibSearch))
$lvLib = New-Object System.Windows.Forms.ListView
$lvLib.Dock = 'Fill'; $lvLib.View = 'Details'; $lvLib.FullRowSelect = $true; $lvLib.HideSelection = $false; $lvLib.MultiSelect = $false
[void]$lvLib.Columns.Add('Name', 260)
[void]$lvLib.Columns.Add('Category', 110)
$splitLib.Panel1.Controls.Add($lvLib)
$splitLib.Panel1.Controls.Add($pnlLibSearch)
$lvLib.BringToFront()

$pnlLibEdit = New-Flow
$txtLibName = New-Object System.Windows.Forms.TextBox
$txtLibName.Width = 300
$txtLibName.Margin = New-Object System.Windows.Forms.Padding(2, 5, 2, 2)
$cboLibCat = New-Object System.Windows.Forms.ComboBox
$cboLibCat.DropDownStyle = 'DropDown'
$cboLibCat.Width = 160
$cboLibCat.Margin = New-Object System.Windows.Forms.Padding(2, 5, 2, 2)
$btnLibNew = New-Button 'New' 55
$btnLibSave = New-Button 'Save' 55
$btnLibDelete = New-Button 'Delete' 60
$btnLibCopy = New-Button 'Copy SQL' 75
$btnLibSend = New-Button 'Send to editor' 105
$btnLibSend.Font = $FontBold
$ToolTip.SetToolTip($btnLibSend, 'Appends to the Query editor and selects it, so F5 runs just this script (double-click an entry does the same)')
$pnlLibEdit.Controls.AddRange(@((New-Label 'Name:'), $txtLibName, (New-Label 'Category:'), $cboLibCat))
$pnlLibEdit.SetFlowBreak($cboLibCat, $true)
$pnlLibEdit.Controls.AddRange(@($btnLibNew, $btnLibSave, $btnLibDelete, $btnLibCopy, $btnLibSend))
$txtLibSql = New-CodeBox
$splitLib.Panel2.Controls.Add($txtLibSql)
$splitLib.Panel2.Controls.Add($pnlLibEdit)
$txtLibSql.BringToFront()
$tabLib.Controls.Add($splitLib)

# ---------------- Settings tab
$tlpSet = New-Object System.Windows.Forms.TableLayoutPanel
$tlpSet.Dock = 'Fill'
$tlpSet.AutoScroll = $true
$tlpSet.ColumnCount = 2
$tlpSet.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$tlpSet.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 230)))
[void]$tlpSet.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tabSettings.Controls.Add($tlpSet)

function Add-SettingRow([string]$Label, $Control, [string]$Tip = '') {
    $l = New-Label $Label
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)
    $row = $tlpSet.RowCount
    $tlpSet.RowCount = $row + 1
    [void]$tlpSet.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $tlpSet.Controls.Add($l, 0, $row)
    $tlpSet.Controls.Add($Control, 1, $row)
    if ($Tip) { $ToolTip.SetToolTip($Control, $Tip); $ToolTip.SetToolTip($l, $Tip) }
}
function New-SetText([int]$Width = 320) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $Width
    $t.Margin = New-Object System.Windows.Forms.Padding(3, 5, 3, 3)
    $t
}
function New-SetCheck([string]$Text) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text; $c.AutoSize = $true
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
    $c
}

$setCentral = New-SetText
$setCentralDb = New-SetText 200
$setQuery = New-CodeBox
$setQuery.Dock = 'None'
$setQuery.Width = 820; $setQuery.Height = 190
$setQuery.Anchor = 'Left,Top,Right'
$setAuth = New-Object System.Windows.Forms.ComboBox
$setAuth.DropDownStyle = 'DropDownList'
[void]$setAuth.Items.AddRange(@('Windows', 'SQL'))
$setAuth.Width = 120
$setSqlUser = New-SetText 200
$setEncrypt = New-SetCheck 'Encrypt connections'
$setTrust = New-SetCheck 'Trust server certificate'
$setIntent = New-SetCheck 'ApplicationIntent=ReadOnly (route AG listeners to a readable secondary)'
$setMsf = New-SetCheck 'MultiSubnetFailover (recommended for multi-subnet AG listeners)'
$setConnTimeout = New-Num 1 300 15
$setQueryTimeout = New-Num 0 86400 300
$setMaxRows = New-Num 1 10000000 50000 100
$setParallel = New-Num 1 64 8
$setLockTimeout = New-Num -1 600000 10000 100
$setProd = New-SetText 200
$setConfirmCount = New-Num 2 100000 10
$setDefaultDb = New-SetText 200

Add-SettingRow 'Central server' $setCentral 'SQL Server that holds your server inventory or CMS'
Add-SettingRow 'Central database' $setCentralDb 'Database the server list query runs in (msdb for CMS)'
Add-SettingRow 'Server list query' $setQuery 'First column must be the server name. Extra columns become filterable columns.'
Add-SettingRow 'Authentication' $setAuth 'Windows = integrated security. SQL = prompts once per session, password never saved.'
Add-SettingRow 'SQL login (SQL auth only)' $setSqlUser
Add-SettingRow '' $setEncrypt
Add-SettingRow '' $setTrust
Add-SettingRow '' $setIntent
Add-SettingRow '' $setMsf
Add-SettingRow 'Connect timeout (s)' $setConnTimeout
Add-SettingRow 'Default query timeout (s)' $setQueryTimeout '0 = no timeout'
Add-SettingRow 'Default max rows per result' $setMaxRows
Add-SettingRow 'Max servers in parallel' $setParallel
Add-SettingRow 'LOCK_TIMEOUT (ms)' $setLockTimeout 'Your session gives up waiting on locks after this. -1 = wait forever.'
Add-SettingRow 'PROD name pattern (regex)' $setProd 'Servers whose name or any column matches are shown in red and counted as PROD'
Add-SettingRow 'Confirm when servers >=' $setConfirmCount 'Ask before a read-only run against this many servers or more'
Add-SettingRow 'Default database' $setDefaultDb

$pnlSetButtons = New-Flow
$pnlSetButtons.Dock = 'None'
$btnSetSave = New-Button 'Save settings' 100
$btnSetSave.Font = $FontBold
$btnSetTest = New-Button 'Test central' 90
$btnSetRefresh = New-Button 'Save + refresh server list' 170
$btnSetFolder = New-Button 'Open data folder' 115
$btnSetForget = New-Button 'Forget SQL credential' 140
$pnlSetButtons.Controls.AddRange(@($btnSetSave, $btnSetTest, $btnSetRefresh, $btnSetFolder, $btnSetForget))
Add-SettingRow '' $pnlSetButtons
$lblSetInfo = New-Label ("Data folder: $DataDir" + "`r`nSettings, server cache, query library, editor draft and execution.log live there. Copy library.json to share queries.")
$lblSetInfo.ForeColor = [System.Drawing.Color]::DimGray
Add-SettingRow '' $lblSetInfo

# ---------------- assemble form
$form.Controls.Add($splitMain)
$form.Controls.Add($statusStrip)
$splitMain.BringToFront()

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 200

# ------------------------------------------------------------------ server list logic
function Get-ServerRows {
    $cols = @($State.Columns)
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $State.Servers) {
        $vals = New-Object string[] $cols.Count
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $p = $s.PSObject.Properties[$cols[$i]]
            $vals[$i] = if ($p -and $null -ne $p.Value) { [string]$p.Value } else { '' }
        }
        if (-not $vals[0] -or -not $seen.Add($vals[0])) { continue }
        $rows.Add([pscustomobject]@{ Name = $vals[0]; Values = $vals; Manual = $false })
    }
    foreach ($m in $State.Manual) {
        if (-not $m -or -not $seen.Add($m)) { continue }
        $vals = New-Object string[] $cols.Count
        for ($i = 0; $i -lt $cols.Count; $i++) { $vals[$i] = '' }
        $vals[0] = $m
        if ($cols.Count -gt 1) { $vals[$cols.Count - 1] = '(manual)' }
        $rows.Add([pscustomobject]@{ Name = $m; Values = $vals; Manual = $true })
    }
    $rows
}

function Get-CheckedServers {
    foreach ($r in (Get-ServerRows)) { if ($State.Checked.Contains($r.Name)) { $r.Name } }
}

function Update-ServerList {
    $terms = @($txtFilter.Text.Split([char[]]@(' '), [System.StringSplitOptions]::RemoveEmptyEntries))
    $cols = @($State.Columns)
    $rx = $null
    try { $rx = New-Object System.Text.RegularExpressions.Regex ([string]$State.Settings.ProdPattern) } catch { $rx = $null }
    $State.Rebuilding = $true
    $lvServers.BeginUpdate()
    try {
        $want = $cols -join '|'
        if ($lvServers.Tag -ne $want) {
            $lvServers.Columns.Clear()
            for ($i = 0; $i -lt $cols.Count; $i++) { [void]$lvServers.Columns.Add($cols[$i], $(if ($i -eq 0) { 170 } else { 100 })) }
            $lvServers.Tag = $want
        }
        $lvServers.Items.Clear()
        $State.ProdSet.Clear()
        $State.VisibleSet.Clear()
        foreach ($srv in (Get-ServerRows)) {
            $joined = $srv.Values -join ' '
            $isProd = [bool]($rx -and $rx.IsMatch($joined))
            if ($isProd) { [void]$State.ProdSet.Add($srv.Name) }
            $ok = $true
            foreach ($t in $terms) { if ($joined.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { $ok = $false; break } }
            if (-not $ok) { continue }
            [void]$State.VisibleSet.Add($srv.Name)
            $it = New-Object System.Windows.Forms.ListViewItem ([string]$srv.Name)
            for ($i = 1; $i -lt $cols.Count; $i++) { [void]$it.SubItems.Add([string]$srv.Values[$i]) }
            $it.Tag = $srv
            $it.Checked = $State.Checked.Contains($srv.Name)
            if ($isProd) { $it.ForeColor = [System.Drawing.Color]::Firebrick; $it.ToolTipText = 'PROD' }
            if ($srv.Manual) { $it.BackColor = [System.Drawing.Color]::FromArgb(238, 244, 255); $it.ToolTipText = ($it.ToolTipText + ' manually added').Trim() }
            [void]$lvServers.Items.Add($it)
        }
    } finally {
        $lvServers.EndUpdate()
        $State.Rebuilding = $false
    }
    Update-ServerSummary
}

function Update-ServerSummary {
    $all = @(Get-ServerRows)
    $checked = @($all | Where-Object { $State.Checked.Contains($_.Name) })
    $cache = if ($State.RefreshedAt) { "cache $($State.RefreshedAt)" } else { 'not refreshed yet' }
    $lblSrvSummary.Text = ' {0} servers | {1} checked | {2}' -f $all.Count, $checked.Count, $cache
    Update-Banner $checked
}

function Update-Banner($CheckedRows) {
    if ($null -eq $CheckedRows) { $CheckedRows = @(Get-ServerRows | Where-Object { $State.Checked.Contains($_.Name) }) }
    $CheckedRows = @($CheckedRows)
    $prod = @($CheckedRows | Where-Object { $State.ProdSet.Contains($_.Name) }).Count
    $hidden = @($CheckedRows | Where-Object { -not $State.VisibleSet.Contains($_.Name) }).Count
    if ($chkWrites.Checked) {
        $lblBanner.BackColor = [System.Drawing.Color]::Firebrick
        $mode = 'WRITES ARMED - the next run commits changes'
    } else {
        $lblBanner.BackColor = [System.Drawing.Color]::SeaGreen
        $mode = 'READ-ONLY - writes blocked, every run rolled back'
    }
    $t = '   {0}   |   Targets: {1} server(s)' -f $mode, $CheckedRows.Count
    if ($prod) { $t += " ($prod PROD)" }
    if ($hidden) { $t += " - $hidden hidden by filter" }
    $db = $cboDatabase.Text.Trim(); if (-not $db) { $db = 'master' }
    $t += "   |   Database: $db"
    $lblBanner.Text = $t
}

function Update-ServersFromCentral {
    $S = $State.Settings
    if (-not $S.CentralServer) {
        $tabsMain.SelectedTab = $tabSettings
        Show-Warn 'Set the central server (and database / query) in Settings first.'
        return
    }
    $g = Test-ReadOnlySql $S.ServerListQuery
    if ($g.Blocked.Count -gt 0) { Show-Warn ('The server list query must be read-only. Detected: ' + ($g.Blocked -join ', ')); return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblStatus.Text = "Querying $($S.CentralServer) for the server list..."
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $dt = Invoke-SqlSync $S.CentralServer $S.CentralDatabase $S.ServerListQuery 60
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    if ($null -eq $dt -or $dt.Columns.Count -eq 0) { Show-Warn 'The server list query returned no columns.'; return }

    $used = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $cols = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $dt.Columns.Count; $i++) {
        $name = if ($i -eq 0) { 'Server' } else { $dt.Columns[$i].ColumnName }
        if (-not $name) { $name = "Col$i" }
        $base = $name; $k = 2
        while (-not $used.Add($name)) { $name = "$base$k"; $k++ }
        $cols.Add($name)
    }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in $dt.Rows) {
        if ($row[0] -is [DBNull]) { continue }
        $name = ([string]$row[0]).Trim()
        if (-not $name) { continue }
        $h = [ordered]@{}
        $h['Server'] = $name
        for ($i = 1; $i -lt $cols.Count; $i++) { $h[$cols[$i]] = if ($row[$i] -is [DBNull]) { '' } else { [string]$row[$i] } }
        $list.Add([pscustomobject]$h)
    }
    $State.Columns = $cols.ToArray()
    $State.Servers = $list
    $State.RefreshedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $State.Source = "$($S.CentralServer)/$($S.CentralDatabase)"
    # drop checks for servers that no longer exist
    $valid = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in (Get-ServerRows)) { [void]$valid.Add($r.Name) }
    $State.Checked.IntersectWith($valid)
    Save-ServerCache
    Update-ServerList
    $lblStatus.Text = "Loaded $($list.Count) servers from $($State.Source) - cached."
}

# ------------------------------------------------------------------ results
function New-ResultGrid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.RowHeadersVisible = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $g.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
    $g.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $g.BackgroundColor = [System.Drawing.Color]::White
    $g.BorderStyle = 'None'
    $g.Font = $FontGrid
    $g.DefaultCellStyle.NullValue = 'NULL'
    $g.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 250)
    $g.RowTemplate.Height = 20
    $g.Tag = $false
    try {
        $p = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
        $p.SetValue($g, $true, $null)
    } catch { }
    $g.add_DataError({ param($s, $e) try { $e.ThrowException = $false } catch { } })
    $g.add_DataBindingComplete({
        param($s, $e)
        Invoke-Safe {
            foreach ($c in $s.Columns) {
                if ($c.ValueType -eq [datetime]) { $c.DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm:ss.fff' }
                if ($c.Index -eq 0) { $c.DefaultCellStyle.ForeColor = [System.Drawing.Color]::SteelBlue }
            }
        } 'grid DataBindingComplete'
    })
    $g.add_VisibleChanged({
        param($s, $e)
        Invoke-Safe {
            if ($s.Visible -and (-not $s.Tag) -and (-not $State.Running)) { Resize-GridColumns $s }
        } 'grid VisibleChanged'
    })
    $g
}

function Resize-GridColumns($g) {
    try {
        $g.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
        foreach ($c in $g.Columns) { if ($c.Width -gt 400) { $c.Width = 400 } }
        $g.Tag = $true
    } catch { }
}

function Reset-Results($Servers) {
    if ($State.ResultMap) {
        foreach ($e in (ConvertTo-PlainArray $State.ResultMap.Values)) {
            $tabResults.TabPages.Remove($e.Tab)
            $e.Grid.DataSource = $null
            $e.Tab.Dispose()
            $e.Table.Dispose()
        }
    }
    $State.ResultMap = [ordered]@{}
    $txtMessages.Clear()
    $State.StatusItems = @{}
    $lvStatus.BeginUpdate()
    $lvStatus.Items.Clear()
    foreach ($s in $Servers) {
        $it = New-Object System.Windows.Forms.ListViewItem ([string]$s)
        [void]$it.SubItems.Add('Queued')
        [void]$it.SubItems.Add('')
        [void]$it.SubItems.Add('')
        [void]$it.SubItems.Add('')
        if ($State.ProdSet.Contains($s)) { $it.ForeColor = [System.Drawing.Color]::Firebrick }
        [void]$lvStatus.Items.Add($it)
        $State.StatusItems[$s] = $it
    }
    $lvStatus.EndUpdate()
    $tabStatus.Text = 'Servers'
    $tabResults.SelectedTab = $tabStatus
}

function Add-ResultTable([System.Data.DataTable]$dt) {
    $sig = '{0}|{1}' -f $dt.ExtendedProperties['Index'], (($dt.Columns | ForEach-Object { '{0}:{1}' -f $_.ColumnName, $_.DataType.Name }) -join '|')
    if (-not $State.ResultMap.Contains($sig)) {
        $n = $State.ResultMap.Count + 1
        $tab = New-Object System.Windows.Forms.TabPage "Results $n"
        $grid = New-ResultGrid
        $tab.Controls.Add($grid)
        $tabResults.TabPages.Insert($n - 1, $tab)
        $bs = New-Object System.Windows.Forms.BindingSource
        $bs.DataSource = $dt
        $grid.DataSource = $bs
        $State.ResultMap[$sig] = @{ Table = $dt; Binding = $bs; Grid = $grid; Tab = $tab; N = $n }
        if ($n -eq 1) { $tabResults.SelectedTab = $tab }
        $entry = $State.ResultMap[$sig]
    } else {
        $entry = $State.ResultMap[$sig]
        $entry.Binding.RaiseListChangedEvents = $false
        try { $entry.Table.Merge($dt, $false, [System.Data.MissingSchemaAction]::Ignore) }
        finally {
            $entry.Binding.RaiseListChangedEvents = $true
            $entry.Binding.ResetBindings($false)
        }
        $dt.Dispose()
    }
    $entry.Tab.Text = 'Results {0} ({1:N0})' -f $entry.N, $entry.Table.Rows.Count
}

function Add-Message([string]$Text) { $txtMessages.AppendText((ConvertTo-CrLf $Text) + "`r`n") }

function Complete-Job($Job, $Res) {
    $it = $State.StatusItems[$Job.Server]
    if (-not $Res) {
        $Res = [pscustomobject]@{ Server = $Job.Server; Tables = @(); Messages = @(); Error = 'Worker returned no result.'; Cancelled = $false; Truncated = $false; Rows = 0; RowsAffected = 0; Ms = 0 }
    }
    $st = if ($Res.Cancelled) { 'Cancelled' } elseif ($Res.Error) { 'Error' } elseif ($Res.Truncated) { 'Truncated' } else { 'OK' }
    $State.RunInfo.Counts[$st]++
    $info = if ($Res.Error) { ($Res.Error -split "`r?`n" | Select-Object -Last 1) } elseif ($Res.RowsAffected -gt 0) { "$($Res.RowsAffected) row(s) affected" } else { '' }
    if ($it) {
        $it.SubItems[1].Text = $st
        $it.SubItems[2].Text = '{0:N0}' -f $Res.Rows
        $it.SubItems[3].Text = '{0:N2}' -f ($Res.Ms / 1000.0)
        $it.SubItems[4].Text = [string]$info
        switch ($st) {
            'OK'        { $it.ForeColor = [System.Drawing.Color]::DarkGreen }
            'Error'     { $it.ForeColor = [System.Drawing.Color]::Firebrick }
            'Truncated' { $it.ForeColor = [System.Drawing.Color]::DarkOrange }
            'Cancelled' { $it.ForeColor = [System.Drawing.Color]::Gray }
        }
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(('---- {0}  [{1}]  rows: {2:N0}  {3:N2}s ----' -f $Job.Server, $st, $Res.Rows, ($Res.Ms / 1000.0)))
    foreach ($m in $Res.Messages) { [void]$sb.AppendLine([string]$m) }
    if ($Res.Error) { [void]$sb.AppendLine($Res.Error) }
    Add-Message $sb.ToString()
    foreach ($dt in (ConvertTo-PlainArray $Res.Tables)) { Add-ResultTable $dt }
}

# ------------------------------------------------------------------ execution
function Set-RunningUi([bool]$Running) {
    $btnExecute.Enabled = -not $Running
    $btnCancel.Enabled = $Running
    $chkWrites.Enabled = -not $Running
    $btnDbList.Enabled = -not $Running
    $btnRefresh.Enabled = -not $Running
}

function Start-Execution {
    if ($State.Running) { return }
    $S = $State.Settings
    $sql = if ($txtEditor.SelectionLength -gt 0) { $txtEditor.SelectedText } else { $txtEditor.Text }
    if ([string]::IsNullOrWhiteSpace($sql)) { Show-Warn 'Nothing to execute.'; return }
    $servers = @(Get-CheckedServers)
    if ($servers.Count -eq 0) { Show-Warn 'Check at least one server in the list on the left.'; return }
    $db = $cboDatabase.Text.Trim(); if (-not $db) { $db = 'master' }
    $writes = $chkWrites.Checked
    $prodCount = @($servers | Where-Object { $State.ProdSet.Contains($_) }).Count

    if (-not $writes) {
        $g = Test-ReadOnlySql $sql
        if ($g.Blocked.Count -gt 0) {
            Show-Warn ("Read-only mode blocked this query.`r`n`r`nDetected: {0}`r`n`r`nTemp tables (#t) and table variables (@t) are allowed. If you really intend to change data, tick 'Allow writes (this run only)'." -f ($g.Blocked -join ', '))
            return
        }
        if ($g.Warnings.Count -gt 0) {
            $q = "This query contains: {0}`r`n`r`nIn read-only mode database changes are rolled back, but effects outside the transaction (linked servers, xp_ procedures, external calls) cannot be undone.`r`n`r`nRun it?" -f ($g.Warnings -join ', ')
            if (-not (Ask-YesNo $q)) { return }
        }
        if ($servers.Count -ge [int]$S.ConfirmServerCount) {
            if (-not (Ask-YesNo ("Run against {0} servers ({1} PROD) in database [{2}]?" -f $servers.Count, $prodCount, $db))) { return }
        }
    } else {
        $preview = ($servers | Select-Object -First 8) -join ', '
        if ($servers.Count -gt 8) { $preview += ", ... (+$($servers.Count - 8) more)" }
        $msg = "WRITES ARE ARMED - no safety rollback.`r`n`r`nServers: $($servers.Count)  ($prodCount PROD)`r`nDatabase: $db`r`n$preview`r`n`r`nType EXECUTE to run it."
        if ($msg.Length -gt 1000) { $msg = $msg.Substring(0, 1000) }
        $ans = [Microsoft.VisualBasic.Interaction]::InputBox($msg, "$AppName - confirm writes", '')
        if ($ans -cne 'EXECUTE') { $lblStatus.Text = 'Write run aborted (EXECUTE not typed).'; return }
    }

    $cred = $null
    if ($S.AuthMode -eq 'SQL') { $cred = Get-SqlCred; if (-not $cred) { return } }
    $batches = [string[]]@(Split-SqlBatches $sql)
    if ($batches.Count -eq 0) { Show-Warn 'Nothing to execute.'; return }

    Reset-Results $servers
    $State.Shared = New-Object 'System.Collections.Concurrent.ConcurrentDictionary[string,object]'
    $State.Pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Max(1, [int]$S.MaxParallel))
    $State.Pool.Open()
    $State.Jobs = New-Object System.Collections.Generic.List[object]
    $State.RunStart = Get-Date
    $runId = [guid]::NewGuid().ToString().Substring(0, 8)
    $State.RunInfo = @{ Id = $runId; Writes = $writes; Database = $db; Servers = $servers; Counts = @{ OK = 0; Error = 0; Truncated = 0; Cancelled = 0 } }
    Write-ExecLog ([ordered]@{ Event = 'start'; RunId = $runId; Mode = $(if ($writes) { 'WRITE' } else { 'READ-ONLY' }); Database = $db; Servers = $servers; Sql = $sql })

    $timeout = [int]$numTimeout.Value
    $maxRows = [int]$numMaxRows.Value
    $script = $WorkerScript.ToString()
    foreach ($srv in $servers) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $State.Pool
        [void]$ps.AddScript($script)
        [void]$ps.AddParameters(@{
            Server = $srv; ConnStr = (New-SqlConnectionString $srv $db); Cred = $cred; Batches = $batches
            ReadOnly = (-not $writes); Timeout = $timeout; MaxRows = $maxRows
            LockTimeoutMs = [int]$S.LockTimeoutMs; Shared = $State.Shared
        })
        $State.Jobs.Add([pscustomobject]@{ Server = $srv; PS = $ps; Handle = $ps.BeginInvoke(); Done = $false })
    }
    $State.Running = $true
    Set-RunningUi $true
    $lblStatus.Text = 'Running on {0} server(s) [{1}] - {2}' -f $servers.Count, $db, $(if ($writes) { 'WRITE MODE' } else { 'read-only' })
    $timer.Start()
}

function Update-Jobs {
    if (-not $State.Running) { $timer.Stop(); return }
    $pending = 0
    $done = 0
    foreach ($job in $State.Jobs) {
        if ($job.Done) { $done++; continue }
        if ($job.Handle.IsCompleted) {
            $job.Done = $true
            $done++
            $res = $null
            try {
                $out = $job.PS.EndInvoke($job.Handle)
                if ($out -and $out.Count -gt 0) { $res = $out[$out.Count - 1] }
                if (-not $res -and $job.PS.Streams.Error.Count -gt 0) {
                    $res = [pscustomobject]@{ Server = $job.Server; Tables = @(); Messages = @(); Error = [string]$job.PS.Streams.Error[0]; Cancelled = $false; Truncated = $false; Rows = 0; RowsAffected = 0; Ms = 0 }
                }
            } catch {
                $res = [pscustomobject]@{ Server = $job.Server; Tables = @(); Messages = @(); Error = (Get-InnerMessage $_); Cancelled = $false; Truncated = $false; Rows = 0; RowsAffected = 0; Ms = 0 }
            }
            try { $job.PS.Dispose() } catch { }
            Complete-Job $job $res
        } else {
            $pending++
            $key = "state:$($job.Server)"
            if ($State.Shared.ContainsKey($key)) {
                $st = [string]$State.Shared[$key]
                $it = $State.StatusItems[$job.Server]
                if ($it -and $it.SubItems[1].Text -ne $st) { $it.SubItems[1].Text = $st }
            }
        }
    }
    $lblElapsed.Text = '{0}/{1} done  {2:N1}s' -f $done, $State.Jobs.Count, ((Get-Date) - $State.RunStart).TotalSeconds
    if ($pending -eq 0) { Complete-Execution }
}

function Complete-Execution {
    $timer.Stop()
    try { $State.Pool.Close(); $State.Pool.Dispose() } catch { }
    $State.Running = $false
    $c = $State.RunInfo.Counts
    $secs = ((Get-Date) - $State.RunStart).TotalSeconds
    $summary = 'Finished in {0:N1}s: {1} OK, {2} error, {3} truncated, {4} cancelled' -f $secs, $c.OK, $c.Error, $c.Truncated, $c.Cancelled
    $tail = if ($State.RunInfo.Writes) { 'WRITE MODE: changes were committed.' } else { 'Read-only mode: every server''s transaction was rolled back.' }
    Add-Message "==== $summary ====`r`n$tail"
    $lblStatus.Text = $summary
    $lblElapsed.Text = '{0:N1}s' -f $secs
    if ($c.Error -gt 0) { $tabStatus.Text = "Servers ($($c.Error) errors)" }
    Set-RunningUi $false
    Write-ExecLog ([ordered]@{ Event = 'end'; RunId = $State.RunInfo.Id; Seconds = [Math]::Round($secs, 1); OK = $c.OK; Errors = $c.Error; Truncated = $c.Truncated; Cancelled = $c.Cancelled })
    if ($State.RunInfo.Writes) { $chkWrites.Checked = $false }   # writes are one-shot

    if ($State.ResultMap.Count -eq 0) {
        $tabResults.SelectedTab = if ($c.Error -gt 0 -and $State.Jobs.Count -gt 1) { $tabStatus } else { $tabMessages }
    } else {
        foreach ($e in $State.ResultMap.Values) { if ($e.Grid.Visible) { Resize-GridColumns $e.Grid } }
    }
}

function Stop-Execution {
    if (-not $State.Running) { return }
    $State.Shared['__cancel'] = $true
    foreach ($k in (ConvertTo-PlainArray $State.Shared.Keys)) {
        $v = $State.Shared[$k]
        if ($v -is [System.Data.SqlClient.SqlCommand]) { try { $v.Cancel() } catch { } }
    }
    $lblStatus.Text = 'Cancelling...'
}

function Export-CurrentResult {
    if (-not $State.ResultMap -or $State.ResultMap.Count -eq 0) { Show-Warn 'No results to export.'; return }
    $entry = $null
    foreach ($e in $State.ResultMap.Values) { if ($e.Tab -eq $tabResults.SelectedTab) { $entry = $e } }
    if (-not $entry) { $entry = (ConvertTo-PlainArray $State.ResultMap.Values)[0] }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.FileName = 'results_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $dt = $entry.Table
    $fmt = {
        param($v)
        if ($null -eq $v -or $v -is [DBNull]) { return '' }
        $s = if ($v -is [datetime]) { $v.ToString('yyyy-MM-dd HH:mm:ss.fff') } else { [string]$v }
        '"' + $s.Replace('"', '""') + '"'
    }
    $w = New-Object System.IO.StreamWriter($dlg.FileName, $false, (New-Object System.Text.UTF8Encoding($true)))
    try {
        $w.WriteLine((@($dt.Columns | ForEach-Object { & $fmt $_.ColumnName }) -join ','))
        foreach ($r in $dt.Rows) { $w.WriteLine((@($r.ItemArray | ForEach-Object { & $fmt $_ }) -join ',')) }
    } finally { $w.Dispose() }
    $lblStatus.Text = "Exported $($dt.Rows.Count) rows to $($dlg.FileName)"
}

# ------------------------------------------------------------------ library logic
function Update-LibraryList([string]$SelectId) {
    $terms = @($txtLibSearch.Text.Split([char[]]@(' '), [System.StringSplitOptions]::RemoveEmptyEntries))
    $lvLib.BeginUpdate()
    try {
        $lvLib.Items.Clear()
        $lvLib.Groups.Clear()
        $groups = @{}
        foreach ($q in ($State.Library | Sort-Object Category, Name)) {
            $hay = '{0} {1} {2}' -f $q.Name, $q.Category, $q.Sql
            $ok = $true
            foreach ($t in $terms) { if ($hay.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { $ok = $false; break } }
            if (-not $ok) { continue }
            $cat = if ($q.Category) { [string]$q.Category } else { 'General' }
            if (-not $groups.ContainsKey($cat)) {
                $grp = New-Object System.Windows.Forms.ListViewGroup($cat, $cat)
                [void]$lvLib.Groups.Add($grp)
                $groups[$cat] = $grp
            }
            $it = New-Object System.Windows.Forms.ListViewItem([string]$q.Name, $groups[$cat])
            [void]$it.SubItems.Add($cat)
            $it.Tag = $q.Id
            $it.ToolTipText = [string]$q.Sql
            [void]$lvLib.Items.Add($it)
            if ($SelectId -and $q.Id -eq $SelectId) { $it.Selected = $true; $it.EnsureVisible() }
        }
    } finally { $lvLib.EndUpdate() }
    $curCat = $cboLibCat.Text
    $cboLibCat.Items.Clear()
    foreach ($c in ($State.Library | ForEach-Object { if ($_.Category) { [string]$_.Category } else { 'General' } } | Sort-Object -Unique)) { [void]$cboLibCat.Items.Add($c) }
    $cboLibCat.Text = $curCat
}

function Send-ToEditor([string]$Sql) {
    if ([string]::IsNullOrWhiteSpace($Sql)) { return }
    $Sql = ConvertTo-CrLf $Sql.TrimEnd()
    $tabsMain.SelectedTab = $tabExec
    if ([string]::IsNullOrWhiteSpace($txtEditor.Text)) {
        $txtEditor.Text = $Sql
        $txtEditor.SelectAll()
    } else {
        $prefix = "`r`n`r`n"
        $start = $txtEditor.TextLength + $prefix.Length
        $txtEditor.AppendText($prefix + $Sql)
        $txtEditor.Select($start, $Sql.Length)
        $txtEditor.ScrollToCaret()
    }
    $txtEditor.Focus()
    $lblStatus.Text = 'Inserted into the editor and selected - F5 runs just the selection.'
}

# ------------------------------------------------------------------ settings logic
function Show-SettingsInUi {
    $S = $State.Settings
    $setCentral.Text = [string]$S.CentralServer
    $setCentralDb.Text = [string]$S.CentralDatabase
    $setQuery.Text = ConvertTo-CrLf ([string]$S.ServerListQuery)
    $setAuth.SelectedItem = $(if ($S.AuthMode -eq 'SQL') { 'SQL' } else { 'Windows' })
    $setSqlUser.Text = [string]$S.SqlUser
    $setEncrypt.Checked = [bool]$S.Encrypt
    $setTrust.Checked = [bool]$S.TrustServerCertificate
    $setIntent.Checked = [bool]$S.ReadOnlyIntent
    $setMsf.Checked = [bool]$S.MultiSubnetFailover
    $setConnTimeout.Value = [Math]::Min([Math]::Max([int]$S.ConnectTimeout, 1), 300)
    $setQueryTimeout.Value = [Math]::Min([Math]::Max([int]$S.QueryTimeout, 0), 86400)
    $setMaxRows.Value = [Math]::Min([Math]::Max([int]$S.MaxRows, 1), 10000000)
    $setParallel.Value = [Math]::Min([Math]::Max([int]$S.MaxParallel, 1), 64)
    $setLockTimeout.Value = [Math]::Min([Math]::Max([int]$S.LockTimeoutMs, -1), 600000)
    $setProd.Text = [string]$S.ProdPattern
    $setConfirmCount.Value = [Math]::Min([Math]::Max([int]$S.ConfirmServerCount, 2), 100000)
    $setDefaultDb.Text = [string]$S.DefaultDatabase
}

function Save-SettingsFromUi {
    try { [void][regex]::new($setProd.Text) } catch { Show-Warn 'PROD name pattern is not a valid regular expression.'; return $false }
    $S = $State.Settings
    $oldAuth = "$($S.AuthMode)|$($S.SqlUser)"
    $S.CentralServer = $setCentral.Text.Trim()
    $S.CentralDatabase = $(if ($setCentralDb.Text.Trim()) { $setCentralDb.Text.Trim() } else { 'master' })
    $S.ServerListQuery = $setQuery.Text
    $S.AuthMode = [string]$setAuth.SelectedItem
    $S.SqlUser = $setSqlUser.Text.Trim()
    $S.Encrypt = $setEncrypt.Checked
    $S.TrustServerCertificate = $setTrust.Checked
    $S.ReadOnlyIntent = $setIntent.Checked
    $S.MultiSubnetFailover = $setMsf.Checked
    $S.ConnectTimeout = [int]$setConnTimeout.Value
    $S.QueryTimeout = [int]$setQueryTimeout.Value
    $S.MaxRows = [int]$setMaxRows.Value
    $S.MaxParallel = [int]$setParallel.Value
    $S.LockTimeoutMs = [int]$setLockTimeout.Value
    $S.ProdPattern = $setProd.Text
    $S.ConfirmServerCount = [int]$setConfirmCount.Value
    $S.DefaultDatabase = $(if ($setDefaultDb.Text.Trim()) { $setDefaultDb.Text.Trim() } else { 'master' })
    if ("$($S.AuthMode)|$($S.SqlUser)" -ne $oldAuth) { $State.SqlCredential = $null }
    Save-Settings
    $numTimeout.Value = $S.QueryTimeout
    $numMaxRows.Value = $S.MaxRows
    Update-ServerList
    $lblStatus.Text = 'Settings saved.'
    $true
}

# ------------------------------------------------------------------ event wiring
$timer.add_Tick({ Invoke-Safe { Update-Jobs } })

$btnExecute.add_Click({ Invoke-Safe { Start-Execution } })
$btnCancel.add_Click({ Invoke-Safe { Stop-Execution } })
$btnExport.add_Click({ Invoke-Safe { Export-CurrentResult } })
$btnClear.add_Click({ Invoke-Safe { if (-not $State.Running) { Reset-Results @(); $lblStatus.Text = 'Results cleared.' } } })
$chkWrites.add_CheckedChanged({ Invoke-Safe { Update-Banner $null } })
$cboDatabase.add_TextChanged({ Invoke-Safe { Update-Banner $null } })

$btnDbList.add_Click({
    Invoke-Safe {
        $first = @(Get-CheckedServers) | Select-Object -First 1
        if (-not $first) { Show-Warn 'Check a server first.'; return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { $dt = Invoke-SqlSync $first 'master' 'SELECT name FROM sys.databases WHERE state = 0 ORDER BY name;' 30 }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
        $cur = $cboDatabase.Text
        $cboDatabase.Items.Clear()
        foreach ($r in $dt.Rows) { [void]$cboDatabase.Items.Add([string]$r[0]) }
        $cboDatabase.Text = $cur
        $cboDatabase.DroppedDown = $true
        $lblStatus.Text = "Loaded $($dt.Rows.Count) database names from $first."
    }
})

$btnSaveToLib.add_Click({
    Invoke-Safe {
        $sql = if ($txtEditor.SelectionLength -gt 0) { $txtEditor.SelectedText } else { $txtEditor.Text }
        if ([string]::IsNullOrWhiteSpace($sql)) { Show-Warn 'The editor is empty.'; return }
        $name = [Microsoft.VisualBasic.Interaction]::InputBox('Name for this query:', "$AppName - save to library", '')
        if (-not $name.Trim()) { return }
        $cat = [Microsoft.VisualBasic.Interaction]::InputBox('Category:', "$AppName - save to library", 'General')
        if (-not $cat.Trim()) { $cat = 'General' }
        $q = New-LibEntry $name.Trim() $cat.Trim() $sql
        $State.Library.Add($q)
        Save-Library
        Update-LibraryList $q.Id
        $lblStatus.Text = "Saved '$($q.Name)' to the query library."
    }
})

# servers panel
$txtFilter.add_TextChanged({ Invoke-Safe { Update-ServerList } })
$btnRefresh.add_Click({ Invoke-Safe { Update-ServersFromCentral } })
$btnAll.add_Click({ Invoke-Safe { foreach ($n in $State.VisibleSet) { [void]$State.Checked.Add($n) }; Update-ServerList } })
$btnNone.add_Click({ Invoke-Safe { $State.Checked.Clear(); Update-ServerList } })
$btnAddSrv.add_Click({
    Invoke-Safe {
        $txt = [Microsoft.VisualBasic.Interaction]::InputBox('Server name(s), separated by commas or semicolons (e.g. SQL01, SQL02\INST, sqlag01.corp.local,1433):', "$AppName - add servers", '')
        $added = 0
        foreach ($n in ($txt -split '[;\r\n]|,(?!\d)')) {
            $n = $n.Trim()
            if ($n -and -not ($State.Manual -contains $n)) { $State.Manual.Add($n); [void]$State.Checked.Add($n); $added++ }
        }
        if ($added) { Save-ServerCache; Update-ServerList; $lblStatus.Text = "Added $added server(s) and checked them." }
    }
})
$lvServers.add_ItemChecked({
    param($s, $e)
    Invoke-Safe {
        if (-not $State.Rebuilding) {
            $name = $e.Item.Text
            if ($e.Item.Checked) { [void]$State.Checked.Add($name) } else { [void]$State.Checked.Remove($name) }
            Update-ServerSummary
        }
    } 'server ItemChecked'
})
$miCheckOnly.add_Click({
    Invoke-Safe {
        if ($lvServers.SelectedItems.Count -eq 0) { return }
        $State.Checked.Clear()
        foreach ($it in $lvServers.SelectedItems) { [void]$State.Checked.Add($it.Text) }
        Update-ServerList
    }
})
$miCheckSel.add_Click({ Invoke-Safe { foreach ($it in $lvServers.SelectedItems) { [void]$State.Checked.Add($it.Text) }; Update-ServerList } })
$miUncheck.add_Click({ Invoke-Safe { foreach ($it in $lvServers.SelectedItems) { [void]$State.Checked.Remove($it.Text) }; Update-ServerList } })
$miCopyName.add_Click({
    Invoke-Safe {
        $names = @($lvServers.SelectedItems | ForEach-Object { $_.Text })
        if ($names.Count) { [System.Windows.Forms.Clipboard]::SetText(($names -join "`r`n")) }
    }
})
$miRemove.add_Click({
    Invoke-Safe {
        $removed = 0
        foreach ($it in @($lvServers.SelectedItems)) {
            if ($it.Tag.Manual) { [void]$State.Manual.Remove($it.Text); [void]$State.Checked.Remove($it.Text); $removed++ }
        }
        if ($removed) { Save-ServerCache; Update-ServerList } else { Show-Info 'Only manually added servers can be removed. Central servers come from the server list query.' }
    }
})

# library
$txtLibSearch.add_TextChanged({ Invoke-Safe { Update-LibraryList $State.LibSelectedId } })
$lvLib.add_SelectedIndexChanged({
    Invoke-Safe {
        if ($lvLib.SelectedItems.Count -eq 1) {
            $q = Find-LibEntry ([string]$lvLib.SelectedItems[0].Tag)
            if ($q) {
                $State.LibSelectedId = $q.Id
                $txtLibName.Text = [string]$q.Name
                $cboLibCat.Text = [string]$q.Category
                $txtLibSql.Text = ConvertTo-CrLf ([string]$q.Sql)
            }
        }
    } 'library SelectedIndexChanged'
})
$lvLib.add_DoubleClick({ Invoke-Safe { Send-ToEditor $txtLibSql.Text } })
$btnLibNew.add_Click({
    Invoke-Safe {
        $State.LibSelectedId = $null
        $lvLib.SelectedItems.Clear()
        $txtLibName.Text = ''; $txtLibSql.Text = ''
        [void]$txtLibName.Focus()
    } 'library New'
})
$btnLibSave.add_Click({
    Invoke-Safe {
        $name = $txtLibName.Text.Trim()
        if (-not $name) { Show-Warn 'Give the query a name.'; return }
        if ([string]::IsNullOrWhiteSpace($txtLibSql.Text)) { Show-Warn 'The SQL is empty.'; return }
        $cat = $cboLibCat.Text.Trim(); if (-not $cat) { $cat = 'General' }
        $q = if ($State.LibSelectedId) { Find-LibEntry $State.LibSelectedId } else { $null }
        if ($q) {
            $q.Name = $name; $q.Category = $cat; $q.Sql = $txtLibSql.Text; $q.Updated = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        } else {
            $q = New-LibEntry $name $cat $txtLibSql.Text
            $State.Library.Add($q)
            $State.LibSelectedId = $q.Id
        }
        Save-Library
        Update-LibraryList $q.Id
        $lblStatus.Text = "Saved '$name'."
    }
})
$btnLibDelete.add_Click({
    Invoke-Safe {
        $q = if ($State.LibSelectedId) { Find-LibEntry $State.LibSelectedId } else { $null }
        if (-not $q) { return }
        if (-not (Ask-YesNo "Delete '$($q.Name)' from the library?")) { return }
        [void]$State.Library.Remove($q)
        $State.LibSelectedId = $null
        $txtLibName.Text = ''; $txtLibSql.Text = ''
        Save-Library
        Update-LibraryList $null
    }
})
$btnLibCopy.add_Click({
    Invoke-Safe {
        if ([string]::IsNullOrWhiteSpace($txtLibSql.Text)) { return }
        [System.Windows.Forms.Clipboard]::SetText($txtLibSql.Text)
        $lblStatus.Text = 'Query copied to the clipboard.'
    }
})
$btnLibSend.add_Click({ Invoke-Safe { Send-ToEditor $txtLibSql.Text } })

# settings
$btnSetSave.add_Click({ Invoke-Safe { [void](Save-SettingsFromUi) } })
$btnSetRefresh.add_Click({ Invoke-Safe { if (Save-SettingsFromUi) { Update-ServersFromCentral; $tabsMain.SelectedTab = $tabExec } } })
$btnSetTest.add_Click({
    Invoke-Safe {
        if (-not (Save-SettingsFromUi)) { return }
        $S = $State.Settings
        if (-not $S.CentralServer) { Show-Warn 'Enter the central server first.'; return }
        $dt = Invoke-SqlSync $S.CentralServer $S.CentralDatabase 'SELECT @@SERVERNAME AS server_name, DB_NAME() AS database_name, SUSER_SNAME() AS login_name;' 30
        $r = $dt.Rows[0]
        Show-Info ("Connected.`r`n`r`nServer: {0}`r`nDatabase: {1}`r`nLogin: {2}" -f $r[0], $r[1], $r[2])
    }
})
$btnSetFolder.add_Click({ Invoke-Safe { Start-Process explorer.exe -ArgumentList ('"{0}"' -f $DataDir) } 'open folder' })
$btnSetForget.add_Click({ Invoke-Safe { $State.SqlCredential = $null; $lblStatus.Text = 'SQL credential cleared from memory.' } 'forget credential' })

# form-level keys and lifecycle
$form.add_KeyDown({
    param($s, $e)
    Invoke-Safe {
        $key = $e.KeyCode
        $isRun = ($key -eq [System.Windows.Forms.Keys]::F5)
        if ($e.Control -and $key -eq [System.Windows.Forms.Keys]::E) { $isRun = $true }
        if ($isRun) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            if ($tabsMain.SelectedTab -eq $tabExec) { Start-Execution }
        } elseif ($key -eq [System.Windows.Forms.Keys]::Escape) {
            if ($State.Running) { $e.Handled = $true; Stop-Execution }
        }
    } 'form KeyDown'
})

$form.add_Shown({
  Invoke-Safe {
    try {
        $splitMain.SplitterDistance = 360
        $splitExec.SplitterDistance = [int]($splitExec.Height * 0.42)
        $splitLib.SplitterDistance = 400
    } catch { }
    $State.Rebuilding = $false   # ListView raises ItemChecked while creating its handle; ignore those
    Update-ServerSummary
    [void]$txtEditor.Focus()
  } 'form Shown'
})

$form.add_FormClosing({
    param($s, $e)
    Invoke-Safe {
        $closeIt = $true
        if ($State.Running) {
            if (Ask-YesNo 'A query is still running. Cancel it and exit?') {
                Stop-Execution
                foreach ($j in $State.Jobs) { if (-not $j.Done) { try { [void]$j.PS.BeginStop($null, $null) } catch { } } }
            } else {
                $e.Cancel = $true
                $closeIt = $false
            }
        }
        if ($closeIt) {
            try { Save-ServerCache } catch { }
            try { Set-Content -Path $Paths.Draft -Value $txtEditor.Text -Encoding UTF8 } catch { }
        }
    } 'form Closing'
})

# ------------------------------------------------------------------ startup
$State.Settings = Load-Settings
Load-ServerCache
Load-Library
Show-SettingsInUi
$numTimeout.Value = [Math]::Min([Math]::Max([int]$State.Settings.QueryTimeout, 0), 86400)
$numMaxRows.Value = [Math]::Min([Math]::Max([int]$State.Settings.MaxRows, 1), 10000000)
$cboDatabase.Text = [string]$State.Settings.DefaultDatabase
if (Test-Path $Paths.Draft) { try { $txtEditor.Text = ConvertTo-CrLf (Get-Content -Path $Paths.Draft -Raw -Encoding UTF8) } catch { } }
if (-not $txtEditor.Text.Trim()) { $txtEditor.Text = "-- F5 runs the selection (or everything). Check servers on the left first.`r`nSELECT @@SERVERNAME AS server_name, @@VERSION AS version;" }
Update-ServerList
Update-LibraryList $null

if (-not $State.Settings.CentralServer) {
    $tabsMain.SelectedTab = $tabSettings
    $lblStatus.Text = 'First run: enter your central server, then click "Save + refresh server list". Or add servers manually with "Add...".'
} elseif ($State.Servers.Count -eq 0) {
    $lblStatus.Text = 'No cached servers yet - click Refresh.'
} else {
    $lblStatus.Text = "Loaded $($State.Servers.Count) cached servers ($($State.Source), $($State.RefreshedAt)). Click Refresh to update."
}

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-ErrorLog $e.Exception 'unhandled UI exception'
    try { Show-Err ($e.Exception.Message + "`r`n`r`nThe app is still running. Details written to $($Paths.ErrorLog)") } catch { }
})

$State.Rebuilding = $true
try {
    [void]$form.ShowDialog()
} catch {
    Write-ErrorLog $_ 'ShowDialog'
    [void][System.Windows.Forms.MessageBox]::Show(("SQL Fleet Studio stopped: " + (Get-InnerMessage $_) + "`r`n`r`nDetails: $($Paths.ErrorLog)"), $AppName)
} finally {
    $form.Dispose()
}

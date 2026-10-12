<#
.SYNOPSIS
    SQL Server Blocking & Deadlock Monitor (WinForms + PowerShell)

.DESCRIPTION
    Tracks ONLY blocking and deadlocks, and explains them in plain English with recommended fixes.

    * Every refresh cycle: CONNECT -> QUERY -> DISCONNECT.
      Connection pooling is turned OFF, so the connection is really closed between cycles.
      The next cycle is scheduled only after the previous one has finished (no overlap, no
      non-stop hammering). Optional "Stop after" limit ends monitoring automatically.
    * Monitoring queries are written so they neither cause blocking nor get stuck behind it:
        SET LOCK_TIMEOUT 500           -> give up after 0.5 s if anything is locked
        SET DEADLOCK_PRIORITY LOW      -> the monitor is always the preferred deadlock victim
        READ UNCOMMITTED + DMVs only   -> no shared locks on user data
        MAXDOP 1, short command timeouts
    * Blocking tab : live chains, head blocker, who/what/why + fixes, KILL guidance.
    * Deadlocks tab: pulls deadlocks from the built-in system_health session
                     (ring buffer = fast, event files = complete), or Open / Paste any
                     deadlock graph (.xdl / XML) and get a full dissection + recommendations.
    * AI Prompt tab: ready-to-paste prompt for any AI assistant if you want a deeper review.

.REQUIREMENTS
    Windows PowerShell 5.1 (built into Windows) or PowerShell 7+ on Windows.
    The login needs VIEW SERVER STATE (SQL 2022+: VIEW SERVER PERFORMANCE STATE is enough).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SqlBlockingDeadlockMonitor.ps1
#>

#region ===================== Bootstrap (WinForms needs an STA thread) =====================
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
try { Add-Type -AssemblyName System.Data.SqlClient -ErrorAction Stop } catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
#endregion

#region ===================== State =====================
$script:S = @{
    Running         = $false
    Busy            = $false
    RunStart        = $null
    Cycles          = 0
    Failures        = 0
    Sessions        = @()
    LastBlockingRep = $null
    LastBlockPrompt = ''
    HeadSeen        = @{}
    Episode         = $null
    LastEpisodeText = ''
    DbById          = @{}
    DbInfo          = @{}
    ServerName      = ''
    ServerVersion   = ''
    Deadlocks       = @{}
    DlBaselineDone  = $false
    NewDeadlocks    = 0
    CurrentDlKey    = $null
}
$script:ResCache  = @{}
$script:ResBudget = 20
#endregion

#region ===================== SQL (all read-only, DMV based) =====================
# Prepended to EVERY batch the tool sends.
$Q_SessionSettings = 'SET NOCOUNT ON; SET LOCK_TIMEOUT 500; SET DEADLOCK_PRIORITY LOW; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;'

$Q_Blocking = @'
;WITH r AS (
    SELECT er.session_id, er.blocking_session_id, er.wait_type, er.wait_time, er.wait_resource,
           er.status, er.command, er.database_id, er.sql_handle,
           er.statement_start_offset, er.statement_end_offset,
           er.cpu_time, er.logical_reads, er.transaction_isolation_level
    FROM sys.dm_exec_requests AS er
    WHERE er.session_id <> @@SPID
),
inv AS (
    SELECT session_id FROM r WHERE blocking_session_id > 0
    UNION
    SELECT blocking_session_id FROM r WHERE blocking_session_id > 0
)
SELECT
    s.session_id                                                        AS SessionId,
    CASE WHEN r.blocking_session_id > 0 THEN r.blocking_session_id ELSE 0 END AS BlockedBy,
    COALESCE(r.status, s.status)                                        AS Status,
    ISNULL(r.wait_type, N'')                                            AS WaitType,
    CAST(ISNULL(r.wait_time, 0) / 1000.0 AS decimal(18,1))              AS WaitSec,
    ISNULL(r.wait_resource, N'')                                        AS WaitResource,
    ISNULL(DB_NAME(COALESCE(r.database_id, s.database_id)), N'')        AS DatabaseName,
    ISNULL(s.login_name, N'')                                           AS LoginName,
    ISNULL(s.host_name, N'')                                            AS HostName,
    ISNULL(s.program_name, N'')                                         AS ProgramName,
    ISNULL(r.command, N'')                                              AS Command,
    s.open_transaction_count                                            AS OpenTran,
    CASE COALESCE(r.transaction_isolation_level, s.transaction_isolation_level)
        WHEN 1 THEN N'Read Uncommitted' WHEN 2 THEN N'Read Committed'
        WHEN 3 THEN N'Repeatable Read'  WHEN 4 THEN N'Serializable'
        WHEN 5 THEN N'Snapshot' ELSE N'Unspecified' END                 AS IsolationLevel,
    CASE WHEN r.session_id IS NULL
         THEN DATEDIFF(SECOND, s.last_request_end_time, GETDATE()) ELSE 0 END AS IdleSec,
    ISNULL(DATEDIFF(SECOND, tx.TranBegin, GETDATE()), 0)                AS TranAgeSec,
    COALESCE(r.cpu_time, s.cpu_time)                                    AS CpuMs,
    COALESCE(r.logical_reads, s.logical_reads)                          AS LogicalReads,
    t.objectid                                                          AS ObjId,
    t.dbid                                                              AS ObjDbId,
    ISNULL(CASE
        WHEN r.sql_handle IS NULL THEN t.text
        WHEN r.statement_end_offset > r.statement_start_offset
             THEN SUBSTRING(t.text, (r.statement_start_offset / 2) + 1,
                            ((r.statement_end_offset - r.statement_start_offset) / 2) + 1)
        ELSE SUBSTRING(t.text, (r.statement_start_offset / 2) + 1, 2147483647)
    END, N'')                                                           AS SqlStatement
FROM inv
JOIN sys.dm_exec_sessions AS s ON s.session_id = inv.session_id
LEFT JOIN r ON r.session_id = s.session_id
LEFT JOIN sys.dm_exec_connections AS c ON c.session_id = s.session_id AND c.parent_connection_id IS NULL
OUTER APPLY sys.dm_exec_sql_text(COALESCE(r.sql_handle, c.most_recent_sql_handle)) AS t
OUTER APPLY (SELECT MIN(at.transaction_begin_time) AS TranBegin
             FROM sys.dm_tran_session_transactions AS st
             JOIN sys.dm_tran_active_transactions AS at ON at.transaction_id = st.transaction_id
             WHERE st.session_id = s.session_id) AS tx
ORDER BY BlockedBy, SessionId
OPTION (MAXDOP 1);
'@

$Q_DbInfo = @'
SELECT database_id, name, is_read_committed_snapshot_on, snapshot_isolation_state_desc,
       CAST(@@SERVERNAME AS nvarchar(256)) AS ServerName,
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(50)) AS ProductVersion
FROM sys.databases OPTION (MAXDOP 1);
'@

$Q_RingBuffer = @'
SELECT CAST(t.target_data AS nvarchar(max)) AS TargetData
FROM sys.dm_xe_session_targets AS t
JOIN sys.dm_xe_sessions AS s ON s.address = t.event_session_address
WHERE s.name = N'system_health' AND t.target_name = N'ring_buffer'
OPTION (MAXDOP 1);
'@

$Q_EventFile = @'
SELECT CAST(event_data AS nvarchar(max)) AS EventData
FROM sys.fn_xe_file_target_read_file(N'system_health*.xel', NULL, NULL, NULL)
WHERE object_name = N'xml_deadlock_report'
OPTION (MAXDOP 1);
'@
#endregion

#region ===================== Generic helpers =====================
function Get-Val($Row, [string]$Col, $Default) {
    $v = $Row[$Col]
    if ($null -eq $v -or $v -is [System.DBNull]) { return $Default }
    return $v
}
function Format-SqlText([string]$s, [int]$Max = 0) {
    if (-not $s) { return '' }
    $t = ($s -replace '\s+', ' ').Trim()
    if ($Max -gt 0 -and $t.Length -gt $Max) { $t = $t.Substring(0, $Max) + ' ...' }
    return $t
}
function Get-SqlVerb([string]$s) {
    if (-not $s) { return '' }
    $x = [regex]::Replace($s, '/\*[\s\S]*?\*/', ' ')
    $x = [regex]::Replace($x, '--[^\r\n]*', ' ')
    $m = [regex]::Match($x, '\b(SELECT|INSERT|UPDATE|DELETE|MERGE|ALTER|CREATE|TRUNCATE|DBCC|EXEC|EXECUTE)\b', 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value.ToUpperInvariant() }
    return ''
}
function Format-Dur([double]$Sec) {
    if ($Sec -lt 60) { return ('{0:0}s' -f $Sec) }
    if ($Sec -lt 3600) { return ('{0}m {1}s' -f [math]::Floor($Sec / 60), [math]::Floor($Sec % 60)) }
    return ('{0}h {1}m' -f [math]::Floor($Sec / 3600), [math]::Floor(($Sec % 3600) / 60))
}
function Coalesce($a, $b) { if ($a) { return $a } return $b }
function QN([string]$n) { return '[' + $n.Replace(']', ']]') + ']' }
function Get-TextHash([string]$s) {
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try { return [System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s))).Replace('-', '') }
    finally { $sha.Dispose() }
}
function Get-DbNameById($Id) {
    if ($null -eq $Id -or "$Id" -eq '') { return '' }
    $i = 0
    if ([int]::TryParse("$Id", [ref]$i) -and $script:S.DbById.ContainsKey($i)) { return $script:S.DbById[$i] }
    return "database id $Id"
}

# Report = list of {T = text, S = style}. Styles: h1 h2 b n bad warn good code dim
function New-Report { $r = New-Object System.Collections.ArrayList; return , $r }
function Add-R($Rep, [string]$Text = '', [string]$Style = 'n') { [void]$Rep.Add([pscustomobject]@{ T = $Text; S = $Style }) }
function Get-ReportText($Rep) { return (@($Rep | ForEach-Object { $_.T }) -join "`r`n") }
#endregion

#region ===================== Plain-English dictionaries =====================
function Get-LockModeText([string]$m) {
    if (-not $m) { return 'unknown lock' }
    switch -Regex ($m) {
        '^S$'     { return 'shared (S = read) lock' }
        '^U$'     { return 'update (U = about to change) lock' }
        '^X$'     { return 'exclusive (X = change) lock' }
        '^IS$'    { return 'intent-shared (IS = reading something inside) lock' }
        '^IX$'    { return 'intent-exclusive (IX = changing something inside) lock' }
        '^IU$'    { return 'intent-update (IU) lock' }
        '^SIU$'   { return 'shared intent-update (SIU) lock' }
        '^SIX$'   { return 'shared + intent-exclusive (SIX = reading all, changing some) lock' }
        '^UIX$'   { return 'update intent-exclusive (UIX) lock' }
        '^Sch-S$' { return 'schema-stability (Sch-S = query is using the table definition) lock' }
        '^Sch-M$' { return 'schema-modification (Sch-M = changing the table structure) lock' }
        '^BU$'    { return 'bulk-update (BU) lock' }
        '^Range'  { return "$m range lock (protects a RANGE of keys - Serializable isolation)" }
        default   { return "$m lock" }
    }
}

function Get-WaitText([string]$wt) {
    if (-not $wt) { return 'running (not waiting on anything)' }
    switch -Regex ($wt) {
        '^LCK_M_(S|IS)$'            { return 'waiting to READ data that another session is changing' }
        '^LCK_M_(U|IU)$'            { return 'waiting to start CHANGING data (update lock)' }
        '^LCK_M_(X|IX|SIX|UIX|SIU)$' { return 'waiting to CHANGE data that another session is using' }
        '^LCK_M_SCH_S$'             { return 'waiting for a table STRUCTURE change (index rebuild / ALTER) to finish' }
        '^LCK_M_SCH_M$'             { return 'waiting to CHANGE a table structure (DDL / index maintenance) - must wait for everyone using the table' }
        '^LCK_M_R'                  { return 'waiting on a RANGE lock (Serializable isolation level)' }
        '^LCK_M_BU'                 { return 'waiting for a bulk-load lock' }
        '^LCK_M_'                   { return 'waiting for a lock' }
        '^PAGEIOLATCH'              { return 'reading data from disk into memory' }
        '^PAGELATCH'                { return 'in-memory page contention (hot page - not a lock)' }
        '^ASYNC_NETWORK_IO$'        { return 'waiting for the CLIENT APPLICATION to read the results' }
        '^WRITELOG$'                { return 'waiting for the transaction log disk to save changes' }
        '^(CXPACKET|CXCONSUMER|CXSYNC)' { return 'parallel query threads coordinating with each other' }
        '^SOS_SCHEDULER_YIELD$'     { return 'busy using CPU' }
        '^RESOURCE_SEMAPHORE'       { return 'waiting for query memory' }
        '^(OLEDB|MSQL_DQ|PREEMPTIVE)' { return 'waiting on an external call (linked server / OS / network)' }
        default                     { return "waiting on $wt" }
    }
}
#endregion

#region ===================== Connection + queries =====================
function Open-MonitorConnection {
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source'] = $txtServer.Text.Trim()
    $b['Initial Catalog'] = 'master'
    $b['Application Name'] = 'BlockingDeadlockMonitor'
    $b['Connect Timeout'] = 8
    $b['Pooling'] = $false          # really disconnect between refreshes
    if ($cboAuth.SelectedIndex -eq 0) {
        $b['Integrated Security'] = $true
    }
    else {
        $b['Integrated Security'] = $false
        $b['User ID'] = $txtUser.Text
        $b['Password'] = $txtPwd.Text
    }
    $b['Encrypt'] = $chkEncrypt.Checked
    $b['TrustServerCertificate'] = $chkTrust.Checked
    $c = New-Object System.Data.SqlClient.SqlConnection ($b.ConnectionString)
    $c.Open()
    return $c
}

function Invoke-MonitorQuery($Conn, [string]$Sql, [int]$Timeout = 15) {
    $cmd = $Conn.CreateCommand()
    $cmd.CommandText = $Q_SessionSettings + "`r`n" + $Sql
    $cmd.CommandTimeout = $Timeout
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    try { [void]$da.Fill($dt) } finally { $da.Dispose(); $cmd.Dispose() }
    return , $dt
}

function Invoke-SafeScalar($Conn, [string]$Sql) {
    try {
        $cmd = $Conn.CreateCommand()
        $cmd.CommandText = $Q_SessionSettings + "`r`n" + $Sql
        $cmd.CommandTimeout = 5
        $v = $cmd.ExecuteScalar()
        $cmd.Dispose()
        if ($null -eq $v -or $v -is [System.DBNull]) { return $null }
        return [string]$v
    }
    catch { return $null }
}

function Update-DbInfo($Conn) {
    $dt = Invoke-MonitorQuery $Conn $Q_DbInfo 10
    $byId = @{}; $info = @{}
    foreach ($r in $dt.Rows) {
        $id = [int]$r['database_id']; $n = [string]$r['name']
        $byId[$id] = $n
        $info[$n] = [pscustomobject]@{ Id = $id; Rcsi = [bool](Get-Val $r 'is_read_committed_snapshot_on' $false); Snapshot = [string](Get-Val $r 'snapshot_isolation_state_desc' '') }
        $script:S.ServerName = [string](Get-Val $r 'ServerName' '')
        $script:S.ServerVersion = [string](Get-Val $r 'ProductVersion' '')
    }
    $script:S.DbById = $byId
    $script:S.DbInfo = $info
}

function Get-CachedName($Conn, [string]$Key, [int]$DbId, [string]$SqlTemplate) {
    if ($script:ResCache.ContainsKey($Key)) { return $script:ResCache[$Key] }
    if ($script:ResBudget -le 0) { return $null }
    $script:ResBudget--
    if (-not $script:S.DbById.ContainsKey($DbId)) { return $null }
    $v = Invoke-SafeScalar $Conn ($SqlTemplate -f (QN $script:S.DbById[$DbId]))
    $script:ResCache[$Key] = $v
    return $v
}

# Turns "KEY: 5:72057594043236352 (8194443284a0)" into "row(s) in [Sales] dbo.Orders (index PK_Orders)"
function Convert-WaitResource($Conn, [string]$Wr) {
    if (-not $Wr) { return '' }
    $w = $Wr.Trim()
    if ($w -match '^KEY:\s*(\d+):(\d+)') {
        $dbid = [int]$Matches[1]; $hobt = $Matches[2]; $dbn = Get-DbNameById $dbid
        $sql = "SELECT TOP (1) s.name + N'.' + o.name + N' (index ' + ISNULL(i.name, N'heap') + N')' FROM {0}.sys.partitions AS p JOIN {0}.sys.objects AS o ON o.object_id = p.object_id JOIN {0}.sys.schemas AS s ON s.schema_id = o.schema_id LEFT JOIN {0}.sys.indexes AS i ON i.object_id = p.object_id AND i.index_id = p.index_id WHERE p.hobt_id = $hobt;"
        $n = Get-CachedName $Conn "H|$dbid|$hobt" $dbid $sql
        if ($n) { return "row(s) in [$dbn] $n" }
        return "row(s) in database [$dbn] (hobt $hobt)"
    }
    if ($w -match '^(PAGE|RID):\s*(\d+):(\d+):(\d+)') {
        $kind = $Matches[1]; $dbid = [int]$Matches[2]; $file = $Matches[3]; $page = $Matches[4]; $dbn = Get-DbNameById $dbid
        $sql = "SELECT TOP (1) s.name + N'.' + o.name FROM sys.dm_db_page_info($dbid, $file, $page, 'LIMITED') AS pi JOIN {0}.sys.objects AS o ON o.object_id = pi.object_id JOIN {0}.sys.schemas AS s ON s.schema_id = o.schema_id;"
        $n = Get-CachedName $Conn "P|$dbid|$file|$page" $dbid $sql
        $what = if ($kind -eq 'RID') { 'a row (heap table - no clustered index)' } else { 'a whole page (8 KB block of rows)' }
        if ($n) { return "$what in [$dbn] $n" }
        return "$what in database [$dbn] (file $file, page $page)"
    }
    if ($w -match '^(?:OBJECT|TAB):\s*(\d+):(\d+)') {
        $dbid = [int]$Matches[1]; $obj = $Matches[2]; $dbn = Get-DbNameById $dbid
        $sql = "SELECT s.name + N'.' + o.name FROM {0}.sys.objects AS o JOIN {0}.sys.schemas AS s ON s.schema_id = o.schema_id WHERE o.object_id = $obj;"
        $n = Get-CachedName $Conn "O|$dbid|$obj" $dbid $sql
        if ($n) { return "the WHOLE table [$dbn] $n" }
        return "a whole object in database [$dbn] (object id $obj)"
    }
    if ($w -match '^DATABASE:\s*(\d+)') { return 'the whole database ' + (Get-DbNameById ([int]$Matches[1])) }
    if ($w -match '^METADATA:') { return 'object definitions (metadata) - ' + $w }
    if ($w -match '^APPLICATION:') { return 'an application lock (sp_getapplock) - ' + $w }
    return $w
}

function Convert-BlockingRows($Dt) {
    $list = New-Object System.Collections.ArrayList
    foreach ($r in $Dt.Rows) {
        [void]$list.Add([pscustomobject]@{
                SessionId      = [int](Get-Val $r 'SessionId' 0)
                BlockedBy      = [int](Get-Val $r 'BlockedBy' 0)
                Status         = [string](Get-Val $r 'Status' '')
                WaitType       = [string](Get-Val $r 'WaitType' '')
                WaitSec        = [double](Get-Val $r 'WaitSec' 0)
                WaitResource   = [string](Get-Val $r 'WaitResource' '')
                DatabaseName   = [string](Get-Val $r 'DatabaseName' '')
                LoginName      = [string](Get-Val $r 'LoginName' '')
                HostName       = [string](Get-Val $r 'HostName' '')
                ProgramName    = [string](Get-Val $r 'ProgramName' '')
                Command        = [string](Get-Val $r 'Command' '')
                OpenTran       = [int](Get-Val $r 'OpenTran' 0)
                IsolationLevel = [string](Get-Val $r 'IsolationLevel' '')
                IdleSec        = [int](Get-Val $r 'IdleSec' 0)
                TranAgeSec     = [int](Get-Val $r 'TranAgeSec' 0)
                CpuMs          = [long](Get-Val $r 'CpuMs' 0)
                LogicalReads   = [long](Get-Val $r 'LogicalReads' 0)
                ObjId          = (Get-Val $r 'ObjId' $null)
                ObjDbId        = (Get-Val $r 'ObjDbId' $null)
                SqlStatement   = [string](Get-Val $r 'SqlStatement' '')
                WaitingFor     = ''
                ProcName       = ''
            })
    }
    return , $list
}

function Resolve-SessionNames($Conn, $Sessions) {
    $script:ResBudget = 20
    if ($script:ResCache.Count -gt 2000) { $script:ResCache.Clear() }
    foreach ($s in $Sessions) {
        if ($s.WaitResource) { $s.WaitingFor = Convert-WaitResource $Conn $s.WaitResource }
        if ($null -ne $s.ObjId -and $null -ne $s.ObjDbId) {
            $dbid = [int]$s.ObjDbId; $oid = [int]$s.ObjId
            $sql = "SELECT s.name + N'.' + o.name FROM {0}.sys.objects AS o JOIN {0}.sys.schemas AS s ON s.schema_id = o.schema_id WHERE o.object_id = $oid;"
            $n = Get-CachedName $Conn "O|$dbid|$oid" $dbid $sql
            if ($n) { $s.ProcName = (Get-DbNameById $dbid) + '.' + $n }
        }
    }
}
#endregion

#region ===================== Blocking analysis =====================
function Get-Descendants($Id, $Kids) {
    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    $q = New-Object System.Collections.Queue
    $q.Enqueue($Id)
    while ($q.Count -gt 0) {
        $cur = $q.Dequeue()
        if ($Kids.ContainsKey($cur)) {
            foreach ($c in $Kids[$cur]) {
                if (-not $seen.ContainsKey($c.SessionId)) { $seen[$c.SessionId] = $true; [void]$out.Add($c); $q.Enqueue($c.SessionId) }
            }
        }
    }
    return , $out
}

function Add-BlockTree($Rep, $Id, $Kids, [int]$Depth, $Visited) {
    if ($Depth -gt 12 -or -not $Kids.ContainsKey($Id)) { return }
    foreach ($c in @($Kids[$Id] | Sort-Object WaitSec -Descending)) {
        if ($Visited.ContainsKey($c.SessionId)) { continue }
        $Visited[$c.SessionId] = $true
        $pad = '    ' * $Depth
        $what = Coalesce $c.WaitingFor $c.WaitResource
        Add-R $Rep ('{0}+-- Session {1}  waiting {2}  - {3}' -f $pad, $c.SessionId, (Format-Dur $c.WaitSec), (Get-WaitText $c.WaitType)) 'code'
        Add-R $Rep ('{0}      on: {1}   [{2} | {3} | {4}]' -f $pad, $what, $c.LoginName, $c.HostName, $c.ProgramName) 'code'
        Add-BlockTree $Rep $c.SessionId $Kids ($Depth + 1) $Visited
    }
}

function New-Cause([string]$Title, [string]$Plain, [string[]]$Fixes) {
    return [pscustomobject]@{ Title = $Title; Plain = $Plain; Fixes = $Fixes }
}

function Get-HeadCauses($H, $Desc) {
    $list = New-Object System.Collections.ArrayList
    $app = $H.ProgramName
    $isTool = $app -match 'Management Studio|Azure Data Studio|SQLCMD|DBeaver|Toad|LINQPad|DataGrip|Query Analyzer'
    $isAgent = $app -match '^SQLAgent'
    $stmtU = ([string]$H.SqlStatement).ToUpperInvariant()
    $killLine = "Last resort (only after you know what it is): KILL $($H.SessionId);  -- its unsaved work is rolled back (undo can take as long as the work took) and the user/app gets an error."

    if ($H.Status -eq 'unknown') {
        [void]$list.Add((New-Cause 'Blocker not visible' "Session $($H.SessionId) is blocking others but no longer shows up as a session (it may have just ended, or it is an orphaned / distributed transaction)." @('Refresh once more - it may already be gone.', 'If it persists, check sys.dm_tran_active_transactions and the DTC / linked-server activity for orphaned distributed transactions.')))
        return , $list
    }

    if ($H.Status -eq 'sleeping' -and $H.OpenTran -gt 0) {
        $plain = "Session $($H.SessionId) started a change and then went quiet - it has not SAVED (COMMIT) or UNDONE (ROLLBACK) its work. " +
        "Everything it touched stays locked until it finishes, so others queue up behind it. " +
        "It has been idle for $(Format-Dur $H.IdleSec)" + $(if ($H.TranAgeSec -gt 0) { " and its transaction has been open for $(Format-Dur $H.TranAgeSec)." } else { '.' })
        $fixes = New-Object System.Collections.ArrayList
        if ($isTool) {
            [void]$fixes.Add("This is a PERSON using a query tool ($app) on computer '$($H.HostName)' as '$($H.LoginName)'. Ask them to run COMMIT (keep changes) or ROLLBACK (discard), or close that query window.")
            [void]$fixes.Add('Remind tool users: in SSMS, a BEGIN TRAN (or SET IMPLICIT_TRANSACTIONS ON) without COMMIT keeps locks forever.')
        }
        elseif ($isAgent) {
            [void]$fixes.Add("A SQL Agent job step ($app) left a transaction open. Review that job step's error handling (BEGIN TRY / CATCH with ROLLBACK, SET XACT_ABORT ON).")
        }
        else {
            [void]$fixes.Add("The application '$app' on '$($H.HostName)' did not finish its transaction. Developers: every BEGIN TRAN must end in COMMIT or ROLLBACK - including when an error happens (use TRY/CATCH and SET XACT_ABORT ON).")
            [void]$fixes.Add('Never wait for user input, web calls or file I/O while a transaction is open.')
            [void]$fixes.Add('Check whether the driver runs with auto-commit OFF / IMPLICIT_TRANSACTIONS ON (common with JDBC/ODBC) - then every statement silently opens a transaction.')
        }
        [void]$fixes.Add($killLine)
        [void]$list.Add((New-Cause 'Unfinished (idle) transaction - the #1 cause of long blocking' $plain $fixes))
    }
    elseif ($H.Status -eq 'sleeping') {
        [void]$list.Add((New-Cause 'Idle session still holding locks' "Session $($H.SessionId) is idle and reports no open transaction, yet others wait on it. This usually means the picture changed between snapshots, or it is part of a distributed (DTC/linked server) transaction." @('Refresh - if it disappears it was a timing artifact.', 'If it persists, look for distributed transactions (sys.dm_tran_active_transactions, transaction_type = 4).')))
    }
    else {
        $wt = $H.WaitType
        if ($H.Command -match 'ALTER INDEX|CREATE INDEX|DBCC|UPDATE STATISTICS|ALTER TABLE' -or $stmtU -match 'ALTER\s+INDEX|CREATE\s+(UNIQUE\s+)?(CLUSTERED\s+|NONCLUSTERED\s+)?INDEX|UPDATE\s+STATISTICS|ALTER\s+TABLE|TRUNCATE\s+TABLE|DBCC\s') {
            [void]$list.Add((New-Cause 'Maintenance / structure change is running' "Session $($H.SessionId) is changing a table's structure or doing maintenance ($($H.Command)). That needs exclusive access, so normal work on that table waits." @('Run index/statistics maintenance and schema changes in a quiet window.', 'For index rebuilds use ONLINE = ON (Enterprise / Azure) with WAIT_AT_LOW_PRIORITY so it steps aside for user work.', 'If business is impacted now, consider stopping the maintenance job and rescheduling it.')))
        }
        if ($wt -eq 'ASYNC_NETWORK_IO') {
            [void]$list.Add((New-Cause 'Application is slow to read its results' "SQL Server finished the work for session $($H.SessionId) but the APPLICATION ('$app' on '$($H.HostName)') is reading the results slowly, so the locks stay held while it waits." @('Fix the app to read results fully and quickly (do not process row-by-row while the reader is open).', 'Return fewer rows (paging, filtering) - is someone pulling a huge result set into a grid/Excel?', 'Check network/client machine load.')))
        }
        elseif ($wt -eq 'WRITELOG') {
            [void]$list.Add((New-Cause 'Slow transaction-log disk' 'The head blocker is waiting for the log disk to save its changes, so it holds its locks longer.' @('Check log-drive latency (should be < 2-5 ms for writes).', 'Avoid committing one row at a time in big loops - commit in batches.', 'Move the log file to faster storage.')))
        }
        elseif ($wt -match '^PAGEIOLATCH') {
            [void]$list.Add((New-Cause 'Reading lots of data from disk' 'The head blocker is reading a large amount of data from disk while holding locks - usually a scan caused by a missing index.' @('Check the statement''s execution plan for scans; add a suitable index.', 'Check memory pressure (page life expectancy) and disk latency.')))
        }
        elseif ($wt -match '^(CXPACKET|CXCONSUMER|CXSYNC)') {
            [void]$list.Add((New-Cause 'Large parallel query' 'The head blocker is a big query running on many threads, holding locks while it works.' @('Tune the query (indexes, updated statistics) so it touches less data.', 'Run heavy reports on a readable secondary / reporting copy, or outside business hours.')))
        }
        elseif ($wt -match '^RESOURCE_SEMAPHORE') {
            [void]$list.Add((New-Cause 'Waiting for memory' 'The head blocker is waiting for memory to run its query while still holding locks.' @('Look for queries with huge memory grants; fix estimates with statistics/indexes.', 'Check max server memory settings.')))
        }
        elseif ($list.Count -eq 0) {
            [void]$list.Add((New-Cause 'Long-running statement' "Session $($H.SessionId) is actively working ($($H.Command), $(Get-WaitText $wt)) and holds locks until it finishes. CPU so far: $($H.CpuMs) ms, logical reads: $($H.LogicalReads)." @('Look at this statement''s execution plan - big scans / many reads usually mean a missing index.', 'Break large updates/deletes into small batches (e.g. 1,000-4,000 rows per transaction).', 'Keep transactions short: prepare data first, then BEGIN TRAN ... COMMIT quickly.')))
        }
    }

    if ($H.IsolationLevel -match 'Serializable|Repeatable') {
        [void]$list.Add((New-Cause 'Strict isolation level' "Session $($H.SessionId) runs at $($H.IsolationLevel). That level keeps read locks until the end of the transaction (and Serializable also locks ranges), so it blocks much more." @('If the app uses .NET TransactionScope, set IsolationLevel = ReadCommitted (the default is Serializable!).', 'Only use Serializable/Repeatable Read where it is truly required.')))
    }
    if ($H.Status -ne 'sleeping' -and $H.TranAgeSec -ge 60) {
        [void]$list.Add((New-Cause 'Long transaction' "Its transaction has been open for $(Format-Dur $H.TranAgeSec). Locks are held for the whole transaction, not just the current statement." @('Commit more often / split the work into smaller transactions.')))
    }

    $waitTypes = @($Desc | ForEach-Object { $_.WaitType })
    $dbn = $H.DatabaseName
    $rcsiOn = $false
    if ($dbn -and $script:S.DbInfo.ContainsKey($dbn)) { $rcsiOn = $script:S.DbInfo[$dbn].Rcsi }
    if (@($waitTypes | Where-Object { $_ -match '^LCK_M_(S|IS)$' }).Count -gt 0) {
        if ($rcsiOn) {
            [void]$list.Add((New-Cause 'Readers blocked by writers (RCSI already ON)' "READ_COMMITTED_SNAPSHOT is already ON for [$dbn], yet readers still wait - they use a stricter level or lock hints (HOLDLOCK, REPEATABLEREAD, UPDLOCK, SERIALIZABLE)." @('Find and remove unnecessary locking hints / isolation settings in the reading code.')))
        }
        else {
            [void]$list.Add((New-Cause 'Readers are blocked by writers' 'Some sessions only want to READ, but must wait because data is being changed.' @("Consider READ_COMMITTED_SNAPSHOT for [$(Coalesce $dbn 'the database')]: readers then see the last committed version instead of waiting. ALTER DATABASE $(QN (Coalesce $dbn 'YourDb')) SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;  -- test first: adds tempdb version-store load and changes read behaviour slightly.", 'Do NOT use NOLOCK as a fix - it can return wrong/duplicated rows.')))
        }
    }
    if ($waitTypes -contains 'LCK_M_SCH_M') {
        [void]$list.Add((New-Cause 'A structure change is stuck in the queue' 'Someone is trying to change a table (ALTER / index rebuild / partition switch). It waits for the head blocker, and EVERY new query on that table now waits behind it too - that is why the queue grows fast.' @('Cancel or reschedule the DDL / maintenance until the head blocker is done.', 'Use WAIT_AT_LOW_PRIORITY (MAX_DURATION, ABORT_AFTER_WAIT = SELF) for online index operations.')))
    }
    if (@($waitTypes | Where-Object { $_ -match '^LCK_M_R' }).Count -gt 0) {
        [void]$list.Add((New-Cause 'Range locks (Serializable)' 'Waiters are asking for range locks - something runs at Serializable isolation (TransactionScope default, HOLDLOCK hint, MERGE ... HOLDLOCK).' @('Use Read Committed unless Serializable is truly required; add indexes so ranges stay small.')))
    }
    return , $list
}

function Get-BlockingAnalysis($Sessions) {
    $rep = New-Report
    $res = [pscustomobject]@{ Report = $rep; Prompt = ''; Roles = @{}; Counts = @{}; Blocked = 0; MaxWait = 0.0; HeadIds = @(); Severity = 'None'; MainCause = '' }
    if (-not $Sessions -or $Sessions.Count -eq 0) { return $res }

    $byId = @{}
    foreach ($s in $Sessions) { $byId[$s.SessionId] = $s }
    # blockers that are no longer visible as sessions
    foreach ($s in @($Sessions)) {
        if ($s.BlockedBy -gt 0 -and -not $byId.ContainsKey($s.BlockedBy)) {
            $ph = [pscustomobject]@{ SessionId = $s.BlockedBy; BlockedBy = 0; Status = 'unknown'; WaitType = ''; WaitSec = 0.0; WaitResource = ''; DatabaseName = ''; LoginName = '?'; HostName = '?'; ProgramName = '?'; Command = ''; OpenTran = 0; IsolationLevel = ''; IdleSec = 0; TranAgeSec = 0; CpuMs = 0; LogicalReads = 0; ObjId = $null; ObjDbId = $null; SqlStatement = ''; WaitingFor = ''; ProcName = '' }
            $byId[$ph.SessionId] = $ph
        }
    }
    $all = @($byId.Values)
    $kids = @{}
    foreach ($s in $all) {
        if ($s.BlockedBy -gt 0) {
            if (-not $kids.ContainsKey($s.BlockedBy)) { $kids[$s.BlockedBy] = New-Object System.Collections.ArrayList }
            [void]$kids[$s.BlockedBy].Add($s)
        }
    }
    $blocked = @($all | Where-Object { $_.BlockedBy -gt 0 })
    $res.Blocked = $blocked.Count
    $res.MaxWait = [double](($blocked | Measure-Object -Property WaitSec -Maximum).Maximum)
    $heads = @($all | Where-Object { $kids.ContainsKey($_.SessionId) -and $_.BlockedBy -eq 0 })

    foreach ($s in $all) {
        $role = 'Blocked'
        if ($kids.ContainsKey($s.SessionId)) { if ($s.BlockedBy -eq 0) { $role = 'HEAD BLOCKER' } else { $role = 'Blocked + blocking' } }
        $res.Roles[$s.SessionId] = $role
        $res.Counts[$s.SessionId] = if ($kids.ContainsKey($s.SessionId)) { (Get-Descendants $s.SessionId $kids).Count } else { 0 }
    }

    if ($res.Blocked -ge 10 -or $res.MaxWait -ge 60) { $res.Severity = 'High' }
    elseif ($res.Blocked -ge 3 -or $res.MaxWait -ge 15) { $res.Severity = 'Medium' }
    else { $res.Severity = 'Low' }

    # how long has each head blocker been blocking?
    $now = Get-Date
    $present = @{}
    foreach ($h in $heads) {
        $k = '{0}|{1}|{2}' -f $h.SessionId, $h.LoginName, $h.HostName
        $present[$k] = $true
        if (-not $script:S.HeadSeen.ContainsKey($k)) { $script:S.HeadSeen[$k] = $now }
    }
    foreach ($k in @($script:S.HeadSeen.Keys)) { if (-not $present.ContainsKey($k)) { $script:S.HeadSeen.Remove($k) } }

    $sevStyle = if ($res.Severity -eq 'High') { 'bad' } else { 'warn' }
    Add-R $rep ("BLOCKING DETECTED  -  {0} severity" -f $res.Severity.ToUpper()) 'h1'
    $headList = (@($heads | ForEach-Object { $_.SessionId }) -join ', ')
    Add-R $rep ("{0} session(s) are waiting. Longest wait so far: {1}. Root cause session(s): {2}" -f $res.Blocked, (Format-Dur $res.MaxWait), (Coalesce $headList 'none - circular wait')) $sevStyle
    Add-R $rep ''
    Add-R $rep 'What is blocking? (plain English)' 'h2'
    Add-R $rep 'SQL Server locks data while it is being read or changed, so two people cannot corrupt the same data. If one session holds a lock for a long time, everyone who needs the same data must queue behind it - like cars stuck behind one car that will not move at a green light. The car at the front is the HEAD BLOCKER. Fix the head blocker and the whole queue moves.' 'dim'
    Add-R $rep ''

    if ($heads.Count -eq 0) {
        Add-R $rep 'Sessions are waiting on each other in a circle' 'h2'
        Add-R $rep 'No session is at the front of the queue - they are waiting on each other. That is a DEADLOCK forming. SQL Server detects these within about 5 seconds and cancels one of them. Check the Deadlocks tab shortly (Fetch from server).' 'bad'
    }

    $n = 0
    foreach ($h in @($heads | Sort-Object { $res.Counts[$_.SessionId] } -Descending)) {
        $n++
        $desc = Get-Descendants $h.SessionId $kids
        $maxW = [double](($desc | Measure-Object -Property WaitSec -Maximum).Maximum)
        Add-R $rep ("Root cause #{0}: Session {1}   ({2} session(s) waiting behind it, longest {3})" -f $n, $h.SessionId, $desc.Count, (Format-Dur $maxW)) 'h2'
        Add-R $rep ("Who            : login '{0}' on computer '{1}' using '{2}'" -f $h.LoginName, $h.HostName, $h.ProgramName)
        Add-R $rep ("Database       : {0}      Isolation level: {1}" -f $h.DatabaseName, $h.IsolationLevel)
        if ($h.Status -eq 'sleeping') { $state = "IDLE - not running anything right now (idle for $(Format-Dur $h.IdleSec))" }
        elseif ($h.Status -eq 'unknown') { $state = 'not visible' }
        else { $state = "ACTIVE ($($h.Status)) - $(Get-WaitText $h.WaitType)" }
        Add-R $rep ("State          : {0}" -f $state)
        Add-R $rep ("Open transactions: {0}{1}" -f $h.OpenTran, $(if ($h.TranAgeSec -gt 0) { "   (oldest open for $(Format-Dur $h.TranAgeSec))" } else { '' }))
        $k = '{0}|{1}|{2}' -f $h.SessionId, $h.LoginName, $h.HostName
        if ($script:S.HeadSeen.ContainsKey($k)) {
            $since = $script:S.HeadSeen[$k]
            Add-R $rep ("Seen blocking  : since {0:HH:mm:ss} (about {1} by this monitor)" -f $since, (Format-Dur ($now - $since).TotalSeconds))
        }
        if ($h.ProcName) { Add-R $rep ("Code object    : {0}" -f $h.ProcName) }
        $stmtLabel = if ($h.Status -eq 'sleeping') { 'LAST statement it ran (may not be the one that took the locks):' } else { 'Statement it is running now:' }
        Add-R $rep $stmtLabel 'b'
        Add-R $rep ('    ' + (Coalesce (Format-SqlText $h.SqlStatement 700) '(not available)')) 'code'
        Add-R $rep 'Queue behind it:' 'b'
        Add-R $rep ("Session {0} (head blocker)" -f $h.SessionId) 'code'
        Add-BlockTree $rep $h.SessionId $kids 0 @{}
        $causes = Get-HeadCauses $h $desc
        if ($n -eq 1 -and $causes.Count -gt 0) { $res.MainCause = $causes[0].Title }
        Add-R $rep 'Why this is happening:' 'b'
        foreach ($c in $causes) {
            Add-R $rep ('* ' + $c.Title) 'warn'
            Add-R $rep ('  ' + $c.Plain)
        }
        Add-R $rep 'What you can do:' 'b'
        $i = 0
        $seenFix = @{}
        foreach ($c in $causes) { foreach ($f in $c.Fixes) { if (-not $seenFix.ContainsKey($f)) { $seenFix[$f] = $true; $i++; Add-R $rep ("  {0}. {1}" -f $i, $f) } } }
        Add-R $rep ''
    }
    $res.HeadIds = @($heads | ForEach-Object { $_.SessionId })

    Add-R $rep 'General guidance' 'h2'
    Add-R $rep '- Short blocking (a second or two) is normal in a busy database. Act when it lasts or keeps growing.' 'dim'
    Add-R $rep '- Never KILL a session before you know what it is: its work is rolled back and the user/app gets an error.' 'dim'
    Add-R $rep '- Blocking that comes back from the same code is a design issue: shorter transactions, the right indexes, and (for reader/writer conflicts) READ_COMMITTED_SNAPSHOT.' 'dim'
    Add-R $rep '- Double-click a row in the grid above for that session''s full details.' 'dim'

    # ---------- AI prompt ----------
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('You are a senior SQL Server DBA. Below is a live BLOCKING snapshot captured from SQL Server DMVs.')
    [void]$sb.AppendLine('Please provide:')
    [void]$sb.AppendLine('1. A plain-English explanation a non-technical manager can understand (what is happening and the business impact).')
    [void]$sb.AppendLine('2. The technical root cause (head blocker, lock types, isolation level, transaction behaviour, likely application pattern).')
    [void]$sb.AppendLine('3. Immediate actions (including whether killing a session is advisable and the risks) and long-term fixes ranked by impact/risk: code/transaction changes, index suggestions with CREATE INDEX scripts, isolation/RCSI options, maintenance scheduling.')
    [void]$sb.AppendLine('4. What extra information (execution plans, table/index definitions, wait stats) you would need to be certain.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('Server: {0}   Version: {1}   Captured: {2:yyyy-MM-dd HH:mm:ss}' -f $script:S.ServerName, $script:S.ServerVersion, $now))
    [void]$sb.AppendLine(('Blocked sessions: {0}   Longest wait: {1}s   Head blocker(s): {2}' -f $res.Blocked, $res.MaxWait, $headList))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Sessions involved (BlockedBy = 0 means it is not blocked itself) ---')
    foreach ($s in @($all | Sort-Object BlockedBy, SessionId)) {
        [void]$sb.AppendLine(('Session={0} Role={1} BlockedBy={2} Status={3} WaitType={4} WaitSec={5} WaitResource="{6}" Resolved="{7}" DB={8} Login={9} Host={10} App="{11}" Command={12} OpenTran={13} TranAgeSec={14} IdleSec={15} Isolation={16} CpuMs={17} Reads={18} Object={19}' -f $s.SessionId, $res.Roles[$s.SessionId], $s.BlockedBy, $s.Status, $s.WaitType, $s.WaitSec, $s.WaitResource, $s.WaitingFor, $s.DatabaseName, $s.LoginName, $s.HostName, $s.ProgramName, $s.Command, $s.OpenTran, $s.TranAgeSec, $s.IdleSec, $s.IsolationLevel, $s.CpuMs, $s.LogicalReads, $s.ProcName))
        [void]$sb.AppendLine('   SQL: ' + (Format-SqlText $s.SqlStatement 1500))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Monitoring tool''s own analysis (for reference) ---')
    [void]$sb.AppendLine((Get-ReportText $rep))
    $res.Prompt = $sb.ToString()
    return $res
}
#endregion

#region ===================== Deadlock parsing + analysis =====================
function Get-DlResourceText($Res) {
    $o = Coalesce $Res.Object 'an object'
    $i = if ($Res.Index) { " [index $($Res.Index)]" } else { '' }
    switch ($Res.Type) {
        'keylock'         { return "row(s) in $o$i" }
        'ridlock'         { return "a row in $o (heap table - no clustered index)" }
        'pagelock'        { return "a whole page (8 KB block of rows) in $o$i" }
        'objectlock'      { return "the WHOLE table $o" }
        'hobtlock'        { return "a whole partition/index of $o$i" }
        'exchangeEvent'   { return 'a parallel-query exchange point (threads of ONE query waiting on each other)' }
        'threadpool'      { return 'worker threads (server ran out of workers)' }
        'resourceWait'    { return 'query memory (memory grant)' }
        'metadatalock'    { return "metadata (definition) of $o" }
        'applicationlock' { return 'an application lock (sp_getapplock)' }
        'databaselock'    { return 'the database' }
        'filelock'        { return 'a database file' }
        'xactlock'        { return 'another transaction' }
        default           { return "$($Res.Type) on $o" }
    }
}

function Get-ShortObj($Res) {
    if ($Res.Object) {
        $p = $Res.Object.Split('.')
        $name = if ($p.Count -ge 2) { $p[-2] + '.' + $p[-1] } else { $Res.Object }
        if ($Res.Index -and $Res.Type -in @('keylock', 'pagelock', 'hobtlock')) { $name += " (index $($Res.Index))" }
        return $name
    }
    return (Get-DlResourceText $Res)
}

function Get-ProcLabel($Procs, [string]$Id) {
    if ($Procs.Contains($Id)) {
        $p = $Procs[$Id]
        $l = 'Session ' + $p.Spid
        if ($p.Ecid -and $p.Ecid -ne '0') { $l += ' (thread ' + $p.Ecid + ')' }
        return $l
    }
    return 'Process ' + $Id
}

function Convert-XeTime([string]$Ts) {
    if (-not $Ts) { return $null }
    try {
        $st = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
        return ([DateTime]::Parse($Ts, [System.Globalization.CultureInfo]::InvariantCulture, $st)).ToLocalTime()
    }
    catch { return $null }
}

function New-DlRecord($Node, $Time, [string]$Source) {
    if (-not $Time) {
        foreach ($p in $Node.SelectNodes('process-list/process')) {
            $t = Coalesce $p.GetAttribute('lastbatchstarted') $p.GetAttribute('lasttranstarted')
            if ($t) { try { $Time = [DateTime]::Parse($t, [System.Globalization.CultureInfo]::InvariantCulture); break } catch { } }
        }
    }
    if (-not $Time) { $Time = Get-Date }
    return [pscustomobject]@{ Key = (Get-TextHash $Node.OuterXml); Time = $Time; Node = $Node; Source = $Source; Analysis = $null }
}

function Add-EventDeadlock($List, $Ev, [string]$Source) {
    $time = Convert-XeTime $Ev.GetAttribute('timestamp')
    $dl = $Ev.SelectSingleNode('.//deadlock')
    if (-not $dl) {
        $v = $Ev.SelectSingleNode('.//data[@name="xml_report"]/value')
        if ($v -and $v.InnerText -match '<deadlock') {
            try { $d2 = New-Object System.Xml.XmlDocument; $d2.LoadXml($v.InnerText); $dl = $d2.SelectSingleNode('//deadlock') } catch { }
        }
    }
    if ($dl) { [void]$List.Add((New-DlRecord $dl $time $Source)) }
}

function Import-DeadlockText([string]$Text, [string]$Source) {
    $list = New-Object System.Collections.ArrayList
    $t = $Text.Trim()
    $t = [regex]::Replace($t, '<\?xml[^>]*\?>', '')
    $t = [regex]::Replace($t, '\sxmlns(:\w+)?="[^"]*"', '')
    $doc = New-Object System.Xml.XmlDocument
    try { $doc.LoadXml($t) }
    catch { $doc = New-Object System.Xml.XmlDocument; $doc.LoadXml('<root>' + $t + '</root>') }
    $events = $doc.SelectNodes('//event[@name="xml_deadlock_report"]')
    if ($events.Count -gt 0) { foreach ($e in $events) { Add-EventDeadlock $list $e $Source } }
    else { foreach ($d in $doc.SelectNodes('//deadlock')) { [void]$list.Add((New-DlRecord $d $null $Source)) } }
    return , $list
}

function Get-ServerDeadlocks($Conn, [string]$Mode) {
    $out = New-Object System.Collections.ArrayList
    if ($Mode -eq 'file') {
        $dt = Invoke-MonitorQuery $Conn $Q_EventFile 120
        foreach ($row in $dt.Rows) {
            $doc = New-Object System.Xml.XmlDocument
            try { $doc.LoadXml([string]$row['EventData']) } catch { continue }
            $ev = $doc.SelectSingleNode('//event')
            if ($ev) { Add-EventDeadlock $out $ev 'system_health (event file)' }
        }
    }
    else {
        $dt = Invoke-MonitorQuery $Conn $Q_RingBuffer 30
        foreach ($row in $dt.Rows) {
            $txt = [string](Get-Val $row 'TargetData' '')
            if (-not $txt) { continue }
            $doc = New-Object System.Xml.XmlDocument
            $doc.LoadXml($txt)
            foreach ($ev in $doc.SelectNodes('//event[@name="xml_deadlock_report"]')) { Add-EventDeadlock $out $ev 'system_health (ring buffer)' }
        }
    }
    return , $out
}

function New-Pattern([string]$Name, [string]$Short, [string]$Plain, [string]$Tech, [string[]]$Fixes) {
    return [pscustomobject]@{ Name = $Name; Short = $Short; Plain = $Plain; Tech = $Tech; Fixes = $Fixes }
}

function Get-DeadlockAnalysis($Dl, $When, [string]$Source) {
    $rep = New-Report
    # ---------- parse ----------
    $victims = New-Object System.Collections.ArrayList
    foreach ($v in $Dl.SelectNodes('victim-list/victimProcess')) { [void]$victims.Add($v.GetAttribute('id')) }
    if ($victims.Count -eq 0 -and $Dl.GetAttribute('victim')) { [void]$victims.Add($Dl.GetAttribute('victim')) }

    $procs = [ordered]@{}
    foreach ($p in $Dl.SelectNodes('process-list/process')) {
        $id = [string]$p.GetAttribute('id')
        $procName = ''; $stmt = ''; $rawStmt = ''
        foreach ($f in $p.SelectNodes('executionStack/frame')) {
            $pn = $f.GetAttribute('procname')
            if (-not $procName -and $pn -and $pn -notin @('adhoc', 'unknown')) { $procName = $pn }
            $ft = Format-SqlText $f.InnerText
            if (-not $stmt -and $ft -and $ft -ne 'unknown') { $stmt = $ft; $rawStmt = $f.InnerText }
        }
        $ib = ''
        $ibn = $p.SelectSingleNode('inputbuf')
        if ($ibn) { $ib = Format-SqlText $ibn.InnerText }
        if (-not $rawStmt) { $rawStmt = $ib }
        $db = $p.GetAttribute('currentdbname')
        if (-not $db) { $db = Get-DbNameById $p.GetAttribute('currentdb') }
        $procs[$id] = [pscustomobject]@{
            Id = $id; Spid = $p.GetAttribute('spid'); Ecid = $p.GetAttribute('ecid')
            Login = $p.GetAttribute('loginname'); Host = $p.GetAttribute('hostname'); App = $p.GetAttribute('clientapp')
            Isolation = $p.GetAttribute('isolationlevel'); LockMode = $p.GetAttribute('lockMode'); WaitRes = $p.GetAttribute('waitresource')
            WaitMs = $p.GetAttribute('waittime'); TranName = $p.GetAttribute('transactionname'); TranCount = $p.GetAttribute('trancount')
            LogUsed = $p.GetAttribute('logused'); Priority = (Coalesce $p.GetAttribute('priority') $p.GetAttribute('taskpriority'))
            LastTranStarted = $p.GetAttribute('lasttranstarted'); LastBatchStarted = $p.GetAttribute('lastbatchstarted')
            Db = $db; Proc = $procName; Statement = $stmt; InputBuf = $ib; Verb = (Get-SqlVerb $rawStmt)
            IsVictim = ($victims -contains $id)
            Holds = [System.Collections.ArrayList]@(); Wants = [System.Collections.ArrayList]@()
        }
    }

    $resources = New-Object System.Collections.ArrayList
    $rl = $Dl.SelectSingleNode('resource-list')
    if ($rl) {
        foreach ($r in $rl.ChildNodes) {
            if ($r.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            $obj = $r.GetAttribute('objectname')
            if (-not $obj -and $r.LocalName -eq 'metadatalock') { $obj = $r.GetAttribute('subresource') }
            $res = [pscustomobject]@{ Type = $r.LocalName; Object = $obj; Index = $r.GetAttribute('indexname'); Friendly = ''; Owners = [System.Collections.ArrayList]@(); Waiters = [System.Collections.ArrayList]@() }
            $res.Friendly = Get-DlResourceText $res
            foreach ($o in $r.SelectNodes('owner-list/owner')) {
                $e = [pscustomobject]@{ Id = [string]$o.GetAttribute('id'); Mode = $o.GetAttribute('mode') }
                [void]$res.Owners.Add($e)
                if ($procs.Contains($e.Id)) { [void]$procs[$e.Id].Holds.Add([pscustomobject]@{ Res = $res; Mode = $e.Mode }) }
            }
            foreach ($w in $r.SelectNodes('waiter-list/waiter')) {
                $e = [pscustomobject]@{ Id = [string]$w.GetAttribute('id'); Mode = $w.GetAttribute('mode') }
                [void]$res.Waiters.Add($e)
                if ($procs.Contains($e.Id)) { [void]$procs[$e.Id].Wants.Add([pscustomobject]@{ Res = $res; Mode = $e.Mode }) }
            }
            [void]$resources.Add($res)
        }
    }

    $plist = @($procs.Values)
    $spids = @($plist | ForEach-Object { $_.Spid } | Select-Object -Unique)
    $objs = @($resources | Where-Object { $_.Object } | ForEach-Object { $_.Object } | Select-Object -Unique)
    $types = @($resources | ForEach-Object { $_.Type } | Select-Object -Unique)
    $modes = @(); foreach ($r in $resources) { $modes += @($r.Owners | ForEach-Object { $_.Mode }); $modes += @($r.Waiters | ForEach-Object { $_.Mode }) }
    $vProc = @($plist | Where-Object { $_.IsVictim })
    $victimText = (@($vProc | ForEach-Object { $_.Spid } | Select-Object -Unique) -join ', ')
    $dbs = @($plist | ForEach-Object { $_.Db } | Where-Object { $_ } | Select-Object -Unique)

    # ---------- pattern detection ----------
    $isParallel = ($types -contains 'exchangeEvent') -or ($spids.Count -eq 1 -and $plist.Count -gt 1)
    $isResource = ($types -contains 'threadpool') -or ($types -contains 'resourceWait')
    $isSchema = ($types -contains 'metadatalock') -or (@($modes | Where-Object { $_ -like 'Sch-*' }).Count -gt 0)
    $isAppLock = $types -contains 'applicationlock'
    $isConversion = $false
    foreach ($r in $resources) {
        $oids = @($r.Owners | ForEach-Object { $_.Id })
        foreach ($w in $r.Waiters) { if ($oids -contains $w.Id) { $isConversion = $true } }
    }
    $rowRes = @($resources | Where-Object { $_.Type -in @('keylock', 'ridlock', 'pagelock') })
    $rowObjs = @($rowRes | ForEach-Object { $_.Object } | Select-Object -Unique)
    $rowIdx = @($rowRes | ForEach-Object { $_.Index } | Select-Object -Unique)
    $hasSelect = @($plist | Where-Object { $_.Verb -eq 'SELECT' }).Count -gt 0
    $isLookup = ($rowObjs.Count -eq 1 -and $rowIdx.Count -ge 2 -and $hasSelect)
    $isRange = (@($modes | Where-Object { $_ -like 'Range*' }).Count -gt 0) -or (@($plist | Where-Object { $_.Isolation -match 'serializable|repeatable' }).Count -gt 0)
    $hasRid = $types -contains 'ridlock'
    $hasCoarse = ($types -contains 'pagelock') -or ($types -contains 'objectlock') -or ($types -contains 'hobtlock')
    $writerWantsRead = @($plist | Where-Object { $_.Verb -in @('UPDATE', 'DELETE', 'INSERT', 'MERGE') -and @($_.Wants | Where-Object { $_.Mode -match '^(S|IS|RangeS-S|RangeS-U)$' }).Count -gt 0 }).Count -gt 0
    $implicit = @($plist | Where-Object { $_.TranName -eq 'implicit_transaction' })
    $procNames = @($plist | ForEach-Object { $_.Proc } | Where-Object { $_ } | Select-Object -Unique)
    $sameProc = ($procNames.Count -eq 1 -and @($plist | Where-Object { $_.Proc }).Count -ge 2 -and -not $isParallel)
    $readersRC = @($plist | Where-Object { ($_.Verb -eq 'SELECT' -or @($_.Wants | Where-Object { $_.Mode -match '^(S|IS)$' }).Count -gt 0) -and $_.Isolation -match 'read committed' })

    $patterns = New-Object System.Collections.ArrayList
    if ($isParallel) {
        [void]$patterns.Add((New-Pattern 'Parallel-query (intra-query) deadlock' 'Parallel query' `
                    'One big query was split into several threads to run faster, and those threads ended up waiting on each other (and sometimes on another session). It is mostly the same query tripping over itself.' `
                    'exchangeEvent resources and/or the same SPID appearing with several ECIDs. Usually a big scan or a poor row estimate that produced a parallel plan.' `
                    @('Look at the execution plan of the statement shown above; add an index so it SEEKS a few rows instead of scanning many.', 'Update statistics on the tables involved (e.g. UPDATE STATISTICS dbo.YourTable WITH FULLSCAN).', 'Targeted workaround: add OPTION (MAXDOP 1) to that query, or raise "cost threshold for parallelism" (e.g. 50) server-wide after testing.', 'Apply the latest Cumulative Update - several parallel-deadlock bugs were fixed in CUs.')))
    }
    if ($isResource) {
        [void]$patterns.Add((New-Pattern 'Resource deadlock (worker threads / memory)' 'Resource (threads/memory)' `
                    'The server ran short of workers or query memory while sessions also held locks, so they could not continue.' `
                    'threadpool / resourceWait resources in the graph. Usually a side-effect of heavy blocking or huge memory grants.' `
                    @('Find and fix the underlying blocking (Blocking tab) - big blocking chains consume worker threads.', 'Look for queries with very large memory grants (sys.dm_exec_query_memory_grants) and fix their estimates/indexes.', 'Review max worker threads and max server memory settings (do not raise blindly).')))
    }
    if ($isSchema) {
        [void]$patterns.Add((New-Pattern 'Schema / metadata deadlock' 'Schema change (DDL)' `
                    'Someone changed the STRUCTURE of a table (index rebuild, ALTER, create index, statistics, partition switch, truncate) while normal work was using it.' `
                    'Sch-M / Sch-S lock modes or metadatalock resources.' `
                    @('Run DDL and index maintenance in a quiet window.', 'Use ONLINE = ON with WAIT_AT_LOW_PRIORITY for index rebuilds where available.', 'Avoid DDL inside busy transactions (ALTER TABLE, CREATE INDEX on permanent tables inside app code).', 'For huge tables consider AUTO_UPDATE_STATISTICS_ASYNC = ON so statistics updates do not run inside user queries.')))
    }
    if ($isAppLock) {
        [void]$patterns.Add((New-Pattern 'Application-lock deadlock' 'Application lock' `
                    'The code uses its own named locks (sp_getapplock) and two sessions requested them in a different order.' `
                    'applicationlock resources in the graph.' `
                    @('Always acquire application locks in the same order everywhere.', 'Use @LockTimeout so callers give up instead of waiting forever, and release locks as early as possible.')))
    }
    if ($isConversion) {
        [void]$patterns.Add((New-Pattern 'Lock-conversion (read-then-update) deadlock' 'Read-then-update' `
                    'Two sessions both READ the same row first (sharing it nicely), then BOTH tried to CHANGE it. Each needs the other to let go of its read before it can write - so they are stuck.' `
                    'A process owns a lock (S/U/IX) on a resource and waits to convert it to a stronger mode (X) on the SAME resource. Classic "check if exists, then insert/update" (upsert) or SELECT-then-UPDATE in one transaction.' `
                    @('Take the right lock at the start: SELECT ... FROM dbo.T WITH (UPDLOCK, ROWLOCK) WHERE ... - then only one session at a time enters the read-then-update section.', 'Better: do it in ONE statement. Upsert pattern: UPDATE dbo.T WITH (UPDLOCK, SERIALIZABLE) SET ... WHERE Key = @k; IF @@ROWCOUNT = 0 INSERT dbo.T (...) VALUES (...);', 'If you use MERGE for upserts, add WITH (HOLDLOCK) and make sure a unique index exists on the join key.', 'Keep the transaction as short as possible.')))
    }
    if ($isLookup) {
        [void]$patterns.Add((New-Pattern 'Key-lookup (bookmark) deadlock' 'Key lookup (index)' `
                    'A reading query found rows through a secondary index and then went to the main table to fetch extra columns. At the same moment an update changed those rows in the opposite order (main table first, then the secondary index). Each held one half and wanted the other.' `
                    ('Same table ({0}), different indexes ({1}). Typically: SELECT holds S on the nonclustered key and wants S on the clustered key, while UPDATE holds X on the clustered key and wants X on the nonclustered key.' -f ($rowObjs -join ', '), ($rowIdx -join ', ')) `
                    @('Make the reader''s index COVERING: add the extra columns it needs as INCLUDE columns so it never visits the main table. Example: CREATE INDEX IX_Name ON dbo.Table (SearchCols) INCLUDE (SelectedCols) WITH (ONLINE = ON);', 'Select only the columns you really need (avoid SELECT *).', 'Or enable READ_COMMITTED_SNAPSHOT so readers stop taking shared locks (test first).')))
    }
    if ($isRange) {
        [void]$patterns.Add((New-Pattern 'Range-lock (Serializable) deadlock' 'Serializable / range locks' `
                    'At least one session used the strictest safety setting (Serializable or Repeatable Read). That setting locks not only the rows it reads but also the "gaps" between them, so far more is locked and collisions are much more likely.' `
                    'Range* lock modes or isolation level serializable(4)/repeatable read(3). Very often caused by .NET TransactionScope (default = Serializable), MERGE ... HOLDLOCK, or an explicit SET TRANSACTION ISOLATION LEVEL SERIALIZABLE.' `
                    @('If the app uses .NET TransactionScope, create it with TransactionOptions { IsolationLevel = ReadCommitted } - the default is Serializable.', 'Use Serializable only where truly required (e.g. a protected upsert), and add an index so the locked range is tiny.', 'On SQL 2014 and older a pooled connection keeps the last isolation level - set it explicitly at the start of each transaction.')))
    }
    if ($objs.Count -ge 2 -and -not $isLookup -and -not $isParallel) {
        [void]$patterns.Add((New-Pattern 'Opposite-order (classic) deadlock' 'Opposite access order' `
                    'Two pieces of work touched the same tables in a DIFFERENT ORDER. Like two people where one walks through door A then door B and the other walks B then A - they meet in the middle, each holding the door the other needs.' `
                    ('Lock cycle across different objects: {0}.' -f ($objs -join ', ')) `
                    @(('Change the code so every transaction touches these tables in the SAME order (for example always: {0}).' -f ($objs -join ' -> ')), 'Make transactions shorter: do reads/calculations before BEGIN TRAN and COMMIT as soon as possible.', 'Add indexes so each statement locks only the rows it needs (fewer locked rows = fewer collisions).', 'If one side only reads, READ_COMMITTED_SNAPSHOT removes reader/writer conflicts.')))
    }
    if ($objs.Count -eq 1 -and -not $isLookup -and -not $isConversion -and -not $isParallel) {
        [void]$patterns.Add((New-Pattern 'Same-table (hot rows) deadlock' 'Same table, different order' `
                    'Two sessions worked on overlapping rows of the same table and picked them up in a different sequence - or locked more rows than needed because they had to scan.' `
                    ('All lock resources are on {0}. Typical causes: missing index (scan locks many rows), multi-row updates in different order, overlapping batches, triggers.' -f $objs[0]) `
                    @('Check each statement''s execution plan: if it scans, add an index on the WHERE/JOIN columns so it seeks.', 'Process rows in a consistent order (e.g. by primary key) and in smaller batches.', 'Use WITH (UPDLOCK) when reading rows you will update later in the same transaction.')))
    }
    if ($patterns.Count -eq 0) {
        [void]$patterns.Add((New-Pattern 'General lock deadlock' 'General' `
                    'Sessions each held a lock the other needed.' 'No single classic pattern matched strongly - review the statements and lock list below.' `
                    @('Keep transactions short and access objects in a consistent order.', 'Index the WHERE/JOIN columns so fewer rows are locked.')))
    }

    # ---------- report ----------
    $whenText = if ($When) { '{0:yyyy-MM-dd HH:mm:ss}' -f $When } else { 'unknown time' }
    Add-R $rep ("DEADLOCK at {0}   ({1})" -f $whenText, $Source) 'h1'
    Add-R $rep ("Most likely cause: {0}" -f $patterns[0].Name) 'bad'
    Add-R $rep ''
    Add-R $rep 'In one sentence' 'h2'
    Add-R $rep ("{0} session(s) each held something the other needed, so none could continue. SQL Server broke the knot by cancelling Session {1} (the 'victim'); the others carried on." -f $spids.Count, (Coalesce $victimText '?'))
    Add-R $rep ''
    Add-R $rep 'Everyday analogy' 'h2'
    Add-R $rep 'Two people each hold one key and need the other person''s key to finish. Neither will give theirs up, so they would wait forever. The building manager (SQL Server) notices after a few seconds and takes the key from one of them - that person has to start over.' 'dim'
    Add-R $rep ''
    Add-R $rep 'What happened, in plain words' 'h2'
    foreach ($p in $plist) {
        $had = (@($p.Holds | ForEach-Object { Get-ShortObj $_.Res } | Select-Object -Unique) -join ', ')
        $need = (@($p.Wants | ForEach-Object { Get-ShortObj $_.Res } | Select-Object -Unique) -join ', ')
        if ($had -or $need) {
            Add-R $rep ("- {0} ({1}) had {2} and needed {3}." -f (Get-ProcLabel $procs $p.Id), (Coalesce $p.Verb 'statement'), (Coalesce $had 'nothing listed'), (Coalesce $need 'nothing listed'))
        }
    }
    if ($vProc.Count -gt 0) {
        $others = @($plist | Where-Object { -not $_.IsVictim })
        $prioDiff = @($plist | ForEach-Object { $_.Priority } | Select-Object -Unique).Count -gt 1
        if ($prioDiff) { $why = 'it had a lower DEADLOCK_PRIORITY than the others' }
        else { $why = ('it was the cheapest to undo (transaction log used: {0} bytes vs {1})' -f (($vProc | ForEach-Object { $_.LogUsed }) -join '/'), (Coalesce (($others | ForEach-Object { $_.LogUsed }) -join '/') 'n/a')) }
        Add-R $rep ("- SQL Server chose Session {0} as the victim because {1}. Its work was rolled back and its application received error 1205 ('Transaction was deadlocked ... Rerun the transaction')." -f $victimText, $why) 'warn'
    }
    Add-R $rep ''
    Add-R $rep 'Session details' 'h2'
    foreach ($p in $plist) {
        $tag = if ($p.IsVictim) { 'VICTIM - cancelled and rolled back' } else { 'survived - completed its work' }
        Add-R $rep ("{0}   [{1}]" -f (Get-ProcLabel $procs $p.Id), $tag) $(if ($p.IsVictim) { 'bad' } else { 'good' })
        Add-R $rep ("   Who        : login '{0}' on '{1}' using '{2}'" -f $p.Login, $p.Host, $p.App)
        Add-R $rep ("   Database   : {0}    Isolation: {1}    Transaction: {2} (count {3}, started {4})" -f $p.Db, $p.Isolation, $p.TranName, $p.TranCount, $p.LastTranStarted)
        if ($p.Proc) { Add-R $rep ("   Code object: {0}" -f $p.Proc) }
        Add-R $rep ("   Statement  : {0}" -f (Coalesce (Format-SqlText $p.Statement 600) '(not captured)')) 'code'
        if ($p.InputBuf -and $p.InputBuf -ne $p.Statement) { Add-R $rep ("   Called as  : {0}" -f (Format-SqlText $p.InputBuf 300)) 'code' }
        foreach ($h in $p.Holds) { Add-R $rep ("   HELD       : {0} on {1}" -f (Get-LockModeText $h.Mode), $h.Res.Friendly) }
        foreach ($w in $p.Wants) {
            $owners = (@($w.Res.Owners | Where-Object { $_.Id -ne $p.Id } | ForEach-Object { Get-ProcLabel $procs $_.Id } | Select-Object -Unique) -join ', ')
            Add-R $rep ("   WANTED     : {0} on {1}  -> blocked by {2}" -f (Get-LockModeText $w.Mode), $w.Res.Friendly, (Coalesce $owners 'its own earlier lock')) 'warn'
        }
        Add-R $rep ''
    }
    Add-R $rep 'Who was waiting on whom (the cycle)' 'h2'
    $edges = [ordered]@{}
    foreach ($r in $resources) {
        foreach ($w in $r.Waiters) {
            foreach ($o in $r.Owners) {
                if ($o.Id -eq $w.Id) { continue }
                $line = '{0} waits for {1}   (wants {2}, other holds {3}, on {4})' -f (Get-ProcLabel $procs $w.Id), (Get-ProcLabel $procs $o.Id), $w.Mode, $o.Mode, (Get-ShortObj $r)
                if (-not $edges.Contains($line)) { $edges[$line] = $true }
            }
        }
    }
    foreach ($e in $edges.Keys) { Add-R $rep ('  ' + $e) 'code' }
    Add-R $rep ''
    Add-R $rep ("Root cause: {0}" -f $patterns[0].Name) 'h2'
    Add-R $rep $patterns[0].Plain
    Add-R $rep ('Technical: ' + $patterns[0].Tech) 'dim'
    if ($patterns.Count -gt 1) {
        Add-R $rep 'Also contributing:' 'b'
        for ($i = 1; $i -lt $patterns.Count; $i++) { Add-R $rep ('* ' + $patterns[$i].Name + ' - ' + $patterns[$i].Plain) }
    }
    Add-R $rep ''
    Add-R $rep 'Recommended fixes (most effective first)' 'h2'
    $k = 0; $seen = @{}
    foreach ($pt in $patterns) { foreach ($f in $pt.Fixes) { if (-not $seen.ContainsKey($f)) { $seen[$f] = $true; $k++; Add-R $rep ("  {0}. {1}" -f $k, $f) } } }

    $notes = New-Object System.Collections.ArrayList
    if ($hasRid) { [void]$notes.Add('A table without a clustered index (heap, RID locks) is involved. Add a clustered index - heaps make updates lock more and behave unpredictably.') }
    if ($hasCoarse) { [void]$notes.Add('Page- or table-level locks were involved. That usually means a statement touched many rows (lock escalation kicks in around 5,000 locks) or used TABLOCK/PAGLOCK hints. Fix with indexes and smaller batches (1,000-4,000 rows per transaction).') }
    if ($writerWantsRead) { [void]$notes.Add('A data-change statement (INSERT/UPDATE/DELETE) needed to READ-lock another table. That is typical of FOREIGN KEY checks, cascading deletes or triggers: make sure foreign-key columns on child tables are indexed, and review triggers on the tables involved.') }
    if ($implicit.Count -gt 0) { [void]$notes.Add(('Session(s) {0} ran in implicit_transaction mode (driver auto-commit OFF or SET IMPLICIT_TRANSACTIONS ON). Transactions then stay open until the app commits - usually far longer than intended. Turn auto-commit on or commit promptly.' -f (($implicit | ForEach-Object { $_.Spid }) -join ', '))) }
    if ($sameProc) { [void]$notes.Add(("Both sides ran the same code ({0}). It deadlocks with ITSELF when called at the same time - fix that procedure (see root cause) rather than the callers." -f $procNames[0])) }
    foreach ($p in $plist) {
        try {
            if ($p.LastTranStarted -and $p.LastBatchStarted) {
                $ts = [DateTime]::Parse($p.LastTranStarted, [System.Globalization.CultureInfo]::InvariantCulture)
                $bs = [DateTime]::Parse($p.LastBatchStarted, [System.Globalization.CultureInfo]::InvariantCulture)
                if (($bs - $ts).TotalSeconds -gt 1) { [void]$notes.Add(('Session {0}: its transaction started {1:0.#}s before the current batch - the application keeps the transaction open across several round trips. Do the whole unit of work in one stored procedure/batch, or commit sooner.' -f $p.Spid, ($bs - $ts).TotalSeconds)) }
            }
        }
        catch { }
    }
    if ($readersRC.Count -gt 0) {
        $rcsiDbs = @($dbs | Where-Object { $script:S.DbInfo.ContainsKey($_) -and $script:S.DbInfo[$_].Rcsi })
        if ($rcsiDbs.Count -gt 0) { [void]$notes.Add(('READ_COMMITTED_SNAPSHOT is already ON for {0}, yet readers took shared locks - look for lock hints (HOLDLOCK, UPDLOCK, REPEATABLEREAD) or stricter isolation in the reading code.' -f ($rcsiDbs -join ', '))) }
        else { [void]$notes.Add(('Readers were part of this deadlock. With READ_COMMITTED_SNAPSHOT ON, ordinary readers stop taking shared locks and this class of deadlock disappears: ALTER DATABASE {0} SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;  (test first - adds tempdb version-store load).' -f (QN (Coalesce ($dbs | Select-Object -First 1) 'YourDb')))) }
    }
    [void]$notes.Add('Add retry logic for error 1205 in the application: retry the WHOLE transaction 2-3 times with a short random delay. Occasional deadlocks are normal in busy systems - the goal is to make them rare and harmless.')
    [void]$notes.Add('If one side is less important (a report, a batch job), put SET DEADLOCK_PRIORITY LOW in it so it is always the one cancelled.')
    [void]$notes.Add('Use "Save as .xdl" and open the file in SSMS to see the graphical deadlock graph.')
    Add-R $rep ''
    Add-R $rep 'Also worth checking' 'h2'
    foreach ($nt in $notes) { Add-R $rep ('- ' + $nt) }
    Add-R $rep ''
    Add-R $rep 'Business impact' 'h2'
    Add-R $rep ("The victim's work (Session {0}) was undone. If the application retries automatically, users likely noticed only a short delay; if not, a user or job saw an error and that action did NOT complete." -f (Coalesce $victimText '?'))
    Add-R $rep ''
    Add-R $rep 'Resources involved (technical)' 'h2'
    foreach ($r in $resources) {
        $own = (@($r.Owners | ForEach-Object { '{0} [{1}]' -f (Get-ProcLabel $procs $_.Id), $_.Mode }) -join ', ')
        $wai = (@($r.Waiters | ForEach-Object { '{0} [{1}]' -f (Get-ProcLabel $procs $_.Id), $_.Mode }) -join ', ')
        Add-R $rep ("  {0}: {1}" -f $r.Type, $r.Friendly) 'code'
        Add-R $rep ("      held by: {0}   |   wanted by: {1}" -f (Coalesce $own '-'), (Coalesce $wai '-')) 'code'
    }

    # ---------- AI prompt ----------
    $xml = $Dl.OuterXml
    if ($xml.Length -gt 60000) { $xml = $xml.Substring(0, 60000) + "`r`n<!-- truncated -->" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('You are a senior SQL Server DBA. Analyze the SQL Server deadlock below.')
    [void]$sb.AppendLine('Please provide:')
    [void]$sb.AppendLine('1. A plain-English explanation a non-technical manager can understand (what happened and the business impact).')
    [void]$sb.AppendLine('2. The exact technical root cause: lock modes, resources, access order, isolation levels, indexes, and the application pattern that most likely produced it.')
    [void]$sb.AppendLine('3. Concrete fixes ranked by impact and risk: index changes (with CREATE INDEX scripts), T-SQL / stored procedure changes (show rewritten code), transaction scope / isolation changes, and application retry logic for error 1205.')
    [void]$sb.AppendLine('4. How to verify the fix and how to monitor for recurrence.')
    [void]$sb.AppendLine('Ask me for table/index definitions or execution plans if you need them to be certain.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(('Server: {0}  Version: {1}  Deadlock time: {2}  Source: {3}' -f $script:S.ServerName, $script:S.ServerVersion, $whenText, $Source))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Summary extracted by my monitoring tool ---')
    [void]$sb.AppendLine((Get-ReportText $rep))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Raw deadlock graph XML ---')
    [void]$sb.AppendLine($xml)

    return [pscustomobject]@{
        Report      = $rep
        Prompt      = $sb.ToString()
        Short       = $patterns[0].Short
        VictimText  = $victimText
        SessionText = ($spids -join ', ')
        ObjectText  = (@($objs | ForEach-Object { $p = $_.Split('.'); if ($p.Count -ge 2) { $p[-2] + '.' + $p[-1] } else { $_ } }) -join ', ')
    }
}
#endregion

#region ===================== UI helpers =====================
$F = @{
    Normal = New-Object System.Drawing.Font('Segoe UI', 9.5)
    Bold   = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    H1     = New-Object System.Drawing.Font('Segoe UI', 12.5, [System.Drawing.FontStyle]::Bold)
    H2     = New-Object System.Drawing.Font('Segoe UI', 10.5, [System.Drawing.FontStyle]::Bold)
    Code   = New-Object System.Drawing.Font('Consolas', 9.5)
}

function Show-Report([System.Windows.Forms.RichTextBox]$Box, $Rep) {
    $Box.SuspendLayout()
    $Box.Clear()
    foreach ($l in $Rep) {
        $font = $F.Normal; $color = [System.Drawing.Color]::Black
        switch ($l.S) {
            'h1'   { $font = $F.H1; $color = [System.Drawing.Color]::FromArgb(20, 60, 120) }
            'h2'   { $font = $F.H2; $color = [System.Drawing.Color]::FromArgb(30, 90, 150) }
            'b'    { $font = $F.Bold }
            'bad'  { $font = $F.Bold; $color = [System.Drawing.Color]::Firebrick }
            'warn' { $color = [System.Drawing.Color]::FromArgb(170, 85, 0) }
            'good' { $font = $F.Bold; $color = [System.Drawing.Color]::ForestGreen }
            'code' { $font = $F.Code; $color = [System.Drawing.Color]::FromArgb(60, 60, 60) }
            'dim'  { $color = [System.Drawing.Color]::DimGray }
        }
        $Box.SelectionStart = $Box.TextLength
        $Box.SelectionLength = 0
        $Box.SelectionFont = $font
        $Box.SelectionColor = $color
        $Box.AppendText($l.T + "`n")
    }
    $Box.SelectionStart = 0
    $Box.ScrollToCaret()
    $Box.ResumeLayout()
}

function Write-Log([string]$Msg) {
    $txtLog.AppendText(('[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $Msg) + "`r`n")
}

function Set-Status([string]$Text, [string]$Kind = 'n') {
    $lblStatus.Text = $Text
    switch ($Kind) {
        'bad'  { $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick }
        'warn' { $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(170, 85, 0) }
        'good' { $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen }
        default { $lblStatus.ForeColor = [System.Drawing.Color]::Black }
    }
}

function Set-Clip([string]$Text) {
    if (-not $Text) { Set-Status 'Nothing to copy yet.' 'warn'; return }
    try { [System.Windows.Forms.Clipboard]::SetText($Text); Set-Status 'Copied to clipboard.' 'good' }
    catch { Set-Status ('Could not copy: ' + $_.Exception.Message) 'bad' }
}

function Show-TextDialog([string]$Title, [string]$Text = '', [switch]$Editable, [string]$OkText = 'Close') {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(950, 650)
    $dlg.StartPosition = 'CenterParent'
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true; $tb.ScrollBars = 'Both'; $tb.WordWrap = $false; $tb.Dock = 'Fill'
    $tb.Font = $F.Code; $tb.MaxLength = 0; $tb.Text = $Text; $tb.ReadOnly = -not $Editable
    $pnl = New-Object System.Windows.Forms.FlowLayoutPanel
    $pnl.Dock = 'Bottom'; $pnl.Height = 44; $pnl.FlowDirection = 'RightToLeft'; $pnl.Padding = New-Object System.Windows.Forms.Padding(6)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = $OkText; $ok.Width = 130; $ok.Height = 28; $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pnl.Controls.Add($ok)
    if ($Editable) {
        $cn = New-Object System.Windows.Forms.Button
        $cn.Text = 'Cancel'; $cn.Width = 100; $cn.Height = 28; $cn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $pnl.Controls.Add($cn); $dlg.CancelButton = $cn
    }
    $dlg.Controls.Add($tb); $dlg.Controls.Add($pnl); $tb.BringToFront()
    $r = $dlg.ShowDialog($form)
    $t = $tb.Text
    $dlg.Dispose()
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) { return $t }
    return $null
}

function New-Grid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.SelectionMode = 'FullRowSelect'
    $g.MultiSelect = $false
    $g.RowHeadersVisible = $false
    $g.AutoSizeColumnsMode = 'None'
    $g.BackgroundColor = [System.Drawing.Color]::White
    $g.BorderStyle = 'None'
    $g.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 252)
    $g.ColumnHeadersDefaultCellStyle.Font = $F.Bold
    $g.Font = $F.Normal
    return $g
}

function New-Btn([string]$Text, [int]$W = 120) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Width = $W; $b.Height = 28
    $b.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    return $b
}

function New-Rtb {
    $r = New-Object System.Windows.Forms.RichTextBox
    $r.Dock = 'Fill'; $r.ReadOnly = $true; $r.BackColor = [System.Drawing.Color]::White
    $r.Font = $F.Normal; $r.BorderStyle = 'None'; $r.DetectUrls = $false
    return $r
}

function Get-IntervalSeconds { return [int]$script:Intervals[$cboInterval.SelectedItem] }
function Get-StopAfterMinutes { return [int]$script:StopAfter[$cboStopAfter.SelectedItem] }

function Get-InnerSqlException($Err) {
    $e = $Err
    while ($e) { if ($e -is [System.Data.SqlClient.SqlException]) { return $e }; $e = $e.InnerException }
    return $null
}
#endregion

#region ===================== Main form =====================
$script:Intervals = [ordered]@{ '5 sec' = 5; '10 sec' = 10; '15 sec' = 15; '30 sec' = 30; '1 min' = 60; '2 min' = 120; '5 min' = 300; '10 min' = 600 }
$script:StopAfter = [ordered]@{ 'Never' = 0; '15 min' = 15; '30 min' = 30; '1 hour' = 60; '4 hours' = 240; '8 hours' = 480 }

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SQL Server Blocking & Deadlock Monitor'
$form.Size = New-Object System.Drawing.Size(1350, 900)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 650)
$form.StartPosition = 'CenterScreen'
$form.Font = $F.Normal
$form.AutoScaleMode = 'Dpi'

$tip = New-Object System.Windows.Forms.ToolTip

# ---------- top connection panel ----------
$top = New-Object System.Windows.Forms.Panel
$top.Dock = 'Top'; $top.Height = 108
$top.BackColor = [System.Drawing.Color]::FromArgb(240, 244, 250)

function Add-Lbl([string]$t, [int]$x, [int]$y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.AutoSize = $true; $l.Location = New-Object System.Drawing.Point($x, ($y + 4))
    $top.Controls.Add($l)
}

Add-Lbl 'Server:' 10 10
$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(70, 10); $txtServer.Width = 210; $txtServer.Text = $env:COMPUTERNAME
$tip.SetToolTip($txtServer, 'Server or Server\Instance or Server,Port')
$top.Controls.Add($txtServer)

Add-Lbl 'Auth:' 292 10
$cboAuth = New-Object System.Windows.Forms.ComboBox
$cboAuth.Location = New-Object System.Drawing.Point(335, 10); $cboAuth.Width = 160; $cboAuth.DropDownStyle = 'DropDownList'
[void]$cboAuth.Items.Add('Windows authentication'); [void]$cboAuth.Items.Add('SQL login'); $cboAuth.SelectedIndex = 0
$top.Controls.Add($cboAuth)

Add-Lbl 'User:' 505 10
$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(545, 10); $txtUser.Width = 130; $txtUser.Enabled = $false
$top.Controls.Add($txtUser)

Add-Lbl 'Password:' 685 10
$txtPwd = New-Object System.Windows.Forms.TextBox
$txtPwd.Location = New-Object System.Drawing.Point(755, 10); $txtPwd.Width = 130; $txtPwd.UseSystemPasswordChar = $true; $txtPwd.Enabled = $false
$top.Controls.Add($txtPwd)

$chkEncrypt = New-Object System.Windows.Forms.CheckBox
$chkEncrypt.Text = 'Encrypt'; $chkEncrypt.AutoSize = $true; $chkEncrypt.Location = New-Object System.Drawing.Point(898, 12)
$top.Controls.Add($chkEncrypt)
$chkTrust = New-Object System.Windows.Forms.CheckBox
$chkTrust.Text = 'Trust server certificate'; $chkTrust.AutoSize = $true; $chkTrust.Location = New-Object System.Drawing.Point(975, 12)
$top.Controls.Add($chkTrust)

Add-Lbl 'Refresh:' 10 44
$cboInterval = New-Object System.Windows.Forms.ComboBox
$cboInterval.Location = New-Object System.Drawing.Point(70, 44); $cboInterval.Width = 90; $cboInterval.DropDownStyle = 'DropDownList'
foreach ($k in $script:Intervals.Keys) { [void]$cboInterval.Items.Add($k) }
$cboInterval.SelectedItem = '10 sec'
$tip.SetToolTip($cboInterval, 'Wait time between cycles. Each cycle connects, runs the check, then disconnects.')
$top.Controls.Add($cboInterval)

Add-Lbl 'Stop after:' 172 44
$cboStopAfter = New-Object System.Windows.Forms.ComboBox
$cboStopAfter.Location = New-Object System.Drawing.Point(245, 44); $cboStopAfter.Width = 90; $cboStopAfter.DropDownStyle = 'DropDownList'
foreach ($k in $script:StopAfter.Keys) { [void]$cboStopAfter.Items.Add($k) }
$cboStopAfter.SelectedItem = '1 hour'
$tip.SetToolTip($cboStopAfter, 'Safety limit: monitoring stops automatically after this long.')
$top.Controls.Add($cboStopAfter)

$chkDlEach = New-Object System.Windows.Forms.CheckBox
$chkDlEach.Text = 'Also check for new deadlocks each refresh'; $chkDlEach.AutoSize = $true; $chkDlEach.Checked = $true
$chkDlEach.Location = New-Object System.Drawing.Point(350, 47)
$tip.SetToolTip($chkDlEach, 'Reads the system_health ring buffer (in memory, no locks) on each cycle.')
$top.Controls.Add($chkDlEach)

$btnTest = New-Btn 'Test connection' 125; $btnTest.Location = New-Object System.Drawing.Point(640, 42); $top.Controls.Add($btnTest)
$btnStart = New-Btn 'Start' 90; $btnStart.Location = New-Object System.Drawing.Point(772, 42); $top.Controls.Add($btnStart)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(220, 240, 220)
$btnStop = New-Btn 'Stop' 90; $btnStop.Location = New-Object System.Drawing.Point(867, 42); $btnStop.Enabled = $false; $top.Controls.Add($btnStop)
$btnNow = New-Btn 'Refresh now' 110; $btnNow.Location = New-Object System.Drawing.Point(962, 42); $top.Controls.Add($btnNow)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 80); $lblStatus.AutoSize = $true
$lblStatus.Text = 'Ready. Enter a server and press Start (or Refresh now for a single check).'
$top.Controls.Add($lblStatus)

$cboAuth.add_SelectedIndexChanged({
        $sql = ($cboAuth.SelectedIndex -eq 1)
        $txtUser.Enabled = $sql; $txtPwd.Enabled = $sql
    })

# ---------- tabs ----------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabBlk = New-Object System.Windows.Forms.TabPage; $tabBlk.Text = 'Blocking'
$tabDl = New-Object System.Windows.Forms.TabPage; $tabDl.Text = 'Deadlocks'
$tabAi = New-Object System.Windows.Forms.TabPage; $tabAi.Text = 'AI Prompt'
$tabHist = New-Object System.Windows.Forms.TabPage; $tabHist.Text = 'History & Log'
$tabs.TabPages.AddRange(@($tabBlk, $tabDl, $tabAi, $tabHist))

# ===== Blocking tab =====
$lblBanner = New-Object System.Windows.Forms.Label
$lblBanner.Dock = 'Top'; $lblBanner.Height = 34; $lblBanner.TextAlign = 'MiddleLeft'
$lblBanner.Font = $F.H2; $lblBanner.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$lblBanner.Text = 'Not checked yet.'; $lblBanner.BackColor = [System.Drawing.Color]::Gainsboro

$splitBlk = New-Object System.Windows.Forms.SplitContainer
$splitBlk.Dock = 'Fill'; $splitBlk.Orientation = 'Horizontal'

$blkTable = New-Object System.Data.DataTable
foreach ($c in @(
        @('Role', [string]), @('Session', [int]), @('BlockedBy', [int]), @('Waiting behind it', [int]), @('Wait (s)', [double]),
        @('Wait type', [string]), @('Waiting for', [string]), @('Status', [string]), @('Open tran', [int]), @('Tran age (s)', [int]),
        @('Idle (s)', [int]), @('Database', [string]), @('Login', [string]), @('Host', [string]), @('Program', [string]),
        @('Command', [string]), @('Isolation', [string]), @('Object', [string]), @('Statement', [string]))) {
    [void]$blkTable.Columns.Add($c[0], $c[1])
}
$gridBlk = New-Grid
$gridBlk.DataSource = $blkTable
$splitBlk.Panel1.Controls.Add($gridBlk)

$barBlk = New-Object System.Windows.Forms.FlowLayoutPanel
$barBlk.Dock = 'Top'; $barBlk.Height = 38
$btnCopyBlk = New-Btn 'Copy analysis' 120
$btnBlkPrompt = New-Btn 'Copy AI prompt' 130
$btnExportBlk = New-Btn 'Export snapshot (CSV)' 160
$barBlk.Controls.AddRange(@($btnCopyBlk, $btnBlkPrompt, $btnExportBlk))
$rtbBlk = New-Rtb
$splitBlk.Panel2.Controls.Add($rtbBlk)
$splitBlk.Panel2.Controls.Add($barBlk)
$rtbBlk.BringToFront()

$tabBlk.Controls.Add($splitBlk)
$tabBlk.Controls.Add($lblBanner)
$splitBlk.BringToFront()

$gridBlk.add_CellFormatting({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $role = [string]$sender.Rows[$e.RowIndex].Cells['Role'].Value
        if ($role -eq 'HEAD BLOCKER') { $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 215, 215); $e.CellStyle.Font = $F.Bold }
        elseif ($role -eq 'Blocked + blocking') { $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 232, 200) }
        elseif ($role -eq 'Blocked') { $e.CellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 250, 215) }
    })

$gridBlk.add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $sid = [int]$sender.Rows[$e.RowIndex].Cells['Session'].Value
        $s = @($script:S.Sessions | Where-Object { $_.SessionId -eq $sid }) | Select-Object -First 1
        if (-not $s) { return }
        $txt = ($s | Format-List SessionId, BlockedBy, Status, WaitType, WaitSec, WaitResource, WaitingFor, DatabaseName, LoginName, HostName, ProgramName, Command, OpenTran, TranAgeSec, IdleSec, IsolationLevel, CpuMs, LogicalReads, ProcName | Out-String).Trim()
        $txt += "`r`n`r`nStatement:`r`n" + $s.SqlStatement
        [void](Show-TextDialog ("Session {0} details" -f $sid) $txt)
    })

# ===== Deadlocks tab =====
$barDl = New-Object System.Windows.Forms.FlowLayoutPanel
$barDl.Dock = 'Top'; $barDl.Height = 38
$lblSrc = New-Object System.Windows.Forms.Label; $lblSrc.Text = 'Source:'; $lblSrc.AutoSize = $true; $lblSrc.Margin = New-Object System.Windows.Forms.Padding(6, 10, 0, 0)
$cboDlSource = New-Object System.Windows.Forms.ComboBox
$cboDlSource.DropDownStyle = 'DropDownList'; $cboDlSource.Width = 230; $cboDlSource.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
[void]$cboDlSource.Items.Add('system_health ring buffer (fast)'); [void]$cboDlSource.Items.Add('system_health event files (complete, slower)')
$cboDlSource.SelectedIndex = 0
$tip.SetToolTip($cboDlSource, 'Ring buffer = in-memory, recent deadlocks. Event files = everything still kept in the system_health .xel files (can take several seconds).')
$btnDlFetch = New-Btn 'Fetch from server' 135
$btnDlOpen = New-Btn 'Open .xdl / .xml' 125
$btnDlPaste = New-Btn 'Paste XML' 95
$btnDlSave = New-Btn 'Save as .xdl' 105
$btnDlCopy = New-Btn 'Copy analysis' 115
$btnDlPrompt = New-Btn 'Copy AI prompt' 125
$btnDlClear = New-Btn 'Clear list' 90
$barDl.Controls.AddRange(@($lblSrc, $cboDlSource, $btnDlFetch, $btnDlOpen, $btnDlPaste, $btnDlSave, $btnDlCopy, $btnDlPrompt, $btnDlClear))

$splitDl = New-Object System.Windows.Forms.SplitContainer
$splitDl.Dock = 'Fill'; $splitDl.Orientation = 'Horizontal'
$dlTable = New-Object System.Data.DataTable
foreach ($c in @(@('Time', [datetime]), @('Victim', [string]), @('Sessions', [string]), @('Objects', [string]), @('Likely cause', [string]), @('Source', [string]), @('Key', [string]))) {
    [void]$dlTable.Columns.Add($c[0], $c[1])
}
$dlView = New-Object System.Data.DataView($dlTable)
$dlView.Sort = 'Time DESC'
$gridDl = New-Grid
$gridDl.DataSource = $dlView
$splitDl.Panel1.Controls.Add($gridDl)
$rtbDl = New-Rtb
$splitDl.Panel2.Controls.Add($rtbDl)
$tabDl.Controls.Add($splitDl)
$tabDl.Controls.Add($barDl)
$splitDl.BringToFront()

$gridDl.add_DataBindingComplete({
        if ($gridDl.Columns.Contains('Key')) { $gridDl.Columns['Key'].Visible = $false }
        if ($gridDl.Columns.Contains('Time')) { $gridDl.Columns['Time'].DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm:ss'; $gridDl.Columns['Time'].Width = 150 }
        foreach ($n in @('Objects', 'Likely cause')) { if ($gridDl.Columns.Contains($n)) { $gridDl.Columns[$n].Width = 260 } }
        if ($gridDl.Columns.Contains('Source')) { $gridDl.Columns['Source'].Width = 200 }
    })

function Show-SelectedDeadlock {
    $row = $gridDl.CurrentRow
    if (-not $row -or -not $row.DataBoundItem) { return }
    $key = [string]$row.DataBoundItem.Row['Key']
    if (-not $script:S.Deadlocks.ContainsKey($key)) { return }
    $script:S.CurrentDlKey = $key
    Show-Report $rtbDl $script:S.Deadlocks[$key].Analysis.Report
}
$gridDl.add_SelectionChanged({ Show-SelectedDeadlock })

function Select-DeadlockRow([string]$Key) {
    foreach ($r in $gridDl.Rows) {
        if ($r.DataBoundItem -and [string]$r.DataBoundItem.Row['Key'] -eq $Key) {
            $gridDl.CurrentCell = $r.Cells['Time']
            Show-SelectedDeadlock
            return
        }
    }
}

function Add-Deadlocks($Records, [bool]$AlertNew) {
    $new = New-Object System.Collections.ArrayList
    foreach ($r in $Records) {
        if ($script:S.Deadlocks.ContainsKey($r.Key)) { continue }
        try { $r.Analysis = Get-DeadlockAnalysis $r.Node $r.Time $r.Source }
        catch { Write-Log ('Could not analyze a deadlock graph: ' + $_.Exception.Message); continue }
        $script:S.Deadlocks[$r.Key] = $r
        [void]$dlTable.Rows.Add($r.Time, $r.Analysis.VictimText, $r.Analysis.SessionText, $r.Analysis.ObjectText, $r.Analysis.Short, $r.Source, $r.Key)
        [void]$new.Add($r)
    }
    if ($AlertNew -and $new.Count -gt 0) {
        $script:S.NewDeadlocks += $new.Count
        $tabDl.Text = 'Deadlocks ({0} NEW)' -f $script:S.NewDeadlocks
        foreach ($r in $new) { Write-Log ('NEW DEADLOCK at {0:yyyy-MM-dd HH:mm:ss}: victim session {1}, objects {2}, likely cause: {3}' -f $r.Time, $r.Analysis.VictimText, $r.Analysis.ObjectText, $r.Analysis.Short) }
        try { [System.Media.SystemSounds]::Exclamation.Play() } catch { }
    }
    return , $new
}

# ===== AI prompt tab =====
$barAi = New-Object System.Windows.Forms.FlowLayoutPanel
$barAi.Dock = 'Top'; $barAi.Height = 38
$btnAiBlk = New-Btn 'Load blocking prompt' 160
$btnAiDl = New-Btn 'Load selected deadlock prompt' 210
$btnAiCopy = New-Btn 'Copy to clipboard' 140
$lblAiNote = New-Object System.Windows.Forms.Label
$lblAiNote.AutoSize = $true; $lblAiNote.Margin = New-Object System.Windows.Forms.Padding(10, 10, 0, 0)
$lblAiNote.ForeColor = [System.Drawing.Color]::FromArgb(170, 85, 0)
$lblAiNote.Text = 'Review before pasting into an external AI: it contains server, login, host names and SQL text. You can edit it here.'
$barAi.Controls.AddRange(@($btnAiBlk, $btnAiDl, $btnAiCopy, $lblAiNote))
$txtPrompt = New-Object System.Windows.Forms.TextBox
$txtPrompt.Multiline = $true; $txtPrompt.ScrollBars = 'Both'; $txtPrompt.WordWrap = $true; $txtPrompt.Dock = 'Fill'
$txtPrompt.Font = $F.Code; $txtPrompt.MaxLength = 0
$tabAi.Controls.Add($txtPrompt); $tabAi.Controls.Add($barAi); $txtPrompt.BringToFront()

# ===== History tab =====
$splitHist = New-Object System.Windows.Forms.SplitContainer
$splitHist.Dock = 'Fill'; $splitHist.Orientation = 'Horizontal'
$histTable = New-Object System.Data.DataTable
foreach ($c in @(@('Time', [datetime]), @('Blocked sessions', [int]), @('Longest wait (s)', [double]), @('Head blocker(s)', [string]), @('Severity', [string]), @('Main cause', [string]))) {
    [void]$histTable.Columns.Add($c[0], $c[1])
}
$histView = New-Object System.Data.DataView($histTable); $histView.Sort = 'Time DESC'
$gridHist = New-Grid
$gridHist.DataSource = $histView
$gridHist.add_DataBindingComplete({
        if ($gridHist.Columns.Contains('Time')) { $gridHist.Columns['Time'].DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm:ss'; $gridHist.Columns['Time'].Width = 150 }
        if ($gridHist.Columns.Contains('Main cause')) { $gridHist.Columns['Main cause'].Width = 380 }
        if ($gridHist.Columns.Contains('Head blocker(s)')) { $gridHist.Columns['Head blocker(s)'].Width = 300 }
    })
$splitHist.Panel1.Controls.Add($gridHist)
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.ReadOnly = $true; $txtLog.Dock = 'Fill'
$txtLog.Font = $F.Code; $txtLog.BackColor = [System.Drawing.Color]::White; $txtLog.MaxLength = 0
$splitHist.Panel2.Controls.Add($txtLog)
$tabHist.Controls.Add($splitHist)

$form.Controls.Add($tabs)
$form.Controls.Add($top)
$tabs.BringToFront()

$timer = New-Object System.Windows.Forms.Timer
#endregion

#region ===================== Refresh cycle =====================
function Update-BlockingView($Sessions, [long]$Ms) {
    $now = Get-Date
    $script:S.Sessions = $Sessions
    $a = Get-BlockingAnalysis $Sessions
    $blkTable.Clear()
    foreach ($s in @($Sessions | Sort-Object { if ($a.Roles[$_.SessionId] -eq 'HEAD BLOCKER') { 0 } else { 1 } }, BlockedBy, SessionId)) {
        [void]$blkTable.Rows.Add($a.Roles[$s.SessionId], $s.SessionId, $s.BlockedBy, $a.Counts[$s.SessionId], $s.WaitSec, $s.WaitType,
            (Coalesce $s.WaitingFor $s.WaitResource), $s.Status, $s.OpenTran, $s.TranAgeSec, $s.IdleSec, $s.DatabaseName, $s.LoginName,
            $s.HostName, $s.ProgramName, $s.Command, $s.IsolationLevel, $s.ProcName, (Format-SqlText $s.SqlStatement 300))
    }
    if ($gridBlk.Columns.Contains('Statement')) { $gridBlk.Columns['Statement'].Width = 500 }
    if ($gridBlk.Columns.Contains('Waiting for')) { $gridBlk.Columns['Waiting for'].Width = 320 }
    $where = Coalesce $script:S.ServerName $txtServer.Text

    if ($a.Blocked -gt 0) {
        $script:S.LastBlockingRep = $a.Report
        $script:S.LastBlockPrompt = $a.Prompt
        Show-Report $rtbBlk $a.Report
        $lblBanner.Text = ('  BLOCKING: {0} session(s) waiting, longest {1}, head blocker(s): {2}   [{3} severity]   checked {4:HH:mm:ss}' -f $a.Blocked, (Format-Dur $a.MaxWait), (Coalesce ($a.HeadIds -join ', ') 'circular'), $a.Severity, $now)
        $lblBanner.BackColor = if ($a.Severity -eq 'High') { [System.Drawing.Color]::FromArgb(230, 90, 90) } elseif ($a.Severity -eq 'Medium') { [System.Drawing.Color]::FromArgb(245, 170, 80) } else { [System.Drawing.Color]::FromArgb(250, 220, 120) }
        $lblBanner.ForeColor = if ($a.Severity -eq 'High') { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
        [void]$histTable.Rows.Add($now, $a.Blocked, $a.MaxWait, ($a.HeadIds -join ', '), $a.Severity, $a.MainCause)
        while ($histTable.Rows.Count -gt 1000) { $histTable.Rows.RemoveAt(0) }
        if (-not $script:S.Episode) {
            $script:S.Episode = @{ Start = $now; PeakBlocked = 0; PeakWait = 0.0; Heads = @{}; Cause = $a.MainCause }
            Write-Log ('Blocking STARTED: {0} blocked, head blocker(s) {1}. {2}' -f $a.Blocked, ($a.HeadIds -join ', '), $a.MainCause)
        }
        $ep = $script:S.Episode
        if ($a.Blocked -gt $ep.PeakBlocked) { $ep.PeakBlocked = $a.Blocked }
        if ($a.MaxWait -gt $ep.PeakWait) { $ep.PeakWait = $a.MaxWait }
        foreach ($h in $a.HeadIds) { $ep.Heads["$h"] = $true }
        if ($a.MainCause) { $ep.Cause = $a.MainCause }
    }
    else {
        if ($script:S.Episode) {
            $ep = $script:S.Episode
            $script:S.LastEpisodeText = ('From {0:HH:mm:ss} to about {1:HH:mm:ss} (~{2}). Peak: {3} session(s) blocked, longest wait {4}. Head blocker(s): {5}. Main cause: {6}.' -f $ep.Start, $now, (Format-Dur ($now - $ep.Start).TotalSeconds), $ep.PeakBlocked, (Format-Dur $ep.PeakWait), (@($ep.Heads.Keys) -join ', '), (Coalesce $ep.Cause 'n/a'))
            Write-Log ('Blocking CLEARED. ' + $script:S.LastEpisodeText)
            $script:S.Episode = $null
        }
        $rep = New-Report
        Add-R $rep 'No blocking right now' 'good'
        Add-R $rep ('Checked {0} at {1:HH:mm:ss} - the check took {2} ms and the connection was closed afterwards.' -f $where, $now, $Ms)
        if ($script:S.LastEpisodeText) {
            Add-R $rep ''
            Add-R $rep 'Last blocking episode' 'h2'
            Add-R $rep $script:S.LastEpisodeText
            Add-R $rep 'Its full analysis is still available: use "Copy analysis" / "Copy AI prompt", and see History & Log.' 'dim'
        }
        Add-R $rep ''
        Add-R $rep 'Short blocking (a second or two) is normal. This screen fills in automatically when sessions are waiting on each other.' 'dim'
        Show-Report $rtbBlk $rep
        $lblBanner.Text = ('  No blocking  -  {0}  checked {1:HH:mm:ss} ({2} ms)' -f $where, $now, $Ms)
        $lblBanner.BackColor = [System.Drawing.Color]::FromArgb(200, 235, 200)
        $lblBanner.ForeColor = [System.Drawing.Color]::Black
    }
}

function Invoke-RefreshCycle {
    if ($script:S.Busy) { return }
    if (-not $txtServer.Text.Trim()) { Set-Status 'Enter a server name first.' 'warn'; return }
    $script:S.Busy = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $conn = $null; $sessions = $null; $dlRecords = $null; $err = $null; $dlErr = $null
    try {
        Set-Status ('Connecting to {0} ...' -f $txtServer.Text.Trim())
        $lblStatus.Refresh()
        $conn = Open-MonitorConnection
        Update-DbInfo $conn
        $dt = Invoke-MonitorQuery $conn $Q_Blocking 15
        $sessions = Convert-BlockingRows $dt
        Resolve-SessionNames $conn $sessions
        if ($chkDlEach.Checked) {
            try { $dlRecords = Get-ServerDeadlocks $conn 'ring' } catch { $dlErr = $_.Exception.Message }
        }
    }
    catch { $err = $_.Exception }
    finally {
        if ($conn) { try { $conn.Close() } catch { }; $conn.Dispose() }   # always disconnect
        $sw.Stop()
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:S.Busy = $false
    }
    $script:S.Cycles++

    if ($err) {
        $script:S.Failures++
        $sqlEx = Get-InnerSqlException $err
        $msg = $err.Message
        if ($sqlEx) {
            switch ($sqlEx.Number) {
                1222 { $msg = 'The monitoring query hit the 500 ms lock timeout and backed off (by design it never waits on locks). Will try again next cycle.' }
                { $_ -in 297, 300, 229 } { $msg = 'Permission denied - the login needs VIEW SERVER STATE (SQL 2022+: VIEW SERVER PERFORMANCE STATE). ' + $sqlEx.Message }
                18456 { $msg = 'Login failed - check user name / password / authentication type. ' + $sqlEx.Message }
                default { $msg = ('SQL error {0}: {1}' -f $sqlEx.Number, $sqlEx.Message) }
            }
        }
        Set-Status ('{0:HH:mm:ss}  {1}' -f (Get-Date), $msg) 'bad'
        Write-Log $msg
        if ($script:S.Running -and $script:S.Failures -ge 5) { Stop-Monitoring 'Stopped after 5 consecutive failures - fix the problem above and press Start again.' }
        return
    }
    $script:S.Failures = 0
    Update-BlockingView $sessions $sw.ElapsedMilliseconds
    if ($dlErr) { Write-Log ('Deadlock check skipped: ' + $dlErr) }
    if ($dlRecords) {
        [void](Add-Deadlocks $dlRecords $script:S.DlBaselineDone)
        $script:S.DlBaselineDone = $true
    }
    $next = ''
    if ($script:S.Running) { $next = ('   Next check at {0:HH:mm:ss}.' -f (Get-Date).AddSeconds((Get-IntervalSeconds))) }
    $kind = if ($sessions.Count -gt 0) { 'warn' } else { 'good' }
    Set-Status ('{0:HH:mm:ss}  Cycle {1} done in {2} ms (connected, checked, disconnected).{3}' -f (Get-Date), $script:S.Cycles, $sw.ElapsedMilliseconds, $next) $kind
}

function Set-RunningUi([bool]$On) {
    $btnStart.Enabled = -not $On
    $btnStop.Enabled = $On
    foreach ($c in @($txtServer, $cboAuth, $chkEncrypt, $chkTrust, $btnTest)) { $c.Enabled = -not $On }
    $txtUser.Enabled = (-not $On) -and ($cboAuth.SelectedIndex -eq 1)
    $txtPwd.Enabled = $txtUser.Enabled
    $form.Text = if ($On) { 'SQL Server Blocking & Deadlock Monitor  -  MONITORING ' + $txtServer.Text.Trim() } else { 'SQL Server Blocking & Deadlock Monitor' }
}

function Start-Monitoring {
    if (-not $txtServer.Text.Trim()) { Set-Status 'Enter a server name first.' 'warn'; return }
    $script:S.Running = $true
    $script:S.RunStart = Get-Date
    $script:S.Failures = 0
    Set-RunningUi $true
    Write-Log ('Monitoring started on {0}: every {1}, stop after {2}.' -f $txtServer.Text.Trim(), $cboInterval.SelectedItem, $cboStopAfter.SelectedItem)
    Invoke-RefreshCycle
    if ($script:S.Running) {
        $timer.Interval = (Get-IntervalSeconds) * 1000
        $timer.Start()
    }
}

function Stop-Monitoring([string]$Reason = 'Monitoring stopped.') {
    $timer.Stop()
    $script:S.Running = $false
    Set-RunningUi $false
    Write-Log $Reason
    Set-Status $Reason 'warn'
}

# Timer: stop -> run one full connect/query/disconnect cycle -> schedule the next one.
# The wait starts AFTER the cycle finishes, so cycles never overlap.
$timer.add_Tick({
        $timer.Stop()
        if (-not $script:S.Running) { return }
        Invoke-RefreshCycle
        if (-not $script:S.Running) { return }
        $limit = Get-StopAfterMinutes
        if ($limit -gt 0 -and ((Get-Date) - $script:S.RunStart).TotalMinutes -ge $limit) {
            Stop-Monitoring ('Stopped automatically after {0} (Stop after setting).' -f $cboStopAfter.SelectedItem)
            return
        }
        $timer.Interval = (Get-IntervalSeconds) * 1000
        $timer.Start()
    })
#endregion

#region ===================== Button handlers =====================
$btnStart.add_Click({ Start-Monitoring })
$btnStop.add_Click({ Stop-Monitoring 'Monitoring stopped by user.' })
$btnNow.add_Click({ Invoke-RefreshCycle })

$btnTest.add_Click({
        $c = $null
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $c = Open-MonitorConnection
            $dt = Invoke-MonitorQuery $c "SELECT CAST(@@SERVERNAME AS nvarchar(256)) AS ServerName, CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(50)) AS Ver, CAST(SERVERPROPERTY('Edition') AS nvarchar(128)) AS Edition, HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER STATE') AS Vss;" 10
            $r = $dt.Rows[0]
            $vss = if ([int](Get-Val $r 'Vss' 0) -eq 1) { 'YES' } else { 'NO (on SQL 2022+ VIEW SERVER PERFORMANCE STATE also works)' }
            $msg = 'Connected OK to {0} - {1} {2}. VIEW SERVER STATE: {3}. Connection closed.' -f $r['ServerName'], $r['Edition'], $r['Ver'], $vss
            Set-Status $msg 'good'; Write-Log $msg
        }
        catch { Set-Status ('Connection failed: ' + $_.Exception.Message) 'bad'; Write-Log ('Connection failed: ' + $_.Exception.Message) }
        finally {
            if ($c) { try { $c.Close() } catch { }; $c.Dispose() }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

$btnCopyBlk.add_Click({
        if ($script:S.LastBlockingRep) { Set-Clip (Get-ReportText $script:S.LastBlockingRep) } else { Set-Clip $rtbBlk.Text }
    })
$btnBlkPrompt.add_Click({
        if (-not $script:S.LastBlockPrompt) { Set-Status 'No blocking has been captured yet in this session.' 'warn'; return }
        $txtPrompt.Text = $script:S.LastBlockPrompt
        Set-Clip $script:S.LastBlockPrompt
    })
$btnExportBlk.add_Click({
        if (-not $script:S.Sessions -or $script:S.Sessions.Count -eq 0) { Set-Status 'Nothing to export - no blocking in the current snapshot.' 'warn'; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV file (*.csv)|*.csv'; $sfd.FileName = 'blocking_{0:yyyyMMdd_HHmmss}.csv' -f (Get-Date)
        if ($sfd.ShowDialog($form) -eq 'OK') {
            $script:S.Sessions | Select-Object SessionId, BlockedBy, Status, WaitType, WaitSec, WaitResource, WaitingFor, DatabaseName, LoginName, HostName, ProgramName, Command, OpenTran, TranAgeSec, IdleSec, IsolationLevel, CpuMs, LogicalReads, ProcName, SqlStatement |
            Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            Set-Status ('Saved ' + $sfd.FileName) 'good'
        }
    })

$btnDlFetch.add_Click({
        if ($script:S.Busy) { return }
        if (-not $txtServer.Text.Trim()) { Set-Status 'Enter a server name first.' 'warn'; return }
        $mode = if ($cboDlSource.SelectedIndex -eq 1) { 'file' } else { 'ring' }
        $c = $null; $recs = $null
        $script:S.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Set-Status 'Reading deadlocks from system_health ...'; $lblStatus.Refresh()
            $c = Open-MonitorConnection
            Update-DbInfo $c
            $recs = Get-ServerDeadlocks $c $mode
        }
        catch { Set-Status ('Deadlock fetch failed: ' + $_.Exception.Message) 'bad'; Write-Log ('Deadlock fetch failed: ' + $_.Exception.Message) }
        finally {
            if ($c) { try { $c.Close() } catch { }; $c.Dispose() }
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $script:S.Busy = $false
        }
        if ($null -ne $recs) {
            $new = Add-Deadlocks $recs $script:S.DlBaselineDone
            $script:S.DlBaselineDone = $true
            Set-Status ('Found {0} deadlock(s) in system_health ({1} new to this list). Connection closed.' -f $recs.Count, $new.Count) 'good'
            if ($gridDl.Rows.Count -gt 0 -and -not $script:S.CurrentDlKey) { $gridDl.CurrentCell = $gridDl.Rows[0].Cells['Time']; Show-SelectedDeadlock }
            if ($recs.Count -eq 0) {
                $rep = New-Report
                Add-R $rep 'No deadlocks found' 'good'
                Add-R $rep 'The system_health session has no deadlock reports in this source. Try "event files (complete)" to look further back, or Open / Paste a deadlock graph you already have.'
                Show-Report $rtbDl $rep
            }
        }
    })

function Import-AndShow([string]$Text, [string]$Source) {
    try {
        $recs = Import-DeadlockText $Text $Source
        if ($recs.Count -eq 0) { Set-Status 'No <deadlock> element found in that XML.' 'bad'; return }
        [void](Add-Deadlocks $recs $false)
        Select-DeadlockRow $recs[0].Key
        Set-Status ('Analyzed {0} deadlock graph(s).' -f $recs.Count) 'good'
    }
    catch { Set-Status ('Could not read the deadlock XML: ' + $_.Exception.Message) 'bad' }
}

$btnDlOpen.add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Deadlock graph (*.xdl;*.xml)|*.xdl;*.xml|All files (*.*)|*.*'
        if ($ofd.ShowDialog($form) -eq 'OK') {
            Import-AndShow ([System.IO.File]::ReadAllText($ofd.FileName)) ('file: ' + [System.IO.Path]::GetFileName($ofd.FileName))
        }
    })
$btnDlPaste.add_Click({
        $init = ''
        try { $clip = [System.Windows.Forms.Clipboard]::GetText(); if ($clip -match '<deadlock') { $init = $clip } } catch { }
        $t = Show-TextDialog 'Paste deadlock graph XML (from SSMS, system_health, Profiler .xdl, etc.)' $init -Editable -OkText 'Analyze'
        if ($t) { Import-AndShow $t 'pasted XML' }
    })
$btnDlSave.add_Click({
        $k = $script:S.CurrentDlKey
        if (-not $k -or -not $script:S.Deadlocks.ContainsKey($k)) { Set-Status 'Select a deadlock first.' 'warn'; return }
        $r = $script:S.Deadlocks[$k]
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'Deadlock graph (*.xdl)|*.xdl'; $sfd.FileName = 'deadlock_{0:yyyyMMdd_HHmmss}.xdl' -f $r.Time
        if ($sfd.ShowDialog($form) -eq 'OK') {
            [System.IO.File]::WriteAllText($sfd.FileName, $r.Node.OuterXml, (New-Object System.Text.UTF8Encoding($false)))
            Set-Status ('Saved ' + $sfd.FileName + ' - open it in SSMS to see the graphical deadlock graph.') 'good'
        }
    })
$btnDlCopy.add_Click({
        $k = $script:S.CurrentDlKey
        if ($k -and $script:S.Deadlocks.ContainsKey($k)) { Set-Clip (Get-ReportText $script:S.Deadlocks[$k].Analysis.Report) } else { Set-Status 'Select a deadlock first.' 'warn' }
    })
$btnDlPrompt.add_Click({
        $k = $script:S.CurrentDlKey
        if ($k -and $script:S.Deadlocks.ContainsKey($k)) { $txtPrompt.Text = $script:S.Deadlocks[$k].Analysis.Prompt; Set-Clip $txtPrompt.Text } else { Set-Status 'Select a deadlock first.' 'warn' }
    })
$btnDlClear.add_Click({
        $dlTable.Clear(); $script:S.Deadlocks = @{}; $script:S.CurrentDlKey = $null; $rtbDl.Clear()
        $script:S.NewDeadlocks = 0; $tabDl.Text = 'Deadlocks'
    })

$btnAiBlk.add_Click({
        if ($script:S.LastBlockPrompt) { $txtPrompt.Text = $script:S.LastBlockPrompt } else { Set-Status 'No blocking has been captured yet in this session.' 'warn' }
    })
$btnAiDl.add_Click({
        $k = $script:S.CurrentDlKey
        if ($k -and $script:S.Deadlocks.ContainsKey($k)) { $txtPrompt.Text = $script:S.Deadlocks[$k].Analysis.Prompt } else { Set-Status 'Select a deadlock on the Deadlocks tab first.' 'warn' }
    })
$btnAiCopy.add_Click({ Set-Clip $txtPrompt.Text })

$tabs.add_SelectedIndexChanged({
        if ($tabs.SelectedTab -eq $tabDl) { $script:S.NewDeadlocks = 0; $tabDl.Text = 'Deadlocks' }
    })

$form.add_Shown({
        try {
            $splitBlk.SplitterDistance = [int]($splitBlk.Height * 0.38)
            $splitDl.SplitterDistance = [int]($splitDl.Height * 0.30)
            $splitHist.SplitterDistance = [int]($splitHist.Height * 0.55)
        }
        catch { }
        $rep = New-Report
        Add-R $rep 'Welcome' 'h1'
        Add-R $rep '1. Enter the server, choose authentication, pick a refresh interval and press Start (or Refresh now for one check).'
        Add-R $rep '2. Each cycle connects, checks, and disconnects. The check uses LOCK_TIMEOUT 500 ms, DEADLOCK_PRIORITY LOW and READ UNCOMMITTED on system views, so it cannot cause blocking or get stuck behind it.'
        Add-R $rep '3. When sessions block each other, this panel explains who is at the front of the queue, why, and what to do - in plain English.'
        Add-R $rep '4. Deadlocks tab: fetch from system_health, or Open / Paste any deadlock graph for a full dissection.'
        Add-R $rep '5. AI Prompt tab: a ready-made prompt you can paste into any AI assistant for a second opinion.'
        Add-R $rep ''
        Add-R $rep 'Permissions needed: VIEW SERVER STATE (SQL 2022+: VIEW SERVER PERFORMANCE STATE).' 'dim'
        Show-Report $rtbBlk $rep
        Write-Log 'Ready.'
    })

$form.add_FormClosing({ $timer.Stop(); $script:S.Running = $false })
#endregion

[void]$form.ShowDialog()
$timer.Dispose()
$form.Dispose()

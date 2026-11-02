<#
    SQL Wait Monitor - desktop viewer (DPA-style)

    Works with both installs:
      DBAWait         single server   (Install-SqlWaitMonitor.sql)
      DBAWaitCentral  many servers via linked servers (Install-SqlWaitMonitorCentral.sql)
                      -> a server picker appears at the top
    It detects which one it is connected to.
    - Stacked bars of wait time per time slice, by STATEMENT or by WAIT TYPE
    - Click a bar  -> that statement's detail for the clicked time slice
      (exact SQL text, proc + line, waits, who blocked it, who ran it, plan)
    - Wait-type mode: click a bar -> statements that had that wait in that slice
    - Tabs: top statements, head blockers, worker threads, alert history

    Start: double-click the .cmd launcher, or
           powershell.exe -NoProfile -ExecutionPolicy Bypass -File SqlWaitMonitor.ps1 -Server localhost -Database DBAWait
    Uses Windows authentication by default. Add -SqlLogin to be prompted for a SQL login.
#>
param(
    [string]$Server        = 'localhost',
    [string]$Database      = 'DBAWait',
    [switch]$SqlLogin,
    [int]   $WaitThreshold = 10000      # seconds of total wait in 10 minutes = red
)

# The chart library ships with Windows PowerShell 5.1 (.NET Framework), not PowerShell 7.
if ($PSVersionTable.PSEdition -eq 'Core') {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"",
                 '-Server', $Server, '-Database', $Database, '-WaitThreshold', $WaitThreshold)
    if ($SqlLogin) { $argList += '-SqlLogin' }
    Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $argList
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------- connection
$csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
$csb['Data Source']      = $Server
$csb['Initial Catalog']  = $Database
$csb['Application Name'] = 'SQL Wait Monitor viewer'
$csb['Connect Timeout']  = 10
if ($SqlLogin) {
    $cred = Get-Credential -Message "SQL Server login for $Server"
    if (-not $cred) { return }
    $csb['User ID']  = $cred.UserName
    $csb['Password'] = $cred.GetNetworkCredential().Password
} else {
    $csb['Integrated Security'] = $true
}
$script:ConnStr = $csb.ConnectionString

# Queries carry server-filter tokens. Central (multi-server) database: they become
# "AND x.ServerName = @srv". Single-server database (no ServerName column): removed.
#   {S}      -> AND ServerName = @srv
#   {S:w}    -> AND w.ServerName = @srv
#   {J:b:w}  -> AND b.ServerName = w.ServerName
$script:Multi = $false
$script:Srv = $null
$script:LinkedSrv = $null
function Expand-ServerFilter([string]$Sql) {
    if ($script:Multi) {
        $Sql = $Sql -replace '\{S:(\w+)\}', ' AND $1.ServerName = @srv'
        $Sql = $Sql -replace '\{S\}', ' AND ServerName = @srv'
        $Sql = $Sql -replace '\{J:(\w+):(\w+)\}', ' AND $1.ServerName = $2.ServerName'
    } else {
        $Sql = $Sql -replace '\{S:\w+\}', '' -replace '\{S\}', '' -replace '\{J:\w+:\w+\}', ''
    }
    return $Sql
}

function Invoke-Sql {
    param([string]$Sql, [hashtable]$Params = @{})
    if (-not $Params.ContainsKey('srv')) { $Params['srv'] = $script:Srv }
    $conn = New-Object System.Data.SqlClient.SqlConnection $script:ConnStr
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = Expand-ServerFilter $Sql
        $cmd.CommandTimeout = 60
        foreach ($k in $Params.Keys) {
            $v = $Params[$k]
            if ($null -eq $v) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::NVarChar, 128); $p.Value = [DBNull]::Value
            } elseif ($v -is [datetime]) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::DateTime2); $p.Value = $v
            } elseif ($v -is [byte[]]) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::VarBinary, 64); $p.Value = $v
            } elseif ($v -is [int]) {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::Int); $p.Value = $v
            } else {
                $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::NVarChar, 4000); $p.Value = [string]$v
            }
        }
        $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dt = New-Object System.Data.DataTable
        [void]$da.Fill($dt)
        return , $dt
    } finally {
        $conn.Dispose()
    }
}

function Show-Error([string]$Where, $Err) {
    [void][System.Windows.Forms.MessageBox]::Show("$Where`n`n$($Err.Exception.Message)", 'SQL Wait Monitor',
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}

function Format-Label([string]$Text, [int]$Max = 60) {
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max - 3) + '...' }
    return $t
}

# ---------------------------------------------------------------- which install is this?
try {
    $mode = Invoke-Sql "SELECT IsCentral = CASE WHEN COL_LENGTH('dbo.WaitSamples', 'ServerName') IS NULL THEN 0 ELSE 1 END;"
    $script:Multi = ([int]$mode.Rows[0].IsCentral -eq 1)
    if ($script:Multi) {
        $script:Servers = Invoke-Sql "SELECT ServerName, LinkedServer FROM dbo.MonitoredServers WHERE IsEnabled = 1 ORDER BY ServerName;"
        if ($script:Servers.Rows.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('No enabled servers in dbo.MonitoredServers. Register one with dbo.usp_AddMonitoredServer.', 'SQL Wait Monitor')
            return
        }
    }
} catch {
    Show-Error "Could not open database $Database on $Server." $_
    return
}

function Set-CurrentServer {
    $script:Srv = [string]$cboServer.SelectedItem
    $row = $script:Servers.Select("ServerName = '" + ($script:Srv -replace "'", "''") + "'")
    $script:LinkedSrv = $null
    if ($row.Count -gt 0 -and $row[0].LinkedServer -isnot [DBNull]) { $script:LinkedSrv = [string]$row[0].LinkedServer }
    $via = $(if ($script:LinkedSrv) { "linked server $($script:LinkedSrv)" } else { 'local' })
    $form.Text = "SQL Wait Monitor - $($script:Srv)  ($via, data in $Database on $Server)"
}

# ---------------------------------------------------------------- state
$script:TzMin    = [int][TimeZoneInfo]::Local.GetUtcOffset([DateTime]::Now).TotalMinutes
# range in minutes, slice (bar) size in seconds
$script:Ranges   = [ordered]@{
    'Last 5 minutes'  = @(5, 10)
    'Last 15 minutes' = @(15, 30)
    'Last 1 hour'     = @(60, 60)
    'Last 4 hours'    = @(240, 300)
    'Last 24 hours'   = @(1440, 900)
    'Last 7 days'     = @(10080, 3600)
}
$script:Bucket      = 60
$script:FromUtc     = [DateTime]::UtcNow.AddHours(-1)
$script:ToUtc       = [DateTime]::UtcNow
$script:ByStatement = $true
$script:GridScope   = @{ From = $script:FromUtc; To = $script:ToUtc; Wait = $null }
$script:Palette     = @('#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F',
                        '#EDC948', '#B07AA1', '#FF9DA7', '#9C755F', '#2F4B7C') |
                      ForEach-Object { [System.Drawing.ColorTranslator]::FromHtml($_) }

# ---------------------------------------------------------------- SQL
# Bucket = start of the time slice (@b seconds wide), converted to local time (@tz minutes)
$BucketExpr = "DATEADD(MINUTE, @tz, DATEADD(SECOND, DATEDIFF(SECOND, '20000101', w.SampleTime) / @b * @b, '20000101'))"

$SqlChartByStatement = @"
WITH top10 AS (
    SELECT TOP (10) StatementId FROM dbo.WaitSamples
    WHERE SampleTime >= @from AND SampleTime < @to{S} AND StatementId IS NOT NULL
    GROUP BY StatementId ORDER BY SUM(WeightSec) DESC)
SELECT b.Bucket,
       SeriesKey  = CAST(ISNULL(t.StatementId, 0) AS nvarchar(20)),
       SeriesName = ISNULL(v.Label, N'Other'),
       WaitSec    = SUM(w.WeightSec)
FROM dbo.WaitSamples w
CROSS APPLY (SELECT Bucket = $BucketExpr) b
LEFT JOIN top10 t ON t.StatementId = w.StatementId
LEFT JOIN dbo.vStatements v ON v.StatementId = t.StatementId
WHERE w.SampleTime >= @from AND w.SampleTime < @to{S:w}
GROUP BY b.Bucket, ISNULL(t.StatementId, 0), ISNULL(v.Label, N'Other');
"@

$SqlChartByWait = @"
WITH top10 AS (
    SELECT TOP (10) WaitCategory FROM dbo.WaitSamples
    WHERE SampleTime >= @from AND SampleTime < @to{S}
    GROUP BY WaitCategory ORDER BY SUM(WeightSec) DESC)
SELECT b.Bucket,
       SeriesKey  = ISNULL(t.WaitCategory, N'Other'),
       SeriesName = ISNULL(t.WaitCategory, N'Other'),
       WaitSec    = SUM(w.WeightSec)
FROM dbo.WaitSamples w
CROSS APPLY (SELECT Bucket = $BucketExpr) b
LEFT JOIN top10 t ON t.WaitCategory = w.WaitCategory
WHERE w.SampleTime >= @from AND w.SampleTime < @to{S:w}
GROUP BY b.Bucket, ISNULL(t.WaitCategory, N'Other');
"@

$SqlTopGrid = @"
SELECT TOP (50) w.StatementId,
       [Statement]  = v.Label,
       [Database]   = v.DatabaseName,
       [Wait sec]   = SUM(w.WeightSec),
       [% of total] = CAST(100.0 * SUM(w.WeightSec) / NULLIF(SUM(SUM(w.WeightSec)) OVER (), 0) AS decimal(5,1)),
       [Main wait]  = (SELECT TOP (1) w2.WaitCategory FROM dbo.WaitSamples w2
                       WHERE w2.StatementId = w.StatementId AND w2.SampleTime >= @from AND w2.SampleTime < @to
                         AND (@wait IS NULL OR w2.WaitCategory = @wait)
                       GROUP BY w2.WaitCategory ORDER BY SUM(w2.WeightSec) DESC),
       [Sessions]   = COUNT(DISTINCT w.SessionId),
       [Last seen]  = DATEADD(MINUTE, @tz, MAX(w.SampleTime))
FROM dbo.WaitSamples w
JOIN dbo.vStatements v ON v.StatementId = w.StatementId
WHERE w.SampleTime >= @from AND w.SampleTime < @to{S:w}
  AND (@wait IS NULL OR w.WaitCategory = @wait)
GROUP BY w.StatementId, v.Label, v.DatabaseName
ORDER BY [Wait sec] DESC;
"@

$SqlHeader = @"
SELECT Total10    = ISNULL((SELECT SUM(WeightSec) FROM dbo.WaitSamples
                            WHERE SampleTime >= DATEADD(MINUTE, -10, SYSUTCDATETIME()){S}), 0),
       BlockedNow = (SELECT COUNT(DISTINCT SessionId) FROM dbo.WaitSamples
                     WHERE BlockingSessionId IS NOT NULL{S}
                       AND SampleTime = (SELECT MAX(SampleTime) FROM dbo.WaitSamples WHERE 1 = 1{S})),
       LastSample = DATEADD(MINUTE, @tz, (SELECT MAX(SampleTime) FROM dbo.WaitSamples WHERE 1 = 1{S}));
"@

$SqlBlockers = @"
SELECT [Blocker SPID] = SessionId,
       [Status]       = MAX(Status),
       [Open tran]    = MAX(OpenTranCount),
       [Max blocked]  = MAX(BlockedCount),
       [First seen]   = DATEADD(MINUTE, @tz, MIN(SampleTime)),
       [Last seen]    = DATEADD(MINUTE, @tz, MAX(SampleTime)),
       [Login]        = MAX(LoginName),
       [Host]         = MAX(HostName),
       [Program]      = MAX(ProgramName),
       [Last SQL]     = MAX(LastSqlText)
FROM dbo.BlockerSamples
WHERE SampleTime >= @from AND SampleTime < @to{S}
GROUP BY SessionId
ORDER BY [Max blocked] DESC;
"@

$SqlWorkers = @"
SELECT T = DATEADD(MINUTE, @tz, SampleTime), ActiveWorkers, CurrentWorkers, RunnableTasks, WorkQueueCount
FROM dbo.SchedulerSamples
WHERE SampleTime >= @from AND SampleTime < @to{S}
ORDER BY SampleTime;
"@

$SqlAlerts = @"
SELECT [Time] = DATEADD(MINUTE, @tz, AlertTime), [Type] = AlertType, [Total wait sec] = TotalWaitSec,
       [Blocked] = BlockedCount, [Head blocker SPID] = HeadBlocker, Details
FROM dbo.AlertLog
WHERE AlertTime >= @from AND AlertTime < @to{S}
ORDER BY AlertTime DESC;
"@

# ---- statement detail
$SqlStmtInfo = @"
SELECT s.StatementId, v.Label, s.DatabaseName, s.ObjectId, s.ObjectName, s.StartLine, s.StatementText,
       s.LastCalledFrom, s.LastPlanHandle,
       FirstSeen = DATEADD(MINUTE, @tz, s.FirstSeen), LastSeen = DATEADD(MINUTE, @tz, s.LastSeen)
FROM dbo.SqlStatements s
JOIN dbo.vStatements v ON v.StatementId = s.StatementId
WHERE s.StatementId = @id;
"@

$SqlStmtChart = @"
SELECT b.Bucket, SeriesKey = w.WaitCategory, SeriesName = w.WaitCategory, WaitSec = SUM(w.WeightSec)
FROM dbo.WaitSamples w
CROSS APPLY (SELECT Bucket = $BucketExpr) b
WHERE w.StatementId = @id AND w.SampleTime >= @from AND w.SampleTime < @to
GROUP BY b.Bucket, w.WaitCategory;
"@

$SqlStmtWaits = @"
SELECT [Wait type]        = WaitCategory,
       [Wait sec]         = SUM(WeightSec),
       [% of statement]   = CAST(100.0 * SUM(WeightSec) / NULLIF(SUM(SUM(WeightSec)) OVER (), 0) AS decimal(5,1)),
       [Example resource] = MAX(WaitResource)
FROM dbo.WaitSamples
WHERE StatementId = @id AND SampleTime >= @from AND SampleTime < @to
GROUP BY WaitCategory
ORDER BY [Wait sec] DESC;
"@

$SqlStmtBlockers = @"
SELECT [Blocker SPID]      = w.BlockingSessionId,
       [Wait sec]          = SUM(w.WeightSec),
       [First]             = DATEADD(MINUTE, @tz, MIN(w.SampleTime)),
       [Last]              = DATEADD(MINUTE, @tz, MAX(w.SampleTime)),
       [Blocker statement] = MAX(COALESCE(bv.Label, LEFT(bs.LastSqlText, 300))),
       [Blocker status]    = MAX(bs.Status),
       [Blocker open tran] = MAX(bs.OpenTranCount),
       [Blocker login]     = MAX(bs.LoginName),
       [Blocker host]      = MAX(bs.HostName)
FROM dbo.WaitSamples w
OUTER APPLY (SELECT TOP (1) b.StatementId FROM dbo.WaitSamples b
             WHERE b.SampleTime = w.SampleTime AND b.SessionId = w.BlockingSessionId{J:b:w}) bw
LEFT JOIN dbo.vStatements bv ON bv.StatementId = bw.StatementId
LEFT JOIN dbo.BlockerSamples bs ON bs.SampleTime = w.SampleTime AND bs.SessionId = w.BlockingSessionId{J:bs:w}
WHERE w.StatementId = @id AND w.BlockingSessionId IS NOT NULL
  AND w.SampleTime >= @from AND w.SampleTime < @to
GROUP BY w.BlockingSessionId
ORDER BY [Wait sec] DESC;
"@

$SqlStmtWho = @"
SELECT [Login] = LoginName, [Host] = HostName, [Program] = ProgramName,
       [Wait sec] = SUM(WeightSec), [Sessions] = COUNT(DISTINCT SessionId)
FROM dbo.WaitSamples
WHERE StatementId = @id AND SampleTime >= @from AND SampleTime < @to
GROUP BY LoginName, HostName, ProgramName
ORDER BY [Wait sec] DESC;
"@

$SqlPlan = "SELECT QueryPlan = CAST(query_plan AS nvarchar(max)) FROM sys.dm_exec_query_plan(@h);"

$SqlCollectorStatus = @"
SELECT [Server] = ServerName, [Linked server] = ISNULL(LinkedServer, N'(central server itself)'), [Enabled] = IsEnabled,
       [Last collected] = DATEADD(MINUTE, @tz, LastCollected), [Duration ms] = LastDurationMs,
       [Last error] = LastError, [Error time] = DATEADD(MINUTE, @tz, LastErrorTime)
FROM dbo.MonitoredServers
ORDER BY ServerName;
"@

# ---------------------------------------------------------------- chart helpers
function New-WaitChart {
    $c = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $c.Dock = 'Fill'
    $c.BackColor = [System.Drawing.Color]::White
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 'main'
    $area.AxisX.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $area.AxisY.LabelStyle.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $area.AxisY.Title = 'Wait seconds'
    $area.AxisY.TitleFont = New-Object System.Drawing.Font('Segoe UI', 8)
    $area.AxisX.IntervalAutoMode = 'VariableCount'
    $c.ChartAreas.Add($area)
    $lg = New-Object System.Windows.Forms.DataVisualization.Charting.Legend 'legend'
    $lg.Docking = 'Right'
    $lg.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lg.IsTextAutoFit = $false
    $c.Legends.Add($lg)
    return $c
}

function Get-Buckets([datetime]$FromUtc, [datetime]$ToUtc, [int]$BucketSec) {
    $base  = [datetime]'2000-01-01'
    $start = [math]::Floor(($FromUtc - $base).TotalSeconds / $BucketSec) * $BucketSec
    $t     = $base.AddSeconds($start).AddMinutes($script:TzMin)
    $end   = $ToUtc.AddMinutes($script:TzMin)
    $list  = New-Object System.Collections.Generic.List[datetime]
    while ($t -lt $end) { $list.Add($t); $t = $t.AddSeconds($BucketSec) }
    return , $list
}

function Get-SliceText([int]$BucketSec) {
    if ($BucketSec -lt 60) { return "$BucketSec-second" }
    if ($BucketSec -lt 3600) { return "$($BucketSec / 60)-minute" }
    return "$($BucketSec / 3600)-hour"
}

function Get-TimeFormat([int]$BucketSec) {
    if ($BucketSec -lt 60) { return 'HH:mm:ss' }
    return 'HH:mm'
}

# Fill a chart with stacked columns from rows of (Bucket, SeriesKey, SeriesName, WaitSec)
function Set-StackedChart($Chart, $Table, $Buckets, [string]$OtherKey, [string]$ClickHint) {
    $Chart.Series.Clear()
    $totals = @{}; $names = @{}; $lookup = @{}
    foreach ($r in $Table.Rows) {
        $k = [string]$r.SeriesKey
        $totals[$k] = [double]$totals[$k] + [double]$r.WaitSec
        $names[$k]  = [string]$r.SeriesName
        $lookup["$k|$(([datetime]$r.Bucket).Ticks)"] = [double]$r.WaitSec
    }
    $tipFmt = 'HH:mm'
    if ($Buckets.Count -gt 1 -and ($Buckets[1] - $Buckets[0]).TotalSeconds -lt 60) { $tipFmt = 'HH:mm:ss' }
    $keys = @($totals.Keys | Sort-Object -Descending -Property { if ($_ -eq $OtherKey) { -1 } else { $totals[$_] } })
    $i = 0
    foreach ($k in $keys) {
        $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series ("s_" + $k)
        $s.ChartType   = 'StackedColumn'
        $s.XValueType  = 'DateTime'
        $s['PointWidth'] = '0.85'
        $s.Tag         = $k
        $s.LegendText  = Format-Label $names[$k] 55
        if ($k -eq $OtherKey) { $s.Color = [System.Drawing.Color]::Silver }
        else { $s.Color = $script:Palette[$i % $script:Palette.Count]; $i++ }
        $safeName = Format-Label $names[$k] 90
        $s.ToolTip = "#VALX{$tipFmt}`n$safeName`n#VAL{N0} s wait`n$ClickHint"
        foreach ($t in $Buckets) {
            $v = $lookup["$k|$($t.Ticks)"]
            if ($null -eq $v) { $v = 0 }
            [void]$s.Points.AddXY($t, $v)
        }
        $Chart.Series.Add($s)
    }
    $fmt = $tipFmt
    if ($Buckets.Count -gt 0 -and ($Buckets[$Buckets.Count - 1] - $Buckets[0]).TotalHours -gt 20) { $fmt = 'MM-dd HH:mm' }
    $Chart.ChartAreas[0].AxisX.LabelStyle.Format = $fmt
}

function Set-GridStyle($Grid) {
    $Grid.Dock = 'Fill'
    $Grid.ReadOnly = $true
    $Grid.AllowUserToAddRows = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = 'FullRowSelect'
    $Grid.MultiSelect = $false
    $Grid.BackgroundColor = [System.Drawing.Color]::White
    $Grid.BorderStyle = 'None'
    $Grid.AutoSizeColumnsMode = 'DisplayedCells'
    $Grid.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $Grid.Add_DataBindingComplete({
        param($sender, $e)
        foreach ($col in $sender.Columns) {
            if ($col.Name -eq 'StatementId' -or $col.Name -eq 'Details') { $col.Visible = $false }
            if ($col.ValueType -eq [datetime]) { $col.DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm:ss' }
            if ($col.Name -like '*sec*') { $col.DefaultCellStyle.Format = 'N0' }
            if ($col.Name -in @('Statement', 'Last SQL', 'Blocker statement')) {
                $col.AutoSizeMode = 'None'; $col.Width = 520
            }
        }
    })
}

# ---------------------------------------------------------------- main window
$form = New-Object System.Windows.Forms.Form
$form.Text = "SQL Wait Monitor - $Server"
$form.Size = New-Object System.Drawing.Size(1500, 950)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$top = New-Object System.Windows.Forms.FlowLayoutPanel
$top.Dock = 'Top'; $top.Height = 44; $top.Padding = New-Object System.Windows.Forms.Padding(6, 8, 6, 0)
$top.WrapContents = $false

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = "Server: $Server"; $lblServer.AutoSize = $true
$lblServer.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblServer.Margin = New-Object System.Windows.Forms.Padding(0, 4, 16, 0)

$cboServer = New-Object System.Windows.Forms.ComboBox
$cboServer.DropDownStyle = 'DropDownList'; $cboServer.Width = 200
$cboServer.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
$cboServer.Visible = $script:Multi
if ($script:Multi) {
    $lblServer.Text = 'Server:'
    $lblServer.Margin = New-Object System.Windows.Forms.Padding(0, 4, 4, 0)
    foreach ($row in $script:Servers.Rows) { [void]$cboServer.Items.Add([string]$row.ServerName) }
}

$cboRange = New-Object System.Windows.Forms.ComboBox
$cboRange.DropDownStyle = 'DropDownList'; $cboRange.Width = 130
foreach ($k in $script:Ranges.Keys) { [void]$cboRange.Items.Add($k) }
$cboRange.SelectedItem = 'Last 1 hour'

$lblGroup = New-Object System.Windows.Forms.Label
$lblGroup.Text = 'Bars by:'; $lblGroup.AutoSize = $true
$lblGroup.Margin = New-Object System.Windows.Forms.Padding(12, 5, 2, 0)

$cboGroup = New-Object System.Windows.Forms.ComboBox
$cboGroup.DropDownStyle = 'DropDownList'; $cboGroup.Width = 110
[void]$cboGroup.Items.Add('Statement'); [void]$cboGroup.Items.Add('Wait type')
$cboGroup.SelectedIndex = 0

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'; $btnRefresh.Width = 80
$btnRefresh.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)

# auto-refresh choices: label -> seconds (0 = Off)
$script:RefreshOptions = [ordered]@{ 'Off' = 0; '5s' = 5; '10s' = 10; '30s' = 30; '1m' = 60; '5m' = 300; '15m' = 900 }
$lblAuto = New-Object System.Windows.Forms.Label
$lblAuto.Text = 'Auto-refresh:'; $lblAuto.AutoSize = $true
$lblAuto.Margin = New-Object System.Windows.Forms.Padding(12, 5, 2, 0)
$cboRefresh = New-Object System.Windows.Forms.ComboBox
$cboRefresh.DropDownStyle = 'DropDownList'; $cboRefresh.Width = 60
$cboRefresh.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
foreach ($k in $script:RefreshOptions.Keys) { [void]$cboRefresh.Items.Add($k) }
$cboRefresh.SelectedItem = '30s'

$lblTotal = New-Object System.Windows.Forms.Label
$lblTotal.AutoSize = $true; $lblTotal.Padding = New-Object System.Windows.Forms.Padding(8, 3, 8, 3)
$lblTotal.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblTotal.ForeColor = [System.Drawing.Color]::White

$lblBlocked = New-Object System.Windows.Forms.Label
$lblBlocked.AutoSize = $true; $lblBlocked.Margin = New-Object System.Windows.Forms.Padding(12, 5, 0, 0)

$lblUpdated = New-Object System.Windows.Forms.Label
$lblUpdated.AutoSize = $true; $lblUpdated.ForeColor = [System.Drawing.Color]::Gray
$lblUpdated.Margin = New-Object System.Windows.Forms.Padding(12, 5, 0, 0)

$top.Controls.AddRange(@($lblServer, $cboServer, $cboRange, $lblGroup, $cboGroup, $btnRefresh, $lblAuto, $cboRefresh, $lblTotal, $lblBlocked, $lblUpdated))

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'

$chartTitle = New-Object System.Windows.Forms.Label
$chartTitle.Dock = 'Top'; $chartTitle.Height = 24
$chartTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$chartTitle.Padding = New-Object System.Windows.Forms.Padding(6, 4, 0, 0)

$script:mainChart = New-WaitChart
$split.Panel1.Controls.Add($script:mainChart)
$split.Panel1.Controls.Add($chartTitle)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tpTop      = New-Object System.Windows.Forms.TabPage 'Top statements'
$tpBlockers = New-Object System.Windows.Forms.TabPage 'Head blockers'
$tpWorkers  = New-Object System.Windows.Forms.TabPage 'Worker threads'
$tpAlerts   = New-Object System.Windows.Forms.TabPage 'Alert history'
$tabs.TabPages.AddRange(@($tpTop, $tpBlockers, $tpWorkers, $tpAlerts))
$tpStatus = New-Object System.Windows.Forms.TabPage 'Collector status'
if ($script:Multi) { $tabs.TabPages.Add($tpStatus) }
$split.Panel2.Controls.Add($tabs)

# Top statements tab
$gridCaption = New-Object System.Windows.Forms.Panel
$gridCaption.Dock = 'Top'; $gridCaption.Height = 26
$lblGrid = New-Object System.Windows.Forms.Label
$lblGrid.AutoSize = $true; $lblGrid.Location = New-Object System.Drawing.Point(4, 5)
$lnkReset = New-Object System.Windows.Forms.LinkLabel
$lnkReset.Text = 'Show whole range'; $lnkReset.AutoSize = $true; $lnkReset.Visible = $false
$gridCaption.Controls.AddRange(@($lblGrid, $lnkReset))
$script:gridTop = New-Object System.Windows.Forms.DataGridView
Set-GridStyle $script:gridTop
$tpTop.Controls.Add($script:gridTop)
$tpTop.Controls.Add($gridCaption)

$script:gridBlockers = New-Object System.Windows.Forms.DataGridView
Set-GridStyle $script:gridBlockers
$lblBlkHelp = New-Object System.Windows.Forms.Label
$lblBlkHelp.Dock = 'Top'; $lblBlkHelp.Height = 22; $lblBlkHelp.ForeColor = [System.Drawing.Color]::DimGray
$lblBlkHelp.Text = "  Sessions that blocked others but were not blocked themselves. Status 'sleeping' with Open tran > 0 = idle session holding locks (e.g. an open BEGIN TRAN)."
$tpBlockers.Controls.Add($script:gridBlockers)
$tpBlockers.Controls.Add($lblBlkHelp)

$script:workerChart = New-WaitChart
$script:workerChart.ChartAreas[0].AxisY.Title = 'Count'
$tpWorkers.Controls.Add($script:workerChart)

$script:gridAlerts = New-Object System.Windows.Forms.DataGridView
Set-GridStyle $script:gridAlerts
$tpAlerts.Controls.Add($script:gridAlerts)

$script:gridStatus = New-Object System.Windows.Forms.DataGridView
Set-GridStyle $script:gridStatus
$tpStatus.Controls.Add($script:gridStatus)
if ($script:Multi) { $cboServer.SelectedIndex = 0; Set-CurrentServer }

$status = New-Object System.Windows.Forms.StatusStrip
$statusText = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusText.Text = 'Click a bar to drill into that time slice. Click a legend entry for the whole range. Double-click a grid row for statement detail.'
[void]$status.Items.Add($statusText)

$form.Controls.Add($split)
$form.Controls.Add($top)
$form.Controls.Add($status)

# ---------------------------------------------------------------- main refresh
function Update-TopGrid([datetime]$FromUtc, [datetime]$ToUtc, $Wait, [string]$Caption) {
    $script:GridScope = @{ From = $FromUtc; To = $ToUtc; Wait = $Wait }
    $dt = Invoke-Sql $SqlTopGrid @{ from = $FromUtc; to = $ToUtc; wait = $Wait; tz = $script:TzMin }
    $script:gridTop.DataSource = $dt
    $lblGrid.Text = $Caption
    $lnkReset.Location = New-Object System.Drawing.Point(($lblGrid.Right + 12), 5)
    $lnkReset.Visible = ($FromUtc -ne $script:FromUtc -or $ToUtc -ne $script:ToUtc -or $null -ne $Wait)
}

function Update-All {
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $r = $script:Ranges[[string]$cboRange.SelectedItem]
        $script:Bucket      = [int]$r[1]
        $script:ToUtc       = [DateTime]::UtcNow
        $script:FromUtc     = $script:ToUtc.AddMinutes(-[int]$r[0])
        $script:ByStatement = ($cboGroup.SelectedIndex -eq 0)
        $p = @{ from = $script:FromUtc; to = $script:ToUtc; b = $script:Bucket; tz = $script:TzMin }
        $buckets = Get-Buckets $script:FromUtc $script:ToUtc $script:Bucket

        if ($script:ByStatement) {
            $dt = Invoke-Sql $SqlChartByStatement $p
            Set-StackedChart $script:mainChart $dt $buckets '0' '(click for statement detail)'
            $chartTitle.Text = "Top statements by wait time - $(Get-SliceText $script:Bucket) slices (click a bar for detail)"
        } else {
            $dt = Invoke-Sql $SqlChartByWait $p
            Set-StackedChart $script:mainChart $dt $buckets 'Other' '(click for the statements behind it)'
            $chartTitle.Text = "Total wait by wait type - $(Get-SliceText $script:Bucket) slices (click a bar to see which statements)"
        }

        Update-TopGrid $script:FromUtc $script:ToUtc $null 'Top statements in the selected range - double-click a row for detail'

        $script:gridBlockers.DataSource = Invoke-Sql $SqlBlockers @{ from = $script:FromUtc; to = $script:ToUtc; tz = $script:TzMin }
        $script:gridAlerts.DataSource   = Invoke-Sql $SqlAlerts   @{ from = $script:FromUtc; to = $script:ToUtc; tz = $script:TzMin }

        # worker threads
        $w = Invoke-Sql $SqlWorkers @{ from = $script:FromUtc; to = $script:ToUtc; tz = $script:TzMin }
        $script:workerChart.Series.Clear()
        $i = 0
        foreach ($col in @('ActiveWorkers', 'CurrentWorkers', 'RunnableTasks', 'WorkQueueCount')) {
            $s = New-Object System.Windows.Forms.DataVisualization.Charting.Series $col
            $s.ChartType = 'Line'; $s.XValueType = 'DateTime'; $s.BorderWidth = 2
            $s.Color = $script:Palette[$i]; $i++
            $s.ToolTip = "#VALX{HH:mm:ss}  $col = #VAL"
            foreach ($row in $w.Rows) { [void]$s.Points.AddXY([datetime]$row.T, [double]$row[$col]) }
            $script:workerChart.Series.Add($s)
        }
        $script:workerChart.ChartAreas[0].AxisX.LabelStyle.Format = 'HH:mm'

        # header
        $h = (Invoke-Sql $SqlHeader @{ tz = $script:TzMin }).Rows[0]
        $total = [int]$h.Total10
        $lblTotal.Text = ('Total wait last 10 min: {0:N0} s' -f $total)
        if ($total -ge $WaitThreshold) { $lblTotal.BackColor = [System.Drawing.Color]::Firebrick }
        elseif ($total -ge ($WaitThreshold / 2)) { $lblTotal.BackColor = [System.Drawing.Color]::DarkOrange }
        else { $lblTotal.BackColor = [System.Drawing.Color]::SeaGreen }
        $lblBlocked.Text = "Blocked now: $($h.BlockedNow)"
        $lblBlocked.ForeColor = $(if ([int]$h.BlockedNow -gt 0) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black })
        $last = ''
        if ($h.LastSample -isnot [DBNull]) { $last = ([datetime]$h.LastSample).ToString('HH:mm:ss') }
        $lblUpdated.Text = "Last sample $last | refreshed $((Get-Date).ToString('HH:mm:ss'))"
        $lblUpdated.ForeColor = [System.Drawing.Color]::Gray
        if ($script:Multi) {
            $cs = Invoke-Sql $SqlCollectorStatus @{ tz = $script:TzMin }
            $script:gridStatus.DataSource = $cs
            $mine = $cs.Select("Server = '" + ($script:Srv -replace "'", "''") + "'")
            if ($mine.Count -gt 0 -and $mine[0]['Last error'] -isnot [DBNull]) {
                $lblUpdated.Text += ' | COLLECTOR ERROR - see Collector status tab'
                $lblUpdated.ForeColor = [System.Drawing.Color]::Firebrick
            }
        }
    } catch {
        $cboRefresh.SelectedItem = 'Off'
        Show-Error 'Could not load data (auto-refresh paused).' $_
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ---------------------------------------------------------------- statement detail window
function Show-StatementDetail([int]$Id, [datetime]$SliceFromUtc, [datetime]$SliceToUtc) {
    try {
        $info = Invoke-Sql $SqlStmtInfo @{ id = $Id; tz = $script:TzMin }
    } catch { Show-Error 'Could not load the statement.' $_; return }
    if ($info.Rows.Count -eq 0) { return }
    $s = $info.Rows[0]

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Statement detail - $(Format-Label ([string]$s.Label) 90)"
    if ($script:Multi) { $f.Text += "  [$($script:Srv)]" }
    $f.Size = New-Object System.Drawing.Size(1300, 900)
    $f.StartPosition = 'CenterParent'
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $state = @{
        Id = $Id; SliceFrom = $SliceFromUtc; SliceTo = $SliceToUtc
        RangeFrom = $script:FromUtc; RangeTo = $script:ToUtc; Bucket = $script:Bucket
        DbName = [string]$s.DatabaseName; ObjectId = [int]$s.ObjectId
        StartLine = $(if ($s.StartLine -is [DBNull]) { 1 } else { [int]$s.StartLine })
        StmtText = [string]$s.StatementText
        PlanHandle = $(if ($s.LastPlanHandle -is [DBNull]) { $null } else { [byte[]]$s.LastPlanHandle })
        Srv = $script:Srv; LinkedSrv = $script:LinkedSrv
    }

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'; $layout.ColumnCount = 1; $layout.RowCount = 4
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 24)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 38)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 38)))

    # header
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock = 'Fill'; $hdr.Height = 128
    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Text = [string]$s.Label; $lblT.AutoSize = $true; $lblT.Location = New-Object System.Drawing.Point(6, 4)
    $lblT.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $obj = $(if ([int]$s.ObjectId -ne 0) { "$($s.ObjectName)   line $($state.StartLine)" } else { '(ad hoc or dynamic SQL)' })
    $called = Format-Label ([string]$s.LastCalledFrom) 160
    $lblD = New-Object System.Windows.Forms.Label
    $lblD.AutoSize = $true; $lblD.Location = New-Object System.Drawing.Point(8, 32)
    $lblD.Text = "Database: $($s.DatabaseName)      Procedure / object: $obj`n" +
                 "Called from: $called`n" +
                 "First seen: $(([datetime]$s.FirstSeen).ToString('yyyy-MM-dd HH:mm'))      Last seen: $(([datetime]$s.LastSeen).ToString('yyyy-MM-dd HH:mm'))"

    $lblScope = New-Object System.Windows.Forms.Label
    $lblScope.Text = 'Tables below show:'; $lblScope.AutoSize = $true; $lblScope.Location = New-Object System.Drawing.Point(8, 98)
    $cboScope = New-Object System.Windows.Forms.ComboBox
    $cboScope.DropDownStyle = 'DropDownList'; $cboScope.Width = 330; $cboScope.Location = New-Object System.Drawing.Point(125, 94)
    $tf = Get-TimeFormat $script:Bucket
    $sliceText = "Clicked slice  {0} - {1}" -f $SliceFromUtc.AddMinutes($script:TzMin).ToString($tf), $SliceToUtc.AddMinutes($script:TzMin).ToString($tf)
    [void]$cboScope.Items.Add($sliceText)
    [void]$cboScope.Items.Add('Whole selected range')
    $cboScope.SelectedIndex = $(if ($SliceFromUtc -eq $script:FromUtc -and $SliceToUtc -eq $script:ToUtc) { 1 } else { 0 })

    $btnProc = New-Object System.Windows.Forms.Button
    $btnProc.Text = 'Show whole procedure'; $btnProc.Width = 160; $btnProc.Location = New-Object System.Drawing.Point(475, 93)
    $btnProc.Enabled = ([int]$s.ObjectId -ne 0)
    $btnPlan = New-Object System.Windows.Forms.Button
    $btnPlan.Text = 'Open execution plan'; $btnPlan.Width = 150; $btnPlan.Location = New-Object System.Drawing.Point(645, 93)
    $btnPlan.Enabled = ($null -ne $state.PlanHandle)
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy SQL'; $btnCopy.Width = 90; $btnCopy.Location = New-Object System.Drawing.Point(805, 93)
    $hdr.Controls.AddRange(@($lblT, $lblD, $lblScope, $cboScope, $btnProc, $btnPlan, $btnCopy))

    # SQL text
    $txt = New-Object System.Windows.Forms.RichTextBox
    $txt.Dock = 'Fill'; $txt.ReadOnly = $true; $txt.WordWrap = $false
    $txt.Font = New-Object System.Drawing.Font('Consolas', 10)
    $txt.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 245)
    $txt.Text = $state.StmtText

    # chart: this statement's waits over the whole range, clicked slice shaded
    $c = New-WaitChart
    $c.Titles.Add('Waits for this statement by wait type (whole selected range, clicked slice shaded)') | Out-Null
    $c.Titles[0].Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    # tables
    $t2 = New-Object System.Windows.Forms.TabControl
    $t2.Dock = 'Fill'
    $pW = New-Object System.Windows.Forms.TabPage 'Wait breakdown'
    $pB = New-Object System.Windows.Forms.TabPage 'Who blocked it'
    $pR = New-Object System.Windows.Forms.TabPage 'Who ran it'
    $t2.TabPages.AddRange(@($pW, $pB, $pR))
    $gW = New-Object System.Windows.Forms.DataGridView; Set-GridStyle $gW; $pW.Controls.Add($gW)
    $gB = New-Object System.Windows.Forms.DataGridView; Set-GridStyle $gB; $pB.Controls.Add($gB)
    $gR = New-Object System.Windows.Forms.DataGridView; Set-GridStyle $gR; $pR.Controls.Add($gR)

    $layout.Controls.Add($hdr, 0, 0)
    $layout.Controls.Add($txt, 0, 1)
    $layout.Controls.Add($c,   0, 2)
    $layout.Controls.Add($t2,  0, 3)
    $f.Controls.Add($layout)

    $state.Grids = @{ W = $gW; B = $gB; R = $gR }
    $state.Scope = $cboScope
    $f.Tag = $state

    # chart data (whole range)
    try {
        $buckets = Get-Buckets $state.RangeFrom $state.RangeTo $state.Bucket
        $cd = Invoke-Sql $SqlStmtChart @{ id = $Id; from = $state.RangeFrom; to = $state.RangeTo; b = $state.Bucket; tz = $script:TzMin }
        Set-StackedChart $c $cd $buckets '' ''
        if ($cboScope.SelectedIndex -eq 0) {
            $strip = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine
            $strip.IntervalOffset = $SliceFromUtc.AddMinutes($script:TzMin).ToOADate()
            $strip.StripWidth     = ($SliceToUtc - $SliceFromUtc).TotalDays
            $strip.BackColor      = [System.Drawing.Color]::FromArgb(60, 255, 215, 0)
            $c.ChartAreas[0].AxisX.StripLines.Add($strip)
        }
    } catch { Show-Error 'Could not load the statement chart.' $_ }

    $cboScope.Add_SelectedIndexChanged({ Update-DetailTables $this.FindForm() })
    $btnCopy.Add_Click({ $t = [string]$this.FindForm().Tag.StmtText; if ($t) { [System.Windows.Forms.Clipboard]::SetText($t) } })
    $btnProc.Add_Click({ Show-WholeProcedure $this.FindForm().Tag })
    $btnPlan.Add_Click({ Open-Plan $this.FindForm().Tag })

    Update-DetailTables $f
    $f.Show($form)
}

function Update-DetailTables($DetailForm) {
    $st = $DetailForm.Tag
    if ($st.Scope.SelectedIndex -eq 0) { $from = $st.SliceFrom; $to = $st.SliceTo } else { $from = $st.RangeFrom; $to = $st.RangeTo }
    $p = @{ id = $st.Id; from = $from; to = $to; tz = $script:TzMin }
    try {
        $st.Grids.W.DataSource = Invoke-Sql $SqlStmtWaits $p
        $st.Grids.B.DataSource = Invoke-Sql $SqlStmtBlockers $p
        $st.Grids.R.DataSource = Invoke-Sql $SqlStmtWho $p
    } catch { Show-Error 'Could not load statement details.' $_ }
}

function Show-WholeProcedure($St) {
    $db = $St.DbName -replace '\]', ']]'
    $prefix = "[$db]"
    if ($St.LinkedSrv) { $prefix = "[$($St.LinkedSrv -replace '\]', ']]')].[$db]" }   # read it on the remote server
    try {
        $dt = Invoke-Sql "SELECT definition FROM $prefix.sys.sql_modules WHERE object_id = @oid;" @{ oid = $St.ObjectId }
    } catch { Show-Error 'Could not read the procedure definition.' $_; return }
    if ($dt.Rows.Count -eq 0 -or $dt.Rows[0].definition -is [DBNull]) {
        [void][System.Windows.Forms.MessageBox]::Show('Definition not available (object dropped, encrypted, or no permission).', 'SQL Wait Monitor')
        return
    }
    $pf = New-Object System.Windows.Forms.Form
    $pf.Text = "Procedure - statement highlighted at line $($St.StartLine)"
    $pf.Size = New-Object System.Drawing.Size(1100, 800)
    $pf.StartPosition = 'CenterParent'
    $rt = New-Object System.Windows.Forms.RichTextBox
    $rt.Dock = 'Fill'; $rt.ReadOnly = $true; $rt.WordWrap = $false
    $rt.Font = New-Object System.Drawing.Font('Consolas', 10)
    $rt.Text = [string]$dt.Rows[0].definition
    $pf.Controls.Add($rt)
    $pf.Add_Shown({
        $r = $this.Controls[0]
        $st2 = $this.Tag
        $line = [math]::Max(0, [math]::Min($st2.StartLine - 1, $r.Lines.Length - 1))
        $startIdx = $r.GetFirstCharIndexFromLine($line)
        $firstLine = (($st2.StmtText -split "`n")[0]).TrimEnd("`r").Trim()
        $pos = -1
        if ($firstLine.Length -gt 0) { $pos = $r.Find($firstLine, $startIdx, [System.Windows.Forms.RichTextBoxFinds]::None) }
        if ($pos -ge 0) {
            $len = [math]::Min(($st2.StmtText -replace "`r", '').Length, $r.TextLength - $pos)
            $r.Select($pos, $len)
        } else {
            $r.Select($startIdx, $r.Lines[$line].Length)
        }
        $r.SelectionBackColor = [System.Drawing.Color]::Gold
        $r.SelectionLength = 0
        $r.ScrollToCaret()
    })
    $pf.Tag = $St
    $pf.Show($form)
}

function Open-Plan($St) {
    try {
        if ($St.LinkedSrv) {
            # the plan lives in the remote server's plan cache
            $ls = $St.LinkedSrv -replace '\]', ']]'
            $dt = Invoke-Sql "EXEC (N'SELECT CAST(query_plan AS nvarchar(max)) FROM sys.dm_exec_query_plan(?);', @h) AT [$ls];" @{ h = $St.PlanHandle }
        } else {
            $dt = Invoke-Sql $SqlPlan @{ h = $St.PlanHandle }
        }
    } catch { Show-Error 'Could not read the plan (needs VIEW SERVER STATE; remote servers also need RPC Out).' $_; return }
    if ($dt.Rows.Count -eq 0 -or $dt.Rows[0][0] -is [DBNull]) {
        [void][System.Windows.Forms.MessageBox]::Show('The plan is no longer in the plan cache. Check Query Store for this query.', 'SQL Wait Monitor')
        return
    }
    $path = Join-Path $env:TEMP ("SqlWait_statement_{0}.sqlplan" -f $St.Id)
    [System.IO.File]::WriteAllText($path, [string]$dt.Rows[0][0], [System.Text.Encoding]::Unicode)
    try { Invoke-Item $path } catch { [void][System.Windows.Forms.MessageBox]::Show("Plan saved to $path", 'SQL Wait Monitor') }
}

# ---------------------------------------------------------------- events
$script:mainChart.Add_MouseClick({
    param($sender, $e)
    $hit = $sender.HitTest($e.X, $e.Y)
    if ($null -eq $hit.Series) { return }
    $key = [string]$hit.Series.Tag
    if ($hit.ChartElementType -eq [System.Windows.Forms.DataVisualization.Charting.ChartElementType]::DataPoint) {
        $pt = $hit.Series.Points[$hit.PointIndex]
        $sliceLocal = [DateTime]::FromOADate($pt.XValue)
        $fromU = $sliceLocal.AddMinutes(-$script:TzMin)
        $toU   = $fromU.AddSeconds($script:Bucket)
        $tf    = Get-TimeFormat $script:Bucket
        if ($script:ByStatement) {
            if ($key -ne '0') { Show-StatementDetail ([int]$key) $fromU $toU }
            else {
                Update-TopGrid $fromU $toU $null ("Statements outside the top 10 at {0}" -f $sliceLocal.ToString($tf))
                $tabs.SelectedTab = $tpTop
            }
        } else {
            $w = $(if ($key -eq 'Other') { $null } else { $key })
            $what = $(if ($w) { "waiting on $w" } else { 'with other waits' })
            try { Update-TopGrid $fromU $toU $w ("Statements $what at {0} - {1}  (double-click a row for detail)" -f $sliceLocal.ToString($tf), $sliceLocal.AddSeconds($script:Bucket).ToString($tf)) }
            catch { Show-Error 'Could not load statements.' $_ }
            $tabs.SelectedTab = $tpTop
        }
    } elseif ($hit.ChartElementType -eq [System.Windows.Forms.DataVisualization.Charting.ChartElementType]::LegendItem) {
        if ($script:ByStatement) {
            if ($key -ne '0') { Show-StatementDetail ([int]$key) $script:FromUtc $script:ToUtc }
        } else {
            $w = $(if ($key -eq 'Other') { $null } else { $key })
            try { Update-TopGrid $script:FromUtc $script:ToUtc $w "Statements waiting on $key in the selected range" } catch { Show-Error 'Could not load statements.' $_ }
            $tabs.SelectedTab = $tpTop
        }
    }
})

$script:mainChart.Add_MouseMove({
    param($sender, $e)
    $hit = $sender.HitTest($e.X, $e.Y)
    $onItem = ($hit.ChartElementType -eq [System.Windows.Forms.DataVisualization.Charting.ChartElementType]::DataPoint -or
               $hit.ChartElementType -eq [System.Windows.Forms.DataVisualization.Charting.ChartElementType]::LegendItem)
    $sender.Cursor = $(if ($onItem) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default })
})

$script:gridTop.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $id = [int]$sender.Rows[$e.RowIndex].Cells['StatementId'].Value
    Show-StatementDetail $id $script:GridScope.From $script:GridScope.To
})

$script:gridAlerts.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    [void][System.Windows.Forms.MessageBox]::Show([string]$sender.Rows[$e.RowIndex].Cells['Details'].Value, 'Alert details')
})

$lnkReset.Add_LinkClicked({
    try { Update-TopGrid $script:FromUtc $script:ToUtc $null 'Top statements in the selected range - double-click a row for detail' }
    catch { Show-Error 'Could not load statements.' $_ }
})

$btnRefresh.Add_Click({ Update-All })
$cboRange.Add_SelectedIndexChanged({ if ($form.Visible) { Update-All } })
$cboGroup.Add_SelectedIndexChanged({ if ($form.Visible) { Update-All } })
$cboServer.Add_SelectedIndexChanged({ Set-CurrentServer; if ($form.Visible) { Update-All } })

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000
$timer.Add_Tick({ Update-All })

function Set-AutoRefresh {
    $sec = [int]$script:RefreshOptions[[string]$cboRefresh.SelectedItem]
    $timer.Stop()
    if ($sec -gt 0) { $timer.Interval = $sec * 1000; $timer.Start() }
}
$cboRefresh.Add_SelectedIndexChanged({ if ($form.Visible) { Set-AutoRefresh } })

$form.Add_Shown({
    $split.SplitterDistance = [int]($split.Height * 0.55)
    Update-All
    Set-AutoRefresh
})
$form.Add_FormClosed({ $timer.Stop(); $timer.Dispose() })

[void]$form.ShowDialog()

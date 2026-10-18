#requires -Version 5.1
<#
.SYNOPSIS
    SQL Server Operations Dashboard - PowerShell monitoring and reporting engine.

.DESCRIPTION
    Collects read-only operational information from one or more SQL Server instances
    (standalone or Always On Availability Group nodes) and the Windows hosts they run on,
    then produces a single standalone HTML dashboard built with the PSWriteHTML module.

    The dashboard has eight tabs:
      1. Overview                  - KPIs, health cards, instance and database inventory, OS, services
      2. Always On                 - AG topology, replica connectivity/sync, database sync, queues, listeners
      3. Backups                   - Freshness, compliance (across AG replicas), history, restores, gaps
      4. Capacity & Performance    - CPU history, memory, storage, DB sizes, resource donuts, waits, I/O, log/VLF
      5. Cluster & Windows         - Failover Cluster health, groups, resources, OS, event logs, services, accounts
      6. SQL Errors                - Error log, grouped messages, suspect pages, memory dumps, Windows critical events
      7. Security & Configuration  - DB security, orphaned users, privileged roles, compliance, certs, HA config diff
      8. Investigation             - Agent job duration trends, CHECKDB evidence, long-running / blocking, findings,
                                     collection and permission errors, optional metric history

    READ-ONLY: every SQL statement is a SELECT against DMVs / catalog views / msdb history, or a
    read-only system procedure (xp_readerrorlog, DBCC TRACESTATUS). Nothing is created, altered or
    deleted in any user or system database, and no SQL Server configuration is changed.

.NOTES
    Script location : any folder, on any drive (e.g. C:\Apps\SQLServer-Operations\SQLServer-Operations.ps1)
    Report output   : <script folder>\Output\SQLServer-Operations.html   (override with -OutputFolder)
    History (opt.)  : <script folder>\Output\History\MetricHistory.csv

    Requirements
      - Windows PowerShell 5.1
      - PSWriteHTML module            (Install-Module PSWriteHTML -Scope CurrentUser)
      - Windows-integrated authentication to every SQL Server endpoint
      - SQL permissions: VIEW SERVER STATE, VIEW ANY DEFINITION, read on msdb history tables;
        sysadmin is simplest (error log reading and some security checks need it)
      - Windows permissions for remote CIM/WMI, Event Log and service queries on each host
      - FailoverClusters PowerShell module (RSAT) for cluster sections
      - SQL Server 2016 SP2 or later for every section (VLF counts and CHECKDB properties
        need it); older builds still report, failed sections appear under Investigation.

EXECUTIION

Install-Module PSWriteHTML -Scope CurrentUser   # one time



cd A:\SQLServer-Operations\output
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\SQLServer-Operations.ps1 -SqlInstances 'hlmm', 'abcd' #list all servers with comma


.EXAMPLE
    .\SQLServer-Operations.ps1 -SqlInstances 'SQLNODE1','SQLNODE2'

.EXAMPLE
    .\SQLServer-Operations.ps1 -SqlInstances 'SQLPROD01\INST1' -OutputFolder 'D:\Reports' -EnableMetricHistory $true -TimestampedFileName
#>
[CmdletBinding()]
param(
    # ---------------- SQL Server endpoints and instance identification ----------------
    [string[]] $SqlInstances = @('SQLNODE1', 'SQLNODE2'),
    [string]   $EnvironmentName = 'SQL Server Environment',

    # ---------------- SQL connection and query timeouts (seconds) ----------------
    [int] $ConnectionTimeoutSeconds = 15,
    [int] $QueryTimeoutSeconds = 120,

    # ---------------- Backup age thresholds ----------------
    [int] $FullBackupMaxAgeHours = 168,
    [int] $DiffBackupMaxAgeHours = 26,
    [int] $LogBackupMaxAgeMinutes = 60,

    # ---------------- Disk space thresholds (percent free) ----------------
    [int] $DiskFreeWarningPct = 20,
    [int] $DiskFreeCriticalPct = 10,

    # ---------------- Long-running query and blocking thresholds (seconds) ----------------
    [int] $LongRunningQuerySeconds = 300,
    [int] $BlockingWaitSeconds = 30,

    # ---------------- I/O latency and transaction log thresholds ----------------
    [int] $IoLatencyWarningMs = 20,
    [int] $IoLatencyCriticalMs = 50,
    [int] $VlfWarningCount = 1000,
    [int] $LogUsedWarningPct = 75,

    # ---------------- DBCC CHECKDB history age ----------------
    [int] $CheckDbMaxAgeDays = 7,

    # ---------------- Certificate expiration warning ----------------
    [int] $CertificateExpiryWarningDays = 30,

    # ---------------- Availability Group queue thresholds (KB) ----------------
    [long] $AgSendQueueWarningKB = 10240,
    [long] $AgRedoQueueWarningKB = 10240,

    # ---------------- Monitoring lookback periods ----------------
    [int] $LookbackHours = 24,          # error log, Windows events
    [int] $HistoryLookbackDays = 14,    # backup/restore history, Agent job history
    [int] $CpuHistoryMinutes = 240,     # CPU ring buffer samples (1 per minute)
    [int] $MaxEventsPerLog = 300,

    # ---------------- Windows Event Viewer / Failover Cluster collection ----------------
    [bool] $CollectEventViewer = $true,
    [bool] $CollectClusterInfo = $true,

    # ---------------- Optional metric history retention ----------------
    [bool] $EnableMetricHistory = $false,
    [int]  $MetricHistoryRetentionDays = 30,

    # ---------------- HTML report output location and filename ----------------
    [string] $OutputFolder = '',          # default: an Output folder next to this script
    [string] $ReportFileName = 'SQLServer-Operations.html',
    [switch] $TimestampedFileName
)

# =====================================================================================
#  Initialisation
# =====================================================================================
$ErrorActionPreference = 'Continue'
$script:ReportStart      = Get-Date
$script:CollectionErrors = New-Object System.Collections.Generic.List[object]
$script:Findings         = New-Object System.Collections.Generic.List[object]
$script:NodeMap          = @{}     # SqlInstance -> display label
$script:NodeOrder        = @()     # reachable SqlInstances, in the order given
$script:NodeLabels       = @()     # display labels of reachable nodes
$script:Inv              = [Globalization.CultureInfo]::InvariantCulture

if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
    throw 'The PSWriteHTML module is required. Install it with: Install-Module PSWriteHTML -Scope CurrentUser'
}
Import-Module PSWriteHTML -ErrorAction Stop

# Default output: an "Output" folder beside the .ps1 (works on any drive / any server).
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir) -and $MyInvocation.MyCommand.Path) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = (Get-Location).ProviderPath }
    $OutputFolder = Join-Path $scriptDir 'Output'
}
$qualifier = Split-Path -Path $OutputFolder -Qualifier -ErrorAction SilentlyContinue
if ($qualifier -and $qualifier -match '^[A-Za-z]:$' -and -not (Test-Path -LiteralPath "$qualifier\")) {
    throw "Drive $qualifier is not available on $env:COMPUTERNAME. Pass -OutputFolder with a valid path, or omit it to write next to the script."
}
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

if ($TimestampedFileName) {
    $base = [IO.Path]::GetFileNameWithoutExtension($ReportFileName)
    $ReportFileName = '{0}_{1}.html' -f $base, (Get-Date -Format 'yyyyMMdd_HHmm')
}
$ReportPath = Join-Path $OutputFolder $ReportFileName

# =====================================================================================
#  Generic helpers
# =====================================================================================
function ConvertTo-HtmlSafe {
    param([object] $Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function F {
    # Invariant-culture number for SVG/CSS (avoids "12,5" in comma-decimal locales)
    param([double] $Value, [string] $Format = '0.##')
    return $Value.ToString($Format, $script:Inv)
}

function Format-SizeMB {
    param([double] $MB)
    if ($MB -ge 1048576) { return ('{0:N2} TB' -f ($MB / 1048576)) }
    if ($MB -ge 1024)    { return ('{0:N2} GB' -f ($MB / 1024)) }
    return ('{0:N0} MB' -f $MB)
}

function Format-Duration {
    param([double] $Seconds)
    $ts = [TimeSpan]::FromSeconds([math]::Round($Seconds))
    if ($ts.TotalHours -ge 1)   { return ('{0}h {1:00}m {2:00}s' -f [math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds) }
    if ($ts.TotalMinutes -ge 1) { return ('{0}m {1:00}s' -f $ts.Minutes, $ts.Seconds) }
    return ('{0}s' -f $ts.Seconds)
}

function Format-Age {
    param([object] $Date)
    if ($null -eq $Date) { return 'Never' }
    $span = (Get-Date) - [datetime]$Date
    if ($span.TotalDays -ge 1)  { return ('{0:N1} days' -f $span.TotalDays) }
    if ($span.TotalHours -ge 1) { return ('{0:N1} hours' -f $span.TotalHours) }
    return ('{0:N0} min' -f $span.TotalMinutes)
}

function Get-WorstStatus {
    param([string[]] $Status)
    if ($Status -contains 'Critical') { return 'Critical' }
    if ($Status -contains 'Warning')  { return 'Warning' }
    if ($Status -contains 'OK')       { return 'OK' }
    return 'Info'
}

function Get-StatusClass {
    param([string] $Status)
    switch ($Status) { 'OK' { 'st-ok' } 'Warning' { 'st-warning' } 'Critical' { 'st-critical' } default { 'st-info' } }
}

function Get-StatusColor {
    param([string] $Status)
    switch ($Status) { 'OK' { '#2e9d5b' } 'Warning' { '#e0a100' } 'Critical' { '#d64545' } default { '#3a7bd5' } }
}

function Add-CollectionError {
    param([string] $Node, [string] $Section, [string] $Message)
    $category = 'Query / collection'
    if ($Message -match 'permission|denied|not authorized|login failed|VIEW SERVER STATE|does not have') { $category = 'Permission' }
    elseif ($Message -match 'timeout|network|not found or was not accessible|RPC server|WinRM|unavailable') { $category = 'Connectivity' }
    $script:CollectionErrors.Add([pscustomobject]@{
        Time     = Get-Date
        SqlNode  = $Node
        Section  = $Section
        Category = $category
        Message  = $Message
    })
}

function Add-Finding {
    param(
        [ValidateSet('Critical', 'Warning', 'Info')][string] $Status,
        [string] $Area, [string] $Node, [string] $Item, [string] $Detail
    )
    $script:Findings.Add([pscustomobject]@{
        Status  = $Status
        Area    = $Area
        SqlNode = $Node
        Item    = $Item
        Detail  = $Detail
    })
}

# =====================================================================================
#  SQL Server data collection functions
# =====================================================================================
function Get-ReportNodeLabel {
    <# Returns the display label used for a SQL Server node, e.g. "SQLNODE1 (Primary)". #>
    param([string] $SqlInstance)
    if ($script:NodeMap.ContainsKey($SqlInstance)) { return $script:NodeMap[$SqlInstance] }
    return $SqlInstance
}

function Invoke-DbQuery {
    <# Executes a read-only monitoring query and returns PSCustomObjects tagged with SqlNode. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SqlInstance,
        [Parameter(Mandatory)][string] $Query,
        [string]    $Database = 'master',
        [hashtable] $Parameters,
        [string]    $Section = 'General',
        [int]       $Timeout = $QueryTimeoutSeconds
    )
    $connString = "Server=$SqlInstance;Database=$Database;Integrated Security=SSPI;" +
                  "Application Name=SQLServer-Operations-Dashboard;Connect Timeout=$ConnectionTimeoutSeconds"
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    $table = New-Object System.Data.DataTable
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        if ($Parameters) {
            foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue("@$k", $Parameters[$k]) }
        }
        $reader = $cmd.ExecuteReader()
        $table.Load($reader)
        $reader.Close()
    }
    catch {
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
        Add-CollectionError -Node (Get-ReportNodeLabel $SqlInstance) -Section $Section -Message $msg
        return
    }
    finally { $conn.Dispose() }

    $label   = Get-ReportNodeLabel $SqlInstance
    $columns = @($table.Columns | ForEach-Object { $_.ColumnName })
    foreach ($row in $table.Rows) {
        $o = [ordered]@{ SqlNode = $label }
        foreach ($c in $columns) {
            $v = $row[$c]
            if ($v -is [DBNull]) { $v = $null }
            $o[$c] = $v
        }
        [pscustomobject]$o
    }
}

function Get-DatabaseCategory {
    <# Classifies databases for structured reporting. #>
    param([string] $DatabaseName, [string] $AgName, [bool] $IsDistributor)
    if ($DatabaseName -in @('master', 'model', 'msdb', 'tempdb')) { return 'System databases' }
    if ($IsDistributor)                     { return 'Replication distribution' }
    if ($AgName)                            { return "Availability Group: $AgName" }
    if ($DatabaseName -eq 'SSISDB')         { return 'Integration Services catalog' }
    if ($DatabaseName -like 'ReportServer*'){ return 'Reporting Services' }
    return 'User databases (standalone)'
}

function Select-ReportColumns {
    <# Selects the relevant columns for report presentation. #>
    param([object[]] $Data, [object[]] $Columns, [switch] $ExcludeNode)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return @() }
    $cols = $Columns
    if ($ExcludeNode) { $cols = @($Columns | Where-Object { $_ -ne 'SqlNode' }) }
    return @($rows | Select-Object -Property $cols)
}

function Get-PerDatabaseQuery {
    <# Builds a read-only batch that runs $InnerQuery in every accessible online database.
       Placeholders: {DB} = QUOTENAME(db), {DBNAME} = db name for use inside N'...' literals. #>
    param([string] $TableDefinition, [string] $InnerQuery, [string] $OrderBy = '')
    $tpl = $InnerQuery.Replace("'", "''")
    return @"
SET NOCOUNT ON;
DECLARE @r TABLE ($TableDefinition);
DECLARE @db sysname, @sql nvarchar(max);
DECLARE @tpl nvarchar(max) = N'$tpl';
DECLARE dbs CURSOR LOCAL FAST_FORWARD FOR
    SELECT d.name FROM sys.databases d
    WHERE d.state = 0 AND d.name <> N'tempdb' AND HAS_DBACCESS(d.name) = 1;
OPEN dbs; FETCH NEXT FROM dbs INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = REPLACE(REPLACE(@tpl, N'{DB}', QUOTENAME(@db)), N'{DBNAME}', REPLACE(@db, N'''', N''''''));
    BEGIN TRY INSERT INTO @r EXEC sys.sp_executesql @sql; END TRY BEGIN CATCH END CATCH;
    FETCH NEXT FROM dbs INTO @db;
END
CLOSE dbs; DEALLOCATE dbs;
SELECT * FROM @r $OrderBy;
"@
}

# =====================================================================================
#  HTML reporting functions
# =====================================================================================
$script:StatusPalette = @(
    @{ Values = @('OK', 'Healthy', 'HEALTHY', 'CONNECTED', 'SYNCHRONIZED', 'ONLINE', 'Online', 'Running', 'Up', 'Succeeded', 'Match', 'Present', 'Enabled'); Back = '#d8f3e3'; Fore = '#1e6b3e' }
    @{ Values = @('Warning', 'PARTIALLY_HEALTHY', 'SYNCHRONIZING', 'Retry', 'Canceled', 'Paused', 'Degraded', 'Overdue', 'NOT_HEALTHY_SECONDARY'); Back = '#fff1cc'; Fore = '#7a5b00' }
    @{ Values = @('Critical', 'NOT_HEALTHY', 'DISCONNECTED', 'NOT SYNCHRONIZING', 'NOT_SYNCHRONIZING', 'Stopped', 'Down', 'Failed', 'Missing', 'Mismatch', 'SUSPECT', 'OFFLINE', 'Offline', 'RECOVERY_PENDING', 'EMERGENCY', 'Never'); Back = '#f9d6d6'; Fore = '#8a1f1f' }
    @{ Values = @('Info', 'N/A', 'Unknown'); Back = '#e3edfb'; Fore = '#1f4f8f' }
)

function New-SafeHTMLTable {
    <# Generates a report table; renders a friendly message when there is no data. #>
    param(
        [object[]] $Data,
        [string]   $EmptyMessage = 'No data was returned for this section.',
        [int]      $PagingLength = 10,
        [string[]] $StatusColumns
    )
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>$(ConvertTo-HtmlSafe $EmptyMessage)</div>" }

    if (-not $StatusColumns) {
        $StatusColumns = @($rows[0].PSObject.Properties.Name | Where-Object { $_ -match '(Status|Health|State|Outcome|Compliance|Comparison|Role)$' })
    }
    $present = @{}
    foreach ($c in $StatusColumns) { $present[$c] = @($rows | ForEach-Object { [string]$_.$c } | Sort-Object -Unique) }
    $palette = $script:StatusPalette

    New-HTMLTable -DataTable $rows -Filtering -PagingLength $PagingLength -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') {
        foreach ($c in $StatusColumns) {
            foreach ($p in $palette) {
                foreach ($v in $p.Values) {
                    if ($present[$c] -ccontains $v) {
                        New-TableCondition -Name $c -ComparisonType string -Operator eq -Value $v -BackgroundColor $p.Back -Color $p.Fore
                    }
                }
            }
        }
    }
}

function New-NodeAwarePanels {
    <# Renders the same content block side by side for every SQL Server node. #>
    param([scriptblock] $ContentForNode)
    $labels = @($script:NodeLabels)
    New-HTMLSection -Invisible {
        foreach ($lbl in $labels) {
            New-HTMLPanel {
                "<div class='ops-node-title'>$(ConvertTo-HtmlSafe $lbl)</div>"
                & $ContentForNode $lbl
            }
        }
    }
}

function New-NodeAwareTable {
    <# Organizes information per SQL Server node (side by side in HA environments). #>
    param(
        [object[]] $Data,
        [object[]] $Columns,
        [int]      $PagingLength = 10,
        [string]   $EmptyMessage = 'No data was returned for this node.'
    )
    $rows = @($Data | Where-Object { $null -ne $_ })
    if (@($script:NodeLabels).Count -le 1) {
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $rows -Columns $Columns -ExcludeNode) -PagingLength $PagingLength -EmptyMessage $EmptyMessage
        return
    }
    New-NodeAwarePanels {
        param($lbl)
        $nodeRows = @($rows | Where-Object { $_.SqlNode -eq $lbl })
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $nodeRows -Columns $Columns -ExcludeNode) -PagingLength $PagingLength -EmptyMessage $EmptyMessage
    }
}

function New-DatabaseGroupedTable {
    <# Groups database-related information by database category. #>
    param([object[]] $Data, [object[]] $Columns, [int] $PagingLength = 10)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No database information was collected.</div>" }
    $order = @{ 'System databases' = 0 }
    $groups = @($rows | Group-Object Category | Sort-Object @{ Expression = { if ($order.ContainsKey($_.Name)) { 0 } else { 1 } } }, Name)
    foreach ($g in $groups) {
        $crit = @($g.Group | Where-Object Status -eq 'Critical').Count
        $warn = @($g.Group | Where-Object Status -eq 'Warning').Count
        $st   = Get-WorstStatus @($g.Group | ForEach-Object Status)
        "<div class='ops-group-head'><span class='ops-pill $(Get-StatusClass $st)'>$st</span> $(ConvertTo-HtmlSafe $g.Name) <span class='ops-muted'>($($g.Count) rows, $crit critical, $warn warning)</span></div>"
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $g.Group -Columns $Columns) -PagingLength $PagingLength
    }
}

function New-BackupComplianceGroupedTable {
    <# Presents backup compliance results grouped by scope. #>
    param([object[]] $Data, [object[]] $Columns)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No backup compliance information was collected.</div>" }
    foreach ($g in @($rows | Group-Object Scope | Sort-Object Name)) {
        $ok   = @($g.Group | Where-Object Status -eq 'OK').Count
        $st   = Get-WorstStatus @($g.Group | ForEach-Object Status)
        "<div class='ops-group-head'><span class='ops-pill $(Get-StatusClass $st)'>$st</span> $(ConvertTo-HtmlSafe $g.Name) <span class='ops-muted'>($ok of $($g.Count) compliant)</span></div>"
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $g.Group -Columns $Columns) -PagingLength 15
    }
}

# =====================================================================================
#  Graphical monitoring functions (pure SVG / HTML, no external chart libraries)
# =====================================================================================
$script:ChartPalette = @('#3a7bd5', '#2e9d5b', '#e0a100', '#d64545', '#8e5bd6', '#17a2b8', '#ff7f50', '#6c757d', '#c2185b', '#00897b')

function New-DonutSvgInternal {
    param(
        [object[]] $Items, [string] $LabelProperty, [string] $ValueProperty,
        [string] $Title, [string] $Unit = '', [int] $Top = 8, [scriptblock] $ValueFormatter
    )
    if (-not $ValueFormatter) { $ValueFormatter = { param($v) '{0:N1}' -f $v } }
    $data = @($Items | Where-Object { $null -ne $_ -and [double]$_.$ValueProperty -gt 0 } |
              Sort-Object -Property @{ Expression = { [double]$_.$ValueProperty } } -Descending)
    if ($data.Count -eq 0) { return "<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-empty'>No data available.</div></div>" }
    if ($data.Count -gt $Top) {
        $others = ($data | Select-Object -Skip $Top | Measure-Object -Property $ValueProperty -Sum).Sum
        $data = @($data | Select-Object -First $Top) + @([pscustomobject]@{ $LabelProperty = 'Others'; $ValueProperty = $others })
    }
    $total = [double](($data | Measure-Object -Property $ValueProperty -Sum).Sum)
    $r = 60; $circ = 2 * [math]::PI * $r; $offset = 0.0; $i = 0
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-donut-body'>")
    [void]$sb.Append("<svg viewBox='0 0 160 160' width='160' height='160' role='img'><g transform='rotate(-90 80 80)'>")
    [void]$sb.Append("<circle cx='80' cy='80' r='$r' fill='none' stroke='#eef1f5' stroke-width='22'/>")
    $legend = New-Object System.Text.StringBuilder
    foreach ($d in $data) {
        $v   = [double]$d.$ValueProperty
        $len = $circ * $v / $total
        $col = $script:ChartPalette[$i % $script:ChartPalette.Count]
        $pct = 100 * $v / $total
        $lbl = ConvertTo-HtmlSafe $d.$LabelProperty
        $val = & $ValueFormatter $v
        [void]$sb.Append("<circle cx='80' cy='80' r='$r' fill='none' stroke='$col' stroke-width='22' stroke-dasharray='$(F $len) $(F ($circ - $len))' stroke-dashoffset='$(F (-$offset))'><title>$lbl : $val $Unit ($('{0:N1}' -f $pct)%)</title></circle>")
        [void]$legend.Append("<li><span class='ops-swatch' style='background:$col'></span><span class='ops-legend-label' title='$lbl'>$lbl</span><span class='ops-legend-val'>$val $Unit &middot; $('{0:N1}' -f $pct)%</span></li>")
        $offset += $len; $i++
    }
    [void]$sb.Append("</g><text x='80' y='78' text-anchor='middle' class='ops-donut-total'>$(& $ValueFormatter $total)</text>")
    [void]$sb.Append("<text x='80' y='96' text-anchor='middle' class='ops-donut-unit'>$(ConvertTo-HtmlSafe $Unit) total</text></svg>")
    [void]$sb.Append("<ul class='ops-legend'>$($legend.ToString())</ul></div></div>")
    return $sb.ToString()
}

function ConvertTo-CpuSvg {
    <# Generates the SQL Server / other process CPU utilization chart. #>
    param([object[]] $Samples, [string] $Title = 'CPU utilization history')
    $s = @($Samples | Where-Object { $null -ne $_.EventTime } | Sort-Object EventTime)
    if ($s.Count -lt 2) { return "<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-empty'>Not enough CPU samples available.</div></div>" }
    $w = 760; $h = 230; $l = 40; $rgt = 12; $t = 14; $b = 36
    $pw = $w - $l - $rgt; $ph = $h - $t - $b; $n = $s.Count
    $sqlPts = New-Object System.Collections.Generic.List[string]
    $othPts = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $n; $i++) {
        $x  = $l + ($pw * $i / ($n - 1))
        $ys = $t + $ph * (1 - [double]$s[$i].SQLProcessUtilization / 100)
        $yo = $t + $ph * (1 - [double]$s[$i].OtherProcessUtilization / 100)
        $sqlPts.Add("$(F $x),$(F $ys)"); $othPts.Add("$(F $x),$(F $yo)")
    }
    $avgSql = ($s | Measure-Object -Property SQLProcessUtilization -Average).Average
    $maxSql = ($s | Measure-Object -Property SQLProcessUtilization -Maximum).Maximum
    $avgOth = ($s | Measure-Object -Property OtherProcessUtilization -Average).Average
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div>")
    [void]$sb.Append("<svg viewBox='0 0 $w $h' class='ops-svg-wide' role='img'>")
    foreach ($g in 0, 25, 50, 75, 100) {
        $y = $t + $ph * (1 - $g / 100)
        [void]$sb.Append("<line x1='$l' x2='$($w - $rgt)' y1='$(F $y)' y2='$(F $y)' stroke='#e5e7eb' stroke-width='1'/>")
        [void]$sb.Append("<text x='$($l - 6)' y='$(F ($y + 4))' text-anchor='end' class='ops-axis'>$g%</text>")
    }
    $area = "$(F $l),$(F ($t + $ph)) " + ($sqlPts -join ' ') + " $(F ($l + $pw)),$(F ($t + $ph))"
    [void]$sb.Append("<polygon points='$area' fill='#3a7bd5' fill-opacity='0.15'/>")
    [void]$sb.Append("<polyline points='$($sqlPts -join ' ')' fill='none' stroke='#3a7bd5' stroke-width='2'/>")
    [void]$sb.Append("<polyline points='$($othPts -join ' ')' fill='none' stroke='#e0a100' stroke-width='1.5' stroke-dasharray='4 3'/>")
    foreach ($idx in 0, [int](($n - 1) / 2), ($n - 1)) {
        $x = $l + ($pw * $idx / ($n - 1))
        $anchor = if ($idx -eq 0) { 'start' } elseif ($idx -eq $n - 1) { 'end' } else { 'middle' }
        [void]$sb.Append("<text x='$(F $x)' y='$($h - 14)' text-anchor='$anchor' class='ops-axis'>$(([datetime]$s[$idx].EventTime).ToString('dd-MMM HH:mm'))</text>")
    }
    [void]$sb.Append("</svg><div class='ops-chart-legend'><span><span class='ops-swatch' style='background:#3a7bd5'></span>SQL Server</span>")
    [void]$sb.Append("<span><span class='ops-swatch' style='background:#e0a100'></span>Other processes</span>")
    [void]$sb.Append("<span class='ops-muted'>SQL avg $('{0:N0}' -f $avgSql)% &middot; SQL max $maxSql% &middot; other avg $('{0:N0}' -f $avgOth)% &middot; $n samples</span></div></div>")
    return $sb.ToString()
}

function ConvertTo-WaitDonutSvg {
    <# Generates the SQL Server wait statistics donut chart. #>
    param([object[]] $Waits, [int] $Top = 8)
    New-DonutSvgInternal -Items $Waits -LabelProperty 'WaitType' -ValueProperty 'WaitSeconds' -Title 'Top waits (share of wait time since startup)' -Unit 's' -Top $Top -ValueFormatter { param($v) '{0:N0}' -f $v }
}

function ConvertTo-DbResourcePie {
    <# Generates a per-database resource distribution donut chart. #>
    param([object[]] $Items, [string] $ValueProperty, [string] $Title, [string] $Unit = 'MB', [scriptblock] $ValueFormatter)
    New-DonutSvgInternal -Items $Items -LabelProperty 'DatabaseName' -ValueProperty $ValueProperty -Title $Title -Unit $Unit -Top 7 -ValueFormatter $ValueFormatter
}

function New-BarRowsHtml {
    param([object[]] $Rows)   # each: Label, Segments(@{Value;Color;Title}), Text, Max
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $Rows) {
        [void]$sb.Append("<div class='ops-bar-row'><div class='ops-bar-label' title='$(ConvertTo-HtmlSafe $r.Label)'>$(ConvertTo-HtmlSafe $r.Label)</div><div class='ops-bar-track'>")
        foreach ($seg in $r.Segments) {
            $pct = 0; if ($r.Max -gt 0) { $pct = [math]::Min(100, 100 * [double]$seg.Value / [double]$r.Max) }
            [void]$sb.Append("<div class='ops-bar-fill' style='width:$(F $pct)%;background:$($seg.Color)' title='$(ConvertTo-HtmlSafe $seg.Title)'></div>")
        }
        [void]$sb.Append("</div><div class='ops-bar-value'>$($r.Text)</div></div>")
    }
    return $sb.ToString()
}

function ConvertTo-DatabaseSizesHtml {
    <# Displays database size information (data vs log) as horizontal bars. #>
    param([object[]] $Databases, [int] $Top = 20)
    $rows = @($Databases | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [double]$_.DataMB + [double]$_.LogMB } } -Descending | Select-Object -First $Top)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No database size information.</div>" }
    $max = ($rows | ForEach-Object { [double]$_.DataMB + [double]$_.LogMB } | Measure-Object -Maximum).Maximum
    $bars = foreach ($d in $rows) {
        [pscustomobject]@{
            Label    = $d.DatabaseName
            Max      = $max
            Text     = Format-SizeMB ([double]$d.DataMB + [double]$d.LogMB)
            Segments = @(
                @{ Value = [double]$d.DataMB; Color = '#3a7bd5'; Title = "Data: $(Format-SizeMB $d.DataMB)" }
                @{ Value = [double]$d.LogMB;  Color = '#e0a100'; Title = "Log: $(Format-SizeMB $d.LogMB)" }
            )
        }
    }
    "<div class='ops-chart'><div class='ops-chart-title'>Database sizes (top $Top)</div>" +
    "<div class='ops-chart-legend'><span><span class='ops-swatch' style='background:#3a7bd5'></span>Data</span><span><span class='ops-swatch' style='background:#e0a100'></span>Log</span></div>" +
    (New-BarRowsHtml $bars) + '</div>'
}

function ConvertTo-FileIoLatencyHtml {
    <# Presents database file I/O latency with threshold colouring. #>
    param([object[]] $Files, [int] $Top = 20)
    $rows = @($Files | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [math]::Max([double]$_.AvgReadMs, [double]$_.AvgWriteMs) } } -Descending | Select-Object -First $Top)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No file I/O statistics.</div>" }
    $max = [math]::Max([double]$IoLatencyCriticalMs, (($rows | ForEach-Object { [math]::Max([double]$_.AvgReadMs, [double]$_.AvgWriteMs) }) | Measure-Object -Maximum).Maximum)
    $colorFor = { param($ms) if ($ms -ge $IoLatencyCriticalMs) { '#d64545' } elseif ($ms -ge $IoLatencyWarningMs) { '#e0a100' } else { '#2e9d5b' } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>File I/O latency (avg ms since startup; warning $IoLatencyWarningMs ms, critical $IoLatencyCriticalMs ms)</div>")
    [void]$sb.Append("<table class='ops-mini-table'><tr><th>Database / file</th><th>Type</th><th>Read latency</th><th>Write latency</th></tr>")
    foreach ($f in $rows) {
        $rd = [double]$f.AvgReadMs; $wr = [double]$f.AvgWriteMs
        $rw = F ([math]::Min(100, 100 * $rd / $max)); $ww = F ([math]::Min(100, 100 * $wr / $max))
        [void]$sb.Append("<tr><td title='$(ConvertTo-HtmlSafe $f.PhysicalName)'>$(ConvertTo-HtmlSafe $f.DatabaseName) / $(ConvertTo-HtmlSafe $f.LogicalName)</td><td>$(ConvertTo-HtmlSafe $f.FileType)</td>")
        [void]$sb.Append("<td><div class='ops-bar-track small'><div class='ops-bar-fill' style='width:$rw%;background:$(& $colorFor $rd)'></div></div><span class='ops-num'>$('{0:N1}' -f $rd) ms</span></td>")
        [void]$sb.Append("<td><div class='ops-bar-track small'><div class='ops-bar-fill' style='width:$ww%;background:$(& $colorFor $wr)'></div></div><span class='ops-num'>$('{0:N1}' -f $wr) ms</span></td></tr>")
    }
    [void]$sb.Append('</table></div>')
    return $sb.ToString()
}

function ConvertTo-MemoryClerkBarsHtml {
    <# Visualizes SQL Server memory clerk usage. #>
    param([object[]] $Clerks)
    $rows = @($Clerks | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [double]$_.MemoryMB } } -Descending)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No memory clerk information.</div>" }
    $max = ($rows | Measure-Object -Property MemoryMB -Maximum).Maximum
    $i = 0
    $bars = foreach ($c in $rows) {
        $col = $script:ChartPalette[$i % $script:ChartPalette.Count]; $i++
        [pscustomobject]@{ Label = $c.ClerkType; Max = $max; Text = Format-SizeMB $c.MemoryMB; Segments = @(@{ Value = [double]$c.MemoryMB; Color = $col; Title = "$($c.ClerkType): $(Format-SizeMB $c.MemoryMB)" }) }
    }
    "<div class='ops-chart'><div class='ops-chart-title'>Memory clerks (top $($rows.Count))</div>" + (New-BarRowsHtml $bars) + '</div>'
}

function ConvertTo-JobDurationSparkline {
    param([double[]] $Durations, [int] $Width = 170, [int] $Height = 34)
    $d = @($Durations)
    if ($d.Count -lt 2) { return "<span class='ops-muted'>1 run</span>" }
    $max = ($d | Measure-Object -Maximum).Maximum; if ($max -le 0) { $max = 1 }
    $pts = for ($i = 0; $i -lt $d.Count; $i++) {
        $x = 2 + ($Width - 4) * $i / ($d.Count - 1)
        $y = 2 + ($Height - 4) * (1 - $d[$i] / $max)
        "$(F $x),$(F $y)"
    }
    $lastX = 2 + ($Width - 4); $lastY = 2 + ($Height - 4) * (1 - $d[-1] / $max)
    return "<svg width='$Width' height='$Height' viewBox='0 0 $Width $Height'><polyline points='$($pts -join ' ')' fill='none' stroke='#3a7bd5' stroke-width='1.5'/><circle cx='$(F $lastX)' cy='$(F $lastY)' r='2.5' fill='#d64545'/></svg>"
}

function ConvertTo-JobDurationTrendHtml {
    param([object[]] $Trends)
    $rows = @($Trends | Where-Object { $null -ne $_ } | Sort-Object JobName)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No SQL Server Agent job history in the lookback period.</div>" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table class='ops-mini-table'><tr><th>Job</th><th>Runs</th><th>Avg</th><th>Max</th><th>Last</th><th>Duration trend</th><th>Last outcome</th></tr>")
    foreach ($j in $rows) {
        [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $j.JobName)</td><td class='ops-num'>$($j.Runs)</td><td class='ops-num'>$($j.AvgDuration)</td><td class='ops-num'>$($j.MaxDuration)</td><td class='ops-num'>$($j.LastDuration)</td>")
        [void]$sb.Append("<td>$(ConvertTo-JobDurationSparkline -Durations $j.Durations)</td><td><span class='ops-pill $(Get-StatusClass $j.Status)'>$(ConvertTo-HtmlSafe $j.LastOutcome)</span></td></tr>")
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

# =====================================================================================
#  Operational health functions (KPI tiles and health cards)
# =====================================================================================
function ConvertTo-KpiHtmlText {
    <# Formats an operational indicator tile. #>
    param([string] $Title, [string] $Value, [string] $Subtitle = '', [string] $Status = 'Info')
    "<div class='ops-kpi $(Get-StatusClass $Status)'><div class='ops-kpi-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-kpi-value'>$(ConvertTo-HtmlSafe $Value)</div><div class='ops-kpi-sub'>$(ConvertTo-HtmlSafe $Subtitle)</div></div>"
}

function New-OpsCardHtml {
    param([string] $Title, [string] $Status, [System.Collections.IDictionary] $Rows, [string] $Footer = '')
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-card $(Get-StatusClass $Status)'><div class='ops-card-head'><span>$(ConvertTo-HtmlSafe $Title)</span><span class='ops-pill $(Get-StatusClass $Status)'>$Status</span></div><table class='ops-card-table'>")
    foreach ($k in $Rows.Keys) { [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $k)</td><td>$(ConvertTo-HtmlSafe $Rows[$k])</td></tr>") }
    [void]$sb.Append('</table>')
    if ($Footer) { [void]$sb.Append("<div class='ops-card-foot'>$(ConvertTo-HtmlSafe $Footer)</div>") }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function ConvertTo-AgHealthCard {
    <# Presents Availability Group health. #>
    param([object[]] $Replicas, [object[]] $Databases)
    $rep = @($Replicas | Where-Object { $null -ne $_ })
    if ($rep.Count -eq 0) { return New-OpsCardHtml -Title 'Always On Availability Groups' -Status 'Info' -Rows ([ordered]@{ 'Configuration' = 'No Availability Groups detected' }) }
    $dbs = @($Databases | Where-Object { $null -ne $_ })
    $agCount   = @($rep | Select-Object -ExpandProperty AgName -Unique).Count
    $repOk     = @($rep | Where-Object SynchronizationHealth -eq 'HEALTHY').Count
    $repConn   = @($rep | Where-Object { $_.ConnectedState -eq 'CONNECTED' -or $_.Role -eq 'PRIMARY' }).Count
    $dbOk      = @($dbs | Where-Object SynchronizationHealth -eq 'HEALTHY').Count
    $maxSend   = [double](($dbs | Measure-Object -Property LogSendQueueKB -Maximum).Maximum)
    $maxRedo   = [double](($dbs | Measure-Object -Property RedoQueueKB -Maximum).Maximum)
    $status    = Get-WorstStatus @(@($rep | ForEach-Object Status) + @($dbs | ForEach-Object Status))
    New-OpsCardHtml -Title 'Always On Availability Groups' -Status $status -Rows ([ordered]@{
        'Availability Groups'        = $agCount
        'Replicas healthy'           = "$repOk of $($rep.Count)"
        'Replicas connected'         = "$repConn of $($rep.Count)"
        'Database replicas healthy'  = "$dbOk of $($dbs.Count)"
        'Max log send queue'         = Format-SizeMB ($maxSend / 1024)
        'Max redo queue'             = Format-SizeMB ($maxRedo / 1024)
    })
}

function ConvertTo-BackupHealthCard {
    <# Presents backup health. #>
    param([object[]] $Compliance)
    $c = @($Compliance | Where-Object { $null -ne $_ -and $_.Status -ne 'Info' })
    if ($c.Count -eq 0) { return New-OpsCardHtml -Title 'Backup compliance' -Status 'Info' -Rows ([ordered]@{ 'Databases evaluated' = 0 }) }
    $status = Get-WorstStatus @($c | ForEach-Object Status)
    New-OpsCardHtml -Title 'Backup compliance' -Status $status -Rows ([ordered]@{
        'Databases evaluated'     = $c.Count
        'Compliant'               = @($c | Where-Object Status -eq 'OK').Count
        'Warning (overdue)'       = @($c | Where-Object Status -eq 'Warning').Count
        'Critical (missing)'      = @($c | Where-Object Status -eq 'Critical').Count
        'Thresholds (full/diff/log)' = "$FullBackupMaxAgeHours h / $DiffBackupMaxAgeHours h / $LogBackupMaxAgeMinutes min"
    })
}

function ConvertTo-HostMemoryCard {
    <# Displays Windows host memory information. #>
    param([object] $Memory, [string] $Node)
    if ($null -eq $Memory) { return New-OpsCardHtml -Title "Host memory - $Node" -Status 'Info' -Rows ([ordered]@{ 'Status' = 'Not collected' }) }
    $total = [double]$Memory.HostTotalMB; $avail = [double]$Memory.HostAvailableMB
    $usedPct = 0; if ($total -gt 0) { $usedPct = 100 * ($total - $avail) / $total }
    $status = 'OK'
    if ($Memory.SystemMemoryState -match 'low' -or $usedPct -ge 97) { $status = 'Critical' } elseif ($usedPct -ge 92) { $status = 'Warning' }
    New-OpsCardHtml -Title "Host memory - $Node" -Status $status -Rows ([ordered]@{
        'Physical memory'   = Format-SizeMB $total
        'Available'         = Format-SizeMB $avail
        'Used'              = ('{0:N1}%' -f $usedPct)
        'Page file free'    = "$(Format-SizeMB $Memory.PageFileAvailableMB) of $(Format-SizeMB $Memory.PageFileTotalMB)"
        'OS memory state'   = $Memory.SystemMemoryState
    })
}

function ConvertTo-SqlProcessMemoryCard {
    <# Displays SQL Server process memory information. #>
    param([object] $Memory, [string] $Node)
    if ($null -eq $Memory) { return New-OpsCardHtml -Title "SQL Server memory - $Node" -Status 'Info' -Rows ([ordered]@{ 'Status' = 'Not collected' }) }
    $status = 'OK'
    if ($Memory.ProcessPhysicalMemoryLow -eq $true -or $Memory.ProcessVirtualMemoryLow -eq $true) { $status = 'Critical' }
    elseif ([double]$Memory.PageLifeExpectancy -lt 300) { $status = 'Warning' }
    $maxMem = [double]$Memory.MaxServerMemoryMB
    $maxText = if ($maxMem -ge 2147483647) { 'Unlimited (default)' } else { Format-SizeMB $maxMem }
    if ($maxMem -ge 2147483647 -and $status -eq 'OK') { $status = 'Warning' }
    New-OpsCardHtml -Title "SQL Server memory - $Node" -Status $status -Rows ([ordered]@{
        'Physical memory in use' = Format-SizeMB $Memory.SqlPhysicalMemoryMB
        'Total / target server'  = "$(Format-SizeMB $Memory.TotalServerMemoryMB) / $(Format-SizeMB $Memory.TargetServerMemoryMB)"
        'Max server memory'      = $maxText
        'Locked pages'           = Format-SizeMB $Memory.LockedPagesMB
        'Page life expectancy'   = ('{0:N0} s' -f [double]$Memory.PageLifeExpectancy)
        'Memory utilization'     = "$($Memory.MemoryUtilizationPct)%"
    })
}

# =====================================================================================
#  Windows / cluster collection
# =====================================================================================
function Get-WindowsHostData {
    param([string] $ComputerName, [string] $Label)
    $r = @{ OS = @(); Volumes = @(); Services = @(); LocalUsers = @(); Admins = @(); Events = @() }
    $session = $null
    try { $session = New-CimSession -ComputerName $ComputerName -OperationTimeoutSec 30 -ErrorAction Stop }
    catch {
        try { $session = New-CimSession -ComputerName $ComputerName -SessionOption (New-CimSessionOption -Protocol Dcom) -OperationTimeoutSec 30 -ErrorAction Stop }
        catch { Add-CollectionError -Node $Label -Section 'Windows (CIM session)' -Message $_.Exception.Message }
    }
    if ($session) {
        try {
            $os  = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem -ErrorAction Stop
            $cs  = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem -ErrorAction Stop
            $cpu = @(Get-CimInstance -CimSession $session -ClassName Win32_Processor -ErrorAction Stop)[0]
            $r.OS = @([pscustomobject]@{
                SqlNode         = $Label
                ComputerName    = $cs.Name
                Domain          = $cs.Domain
                OperatingSystem = $os.Caption
                Version         = $os.Version
                Build           = $os.BuildNumber
                Architecture    = $os.OSArchitecture
                Manufacturer    = $cs.Manufacturer
                Model           = $cs.Model
                Processor       = $cpu.Name
                Sockets         = $cs.NumberOfProcessors
                LogicalCPUs     = $cs.NumberOfLogicalProcessors
                MemoryGB        = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
                FreeMemoryGB    = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
                LastBoot        = $os.LastBootUpTime
                UptimeDays      = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
                InstallDate     = $os.InstallDate
            })
        } catch { Add-CollectionError -Node $Label -Section 'Windows operating system' -Message $_.Exception.Message }

        try {
            $r.Volumes = @(Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType=3' -ErrorAction Stop |
                Where-Object { $_.Capacity -gt 0 -and $_.Name -notlike '\\?\*' } | ForEach-Object {
                    $freePct = [math]::Round(100 * $_.FreeSpace / $_.Capacity, 1)
                    $st = if ($freePct -lt $DiskFreeCriticalPct) { 'Critical' } elseif ($freePct -lt $DiskFreeWarningPct) { 'Warning' } else { 'OK' }
                    [pscustomobject]@{
                        SqlNode = $Label; Volume = $_.Name; Label = $_.Label; FileSystem = $_.FileSystem
                        CapacityGB = [math]::Round($_.Capacity / 1GB, 2); FreeGB = [math]::Round($_.FreeSpace / 1GB, 2)
                        FreePct = $freePct; BlockSizeKB = [math]::Round($_.BlockSize / 1KB, 0); Status = $st
                    }
                })
        } catch { Add-CollectionError -Node $Label -Section 'Windows volumes' -Message $_.Exception.Message }

        try {
            $pattern = '^(MSSQL|SQLAgent|SQLSERVERAGENT|SQLBrowser|SQLWriter|MSSQLFDLauncher|SQLTELEMETRY|SSISTELEMETRY|MsDtsServer|ReportServer|SQLServerReportingServices|PowerBIReportServer|MSSQLLaunchpad|MSOLAP|ClusSvc|MSDTC)'
            $r.Services = @(Get-CimInstance -CimSession $session -ClassName Win32_Service -ErrorAction Stop |
                Where-Object { $_.Name -match $pattern } | ForEach-Object {
                    $st = 'OK'
                    if ($_.State -ne 'Running') { $st = if ($_.StartMode -eq 'Auto') { 'Critical' } else { 'Info' } }
                    [pscustomobject]@{
                        SqlNode = $Label; ServiceName = $_.Name; DisplayName = $_.DisplayName
                        State = $_.State; StartMode = $_.StartMode; ServiceAccount = $_.StartName; Status = $st
                    }
                })
        } catch { Add-CollectionError -Node $Label -Section 'Windows services' -Message $_.Exception.Message }

        try {
            $r.LocalUsers = @(Get-CimInstance -CimSession $session -ClassName Win32_UserAccount -Filter 'LocalAccount=True' -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    SqlNode = $Label; AccountName = $_.Name; FullName = $_.FullName; Disabled = $_.Disabled
                    LockedOut = $_.Lockout; PasswordRequired = $_.PasswordRequired; PasswordExpires = $_.PasswordExpires; SID = $_.SID
                    Status = if (-not $_.Disabled -and $_.SID -like '*-500') { 'Warning' } elseif (-not $_.Disabled -and -not $_.PasswordRequired) { 'Warning' } else { 'OK' }
                }
            })
        } catch { Add-CollectionError -Node $Label -Section 'Windows local accounts' -Message $_.Exception.Message }

        try {
            $grp = Get-CimInstance -CimSession $session -ClassName Win32_Group -Filter "LocalAccount=True AND SID='S-1-5-32-544'" -ErrorAction Stop
            if ($grp) {
                $r.Admins = @(Get-CimAssociatedInstance -CimSession $session -InputObject $grp -Association Win32_GroupUser -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{
                        SqlNode = $Label; Group = $grp.Name; Member = "$($_.Domain)\$($_.Name)"
                        MemberType = $_.CimClass.CimClassName -replace '^Win32_', ''; SID = $_.SID
                    }
                })
            }
        } catch { Add-CollectionError -Node $Label -Section 'Windows administrators' -Message $_.Exception.Message }

        Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
    }

    if ($CollectEventViewer) {
        $start = (Get-Date).AddHours(-$LookbackHours)
        foreach ($log in 'System', 'Application') {
            try {
                $r.Events += @(Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{ LogName = $log; Level = 1, 2, 3; StartTime = $start } -MaxEvents $MaxEventsPerLog -ErrorAction Stop | ForEach-Object {
                    $msg = [string]$_.Message
                    if ($msg.Length -gt 600) { $msg = $msg.Substring(0, 600) + '...' }
                    [pscustomobject]@{
                        SqlNode = $Label; TimeCreated = $_.TimeCreated; LogName = $_.LogName
                        Level = switch ($_.Level) { 1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } default { [string]$_.LevelDisplayName } }
                        Provider = $_.ProviderName; EventId = $_.Id; Message = ($msg -replace '\s+', ' ')
                    }
                })
            } catch {
                if ($_.Exception.Message -notmatch 'No events were found') {
                    Add-CollectionError -Node $Label -Section "Event Viewer ($log)" -Message $_.Exception.Message
                }
            }
        }
    }
    return $r
}

function Get-ClusterData {
    param([string[]] $ComputerNames)
    $out = @{ Clusters = @(); Nodes = @(); Groups = @(); Resources = @() }
    if (-not $CollectClusterInfo) { return $out }
    if (-not (Get-Command Get-ClusterNode -ErrorAction SilentlyContinue)) {
        try { Import-Module FailoverClusters -ErrorAction Stop } catch { }
    }
    if (-not (Get-Command Get-ClusterNode -ErrorAction SilentlyContinue)) {
        if ($script:IsHA) { Add-CollectionError -Node 'All' -Section 'Failover Cluster' -Message 'FailoverClusters PowerShell cmdlets are not available on this machine (install RSAT Failover Clustering Tools).' }
        return $out
    }
    $seen = @{}
    foreach ($cn in $ComputerNames) {
        try { $cl = Get-Cluster -Name $cn -ErrorAction Stop } catch { continue }
        if ($seen.ContainsKey($cl.Name)) { continue }
        $seen[$cl.Name] = $true
        try {
            $q = Get-ClusterQuorum -Cluster $cl.Name -ErrorAction SilentlyContinue
            $out.Clusters += [pscustomobject]@{
                ClusterName = $cl.Name; Domain = $cl.Domain
                QuorumType = [string]$q.QuorumType; QuorumResource = [string]$q.QuorumResource
                SameSubnetThreshold = $cl.SameSubnetThreshold; CrossSubnetThreshold = $cl.CrossSubnetThreshold
            }
            $out.Nodes += @(Get-ClusterNode -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; NodeName = $_.Name; State = [string]$_.State; NodeWeight = $_.NodeWeight
                    DynamicWeight = $_.DynamicWeight; Status = if ([string]$_.State -eq 'Up') { 'OK' } elseif ([string]$_.State -eq 'Paused') { 'Warning' } else { 'Critical' }
                }
            })
            $out.Groups += @(Get-ClusterGroup -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; GroupName = $_.Name; OwnerNode = [string]$_.OwnerNode; State = [string]$_.State
                    GroupType = [string]$_.GroupType; Status = if ([string]$_.State -eq 'Online') { 'OK' } elseif ([string]$_.State -eq 'PartialOnline') { 'Warning' } else { 'Critical' }
                }
            })
            $out.Resources += @(Get-ClusterResource -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; ResourceName = $_.Name; ResourceType = [string]$_.ResourceType
                    OwnerGroup = [string]$_.OwnerGroup; OwnerNode = [string]$_.OwnerNode; State = [string]$_.State
                    Status = if ([string]$_.State -eq 'Online') { 'OK' } elseif ([string]$_.State -eq 'Offline') { 'Warning' } else { 'Critical' }
                }
            })
        } catch { Add-CollectionError -Node $cl.Name -Section 'Failover Cluster' -Message $_.Exception.Message }
    }
    return $out
}

# =====================================================================================
#  Monitoring queries (all read-only)
# =====================================================================================
$Queries = @{}

$Queries.Discovery = @'
SELECT
    CAST(SERVERPROPERTY('ServerName') AS nvarchar(256))                  AS ServerName,
    CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS nvarchar(256)) AS PhysicalHost,
    CAST(SERVERPROPERTY('Edition') AS nvarchar(256))                     AS Edition,
    CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64))               AS ProductVersion,
    CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(64))                 AS ProductLevel,
    CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(64))           AS ProductUpdateLevel,
    CAST(SERVERPROPERTY('Collation') AS nvarchar(128))                   AS Collation,
    CAST(ISNULL(SERVERPROPERTY('IsClustered'), 0) AS int)                AS IsClustered,
    CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS int)              AS IsHadrEnabled,
    CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int)              AS WindowsAuthOnly,
    si.sqlserver_start_time                                              AS StartTime,
    si.cpu_count                                                         AS CpuCount,
    si.physical_memory_kb / 1024                                         AS HostMemoryMB,
    (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'max server memory (MB)')        AS MaxServerMemoryMB,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'max degree of parallelism')        AS MaxDop,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'cost threshold for parallelism')   AS CostThreshold,
    (SELECT COUNT(*) FROM sys.availability_groups)                                                            AS AgCount,
    (SELECT COUNT(*) FROM sys.dm_hadr_availability_replica_states WHERE is_local = 1 AND role = 1)            AS LocalPrimaryCount,
    (SELECT COUNT(*) FROM sys.dm_hadr_availability_replica_states WHERE is_local = 1 AND role = 2)            AS LocalSecondaryCount
FROM sys.dm_os_sys_info si;
'@

$Queries.Databases = @'
SELECT d.name AS DatabaseName, d.database_id AS DatabaseId, d.state_desc AS State, d.recovery_model_desc AS RecoveryModel,
       d.compatibility_level AS CompatLevel, SUSER_SNAME(d.owner_sid) AS Owner, d.user_access_desc AS UserAccess,
       d.is_read_only AS IsReadOnly, d.page_verify_option_desc AS PageVerify, d.is_auto_close_on AS AutoClose,
       d.is_auto_shrink_on AS AutoShrink, d.create_date AS Created, d.collation_name AS Collation,
       d.is_distributor AS IsDistributor, d.log_reuse_wait_desc AS LogReuseWait, ag.name AS AgName,
       CAST(sz.DataMB AS decimal(18,2)) AS DataMB, CAST(sz.LogMB AS decimal(18,2)) AS LogMB
FROM sys.databases d
LEFT JOIN sys.availability_replicas ar ON ar.replica_id = d.replica_id
LEFT JOIN sys.availability_groups ag  ON ag.group_id = ar.group_id
OUTER APPLY (SELECT SUM(CASE WHEN mf.type = 0 THEN CAST(mf.size AS bigint) ELSE 0 END) * 8 / 1024.0 AS DataMB,
                    SUM(CASE WHEN mf.type = 1 THEN CAST(mf.size AS bigint) ELSE 0 END) * 8 / 1024.0 AS LogMB
             FROM sys.master_files mf WHERE mf.database_id = d.database_id) sz
ORDER BY d.name;
'@

$Queries.AgReplicas = @'
SELECT ag.name AS AgName, ar.replica_server_name AS ReplicaServer, ISNULL(ars.role_desc, 'UNKNOWN') AS Role,
       ar.availability_mode_desc AS AvailabilityMode, ar.failover_mode_desc AS FailoverMode, ar.seeding_mode_desc AS SeedingMode,
       ar.secondary_role_allow_connections_desc AS ReadableSecondary, ar.endpoint_url AS EndpointUrl,
       ar.backup_priority AS BackupPriority, ag.automated_backup_preference_desc AS BackupPreference, ar.session_timeout AS SessionTimeout,
       ars.connected_state_desc AS ConnectedState, ars.operational_state_desc AS OperationalState,
       ars.recovery_health_desc AS RecoveryHealth, ars.synchronization_health_desc AS SynchronizationHealth,
       ars.last_connect_error_number AS LastConnectError, ars.last_connect_error_description AS LastConnectErrorDescription,
       ars.last_connect_error_timestamp AS LastConnectErrorTime
FROM sys.availability_groups ag
JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states ars ON ars.replica_id = ar.replica_id
ORDER BY ag.name, ar.replica_server_name;
'@

$Queries.AgDatabases = @'
SELECT ag.name AS AgName, ar.replica_server_name AS ReplicaServer, DB_NAME(drs.database_id) AS DatabaseName,
       drs.is_local AS IsLocal, drs.is_primary_replica AS IsPrimaryReplica,
       drs.synchronization_state_desc AS SynchronizationState, drs.synchronization_health_desc AS SynchronizationHealth,
       drs.database_state_desc AS DatabaseState, drs.is_suspended AS IsSuspended, drs.suspend_reason_desc AS SuspendReason,
       drs.log_send_queue_size AS LogSendQueueKB, drs.log_send_rate AS LogSendRateKBs,
       drs.redo_queue_size AS RedoQueueKB, drs.redo_rate AS RedoRateKBs,
       drs.last_commit_time AS LastCommitTime, drs.last_hardened_time AS LastHardenedTime, drs.last_redone_time AS LastRedoneTime
FROM sys.dm_hadr_database_replica_states drs
JOIN sys.availability_replicas ar ON ar.replica_id = drs.replica_id
JOIN sys.availability_groups ag  ON ag.group_id = drs.group_id
ORDER BY ag.name, DatabaseName, ar.replica_server_name;
'@

$Queries.AgListeners = @'
SELECT ag.name AS AgName, l.dns_name AS ListenerName, l.port AS Port, lip.ip_address AS IpAddress,
       lip.ip_subnet_mask AS SubnetMask, lip.is_dhcp AS IsDhcp, lip.state_desc AS State
FROM sys.availability_group_listeners l
JOIN sys.availability_groups ag ON ag.group_id = l.group_id
LEFT JOIN sys.availability_group_listener_ip_addresses lip ON lip.listener_id = l.listener_id
ORDER BY ag.name, l.dns_name;
'@

$Queries.BackupStatus = @'
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel, d.state_desc AS State, ag.name AS AgName,
       CASE WHEN d.replica_id IS NULL THEN NULL WHEN sys.fn_hadr_is_primary_replica(d.name) = 1 THEN 'PRIMARY' ELSE 'SECONDARY' END AS LocalRole,
       b.LastFull, b.LastDiff, b.LastLog
FROM sys.databases d
LEFT JOIN sys.availability_replicas ar ON ar.replica_id = d.replica_id
LEFT JOIN sys.availability_groups ag  ON ag.group_id = ar.group_id
OUTER APPLY (SELECT MAX(CASE WHEN bs.type = 'D' THEN bs.backup_finish_date END) AS LastFull,
                    MAX(CASE WHEN bs.type = 'I' THEN bs.backup_finish_date END) AS LastDiff,
                    MAX(CASE WHEN bs.type = 'L' THEN bs.backup_finish_date END) AS LastLog
             FROM msdb.dbo.backupset bs WHERE bs.database_name = d.name) b
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.BackupHistory = @'
SELECT TOP (500) bs.database_name AS DatabaseName,
       CASE bs.type WHEN 'D' THEN 'Full' WHEN 'I' THEN 'Differential' WHEN 'L' THEN 'Log' WHEN 'F' THEN 'File' ELSE bs.type END AS BackupType,
       bs.is_copy_only AS CopyOnly, bs.backup_start_date AS Started, bs.backup_finish_date AS Finished,
       DATEDIFF(second, bs.backup_start_date, bs.backup_finish_date) AS DurationSec,
       CAST(bs.backup_size / 1048576.0 AS decimal(18,2)) AS SizeMB,
       CAST(bs.compressed_backup_size / 1048576.0 AS decimal(18,2)) AS CompressedMB,
       bmf.physical_device_name AS Device, bs.user_name AS ExecutedBy
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE bs.backup_finish_date >= DATEADD(day, -@days, GETDATE())
ORDER BY bs.backup_finish_date DESC;
'@

$Queries.RestoreHistory = @'
SELECT TOP (200) rh.destination_database_name AS DatabaseName, rh.restore_date AS RestoreDate,
       CASE rh.restore_type WHEN 'D' THEN 'Database' WHEN 'I' THEN 'Differential' WHEN 'L' THEN 'Log' WHEN 'F' THEN 'File'
                            WHEN 'G' THEN 'Filegroup' WHEN 'V' THEN 'Verify only' ELSE rh.restore_type END AS RestoreType,
       rh.user_name AS RestoredBy, rh.[replace] AS WithReplace, rh.[recovery] AS WithRecovery,
       bs.database_name AS SourceDatabase, bs.server_name AS SourceServer, bs.backup_finish_date AS SourceBackupDate,
       bmf.physical_device_name AS Device
FROM msdb.dbo.restorehistory rh
LEFT JOIN msdb.dbo.backupset bs ON bs.backup_set_id = rh.backup_set_id
LEFT JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE rh.restore_date >= DATEADD(day, -@days, GETDATE())
ORDER BY rh.restore_date DESC;
'@

$Queries.CpuHistory = @'
DECLARE @ts_now bigint = (SELECT cpu_ticks / (cpu_ticks / ms_ticks) FROM sys.dm_os_sys_info);
SELECT TOP (@minutes)
       DATEADD(ms, -1 * (@ts_now - y.[timestamp]), GETDATE()) AS EventTime,
       y.SQLProcessUtilization,
       y.SystemIdle,
       CASE WHEN 100 - y.SystemIdle - y.SQLProcessUtilization < 0 THEN 0 ELSE 100 - y.SystemIdle - y.SQLProcessUtilization END AS OtherProcessUtilization
FROM (SELECT x.[timestamp],
             x.record.value('(./Record/@id)[1]', 'int') AS record_id,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS SystemIdle,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS SQLProcessUtilization
      FROM (SELECT [timestamp], CONVERT(xml, record) AS record
            FROM sys.dm_os_ring_buffers
            WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR' AND record LIKE N'%<SystemHealth>%') x) y
ORDER BY y.record_id DESC;
'@

$Queries.Memory = @'
SELECT pm.physical_memory_in_use_kb / 1024 AS SqlPhysicalMemoryMB, pm.locked_page_allocations_kb / 1024 AS LockedPagesMB,
       pm.memory_utilization_percentage AS MemoryUtilizationPct,
       CAST(pm.process_physical_memory_low AS bit) AS ProcessPhysicalMemoryLow, CAST(pm.process_virtual_memory_low AS bit) AS ProcessVirtualMemoryLow,
       sm.total_physical_memory_kb / 1024 AS HostTotalMB, sm.available_physical_memory_kb / 1024 AS HostAvailableMB,
       sm.total_page_file_kb / 1024 AS PageFileTotalMB, sm.available_page_file_kb / 1024 AS PageFileAvailableMB,
       sm.system_memory_state_desc AS SystemMemoryState,
       (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'max server memory (MB)') AS MaxServerMemoryMB,
       (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'min server memory (MB)') AS MinServerMemoryMB,
       (SELECT TOP 1 cntr_value / 1024 FROM sys.dm_os_performance_counters WHERE counter_name = 'Target Server Memory (KB)') AS TargetServerMemoryMB,
       (SELECT TOP 1 cntr_value / 1024 FROM sys.dm_os_performance_counters WHERE counter_name = 'Total Server Memory (KB)')  AS TotalServerMemoryMB,
       (SELECT TOP 1 cntr_value FROM sys.dm_os_performance_counters WHERE counter_name = 'Page life expectancy' AND object_name LIKE '%Buffer Manager%') AS PageLifeExpectancy
FROM sys.dm_os_process_memory pm CROSS JOIN sys.dm_os_sys_memory sm;
'@

$Queries.DatabaseFiles = @'
SELECT DB_NAME(mf.database_id) AS DatabaseName, mf.name AS LogicalName, mf.type_desc AS FileType, mf.physical_name AS PhysicalName,
       CAST(CAST(mf.size AS bigint) * 8 / 1024.0 AS decimal(18,2)) AS SizeMB,
       CASE mf.max_size WHEN -1 THEN 'Unlimited' WHEN 0 THEN 'No growth' ELSE CAST(CAST(mf.max_size AS bigint) * 8 / 1024 AS varchar(20)) + ' MB' END AS MaxSize,
       CASE mf.is_percent_growth WHEN 1 THEN CAST(mf.growth AS varchar(10)) + ' %' ELSE CAST(CAST(mf.growth AS bigint) * 8 / 1024 AS varchar(20)) + ' MB' END AS Growth,
       mf.state_desc AS FileState
FROM sys.master_files mf
ORDER BY DatabaseName, mf.type, mf.file_id;
'@

$Queries.SqlVolumes = @'
SELECT DISTINCT vs.volume_mount_point AS VolumePath, vs.logical_volume_name AS VolumeLabel, vs.file_system_type AS FileSystem,
       CAST(vs.total_bytes / 1073741824.0 AS decimal(18,2)) AS TotalGB,
       CAST(vs.available_bytes / 1073741824.0 AS decimal(18,2)) AS FreeGB
FROM sys.master_files mf
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs;
'@

$Queries.BufferCache = @'
SELECT CASE database_id WHEN 32767 THEN 'ResourceDb' ELSE DB_NAME(database_id) END AS DatabaseName,
       CAST(COUNT_BIG(*) * 8 / 1024.0 AS decimal(18,2)) AS BufferCacheMB
FROM sys.dm_os_buffer_descriptors
GROUP BY database_id
ORDER BY BufferCacheMB DESC;
'@

$Queries.DbIo = @'
SELECT DB_NAME(vfs.database_id) AS DatabaseName,
       CAST(SUM(vfs.num_of_bytes_read + vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS TotalIoMB,
       CAST(SUM(vfs.num_of_bytes_read) / 1048576.0 AS decimal(18,2)) AS ReadMB,
       CAST(SUM(vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS WrittenMB,
       SUM(vfs.num_of_reads) AS Reads, SUM(vfs.num_of_writes) AS Writes, SUM(vfs.io_stall) AS IoStallMs
FROM sys.dm_io_virtual_file_stats(NULL, NULL) vfs
GROUP BY vfs.database_id
ORDER BY TotalIoMB DESC;
'@

$Queries.DbCpu = @'
WITH DbCpu AS (
    SELECT pa.DatabaseID, SUM(qs.total_worker_time / 1000) AS CpuMs
    FROM sys.dm_exec_query_stats qs
    CROSS APPLY (SELECT CONVERT(int, value) AS DatabaseID FROM sys.dm_exec_plan_attributes(qs.plan_handle) WHERE attribute = N'dbid') pa
    GROUP BY pa.DatabaseID)
SELECT CASE WHEN DatabaseID = 32767 THEN 'ResourceDb' ELSE DB_NAME(DatabaseID) END AS DatabaseName, CpuMs,
       CAST(CpuMs * 100.0 / NULLIF(SUM(CpuMs) OVER (), 0) AS decimal(5,2)) AS CpuPct
FROM DbCpu
WHERE DatabaseID IS NOT NULL AND (DatabaseID = 32767 OR DB_NAME(DatabaseID) IS NOT NULL)
ORDER BY CpuMs DESC;
'@

$Queries.WaitStats = @'
WITH w AS (
    SELECT wait_type, wait_time_ms / 1000.0 AS WaitS, (wait_time_ms - signal_wait_time_ms) / 1000.0 AS ResourceS,
           signal_wait_time_ms / 1000.0 AS SignalS, waiting_tasks_count AS WaitCount
    FROM sys.dm_os_wait_stats
    WHERE waiting_tasks_count > 0
      AND wait_type NOT LIKE N'SLEEP%' AND wait_type NOT LIKE N'PREEMPTIVE_XE%' AND wait_type NOT LIKE N'QDS%'
      AND wait_type NOT IN (
        N'BROKER_EVENTHANDLER', N'BROKER_RECEIVE_WAITFOR', N'BROKER_TASK_STOP', N'BROKER_TO_FLUSH', N'BROKER_TRANSMITTER',
        N'CHECKPOINT_QUEUE', N'CHKPT', N'CLR_AUTO_EVENT', N'CLR_MANUAL_EVENT', N'CLR_SEMAPHORE', N'CXCONSUMER',
        N'DBMIRROR_DBM_EVENT', N'DBMIRROR_EVENTS_QUEUE', N'DBMIRROR_WORKER_QUEUE', N'DBMIRRORING_CMD', N'DIRTY_PAGE_POLL',
        N'DISPATCHER_QUEUE_SEMAPHORE', N'EXECSYNC', N'FSAGENT', N'FT_IFTS_SCHEDULER_IDLE_WAIT', N'FT_IFTSHC_MUTEX',
        N'HADR_CLUSAPI_CALL', N'HADR_FILESTREAM_IOMGR_IOCOMPLETION', N'HADR_LOGCAPTURE_WAIT', N'HADR_NOTIFICATION_DEQUEUE',
        N'HADR_TIMER_TASK', N'HADR_WORK_QUEUE', N'KSOURCE_WAKEUP', N'LAZYWRITER_SLEEP', N'LOGMGR_QUEUE',
        N'MEMORY_ALLOCATION_EXT', N'ONDEMAND_TASK_QUEUE', N'PARALLEL_REDO_DRAIN_WORKER', N'PARALLEL_REDO_LOG_CACHE',
        N'PARALLEL_REDO_TRAN_LIST', N'PARALLEL_REDO_WORKER_SYNC', N'PARALLEL_REDO_WORKER_WAIT_WORK',
        N'PREEMPTIVE_OS_FLUSHFILEBUFFERS', N'PVS_PREALLOCATE', N'PWAIT_ALL_COMPONENTS_INITIALIZED',
        N'PWAIT_DIRECTLOGCONSUMER_GETNEXT', N'PWAIT_EXTENSIBILITY_CLEANUP_TASK', N'REDO_THREAD_PENDING_WORK',
        N'REQUEST_FOR_DEADLOCK_SEARCH', N'RESOURCE_QUEUE', N'SERVER_IDLE_CHECK', N'SNI_HTTP_ACCEPT', N'SOS_WORK_DISPATCHER',
        N'SP_SERVER_DIAGNOSTICS_SLEEP', N'SQLTRACE_BUFFER_FLUSH', N'SQLTRACE_INCREMENTAL_FLUSH_SLEEP', N'SQLTRACE_WAIT_ENTRIES',
        N'STARTUP_DEPENDENCY_MANAGER', N'UCS_SESSION_REGISTRATION', N'VDI_CLIENT_OTHER', N'WAIT_FOR_RESULTS', N'WAITFOR',
        N'WAITFOR_TASKSHUTDOWN', N'WAIT_XTP_RECOVERY', N'WAIT_XTP_HOST_WAIT', N'WAIT_XTP_OFFLINE_CKPT_NEW_LOG',
        N'WAIT_XTP_CKPT_CLOSE', N'XE_DISPATCHER_JOIN', N'XE_DISPATCHER_WAIT', N'XE_TIMER_EVENT', N'XE_LIVE_TARGET_TVF'))
SELECT TOP (15) wait_type AS WaitType, CAST(WaitS AS decimal(18,2)) AS WaitSeconds, CAST(ResourceS AS decimal(18,2)) AS ResourceSeconds,
       CAST(SignalS AS decimal(18,2)) AS SignalSeconds, WaitCount,
       CAST(100.0 * WaitS / SUM(WaitS) OVER () AS decimal(5,2)) AS Percentage,
       CAST(WaitS * 1000 / WaitCount AS decimal(18,2)) AS AvgWaitMs
FROM w
ORDER BY WaitS DESC;
'@

$Queries.FileLatency = @'
SELECT DB_NAME(vfs.database_id) AS DatabaseName, mf.name AS LogicalName, mf.type_desc AS FileType, mf.physical_name AS PhysicalName,
       vfs.num_of_reads AS Reads, vfs.num_of_writes AS Writes,
       CAST(CASE WHEN vfs.num_of_reads = 0 THEN 0 ELSE vfs.io_stall_read_ms * 1.0 / vfs.num_of_reads END AS decimal(18,2)) AS AvgReadMs,
       CAST(CASE WHEN vfs.num_of_writes = 0 THEN 0 ELSE vfs.io_stall_write_ms * 1.0 / vfs.num_of_writes END AS decimal(18,2)) AS AvgWriteMs,
       CAST((vfs.num_of_bytes_read + vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS TotalIoMB
FROM sys.dm_io_virtual_file_stats(NULL, NULL) vfs
JOIN sys.master_files mf ON mf.database_id = vfs.database_id AND mf.file_id = vfs.file_id
ORDER BY (vfs.io_stall_read_ms + vfs.io_stall_write_ms) DESC;
'@

$Queries.LogUsage = @'
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel, d.log_reuse_wait_desc AS LogReuseWait,
       CAST(ls.cntr_value / 1024.0 AS decimal(18,2)) AS LogSizeMB, CAST(lu.cntr_value / 1024.0 AS decimal(18,2)) AS LogUsedMB,
       CAST(CASE WHEN ISNULL(ls.cntr_value, 0) = 0 THEN 0 ELSE lu.cntr_value * 100.0 / ls.cntr_value END AS decimal(5,2)) AS LogUsedPct
FROM sys.databases d
LEFT JOIN sys.dm_os_performance_counters ls
       ON RTRIM(ls.instance_name) COLLATE DATABASE_DEFAULT = d.name COLLATE DATABASE_DEFAULT
      AND ls.counter_name = 'Log File(s) Size (KB)' AND ls.object_name LIKE '%:Databases%'
LEFT JOIN sys.dm_os_performance_counters lu
       ON RTRIM(lu.instance_name) COLLATE DATABASE_DEFAULT = d.name COLLATE DATABASE_DEFAULT
      AND lu.counter_name = 'Log File(s) Used Size (KB)' AND lu.object_name LIKE '%:Databases%'
WHERE d.state = 0
ORDER BY d.name;
'@

$Queries.VlfCounts = @'
SELECT d.name AS DatabaseName, v.VlfCount, v.ActiveVlfCount
FROM sys.databases d
CROSS APPLY (SELECT COUNT(*) AS VlfCount, SUM(CASE WHEN li.vlf_active = 1 THEN 1 ELSE 0 END) AS ActiveVlfCount
             FROM sys.dm_db_log_info(d.database_id) li) v
WHERE d.state = 0 AND HAS_DBACCESS(d.name) = 1
  AND d.database_id NOT IN (SELECT drs.database_id FROM sys.dm_hadr_database_replica_states drs
                            WHERE drs.is_local = 1 AND drs.is_primary_replica = 0);
'@

$Queries.MemoryClerks = @'
SELECT TOP (15) type AS ClerkType, CAST(SUM(pages_kb) / 1024.0 AS decimal(18,2)) AS MemoryMB
FROM sys.dm_os_memory_clerks
GROUP BY type
ORDER BY SUM(pages_kb) DESC;
'@

$Queries.ServerServices = @'
SELECT servicename AS ServiceName, startup_type_desc AS StartMode, status_desc AS State, service_account AS ServiceAccount,
       last_startup_time AS LastStartup, is_clustered AS IsClustered, instant_file_initialization_enabled AS InstantFileInit
FROM sys.dm_server_services;
'@

$Queries.ErrorLog = @'
SET NOCOUNT ON;
DECLARE @s datetime = DATEADD(hour, -@hours, GETDATE()), @e datetime = GETDATE();
DECLARE @log TABLE (LogDate datetime, ProcessInfo nvarchar(100), [Text] nvarchar(max));
INSERT INTO @log EXEC master.dbo.xp_readerrorlog 0, 1, NULL, NULL, @s, @e, N'desc';
BEGIN TRY INSERT INTO @log EXEC master.dbo.xp_readerrorlog 1, 1, NULL, NULL, @s, @e, N'desc'; END TRY BEGIN CATCH END CATCH;
SELECT LogDate, ProcessInfo,
       CASE WHEN [Text] LIKE '%Severity: 1[7-9]%' OR [Text] LIKE '%Severity: 2[0-5]%' OR [Text] LIKE '%stack dump%'
                 OR [Text] LIKE '%corrupt%' OR [Text] LIKE '%I/O requests taking longer%' THEN 'Critical'
            WHEN [Text] LIKE '%error%' OR [Text] LIKE '%fail%' THEN 'Error'
            ELSE 'Warning' END AS Level,
       [Text] AS Message
FROM @log
WHERE ([Text] LIKE '%error%' OR [Text] LIKE '%fail%' OR [Text] LIKE '%warning%' OR [Text] LIKE '%severity%'
       OR [Text] LIKE '%deadlock%' OR [Text] LIKE '%I/O requests taking longer%' OR [Text] LIKE '%stack dump%'
       OR [Text] LIKE '%corrupt%' OR [Text] LIKE '%terminat%' OR [Text] LIKE '%non-yielding%')
  AND [Text] NOT LIKE '%found 0 errors%' AND [Text] NOT LIKE '%\ERRORLOG%' AND [Text] NOT LIKE '%without errors%'
  AND [Text] NOT LIKE 'Logging SQL Server messages%' AND [Text] NOT LIKE '%informational message%'
ORDER BY LogDate DESC;
'@

$Queries.CheckDbLog = @'
SET NOCOUNT ON;
DECLARE @log TABLE (LogDate datetime, ProcessInfo nvarchar(100), [Text] nvarchar(max));
INSERT INTO @log EXEC master.dbo.xp_readerrorlog 0, 1, N'DBCC CHECKDB';
BEGIN TRY INSERT INTO @log EXEC master.dbo.xp_readerrorlog 1, 1, N'DBCC CHECKDB'; END TRY BEGIN CATCH END CATCH;
SELECT TOP (200) LogDate, [Text] AS Message,
       CASE WHEN [Text] LIKE '%found 0 errors and repaired 0 errors%' THEN 'OK' ELSE 'Critical' END AS Status
FROM @log
ORDER BY LogDate DESC;
'@

$Queries.SuspectPages = @'
SELECT DB_NAME(sp.database_id) AS DatabaseName, sp.file_id AS FileId, sp.page_id AS PageId,
       CASE sp.event_type WHEN 1 THEN '823/824 error' WHEN 2 THEN 'Bad checksum' WHEN 3 THEN 'Torn page'
                          WHEN 4 THEN 'Restored' WHEN 5 THEN 'Repaired' WHEN 7 THEN 'Deallocated by DBCC' ELSE CAST(sp.event_type AS varchar(10)) END AS EventType,
       sp.error_count AS ErrorCount, sp.last_update_date AS LastUpdate
FROM msdb.dbo.suspect_pages sp
ORDER BY sp.last_update_date DESC;
'@

$Queries.MemoryDumps = @'
SELECT filename AS DumpFile, creation_time AS Created, CAST(size_in_bytes / 1048576.0 AS decimal(18,2)) AS SizeMB
FROM sys.dm_server_memory_dumps
ORDER BY creation_time DESC;
'@

$Queries.DbSecurity = @'
SELECT d.name AS DatabaseName, SUSER_SNAME(d.owner_sid) AS Owner, d.is_trustworthy_on AS Trustworthy,
       d.is_db_chaining_on AS DbChaining, d.is_encrypted AS TdeEnabled, dek.encryption_state AS EncryptionState,
       d.containment_desc AS Containment, d.is_broker_enabled AS BrokerEnabled, d.state_desc AS State
FROM sys.databases d
LEFT JOIN sys.dm_database_encryption_keys dek ON dek.database_id = d.database_id
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.OrphanedUsers = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, UserName sysname, UserType nvarchar(60), AuthenticationType nvarchar(60), CreateDate datetime' -InnerQuery @'
SELECT N'{DBNAME}', dp.name, dp.type_desc, dp.authentication_type_desc, dp.create_date
FROM {DB}.sys.database_principals dp
LEFT JOIN sys.server_principals sp ON sp.sid = dp.sid
WHERE sp.sid IS NULL AND dp.type IN ('S','U','G') AND dp.authentication_type IN (1,3)
  AND dp.principal_id > 4 AND dp.sid IS NOT NULL AND dp.name NOT LIKE '##%'
'@ -OrderBy 'ORDER BY DatabaseName, UserName'

$Queries.DbRoleMembers = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, RoleName sysname, MemberName sysname, MemberType nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', r.name, m.name, m.type_desc
FROM {DB}.sys.database_role_members rm
JOIN {DB}.sys.database_principals r ON r.principal_id = rm.role_principal_id
JOIN {DB}.sys.database_principals m ON m.principal_id = rm.member_principal_id
WHERE r.name IN ('db_owner','db_securityadmin','db_accessadmin','db_ddladmin','db_backupoperator') AND m.name <> 'dbo'
'@ -OrderBy 'ORDER BY DatabaseName, RoleName, MemberName'

$Queries.GuestAccess = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, PermissionState nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', p.state_desc
FROM {DB}.sys.database_permissions p
WHERE p.grantee_principal_id = 2 AND p.permission_name = 'CONNECT' AND p.state IN ('G','W')
  AND N'{DBNAME}' NOT IN (N'master', N'msdb', N'tempdb')
'@ -OrderBy 'ORDER BY DatabaseName'

$Queries.Certificates = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, CertificateName sysname, Subject nvarchar(4000), ExpiryDate datetime, KeyEncryption nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', c.name, c.subject, c.expiry_date, c.pvt_key_encryption_type_desc
FROM {DB}.sys.certificates c
WHERE c.name NOT LIKE '##%'
'@ -OrderBy 'ORDER BY ExpiryDate'

$Queries.ServerRoleMembers = @'
SELECT r.name AS ServerRole, m.name AS MemberName, m.type_desc AS MemberType, m.is_disabled AS IsDisabled
FROM sys.server_role_members rm
JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id
JOIN sys.server_principals m ON m.principal_id = rm.member_principal_id
WHERE r.name IN ('sysadmin','securityadmin','serveradmin','setupadmin','processadmin','diskadmin','dbcreator','bulkadmin')
ORDER BY r.name, m.name;
'@

$Queries.InstanceSecurity = @'
SELECT
    (SELECT CAST(is_disabled AS int) FROM sys.server_principals WHERE sid = 0x01) AS SaDisabled,
    (SELECT name FROM sys.server_principals WHERE sid = 0x01)                     AS SaLoginName,
    CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int)                         AS WindowsAuthOnly,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'xp_cmdshell')                  AS XpCmdShell,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'clr enabled')                  AS ClrEnabled,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'Ole Automation Procedures')    AS OleAutomation,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'Ad Hoc Distributed Queries')   AS AdHocDistributedQueries,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'cross db ownership chaining')  AS CrossDbOwnershipChaining,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'remote admin connections')     AS RemoteAdminConnections,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'contained database authentication') AS ContainedDbAuth,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE is_disabled = 0 AND is_policy_checked = 0)     AS SqlLoginsWithoutPolicy,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE is_disabled = 0 AND is_expiration_checked = 0) AS SqlLoginsWithoutExpiration,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE PWDCOMPARE('', password_hash) = 1)            AS BlankPasswordLogins,
    (SELECT COUNT(*) FROM sys.server_role_members rm JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id WHERE r.name = 'sysadmin') AS SysadminMembers;
'@

$Queries.Configurations = @'
SELECT name AS Setting, CAST(value_in_use AS nvarchar(50)) AS Value FROM sys.configurations ORDER BY name;
'@

$Queries.Logins = @'
SELECT name AS LoginName, type_desc AS LoginType, is_disabled AS IsDisabled, default_database_name AS DefaultDatabase
FROM sys.server_principals
WHERE type IN ('S','U','G') AND name NOT LIKE '##%' AND name NOT LIKE 'NT SERVICE\%' AND name NOT LIKE 'NT AUTHORITY\%'
ORDER BY name;
'@

$Queries.AgentJobs = @'
SELECT j.name AS JobName, j.enabled AS IsEnabled, SUSER_SNAME(j.owner_sid) AS Owner, c.name AS Category
FROM msdb.dbo.sysjobs j
LEFT JOIN msdb.dbo.syscategories c ON c.category_id = j.category_id
ORDER BY j.name;
'@

$Queries.LinkedServers = @'
SELECT name AS LinkedServer, product AS Product, provider AS Provider, data_source AS DataSource
FROM sys.servers WHERE is_linked = 1 ORDER BY name;
'@

$Queries.TraceFlags = @'
SET NOCOUNT ON;
DECLARE @tf TABLE (TraceFlag int, Status int, Global int, Session int);
INSERT INTO @tf EXEC ('DBCC TRACESTATUS(-1) WITH NO_INFOMSGS');
SELECT TraceFlag, Status, Global, Session FROM @tf ORDER BY TraceFlag;
'@

$Queries.JobHistory = @'
SELECT j.name AS JobName, j.enabled AS IsEnabled,
       msdb.dbo.agent_datetime(h.run_date, h.run_time) AS RunStart,
       (h.run_duration / 10000) * 3600 + ((h.run_duration / 100) % 100) * 60 + (h.run_duration % 100) AS DurationSec,
       CASE h.run_status WHEN 0 THEN 'Failed' WHEN 1 THEN 'Succeeded' WHEN 2 THEN 'Retry' WHEN 3 THEN 'Canceled' WHEN 4 THEN 'In Progress' ELSE 'Unknown' END AS RunStatus
FROM msdb.dbo.sysjobhistory h
JOIN msdb.dbo.sysjobs j ON j.job_id = h.job_id
WHERE h.step_id = 0
  AND h.run_date >= CONVERT(int, CONVERT(char(8), DATEADD(day, -@days, GETDATE()), 112))
ORDER BY j.name, RunStart;
'@

$Queries.CheckDb = @'
SELECT d.name AS DatabaseName, d.state_desc AS State,
       CAST(DATABASEPROPERTYEX(d.name, 'LastGoodCheckDbTime') AS datetime) AS LastGoodCheckDb
FROM sys.databases d
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.LongRunning = @'
SELECT r.session_id AS SessionId, s.login_name AS LoginName, s.host_name AS HostName, s.program_name AS ProgramName,
       DB_NAME(r.database_id) AS DatabaseName, r.status AS RequestStatus, r.command AS Command, r.start_time AS StartTime,
       CAST(r.total_elapsed_time / 1000.0 AS decimal(18,1)) AS ElapsedSec, r.cpu_time AS CpuMs, r.logical_reads AS LogicalReads,
       r.wait_type AS WaitType, r.blocking_session_id AS BlockedBy, LEFT(t.text, 2000) AS SqlText
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE s.is_user_process = 1 AND r.session_id <> @@SPID AND r.total_elapsed_time >= @ms
  AND r.wait_type NOT IN (N'BROKER_RECEIVE_WAITFOR', N'WAITFOR')
ORDER BY r.total_elapsed_time DESC;
'@

$Queries.Blocking = @'
SELECT r.session_id AS BlockedSession, r.blocking_session_id AS BlockingSession, DB_NAME(r.database_id) AS DatabaseName,
       r.wait_type AS WaitType, CAST(r.wait_time / 1000.0 AS decimal(18,1)) AS WaitSec, r.wait_resource AS WaitResource,
       s.login_name AS BlockedLogin, bs.login_name AS BlockingLogin, bs.host_name AS BlockingHost, bs.program_name AS BlockingProgram,
       LEFT(t.text, 1000) AS BlockedSql, LEFT(bt.text, 1000) AS BlockingSql
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
LEFT JOIN sys.dm_exec_sessions bs ON bs.session_id = r.blocking_session_id
LEFT JOIN sys.dm_exec_connections bc ON bc.session_id = r.blocking_session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
OUTER APPLY sys.dm_exec_sql_text(bc.most_recent_sql_handle) bt
WHERE r.blocking_session_id <> 0 AND r.wait_time >= @ms
ORDER BY r.wait_time DESC;
'@

# =====================================================================================
#  Phase 1 - Discovery (reachability, HA role, node labels)
# =====================================================================================
Write-Host "SQL Server Operations Dashboard - collection started $($script:ReportStart)" -ForegroundColor Cyan
$Discovery = @{}
$Unreachable = @()
foreach ($inst in $SqlInstances) {
    Write-Host "  Discovering $inst ..."
    $row = @(Invoke-DbQuery -SqlInstance $inst -Query $Queries.Discovery -Section 'Discovery / connectivity') | Select-Object -First 1
    if ($null -eq $row) {
        $script:NodeMap[$inst] = "$inst (Unreachable)"
        $Unreachable += $inst
        Add-Finding -Status Critical -Area 'Databases' -Node $inst -Item 'Instance unreachable' -Detail 'Connection or discovery query failed; see collection errors.'
        continue
    }
    $role = 'Standalone'
    if ([int]$row.AgCount -gt 0) {
        if ([int]$row.LocalPrimaryCount -gt 0 -and [int]$row.LocalSecondaryCount -gt 0) { $role = 'Primary/Secondary' }
        elseif ([int]$row.LocalPrimaryCount -gt 0) { $role = 'Primary' }
        elseif ([int]$row.LocalSecondaryCount -gt 0) { $role = 'Secondary' }
        else { $role = 'AG Replica' }
    } elseif ([int]$row.IsClustered -eq 1) { $role = 'FCI' }
    $script:NodeMap[$inst] = "$inst ($role)"
    $row.SqlNode = $script:NodeMap[$inst]
    $row | Add-Member -NotePropertyName Role -NotePropertyValue $role
    $Discovery[$inst] = $row
    $script:NodeOrder += $inst
}
$script:NodeLabels = @($script:NodeOrder | ForEach-Object { Get-ReportNodeLabel $_ })
$script:IsHA = (@($Discovery.Values | Where-Object { [int]$_.AgCount -gt 0 -or [int]$_.IsClustered -eq 1 }).Count -gt 0)
$HasAG = (@($Discovery.Values | Where-Object { [int]$_.AgCount -gt 0 }).Count -gt 0)
$EnvironmentMode = if ($HasAG) { 'High Availability - Always On Availability Groups' } elseif ($script:IsHA) { 'High Availability - Failover Cluster Instance' } else { 'Standalone' }

# =====================================================================================
#  Phase 2 - SQL Server collection
# =====================================================================================
$collectionPlan = @(
    @{ Key = 'Databases';         Section = 'Database inventory';          Query = $Queries.Databases }
    @{ Key = 'AgReplicas';        Section = 'AG replicas';                 Query = $Queries.AgReplicas;     AgOnly = $true }
    @{ Key = 'AgDatabases';       Section = 'AG database states';          Query = $Queries.AgDatabases;    AgOnly = $true }
    @{ Key = 'AgListeners';       Section = 'AG listeners';                Query = $Queries.AgListeners;    AgOnly = $true }
    @{ Key = 'BackupStatus';      Section = 'Backup freshness';            Query = $Queries.BackupStatus }
    @{ Key = 'BackupHistory';     Section = 'Backup history';              Query = $Queries.BackupHistory;  Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'RestoreHistory';    Section = 'Restore history';             Query = $Queries.RestoreHistory; Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'CpuHistory';        Section = 'CPU history';                 Query = $Queries.CpuHistory;     Params = @{ minutes = $CpuHistoryMinutes } }
    @{ Key = 'Memory';            Section = 'Memory';                      Query = $Queries.Memory }
    @{ Key = 'DatabaseFiles';     Section = 'Database files';              Query = $Queries.DatabaseFiles }
    @{ Key = 'SqlVolumes';        Section = 'Storage volumes';             Query = $Queries.SqlVolumes }
    @{ Key = 'BufferCache';       Section = 'Buffer cache per database';   Query = $Queries.BufferCache }
    @{ Key = 'DbIo';              Section = 'I/O per database';            Query = $Queries.DbIo }
    @{ Key = 'DbCpu';             Section = 'CPU per database';            Query = $Queries.DbCpu }
    @{ Key = 'WaitStats';         Section = 'Wait statistics';             Query = $Queries.WaitStats }
    @{ Key = 'FileLatency';       Section = 'File I/O latency';            Query = $Queries.FileLatency }
    @{ Key = 'LogUsage';          Section = 'Transaction log usage';       Query = $Queries.LogUsage }
    @{ Key = 'VlfCounts';         Section = 'VLF statistics';              Query = $Queries.VlfCounts }
    @{ Key = 'MemoryClerks';      Section = 'Memory clerks';               Query = $Queries.MemoryClerks }
    @{ Key = 'ServerServices';    Section = 'SQL Server services (DMV)';   Query = $Queries.ServerServices }
    @{ Key = 'ErrorLog';          Section = 'SQL Server error log';        Query = $Queries.ErrorLog;       Params = @{ hours = $LookbackHours } }
    @{ Key = 'CheckDbLog';        Section = 'CHECKDB error log evidence';  Query = $Queries.CheckDbLog }
    @{ Key = 'SuspectPages';      Section = 'Suspect pages';               Query = $Queries.SuspectPages }
    @{ Key = 'MemoryDumps';       Section = 'Memory dumps';                Query = $Queries.MemoryDumps }
    @{ Key = 'DbSecurity';        Section = 'Database security';           Query = $Queries.DbSecurity }
    @{ Key = 'OrphanedUsers';     Section = 'Orphaned users';              Query = $Queries.OrphanedUsers }
    @{ Key = 'DbRoleMembers';     Section = 'Privileged database roles';   Query = $Queries.DbRoleMembers }
    @{ Key = 'GuestAccess';       Section = 'Guest access';                Query = $Queries.GuestAccess }
    @{ Key = 'Certificates';      Section = 'Certificates';                Query = $Queries.Certificates }
    @{ Key = 'ServerRoleMembers'; Section = 'Server role membership';      Query = $Queries.ServerRoleMembers }
    @{ Key = 'InstanceSecurity';  Section = 'Instance security checks';    Query = $Queries.InstanceSecurity }
    @{ Key = 'Configurations';    Section = 'sp_configure values';         Query = $Queries.Configurations }
    @{ Key = 'Logins';            Section = 'Logins';                      Query = $Queries.Logins }
    @{ Key = 'AgentJobs';         Section = 'Agent jobs';                  Query = $Queries.AgentJobs }
    @{ Key = 'LinkedServers';     Section = 'Linked servers';              Query = $Queries.LinkedServers }
    @{ Key = 'TraceFlags';        Section = 'Trace flags';                 Query = $Queries.TraceFlags }
    @{ Key = 'JobHistory';        Section = 'Agent job history';           Query = $Queries.JobHistory;     Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'CheckDb';           Section = 'DBCC CHECKDB history';        Query = $Queries.CheckDb }
    @{ Key = 'LongRunning';       Section = 'Long-running requests';       Query = $Queries.LongRunning;    Params = @{ ms = $LongRunningQuerySeconds * 1000 } }
    @{ Key = 'Blocking';          Section = 'Blocking';                    Query = $Queries.Blocking;       Params = @{ ms = $BlockingWaitSeconds * 1000 } }
)

$SqlData = @{}
foreach ($p in $collectionPlan) { $SqlData[$p.Key] = @() }
foreach ($inst in $script:NodeOrder) {
    Write-Host "  Collecting SQL Server data from $(Get-ReportNodeLabel $inst) ..."
    foreach ($p in $collectionPlan) {
        if ($p.AgOnly -and [int]$Discovery[$inst].AgCount -eq 0) { continue }
        $qArgs = @{ SqlInstance = $inst; Query = $p.Query; Section = $p.Section }
        if ($p.Params) { $qArgs.Parameters = $p.Params }
        $SqlData[$p.Key] += @(Invoke-DbQuery @qArgs)
    }
}

# =====================================================================================
#  Phase 3 - Windows and cluster collection
# =====================================================================================
$WinData = @{ OS = @(); Volumes = @(); Services = @(); LocalUsers = @(); Admins = @(); Events = @() }
$hostSeen = @{}
foreach ($inst in $script:NodeOrder) {
    $hostName = [string]$Discovery[$inst].PhysicalHost
    if (-not $hostName -or $hostSeen.ContainsKey($hostName)) { continue }
    $hostSeen[$hostName] = $true
    Write-Host "  Collecting Windows data from $hostName ..."
    $h = Get-WindowsHostData -ComputerName $hostName -Label (Get-ReportNodeLabel $inst)
    foreach ($k in @($WinData.Keys)) { $WinData[$k] += @($h[$k]) }
}
Write-Host '  Collecting Failover Cluster data ...'
$ClusterData = Get-ClusterData -ComputerNames @($hostSeen.Keys)

# Fall back to sys.dm_server_services if remote service enumeration failed
if (@($WinData.Services).Count -eq 0 -and @($SqlData.ServerServices).Count -gt 0) {
    $WinData.Services = @($SqlData.ServerServices | ForEach-Object {
        [pscustomobject]@{
            SqlNode = $_.SqlNode; ServiceName = $_.ServiceName; DisplayName = $_.ServiceName; State = $_.State
            StartMode = $_.StartMode; ServiceAccount = $_.ServiceAccount
            Status = if ($_.State -eq 'Running') { 'OK' } elseif ($_.StartMode -eq 'Automatic') { 'Critical' } else { 'Info' }
        }
    })
}

# =====================================================================================
#  Phase 4 - Analysis and findings
# =====================================================================================
Write-Host '  Analysing collected data ...'
$now = Get-Date

# ---- Instance inventory ----
$versionNames = @{ '11' = '2012'; '12' = '2014'; '13' = '2016'; '14' = '2017'; '15' = '2019'; '16' = '2022'; '17' = '2025' }
$InstanceInventory = @(
    foreach ($inst in $script:NodeOrder) {
        $d = $Discovery[$inst]; $major = ([string]$d.ProductVersion).Split('.')[0]
        [pscustomobject]@{
            SqlNode = $d.SqlNode; ServerName = $d.ServerName; PhysicalHost = $d.PhysicalHost; Role = $d.Role
            SqlVersion = "SQL Server $($versionNames[$major])"; ProductVersion = $d.ProductVersion
            Level = (@($d.ProductLevel, $d.ProductUpdateLevel) | Where-Object { $_ }) -join ' '
            Edition = $d.Edition; Collation = $d.Collation; HadrEnabled = [bool]$d.IsHadrEnabled; Clustered = [bool]$d.IsClustered
            WindowsAuthOnly = [bool]$d.WindowsAuthOnly; StartTime = $d.StartTime
            UptimeDays = [math]::Round(($now - [datetime]$d.StartTime).TotalDays, 1)
            CPUs = $d.CpuCount; HostMemory = Format-SizeMB $d.HostMemoryMB
            MaxServerMemory = if ([double]$d.MaxServerMemoryMB -ge 2147483647) { 'Unlimited' } else { Format-SizeMB $d.MaxServerMemoryMB }
            MaxDop = $d.MaxDop; CostThreshold = $d.CostThreshold; Status = 'OK'
        }
    }
    foreach ($u in $Unreachable) {
        [pscustomobject]@{ SqlNode = Get-ReportNodeLabel $u; ServerName = $u; Role = 'Unknown'; Status = 'Critical' }
    }
)

# ---- Databases ----
foreach ($db in $SqlData.Databases) {
    $cat = Get-DatabaseCategory -DatabaseName $db.DatabaseName -AgName $db.AgName -IsDistributor ([bool]$db.IsDistributor)
    $notes = @(); $st = 'OK'
    switch -Regex ($db.State) {
        '^ONLINE$'                                  { }
        '^(OFFLINE|RESTORING)$'                     { $st = 'Warning';  $notes += "State $($db.State)" }
        default                                     { $st = 'Critical'; $notes += "State $($db.State)" }
    }
    if ($db.AutoClose -eq $true)  { if ($st -eq 'OK') { $st = 'Warning' }; $notes += 'AUTO_CLOSE on' }
    if ($db.AutoShrink -eq $true) { if ($st -eq 'OK') { $st = 'Warning' }; $notes += 'AUTO_SHRINK on' }
    if ($db.PageVerify -and $db.PageVerify -ne 'CHECKSUM' -and $db.DatabaseName -ne 'tempdb') { if ($st -eq 'OK') { $st = 'Warning' }; $notes += "PAGE_VERIFY $($db.PageVerify)" }
    if (-not $db.Owner) { $notes += 'Owner SID has no login' }
    $db | Add-Member -NotePropertyName Category -NotePropertyValue $cat -Force
    $db | Add-Member -NotePropertyName TotalSize -NotePropertyValue (Format-SizeMB ([double]$db.DataMB + [double]$db.LogMB)) -Force
    $db | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $db | Add-Member -NotePropertyName Notes -NotePropertyValue ($notes -join '; ') -Force
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Databases' -Node $db.SqlNode -Item $db.DatabaseName -Detail ($notes -join '; ') }
}

# ---- Always On ----
$roleOf = @{}; foreach ($inst in $script:NodeOrder) { $roleOf[(Get-ReportNodeLabel $inst)] = $Discovery[$inst].Role }
$AgTopology = @(
    $SqlData.AgReplicas | Group-Object AgName, ReplicaServer | ForEach-Object {
        $best = @($_.Group | Sort-Object @{ Expression = { if ($roleOf[$_.SqlNode] -like 'Primary*') { 0 } elseif ($_.ConnectedState) { 1 } else { 2 } } })[0]
        $st = 'OK'
        if ($best.ConnectedState -eq 'DISCONNECTED' -or $best.SynchronizationHealth -eq 'NOT_HEALTHY') { $st = 'Critical' }
        elseif ($best.SynchronizationHealth -eq 'PARTIALLY_HEALTHY' -or ($best.OperationalState -and $best.OperationalState -ne 'ONLINE' -and $best.Role -eq 'PRIMARY')) { $st = 'Warning' }
        elseif (-not $best.SynchronizationHealth) { $st = 'Info' }
        $best | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
        $best | Add-Member -NotePropertyName ReportedBy -NotePropertyValue $best.SqlNode -Force
        if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Always On' -Node $best.SqlNode -Item "$($best.AgName) / $($best.ReplicaServer)" -Detail "Connected: $($best.ConnectedState); sync health: $($best.SynchronizationHealth); last error: $($best.LastConnectErrorDescription)" }
        $best
    }
)

$AgDbStates = @(
    $SqlData.AgDatabases | Group-Object AgName, ReplicaServer, DatabaseName | ForEach-Object {
        @($_.Group | Sort-Object @{ Expression = { if ($roleOf[$_.SqlNode] -like 'Primary*') { 0 } else { 1 } } })[0]
    }
)
foreach ($g in ($AgDbStates | Group-Object AgName, DatabaseName)) {
    $prim = @($g.Group | Where-Object { $_.IsPrimaryReplica -eq $true }) | Select-Object -First 1
    foreach ($r in $g.Group) {
        $loss = $null
        if ($prim -and $r.IsPrimaryReplica -ne $true -and $prim.LastCommitTime -and $r.LastCommitTime) {
            $loss = [math]::Max(0, [math]::Round(([datetime]$prim.LastCommitTime - [datetime]$r.LastCommitTime).TotalSeconds, 0))
        }
        $rec = $null
        if ([double]$r.RedoRateKBs -gt 0) { $rec = [math]::Round([double]$r.RedoQueueKB / [double]$r.RedoRateKBs, 0) }
        $st = 'OK'; $why = @()
        if ($r.SynchronizationHealth -eq 'NOT_HEALTHY' -or $r.IsSuspended -eq $true) { $st = 'Critical'; $why += "health $($r.SynchronizationHealth); suspended $($r.IsSuspended) $($r.SuspendReason)" }
        elseif ($r.SynchronizationHealth -eq 'PARTIALLY_HEALTHY') { $st = 'Warning'; $why += 'partially healthy' }
        if ([double]$r.LogSendQueueKB -gt $AgSendQueueWarningKB) { if ($st -eq 'OK') { $st = 'Warning' }; $why += "log send queue $([double]$r.LogSendQueueKB) KB" }
        if ([double]$r.RedoQueueKB -gt $AgRedoQueueWarningKB)    { if ($st -eq 'OK') { $st = 'Warning' }; $why += "redo queue $([double]$r.RedoQueueKB) KB" }
        $r | Add-Member -NotePropertyName EstDataLossSec -NotePropertyValue $loss -Force
        $r | Add-Member -NotePropertyName EstRecoverySec -NotePropertyValue $rec -Force
        $r | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
        if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Always On' -Node $r.SqlNode -Item "$($r.AgName) / $($r.DatabaseName) @ $($r.ReplicaServer)" -Detail ($why -join '; ') }
    }
}
$AgListeners = @($SqlData.AgListeners | Group-Object AgName, ListenerName, IpAddress | ForEach-Object {
    $l = $_.Group[0]
    $l | Add-Member -NotePropertyName Status -NotePropertyValue $(if ($l.State -eq 'ONLINE') { 'OK' } elseif ($l.State) { 'Warning' } else { 'Info' }) -Force
    $l })

# ---- Backups (compliance across AG replicas) ----
$BackupCompliance = @(
    $SqlData.BackupStatus | Group-Object { if ($_.AgName) { "AG|$($_.AgName)|$($_.DatabaseName)" } else { "SA|$($_.SqlNode)|$($_.DatabaseName)" } } | ForEach-Object {
        $rows  = @($_.Group); $first = $rows[0]
        $isSys = $first.DatabaseName -in @('master', 'model', 'msdb')
        $scope = if ($isSys) { 'System databases' } elseif ($first.AgName) { 'Availability Group databases (evaluated across queried replicas)' } else { 'Standalone databases' }
        $lastFull = ($rows | Where-Object LastFull | Sort-Object LastFull -Descending | Select-Object -First 1)
        $lastDiff = ($rows | Where-Object LastDiff | Sort-Object LastDiff -Descending | Select-Object -First 1)
        $lastLog  = ($rows | Where-Object LastLog  | Sort-Object LastLog  -Descending | Select-Object -First 1)
        $recModel = @($rows | Where-Object RecoveryModel | Select-Object -ExpandProperty RecoveryModel)[0]
        $states   = @($rows | ForEach-Object State)

        $fullSt = 'OK'; $diffSt = 'N/A'; $logSt = 'N/A'
        if (-not $lastFull) { $fullSt = 'Critical' } elseif (($now - [datetime]$lastFull.LastFull).TotalHours -gt $FullBackupMaxAgeHours) { $fullSt = 'Warning' }
        if ($lastDiff) {
            $diffAge = ($now - [datetime]$lastDiff.LastDiff).TotalHours
            $fullAge = if ($lastFull) { ($now - [datetime]$lastFull.LastFull).TotalHours } else { [double]::MaxValue }
            $diffSt = if ($diffAge -gt $DiffBackupMaxAgeHours -and $fullAge -gt $DiffBackupMaxAgeHours) { 'Warning' } else { 'OK' }
        }
        if (-not $isSys -and $recModel -and $recModel -ne 'SIMPLE') {
            if (-not $lastLog) { $logSt = 'Critical' } elseif (($now - [datetime]$lastLog.LastLog).TotalMinutes -gt $LogBackupMaxAgeMinutes) { $logSt = 'Warning' } else { $logSt = 'OK' }
        }
        $overall = Get-WorstStatus @($fullSt, $diffSt, $logSt)
        if ($states -notcontains 'ONLINE') { $overall = 'Info' }

        $item = [pscustomobject]@{
            Scope = $scope; DatabaseName = $first.DatabaseName; AgName = $first.AgName; RecoveryModel = $recModel
            ReplicasQueried = (@($rows | ForEach-Object SqlNode | Sort-Object -Unique) -join ', ')
            LastFull = if ($lastFull) { $lastFull.LastFull } else { $null }; FullAge = Format-Age $(if ($lastFull) { $lastFull.LastFull }); FullTakenOn = if ($lastFull) { $lastFull.SqlNode } else { '' }; FullStatus = $fullSt
            LastDiff = if ($lastDiff) { $lastDiff.LastDiff } else { $null }; DiffAge = if ($lastDiff) { Format-Age $lastDiff.LastDiff } else { 'Not used' }; DiffStatus = $diffSt
            LastLog = if ($lastLog) { $lastLog.LastLog } else { $null }; LogAge = if ($lastLog) { Format-Age $lastLog.LastLog } else { $(if ($logSt -eq 'N/A') { 'N/A' } else { 'Never' }) }; LogTakenOn = if ($lastLog) { $lastLog.SqlNode } else { '' }; LogStatus = $logSt
            Status = $overall
            SqlNode = if ($first.AgName) { 'AG (all replicas)' } else { $first.SqlNode }
        }
        if ($overall -in 'Critical', 'Warning') {
            $why = @(); if ($fullSt -ne 'OK') { $why += "full: $($item.FullAge)" }; if ($diffSt -eq 'Warning') { $why += "diff: $($item.DiffAge)" }; if ($logSt -in 'Critical', 'Warning') { $why += "log: $($item.LogAge)" }
            Add-Finding -Status $overall -Area 'Backups' -Node $item.SqlNode -Item $item.DatabaseName -Detail ($why -join '; ')
        }
        $item
    }
)
$BackupGaps = @($BackupCompliance | Where-Object { $_.Status -in 'Critical', 'Warning' })

# ---- Capacity ----
$SqlVolumes = @($SqlData.SqlVolumes | ForEach-Object {
    $pct = 0; if ([double]$_.TotalGB -gt 0) { $pct = [math]::Round(100 * [double]$_.FreeGB / [double]$_.TotalGB, 1) }
    $st = if ($pct -lt $DiskFreeCriticalPct) { 'Critical' } elseif ($pct -lt $DiskFreeWarningPct) { 'Warning' } else { 'OK' }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $_.SqlNode -Item "Volume $($_.VolumePath)" -Detail "$pct% free ($($_.FreeGB) GB of $($_.TotalGB) GB)" }
    $_ | Add-Member -NotePropertyName FreePct -NotePropertyValue $pct -Force
    $_ | Add-Member -NotePropertyName UsedGB -NotePropertyValue ([math]::Round([double]$_.TotalGB - [double]$_.FreeGB, 2)) -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

foreach ($f in $SqlData.FileLatency) {
    $worst = [math]::Max([double]$f.AvgReadMs, [double]$f.AvgWriteMs)
    $st = if ($worst -ge $IoLatencyCriticalMs) { 'Critical' } elseif ($worst -ge $IoLatencyWarningMs) { 'Warning' } else { 'OK' }
    $f | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    if ($st -ne 'OK' -and ([long]$f.Reads + [long]$f.Writes) -gt 1000) { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $f.SqlNode -Item "$($f.DatabaseName) / $($f.LogicalName)" -Detail "Avg read $($f.AvgReadMs) ms, avg write $($f.AvgWriteMs) ms" }
}

$vlfIndex = @{}; foreach ($v in $SqlData.VlfCounts) { $vlfIndex["$($v.SqlNode)|$($v.DatabaseName)"] = $v }
$LogUsage = @($SqlData.LogUsage | ForEach-Object {
    $v = $vlfIndex["$($_.SqlNode)|$($_.DatabaseName)"]
    $vlf = if ($v) { [int]$v.VlfCount } else { $null }
    $st = 'OK'; $why = @()
    if ([double]$_.LogUsedPct -ge $LogUsedWarningPct -and $_.LogReuseWait -ne 'NOTHING') { $st = 'Warning'; $why += "log $($_.LogUsedPct)% used, reuse wait $($_.LogReuseWait)" }
    if ($vlf -and $vlf -ge $VlfWarningCount) { $st = 'Warning'; $why += "$vlf VLFs" }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $_.SqlNode -Item "$($_.DatabaseName) transaction log" -Detail ($why -join '; ') }
    $_ | Add-Member -NotePropertyName VlfCount -NotePropertyValue $vlf -Force
    $_ | Add-Member -NotePropertyName ActiveVlfs -NotePropertyValue $(if ($v) { $v.ActiveVlfCount }) -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

foreach ($m in $SqlData.Memory) {
    if ($m.SystemMemoryState -match 'low') { Add-Finding -Status Critical -Area 'Capacity & Performance' -Node $m.SqlNode -Item 'Host memory' -Detail $m.SystemMemoryState }
    if ([double]$m.MaxServerMemoryMB -ge 2147483647) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $m.SqlNode -Item 'max server memory' -Detail 'Left at default (unlimited).' }
    if ([double]$m.PageLifeExpectancy -lt 300) { Add-Finding -Status Warning -Area 'Capacity & Performance' -Node $m.SqlNode -Item 'Page life expectancy' -Detail "$($m.PageLifeExpectancy) seconds" }
}
foreach ($node in $script:NodeLabels) {
    $cpu = @($SqlData.CpuHistory | Where-Object SqlNode -eq $node)
    if ($cpu.Count -gt 0) {
        $avg = ($cpu | Measure-Object -Property SQLProcessUtilization -Average).Average
        if ($avg -ge 80) { Add-Finding -Status Critical -Area 'Capacity & Performance' -Node $node -Item 'SQL Server CPU' -Detail ('Average {0:N0}% over {1} minutes' -f $avg, $cpu.Count) }
        elseif ($avg -ge 60) { Add-Finding -Status Warning -Area 'Capacity & Performance' -Node $node -Item 'SQL Server CPU' -Detail ('Average {0:N0}% over {1} minutes' -f $avg, $cpu.Count) }
    }
}

# ---- Cluster & Windows ----
foreach ($n in $ClusterData.Nodes)     { if ($n.Status -ne 'OK') { Add-Finding -Status $n.Status -Area 'Cluster & Windows' -Node $n.ClusterName -Item "Cluster node $($n.NodeName)" -Detail "State $($n.State)" } }
foreach ($g in $ClusterData.Groups)    { if ($g.Status -ne 'OK') { Add-Finding -Status $g.Status -Area 'Cluster & Windows' -Node $g.ClusterName -Item "Cluster group $($g.GroupName)" -Detail "State $($g.State) on $($g.OwnerNode)" } }
foreach ($r in $ClusterData.Resources) { if ($r.Status -eq 'Critical') { Add-Finding -Status Critical -Area 'Cluster & Windows' -Node $r.ClusterName -Item "Cluster resource $($r.ResourceName)" -Detail "State $($r.State)" } }
foreach ($s in $WinData.Services)  { if ($s.Status -eq 'Critical') { Add-Finding -Status Critical -Area 'Cluster & Windows' -Node $s.SqlNode -Item "Service $($s.ServiceName)" -Detail "Automatic service is $($s.State)" } }
foreach ($v in $WinData.Volumes)   { if ($v.Status -ne 'OK') { Add-Finding -Status $v.Status -Area 'Cluster & Windows' -Node $v.SqlNode -Item "Windows volume $($v.Volume)" -Detail "$($v.FreePct)% free" } }

$clusterProviders = 'Microsoft-Windows-FailoverClustering|Microsoft-Windows-DNS-Client|DNS Client Events|NETLOGON|Microsoft-Windows-Security-Kerberos|LsaSrv|Microsoft-Windows-GroupPolicy|Microsoft-Windows-Time-Service'
$dnsIds = @(1196, 1257, 1578, 1579)
$adIds  = @(1193, 1194, 1206, 1207, 1222, 1223, 1226, 1228, 1683, 1685, 1688)
$ClusterDnsAdEvents = @($WinData.Events | Where-Object { $_.LogName -eq 'System' -and $_.Provider -match $clusterProviders } | ForEach-Object {
    $cat = 'Cluster'
    if ($_.Provider -match 'DNS' -or ($_.Provider -match 'FailoverClustering' -and $dnsIds -contains $_.EventId)) { $cat = 'DNS' }
    elseif ($_.Provider -match 'NETLOGON|Kerberos|LsaSrv|GroupPolicy|Time-Service' -or ($_.Provider -match 'FailoverClustering' -and $adIds -contains $_.EventId)) { $cat = 'Active Directory' }
    $_ | Add-Member -NotePropertyName Category -NotePropertyValue $cat -Force
    $_ })
$WindowsCritical = @($WinData.Events | Where-Object { $_.Level -eq 'Critical' -or ($_.Level -eq 'Error' -and $_.Message -match 'fail|unexpected|terminat|crash|corrupt') })
foreach ($g in ($WindowsCritical | Where-Object Level -eq 'Critical' | Group-Object SqlNode)) {
    Add-Finding -Status Critical -Area 'SQL Errors' -Node $g.Name -Item 'Windows critical events' -Detail "$($g.Count) critical event(s) in the last $LookbackHours h"
}
foreach ($g in ($ClusterDnsAdEvents | Where-Object Level -in 'Critical', 'Error' | Group-Object SqlNode)) {
    Add-Finding -Status Warning -Area 'Cluster & Windows' -Node $g.Name -Item 'Cluster / DNS / AD events' -Detail "$($g.Count) error event(s) in the last $LookbackHours h"
}

# ---- SQL errors ----
$ErrorGroups = @($SqlData.ErrorLog | Group-Object SqlNode, @{ Expression = {
        $t = [string]$_.Message
        $t = $t -replace '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b', '<guid>'
        $t = $t -replace "'[^']*'", "'<x>'" -replace '\[CLIENT: [^\]]*\]', '[CLIENT: <x>]' -replace '\d+', '#'
        $t.Trim() } } | ForEach-Object {
    $items = @($_.Group | Sort-Object LogDate)
    [pscustomobject]@{
        SqlNode = $items[0].SqlNode; Occurrences = $items.Count; Level = Get-WorstStatus @($items | ForEach-Object { if ($_.Level -eq 'Critical') { 'Critical' } elseif ($_.Level -eq 'Error') { 'Warning' } else { 'Info' } })
        FirstSeen = $items[0].LogDate; LastSeen = $items[-1].LogDate; SampleMessage = $items[-1].Message
    } } | Sort-Object Occurrences -Descending)
foreach ($g in ($SqlData.ErrorLog | Where-Object Level -eq 'Critical' | Group-Object SqlNode)) {
    Add-Finding -Status Critical -Area 'SQL Errors' -Node $g.Name -Item 'Critical error log entries' -Detail "$($g.Count) entries (severity 17+, stack dumps, corruption or slow I/O) in $LookbackHours h"
}
foreach ($g in ($SqlData.ErrorLog | Where-Object Message -like '*Login failed*' | Group-Object SqlNode)) {
    if ($g.Count -ge 10) { Add-Finding -Status Warning -Area 'SQL Errors' -Node $g.Name -Item 'Failed logins' -Detail "$($g.Count) failed login messages in $LookbackHours h" }
}
foreach ($s in $SqlData.SuspectPages) {
    if ($s.EventType -notin 'Restored', 'Repaired', 'Deallocated by DBCC') { Add-Finding -Status Critical -Area 'SQL Errors' -Node $s.SqlNode -Item "Suspect page in $($s.DatabaseName)" -Detail "File $($s.FileId) page $($s.PageId): $($s.EventType), $($s.ErrorCount) error(s)" }
}
foreach ($g in ($SqlData.MemoryDumps | Where-Object { $_.Created -ge $now.AddDays(-$HistoryLookbackDays) } | Group-Object SqlNode)) {
    Add-Finding -Status Warning -Area 'SQL Errors' -Node $g.Name -Item 'SQL Server memory dumps' -Detail "$($g.Count) dump(s) in the last $HistoryLookbackDays days"
}

# ---- Security & configuration ----
$DbSecurity = @($SqlData.DbSecurity | ForEach-Object {
    $isSys = $_.DatabaseName -in @('master', 'model', 'msdb')
    $why = @()
    if ($_.Trustworthy -eq $true -and $_.DatabaseName -ne 'msdb') { $why += 'TRUSTWORTHY on' }
    if ($_.DbChaining -eq $true -and -not $isSys) { $why += 'DB chaining on' }
    if (-not $_.Owner) { $why += 'Owner has no matching login' }
    $_ | Add-Member -NotePropertyName Encryption -NotePropertyValue $(switch ([int]$_.EncryptionState) { 0 { 'None' } 1 { 'Unencrypted' } 2 { 'Encrypting' } 3 { 'Encrypted' } 4 { 'Key change' } 5 { 'Decrypting' } 6 { 'Protection change' } default { 'None' } }) -Force
    $_ | Add-Member -NotePropertyName Issues -NotePropertyValue ($why -join '; ') -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $(if ($why.Count) { 'Warning' } else { 'OK' }) -Force
    if ($why.Count) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $_.SqlNode -Item $_.DatabaseName -Detail ($why -join '; ') }
    $_ })
foreach ($g in ($SqlData.OrphanedUsers | Group-Object SqlNode)) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $g.Name -Item 'Orphaned database users' -Detail "$($g.Count) user(s) without a matching login" }
foreach ($g in $SqlData.GuestAccess) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $g.SqlNode -Item $g.DatabaseName -Detail 'guest user has CONNECT permission' }

$SecurityChecks = @(foreach ($s in $SqlData.InstanceSecurity) {
    $chk = @(
        @{ Check = 'sa login disabled';                   Value = [bool]$s.SaDisabled;               Expected = $true;  Sev = 'Warning' }
        @{ Check = 'sa login renamed';                    Value = ($s.SaLoginName -ne 'sa');          Expected = $true;  Sev = 'Info' }
        @{ Check = 'Windows authentication only';         Value = [bool]$s.WindowsAuthOnly;          Expected = $true;  Sev = 'Info' }
        @{ Check = 'xp_cmdshell disabled';                Value = ([int]$s.XpCmdShell -eq 0);         Expected = $true;  Sev = 'Warning' }
        @{ Check = 'CLR disabled';                        Value = ([int]$s.ClrEnabled -eq 0);         Expected = $true;  Sev = 'Info' }
        @{ Check = 'OLE Automation disabled';             Value = ([int]$s.OleAutomation -eq 0);      Expected = $true;  Sev = 'Warning' }
        @{ Check = 'Ad Hoc Distributed Queries disabled'; Value = ([int]$s.AdHocDistributedQueries -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'Cross-database ownership chaining off'; Value = ([int]$s.CrossDbOwnershipChaining -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'Remote admin connections (DAC) enabled'; Value = ([int]$s.RemoteAdminConnections -eq 1); Expected = $true; Sev = 'Info' }
        @{ Check = 'No logins with blank password';       Value = ([int]$s.BlankPasswordLogins -eq 0); Expected = $true; Sev = 'Critical' }
        @{ Check = 'All SQL logins enforce password policy'; Value = ([int]$s.SqlLoginsWithoutPolicy -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'All SQL logins enforce expiration';   Value = ([int]$s.SqlLoginsWithoutExpiration -eq 0); Expected = $true; Sev = 'Info' }
        @{ Check = 'sysadmin members (5 or fewer)';       Value = ([int]$s.SysadminMembers -le 5);    Expected = $true;  Sev = 'Warning' }
    )
    foreach ($c in $chk) {
        $pass = ($c.Value -eq $c.Expected)
        $st = if ($pass) { 'OK' } else { $c.Sev }
        if (-not $pass -and $c.Sev -ne 'Info') { Add-Finding -Status $c.Sev -Area 'Security & Configuration' -Node $s.SqlNode -Item $c.Check -Detail 'Check not met' }
        [pscustomobject]@{ SqlNode = $s.SqlNode; Check = $c.Check; Result = if ($pass) { 'Pass' } else { 'Not met' }; Status = $st }
    }
})
$Certificates = @($SqlData.Certificates | ForEach-Object {
    $days = if ($_.ExpiryDate) { [math]::Floor(([datetime]$_.ExpiryDate - $now).TotalDays) } else { $null }
    $st = if ($null -eq $days) { 'Info' } elseif ($days -lt 0) { 'Critical' } elseif ($days -le $CertificateExpiryWarningDays) { 'Warning' } else { 'OK' }
    if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Security & Configuration' -Node $_.SqlNode -Item "Certificate $($_.CertificateName) ($($_.DatabaseName))" -Detail "Expires $($_.ExpiryDate) ($days days)" }
    $_ | Add-Member -NotePropertyName DaysToExpiry -NotePropertyValue $days -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

function Compare-AcrossNodes {
    param([object[]] $Data, [string] $KeyProperty, [scriptblock] $ValueOf, [string] $Area, [string] $What, [switch] $Presence)
    $labels = @($script:NodeLabels)
    $index = @{}
    foreach ($r in $Data) { $index["$($r.SqlNode)|$($r.$KeyProperty)"] = $r }
    $keys = @($Data | ForEach-Object { [string]$_.$KeyProperty } | Sort-Object -Unique)
    foreach ($k in $keys) {
        $o = [ordered]@{ Item = $k }
        $vals = @()
        foreach ($l in $labels) {
            $r = $index["$l|$k"]
            $v = if ($null -eq $r) { 'Missing' } else { [string](& $ValueOf $r) }
            $o[$l] = $v; $vals += $v
        }
        $same = (@($vals | Sort-Object -Unique).Count -le 1)
        $o['Comparison'] = if ($same) { 'Match' } else { 'Mismatch' }
        if (-not $same -and $labels.Count -gt 1) { Add-Finding -Status Warning -Area $Area -Node 'HA nodes' -Item "$What $k" -Detail (($labels | ForEach-Object { "$_ = $($o[$_])" }) -join ' | ') }
        [pscustomobject]$o
    }
}
$ConfigCompare = @(); $PropertyCompare = @(); $LoginCompare = @(); $JobCompare = @(); $TraceFlagCompare = @(); $LinkedCompare = @()
if (@($script:NodeLabels).Count -gt 1) {
    $ConfigCompare   = @(Compare-AcrossNodes -Data $SqlData.Configurations -KeyProperty Setting -ValueOf { param($r) $r.Value } -Area 'Security & Configuration' -What 'sp_configure')
    $propRows = foreach ($inst in $script:NodeOrder) {
        $d = $Discovery[$inst]
        foreach ($p in 'Edition', 'ProductVersion', 'ProductUpdateLevel', 'Collation', 'MaxServerMemoryMB', 'MaxDop', 'CostThreshold', 'CpuCount', 'HostMemoryMB', 'IsHadrEnabled', 'WindowsAuthOnly') {
            [pscustomobject]@{ SqlNode = $d.SqlNode; Property = $p; Value = $d.$p }
        }
    }
    $PropertyCompare = @(Compare-AcrossNodes -Data $propRows -KeyProperty Property -ValueOf { param($r) $r.Value } -Area 'Security & Configuration' -What 'Server property')
    $LoginCompare    = @(Compare-AcrossNodes -Data $SqlData.Logins -KeyProperty LoginName -ValueOf { param($r) if ($r.IsDisabled) { 'Present (disabled)' } else { 'Present' } } -Area 'Security & Configuration' -What 'Login')
    $JobCompare      = @(Compare-AcrossNodes -Data $SqlData.AgentJobs -KeyProperty JobName -ValueOf { param($r) if ($r.IsEnabled) { 'Enabled' } else { 'Present (disabled)' } } -Area 'Security & Configuration' -What 'Agent job')
    $TraceFlagCompare = @(Compare-AcrossNodes -Data ($SqlData.TraceFlags | Where-Object { $_.Global -eq 1 }) -KeyProperty TraceFlag -ValueOf { param($r) 'Enabled' } -Area 'Security & Configuration' -What 'Trace flag')
    $LinkedCompare   = @(Compare-AcrossNodes -Data $SqlData.LinkedServers -KeyProperty LinkedServer -ValueOf { param($r) 'Present' } -Area 'Security & Configuration' -What 'Linked server')
}

# ---- Investigation ----
$JobTrends = @($SqlData.JobHistory | Group-Object SqlNode, JobName | ForEach-Object {
    $runs = @($_.Group | Sort-Object RunStart)
    $dur  = [double[]]@($runs | ForEach-Object { [double]$_.DurationSec })
    $avg  = ($dur | Measure-Object -Average).Average
    $max  = ($dur | Measure-Object -Maximum).Maximum
    $last = $runs[-1]
    $failed = @($runs | Where-Object RunStatus -eq 'Failed').Count
    $trendPct = if ($avg -gt 0) { [math]::Round(100 * ([double]$last.DurationSec - $avg) / $avg, 0) } else { 0 }
    $st = 'OK'
    if ($last.RunStatus -eq 'Failed') { $st = 'Critical' }
    elseif ($failed -gt 0 -or ($avg -ge 60 -and [double]$last.DurationSec -gt 2 * $avg)) { $st = 'Warning' }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Investigation' -Node $last.SqlNode -Item "Agent job $($last.JobName)" -Detail "Last outcome $($last.RunStatus); $failed failure(s) in $HistoryLookbackDays days; last run $(Format-Duration $last.DurationSec) vs avg $(Format-Duration $avg)" }
    [pscustomobject]@{
        SqlNode = $last.SqlNode; JobName = $last.JobName; Runs = $runs.Count; Failed = $failed
        AvgDuration = Format-Duration $avg; MaxDuration = Format-Duration $max; LastDuration = Format-Duration $last.DurationSec
        LastRun = $last.RunStart; LastOutcome = $last.RunStatus; LastVsAvgPct = $trendPct; Status = $st; Durations = $dur
    } })

$CheckDbStatus = @($SqlData.CheckDb | ForEach-Object {
    $last = $_.LastGoodCheckDb
    $st = 'OK'; $age = $null
    if ($null -eq $last) { $st = 'Info'; $age = 'Unknown (not readable on this replica)' }
    elseif (([datetime]$last).Year -le 1900) { $st = 'Critical'; $age = 'Never'; $last = $null }
    else {
        $days = ($now - [datetime]$last).TotalDays; $age = '{0:N1} days' -f $days
        if ($days -gt $CheckDbMaxAgeDays) { $st = 'Warning' }
    }
    if ($_.State -ne 'ONLINE') { $st = 'Info' }
    if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Investigation' -Node $_.SqlNode -Item "DBCC CHECKDB $($_.DatabaseName)" -Detail "Last known good: $age (threshold $CheckDbMaxAgeDays days)" }
    [pscustomobject]@{ SqlNode = $_.SqlNode; DatabaseName = $_.DatabaseName; State = $_.State; LastKnownGood = $last; Age = $age; Status = $st }
})
foreach ($e in ($SqlData.CheckDbLog | Where-Object Status -eq 'Critical')) { Add-Finding -Status Critical -Area 'Investigation' -Node $e.SqlNode -Item 'DBCC CHECKDB reported errors' -Detail $e.Message }
foreach ($g in ($SqlData.LongRunning | Group-Object SqlNode)) { Add-Finding -Status Warning -Area 'Investigation' -Node $g.Name -Item 'Long-running requests' -Detail "$($g.Count) request(s) running longer than $LongRunningQuerySeconds s" }
foreach ($g in ($SqlData.Blocking | Group-Object SqlNode))    { Add-Finding -Status Warning -Area 'Investigation' -Node $g.Name -Item 'Blocking' -Detail "$($g.Count) blocked request(s) waiting longer than $BlockingWaitSeconds s" }
if ($script:CollectionErrors.Count -gt 0) {
    Add-Finding -Status Info -Area 'Investigation' -Node 'All' -Item 'Collection errors' -Detail "$($script:CollectionErrors.Count) section(s) could not be collected; see 'Collection errors and permission issues'."
}

# ---- Optional metric history ----
$HistoryRows = @()
if ($EnableMetricHistory) {
    $histFolder = Join-Path $OutputFolder 'History'
    if (-not (Test-Path -LiteralPath $histFolder)) { New-Item -ItemType Directory -Path $histFolder -Force | Out-Null }
    $histFile = Join-Path $histFolder 'MetricHistory.csv'
    $stamp = $now.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv)
    $newRows = foreach ($node in $script:NodeLabels) {
        $cpu = @($SqlData.CpuHistory | Where-Object SqlNode -eq $node)
        $mem = @($SqlData.Memory | Where-Object SqlNode -eq $node) | Select-Object -First 1
        $metrics = [ordered]@{
            'SQL CPU avg %'            = if ($cpu) { [math]::Round(($cpu | Measure-Object -Property SQLProcessUtilization -Average).Average, 1) } else { $null }
            'Page life expectancy'     = if ($mem) { $mem.PageLifeExpectancy } else { $null }
            'SQL memory in use MB'     = if ($mem) { $mem.SqlPhysicalMemoryMB } else { $null }
            'Databases'                = @($SqlData.Databases | Where-Object SqlNode -eq $node).Count
            'Total database size GB'   = [math]::Round((($SqlData.Databases | Where-Object SqlNode -eq $node | ForEach-Object { [double]$_.DataMB + [double]$_.LogMB }) | Measure-Object -Sum).Sum / 1024, 2)
            'Critical findings'        = @($script:Findings | Where-Object { $_.SqlNode -eq $node -and $_.Status -eq 'Critical' }).Count
            'Warning findings'         = @($script:Findings | Where-Object { $_.SqlNode -eq $node -and $_.Status -eq 'Warning' }).Count
            'Error log entries'        = @($SqlData.ErrorLog | Where-Object SqlNode -eq $node).Count
        }
        foreach ($k in $metrics.Keys) { [pscustomobject]@{ Timestamp = $stamp; SqlNode = $node; Metric = $k; Value = $metrics[$k] } }
    }
    $existing = @()
    if (Test-Path -LiteralPath $histFile) { $existing = @(Import-Csv -LiteralPath $histFile) }
    $cutoff = $now.AddDays(-$MetricHistoryRetentionDays)
    $kept = @($existing | Where-Object {
        $t = [datetime]::MinValue
        [datetime]::TryParseExact($_.Timestamp, 'yyyy-MM-dd HH:mm:ss', $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$t) -and $t -ge $cutoff })
    $HistoryRows = @($kept) + @($newRows)
    $HistoryRows | Export-Csv -LiteralPath $histFile -NoTypeInformation -Encoding UTF8
}

# =====================================================================================
#  Phase 5 - HTML dashboard
# =====================================================================================
Write-Host '  Building HTML dashboard ...'
$AreaNames = @('Databases', 'Always On', 'Backups', 'Capacity & Performance', 'Cluster & Windows', 'SQL Errors', 'Security & Configuration', 'Investigation')
$AreaSummary = @(foreach ($a in $AreaNames) {
    $f = @($script:Findings | Where-Object Area -eq $a)
    $crit = @($f | Where-Object Status -eq 'Critical').Count; $warn = @($f | Where-Object Status -eq 'Warning').Count
    [pscustomobject]@{ Area = $a; Critical = $crit; Warning = $warn; Status = if ($crit) { 'Critical' } elseif ($warn) { 'Warning' } else { 'OK' } }
})
$critTotal = @($script:Findings | Where-Object Status -eq 'Critical').Count
$warnTotal = @($script:Findings | Where-Object Status -eq 'Warning').Count
$overall   = if ($critTotal) { 'Critical' } elseif ($warnTotal) { 'Warning' } else { 'OK' }
$dbTotal   = @($SqlData.Databases).Count
$dbOnline  = @($SqlData.Databases | Where-Object State -eq 'ONLINE').Count
$duration  = Format-Duration ((Get-Date) - $script:ReportStart).TotalSeconds

$Css = @'
<style>
.ops-banner{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px;padding:14px 18px;border-radius:10px;background:linear-gradient(90deg,#1f3b64,#2f5d9a);color:#fff;margin-bottom:10px}
.ops-banner h1{font-size:20px;margin:0;font-weight:600;color:#fff}
.ops-banner .ops-meta{font-size:12px;opacity:.9}
.ops-kpi-grid,.ops-card-grid{display:flex;flex-wrap:wrap;gap:12px;margin:4px 0 8px}
.ops-kpi{flex:1 1 150px;border-radius:8px;padding:12px 14px;background:#fff;border:1px solid #e3e7ee;border-left:5px solid #3a7bd5}
.ops-kpi.st-ok,.ops-card.st-ok{border-left-color:#2e9d5b}.ops-kpi.st-warning,.ops-card.st-warning{border-left-color:#e0a100}
.ops-kpi.st-critical,.ops-card.st-critical{border-left-color:#d64545}.ops-kpi.st-info,.ops-card.st-info{border-left-color:#3a7bd5}
.ops-kpi-title{font-size:11px;text-transform:uppercase;color:#6b7280;letter-spacing:.05em}
.ops-kpi-value{font-size:26px;font-weight:600;color:#1f2937;margin-top:4px}
.ops-kpi-sub{font-size:12px;color:#6b7280;margin-top:2px}
.ops-card{flex:1 1 260px;border-radius:8px;padding:10px 14px;background:#fff;border:1px solid #e3e7ee;border-left:5px solid #3a7bd5}
.ops-card-head{display:flex;justify-content:space-between;align-items:center;font-weight:600;color:#1f2937;margin-bottom:6px}
.ops-card-table{width:100%;font-size:12px;border-collapse:collapse}
.ops-card-table td{padding:3px 0;border-bottom:1px solid #f1f3f6}.ops-card-table td:last-child{text-align:right;font-weight:500}
.ops-card-foot{font-size:11px;color:#6b7280;margin-top:6px}
.ops-pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600}
.ops-pill.st-ok{background:#d8f3e3;color:#1e6b3e}.ops-pill.st-warning{background:#fff1cc;color:#7a5b00}
.ops-pill.st-critical{background:#f9d6d6;color:#8a1f1f}.ops-pill.st-info{background:#e3edfb;color:#1f4f8f}
.ops-node-title{font-weight:600;font-size:14px;margin:2px 0 8px;color:#1f3b64;border-bottom:2px solid #e3e7ee;padding-bottom:4px}
.ops-group-head{font-weight:600;font-size:13px;margin:14px 0 6px;color:#1f2937}
.ops-muted{color:#6b7280;font-weight:400;font-size:12px}
.ops-empty{color:#6b7280;font-style:italic;padding:8px 2px}
.ops-chart{margin:6px 0 12px}
.ops-chart-title{font-weight:600;font-size:13px;color:#1f2937;margin-bottom:6px}
.ops-chart-legend{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;margin:4px 0}
.ops-svg-wide{width:100%;max-width:900px;height:auto}
.ops-axis{font-size:10px;fill:#6b7280}
.ops-donut-body{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.ops-donut-total{font-size:15px;font-weight:600;fill:#1f2937}.ops-donut-unit{font-size:10px;fill:#6b7280}
.ops-legend{list-style:none;margin:0;padding:0;font-size:12px;min-width:200px}
.ops-legend li{display:flex;align-items:center;gap:6px;margin:3px 0}
.ops-legend-label{max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ops-legend-val{color:#6b7280;margin-left:auto;padding-left:8px;white-space:nowrap}
.ops-swatch{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}
.ops-bar-row{display:flex;align-items:center;gap:8px;margin:3px 0;font-size:12px}
.ops-bar-label{flex:0 0 200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ops-bar-track{flex:1;background:#eef1f5;border-radius:4px;height:14px;overflow:hidden;display:flex}
.ops-bar-track.small{height:8px;display:inline-flex;width:110px;vertical-align:middle;margin-right:6px}
.ops-bar-fill{height:100%}
.ops-bar-value{flex:0 0 90px;text-align:right;color:#374151}
.ops-mini-table{width:100%;border-collapse:collapse;font-size:12px}
.ops-mini-table th{text-align:left;background:#f4f6f9;padding:5px 6px;font-weight:600;color:#374151}
.ops-mini-table td{padding:4px 6px;border-bottom:1px solid #eef1f5;vertical-align:middle}
.ops-num{font-variant-numeric:tabular-nums;white-space:nowrap}
.ops-donut-grid{display:flex;flex-wrap:wrap;gap:18px}
.ops-donut-grid .ops-chart{flex:1 1 360px}
</style>
'@

$Banner = "<div class='ops-banner'><div><h1>SQL Server Operations Dashboard</h1><div class='ops-meta'>$(ConvertTo-HtmlSafe $EnvironmentName) &middot; $(ConvertTo-HtmlSafe $EnvironmentMode)</div></div>" +
          "<div class='ops-meta'>Generated $($now.ToString('yyyy-MM-dd HH:mm')) by $(ConvertTo-HtmlSafe "$env:USERDOMAIN\$env:USERNAME") on $(ConvertTo-HtmlSafe $env:COMPUTERNAME) &middot; collection time $duration<br/>" +
          "Nodes: $(ConvertTo-HtmlSafe (($script:NodeLabels + @($Unreachable | ForEach-Object { Get-ReportNodeLabel $_ })) -join ', ')) &middot; read-only collection</div></div>"

$Kpis = "<div class='ops-kpi-grid'>" +
    (ConvertTo-KpiHtmlText -Title 'Overall status' -Value $overall -Subtitle "$critTotal critical, $warnTotal warning findings" -Status $overall) +
    (ConvertTo-KpiHtmlText -Title 'SQL Server nodes' -Value "$(@($script:NodeOrder).Count) / $(@($SqlInstances).Count)" -Subtitle 'reachable / configured' -Status $(if ($Unreachable.Count) { 'Critical' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Databases online' -Value "$dbOnline / $dbTotal" -Subtitle 'across all nodes' -Status $(if ($dbOnline -lt $dbTotal) { 'Warning' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Availability Groups' -Value "$(@($AgTopology | Select-Object -ExpandProperty AgName -Unique).Count)" -Subtitle "$(@($AgTopology | Where-Object Status -eq 'OK').Count) of $(@($AgTopology).Count) replicas healthy" -Status (Get-WorstStatus @($AgTopology | ForEach-Object Status))) +
    (ConvertTo-KpiHtmlText -Title 'Backup compliance' -Value "$(@($BackupCompliance | Where-Object Status -eq 'OK').Count) / $(@($BackupCompliance | Where-Object Status -ne 'Info').Count)" -Subtitle 'databases compliant' -Status (Get-WorstStatus @($BackupCompliance | ForEach-Object Status))) +
    (ConvertTo-KpiHtmlText -Title 'SQL error log' -Value "$(@($SqlData.ErrorLog).Count)" -Subtitle "entries in last $LookbackHours h" -Status $(if (@($SqlData.ErrorLog | Where-Object Level -eq 'Critical').Count) { 'Critical' } elseif (@($SqlData.ErrorLog).Count) { 'Warning' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Collection errors' -Value "$($script:CollectionErrors.Count)" -Subtitle 'sections not collected' -Status $(if ($script:CollectionErrors.Count) { 'Warning' } else { 'OK' })) +
    '</div>'

$HealthCards = "<div class='ops-card-grid'>" + (ConvertTo-AgHealthCard -Replicas $AgTopology -Databases $AgDbStates) + (ConvertTo-BackupHealthCard -Compliance $BackupCompliance)
foreach ($node in $script:NodeLabels) {
    $m = @($SqlData.Memory | Where-Object SqlNode -eq $node) | Select-Object -First 1
    $HealthCards += (ConvertTo-HostMemoryCard -Memory $m -Node $node) + (ConvertTo-SqlProcessMemoryCard -Memory $m -Node $node)
}
$HealthCards += '</div>'

$dbColumns     = @('SqlNode', 'DatabaseName', 'State', 'Status', 'RecoveryModel', 'CompatLevel', 'Owner', 'TotalSize', 'DataMB', 'LogMB', 'UserAccess', 'IsReadOnly', 'PageVerify', 'LogReuseWait', 'AgName', 'Created', 'Notes')
$backupColumns = @('DatabaseName', 'AgName', 'RecoveryModel', 'Status', 'LastFull', 'FullAge', 'FullTakenOn', 'FullStatus', 'LastDiff', 'DiffAge', 'DiffStatus', 'LastLog', 'LogAge', 'LogTakenOn', 'LogStatus', 'ReplicasQueried')

New-HTML -TitleText "SQL Server Operations Dashboard - $EnvironmentName" -FilePath $ReportPath {

    # ---------------------------------------------------------------- 1. Overview
    New-HTMLTab -Name 'Overview' {
        New-HTMLSection -Invisible -Direction column { $Css; $Banner; $Kpis }
        New-HTMLSection -Direction column -HeaderText 'Operational health' -CanCollapse { $HealthCards }
        New-HTMLSection -Direction column -HeaderText 'Status summary by area' -CanCollapse { New-SafeHTMLTable -Data $AreaSummary -PagingLength 10 }
        New-HTMLSection -Direction column -HeaderText 'SQL Server instance inventory' -CanCollapse { New-SafeHTMLTable -Data $InstanceInventory }
        New-HTMLSection -Direction column -HeaderText 'Database inventory and health' -CanCollapse { New-DatabaseGroupedTable -Data $SqlData.Databases -Columns $dbColumns }
        New-HTMLSection -Direction column -HeaderText 'Operating system per SQL Server node' -CanCollapse { New-SafeHTMLTable -Data $WinData.OS -EmptyMessage 'Operating system information could not be collected (see Investigation).' }
        New-HTMLSection -Direction column -HeaderText 'SQL Server-related Windows services' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Services -Columns @('SqlNode', 'ServiceName', 'DisplayName', 'State', 'StartMode', 'ServiceAccount', 'Status')
        }
    }

    # ---------------------------------------------------------------- 2. Always On
    New-HTMLTab -Name 'Always On' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        if (-not $HasAG) {
            New-HTMLSection -Direction column -HeaderText 'Always On Availability Groups' { "<div class='ops-empty'>No Availability Groups were detected on the queried instances. The environment is reported as $(ConvertTo-HtmlSafe $EnvironmentMode).</div>" }
        } else {
            New-HTMLSection -Direction column -HeaderText 'Availability Group health' -CanCollapse { "<div class='ops-card-grid'>$(ConvertTo-AgHealthCard -Replicas $AgTopology -Databases $AgDbStates)</div>" }
            New-HTMLSection -Direction column -HeaderText 'Availability Group topology and replica roles' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgTopology -Columns @('AgName', 'ReplicaServer', 'Role', 'AvailabilityMode', 'FailoverMode', 'SeedingMode', 'ReadableSecondary', 'BackupPriority', 'BackupPreference', 'SessionTimeout', 'EndpointUrl', 'ReportedBy'))
            }
            New-HTMLSection -Direction column -HeaderText 'Replica connectivity, synchronization state and health' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgTopology -Columns @('AgName', 'ReplicaServer', 'Role', 'ConnectedState', 'OperationalState', 'RecoveryHealth', 'SynchronizationHealth', 'Status', 'LastConnectError', 'LastConnectErrorDescription', 'LastConnectErrorTime'))
            }
            New-HTMLSection -Direction column -HeaderText 'Database synchronization' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgDbStates -Columns @('AgName', 'DatabaseName', 'ReplicaServer', 'IsPrimaryReplica', 'SynchronizationState', 'SynchronizationHealth', 'DatabaseState', 'IsSuspended', 'SuspendReason', 'Status', 'LastCommitTime', 'LastHardenedTime', 'LastRedoneTime')) -PagingLength 15
            }
            New-HTMLSection -Direction column -HeaderText "Log send and redo queues (warning: send > $AgSendQueueWarningKB KB, redo > $AgRedoQueueWarningKB KB)" -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data ($AgDbStates | Where-Object { $_.IsPrimaryReplica -ne $true }) -Columns @('AgName', 'DatabaseName', 'ReplicaServer', 'LogSendQueueKB', 'LogSendRateKBs', 'RedoQueueKB', 'RedoRateKBs', 'EstDataLossSec', 'EstRecoverySec', 'Status')) -PagingLength 15 -EmptyMessage 'No secondary database replica statistics available.'
            }
            New-HTMLSection -Direction column -HeaderText 'Availability Group listener configuration' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgListeners -Columns @('AgName', 'ListenerName', 'Port', 'IpAddress', 'SubnetMask', 'IsDhcp', 'State', 'Status')) -EmptyMessage 'No listeners are configured.'
            }
            New-HTMLSection -Direction column -HeaderText 'Independent database inventory per HA node' -CanCollapse {
                New-NodeAwareTable -Data $SqlData.Databases -Columns @('SqlNode', 'DatabaseName', 'State', 'RecoveryModel', 'AgName', 'TotalSize', 'Status') -PagingLength 15
            }
        }
    }

    # ---------------------------------------------------------------- 3. Backups
    New-HTMLTab -Name 'Backups' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Backup health' -CanCollapse { "<div class='ops-card-grid'>$(ConvertTo-BackupHealthCard -Compliance $BackupCompliance)</div>" }
        New-HTMLSection -Direction column -HeaderText "Backup freshness and compliance (full $FullBackupMaxAgeHours h, differential $DiffBackupMaxAgeHours h, log $LogBackupMaxAgeMinutes min)" -CanCollapse {
            New-BackupComplianceGroupedTable -Data $BackupCompliance -Columns $backupColumns
        }
        New-HTMLSection -Direction column -HeaderText 'Missing or outdated backup information' -CanCollapse {
            New-SafeHTMLTable -Data (Select-ReportColumns -Data $BackupGaps -Columns (@('Scope') + $backupColumns)) -EmptyMessage 'Every evaluated database meets the configured backup thresholds.'
        }
        New-HTMLSection -Direction column -HeaderText 'Backup status across queried Availability Group replicas' -CanCollapse {
            New-NodeAwareTable -Data ($SqlData.BackupStatus | Where-Object AgName) -Columns @('SqlNode', 'DatabaseName', 'AgName', 'LocalRole', 'LastFull', 'LastDiff', 'LastLog') -EmptyMessage 'No Availability Group databases on this node.'
        }
        New-HTMLSection -Direction column -HeaderText "Local backup history (last $HistoryLookbackDays days, diagnostic)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.BackupHistory -Columns @('SqlNode', 'DatabaseName', 'BackupType', 'CopyOnly', 'Started', 'Finished', 'DurationSec', 'SizeMB', 'CompressedMB', 'Device', 'ExecutedBy') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Recent database restore operations (last $HistoryLookbackDays days)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.RestoreHistory -Columns @('SqlNode', 'DatabaseName', 'RestoreDate', 'RestoreType', 'RestoredBy', 'WithReplace', 'WithRecovery', 'SourceDatabase', 'SourceServer', 'SourceBackupDate', 'Device') -EmptyMessage 'No restores recorded in the lookback period.'
        }
    }

    # ---------------------------------------------------------------- 4. Capacity & Performance
    New-HTMLTab -Name 'Capacity & Performance' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText "SQL Server CPU utilization history (last $CpuHistoryMinutes minutes)" -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-CpuSvg -Samples @($SqlData.CpuHistory | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server and operating system memory' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $m = @($SqlData.Memory | Where-Object SqlNode -eq $n) | Select-Object -First 1
                "<div class='ops-card-grid'>$(ConvertTo-SqlProcessMemoryCard -Memory $m -Node $n)$(ConvertTo-HostMemoryCard -Memory $m -Node $n)</div>"
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Database size and file allocation' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-DatabaseSizesHtml -Databases @($SqlData.Databases | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'Database files' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.DatabaseFiles -Columns @('SqlNode', 'DatabaseName', 'LogicalName', 'FileType', 'SizeMB', 'Growth', 'MaxSize', 'FileState', 'PhysicalName') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Storage volume capacity (warning < $DiskFreeWarningPct% free, critical < $DiskFreeCriticalPct% free)" -CanCollapse {
            New-NodeAwareTable -Data $SqlVolumes -Columns @('SqlNode', 'VolumePath', 'VolumeLabel', 'FileSystem', 'TotalGB', 'UsedGB', 'FreeGB', 'FreePct', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Database resource distribution' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $alloc = @($SqlData.Databases | Where-Object SqlNode -eq $n | ForEach-Object { [pscustomobject]@{ DatabaseName = $_.DatabaseName; AllocatedMB = [double]$_.DataMB + [double]$_.LogMB } })
                "<div class='ops-donut-grid'>" +
                (ConvertTo-DbResourcePie -Items $alloc -ValueProperty 'AllocatedMB' -Title 'Allocated storage per database' -Unit 'GB' -ValueFormatter { param($v) '{0:N1}' -f ($v / 1024) }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.BufferCache | Where-Object SqlNode -eq $n) -ValueProperty 'BufferCacheMB' -Title 'Buffer cache utilization per database' -Unit 'MB' -ValueFormatter { param($v) '{0:N0}' -f $v }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.DbIo | Where-Object SqlNode -eq $n) -ValueProperty 'TotalIoMB' -Title 'I/O activity per database (since startup)' -Unit 'GB' -ValueFormatter { param($v) '{0:N1}' -f ($v / 1024) }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.DbCpu | Where-Object SqlNode -eq $n) -ValueProperty 'CpuMs' -Title 'Estimated CPU distribution per database (cached plans)' -Unit 's' -ValueFormatter { param($v) '{0:N0}' -f ($v / 1000) }) +
                '</div>'
            }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server wait statistics' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $w = @($SqlData.WaitStats | Where-Object SqlNode -eq $n)
                ConvertTo-WaitDonutSvg -Waits $w
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $w -Columns @('WaitType', 'Percentage', 'WaitSeconds', 'ResourceSeconds', 'SignalSeconds', 'WaitCount', 'AvgWaitMs')) -PagingLength 15
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Database file I/O latency' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-FileIoLatencyHtml -Files @($SqlData.FileLatency | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'File I/O statistics' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.FileLatency -Columns @('SqlNode', 'DatabaseName', 'LogicalName', 'FileType', 'Reads', 'Writes', 'AvgReadMs', 'AvgWriteMs', 'TotalIoMB', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText "Transaction log usage, recovery models and VLFs (warning: > $LogUsedWarningPct% used, >= $VlfWarningCount VLFs)" -CanCollapse {
            New-NodeAwareTable -Data $LogUsage -Columns @('SqlNode', 'DatabaseName', 'RecoveryModel', 'LogSizeMB', 'LogUsedMB', 'LogUsedPct', 'LogReuseWait', 'VlfCount', 'ActiveVlfs', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server memory clerk consumption' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-MemoryClerkBarsHtml -Clerks @($SqlData.MemoryClerks | Where-Object SqlNode -eq $n) }
        }
    }

    # ---------------------------------------------------------------- 5. Cluster & Windows
    New-HTMLTab -Name 'Cluster & Windows' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Windows Server operating system information' -CanCollapse { New-SafeHTMLTable -Data $WinData.OS -EmptyMessage 'Operating system information could not be collected.' }
        New-HTMLSection -Direction column -HeaderText 'Windows Failover Cluster' -CanCollapse {
            New-SafeHTMLTable -Data $ClusterData.Clusters -EmptyMessage $(if ($CollectClusterInfo) { 'No Windows Failover Cluster was detected for the SQL Server hosts.' } else { 'Cluster collection is disabled (-CollectClusterInfo $false).' })
        }
        New-HTMLSection -Direction column -HeaderText 'Cluster node health' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Nodes -EmptyMessage 'No cluster nodes.' }
        New-HTMLSection -Direction column -HeaderText 'Cluster resource groups and ownership' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Groups -EmptyMessage 'No cluster groups.' }
        New-HTMLSection -Direction column -HeaderText 'Cluster resources' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Resources -EmptyMessage 'No cluster resources.' -PagingLength 15 }
        New-HTMLSection -Direction column -HeaderText "Cluster-related DNS and Active Directory events (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $ClusterDnsAdEvents -Columns @('SqlNode', 'TimeCreated', 'Category', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage 'No cluster, DNS or Active Directory warnings or errors.'
        }
        New-HTMLSection -Direction column -HeaderText "Windows System event log - critical, error, warning (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data ($WinData.Events | Where-Object LogName -eq 'System') -Columns @('SqlNode', 'TimeCreated', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage $(if ($CollectEventViewer) { 'No matching events.' } else { 'Event Viewer collection is disabled.' })
        }
        New-HTMLSection -Direction column -HeaderText "Windows Application event log - critical, error, warning (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data ($WinData.Events | Where-Object LogName -eq 'Application') -Columns @('SqlNode', 'TimeCreated', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage $(if ($CollectEventViewer) { 'No matching events.' } else { 'Event Viewer collection is disabled.' })
        }
        New-HTMLSection -Direction column -HeaderText 'Windows service availability' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Services -Columns @('SqlNode', 'ServiceName', 'DisplayName', 'State', 'StartMode', 'ServiceAccount', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Windows storage volumes' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Volumes -Columns @('SqlNode', 'Volume', 'Label', 'FileSystem', 'CapacityGB', 'FreeGB', 'FreePct', 'BlockSizeKB', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Windows local accounts' -CanCollapse {
            New-NodeAwareTable -Data $WinData.LocalUsers -Columns @('SqlNode', 'AccountName', 'FullName', 'Disabled', 'LockedOut', 'PasswordRequired', 'PasswordExpires', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Local Administrators group membership' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Admins -Columns @('SqlNode', 'Member', 'MemberType', 'SID')
        }
    }

    # ---------------------------------------------------------------- 6. SQL Errors
    New-HTMLTab -Name 'SQL Errors' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText "Grouped SQL Server error messages (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $ErrorGroups -Columns @('SqlNode', 'Occurrences', 'Level', 'FirstSeen', 'LastSeen', 'SampleMessage') -EmptyMessage 'No errors or warnings in the SQL Server error log.'
        }
        New-HTMLSection -Direction column -HeaderText 'Recent operational errors and warnings' -CanCollapse {
            New-NodeAwareTable -Data ($SqlData.ErrorLog | Where-Object Level -in 'Critical', 'Error' | Sort-Object LogDate -Descending | Select-Object -First 200) -Columns @('SqlNode', 'LogDate', 'Level', 'ProcessInfo', 'Message') -EmptyMessage 'No errors in the lookback period.'
        }
        New-HTMLSection -Direction column -HeaderText "SQL Server error log analysis - all matched entries (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.ErrorLog -Columns @('SqlNode', 'LogDate', 'Level', 'ProcessInfo', 'Message') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Suspect pages (msdb.dbo.suspect_pages)' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.SuspectPages -Columns @('SqlNode', 'DatabaseName', 'FileId', 'PageId', 'EventType', 'ErrorCount', 'LastUpdate') -EmptyMessage 'No suspect pages recorded.'
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server memory dump indicators' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.MemoryDumps -Columns @('SqlNode', 'Created', 'SizeMB', 'DumpFile') -EmptyMessage 'No memory dumps recorded.'
        }
        New-HTMLSection -Direction column -HeaderText "Windows critical events and operational failures (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $WindowsCritical -Columns @('SqlNode', 'TimeCreated', 'LogName', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage 'No Windows critical events or failures.'
        }
    }

    # ---------------------------------------------------------------- 7. Security & Configuration
    New-HTMLTab -Name 'Security & Configuration' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Database security compliance' -CanCollapse {
            New-NodeAwareTable -Data $SecurityChecks -Columns @('SqlNode', 'Check', 'Result', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Database security configuration' -CanCollapse {
            New-NodeAwareTable -Data $DbSecurity -Columns @('SqlNode', 'DatabaseName', 'Owner', 'Trustworthy', 'DbChaining', 'TdeEnabled', 'Encryption', 'Containment', 'BrokerEnabled', 'Issues', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Orphaned database users' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.OrphanedUsers -Columns @('SqlNode', 'DatabaseName', 'UserName', 'UserType', 'AuthenticationType', 'CreateDate') -EmptyMessage 'No orphaned users found.'
        }
        New-HTMLSection -Direction column -HeaderText 'Privileged database role membership' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.DbRoleMembers -Columns @('SqlNode', 'DatabaseName', 'RoleName', 'MemberName', 'MemberType') -PagingLength 15 -EmptyMessage 'No members in privileged database roles (other than dbo).'
        }
        New-HTMLSection -Direction column -HeaderText 'Privileged server role membership' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.ServerRoleMembers -Columns @('SqlNode', 'ServerRole', 'MemberName', 'MemberType', 'IsDisabled') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Guest access in user databases' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.GuestAccess -Columns @('SqlNode', 'DatabaseName', 'PermissionState') -EmptyMessage 'guest has no CONNECT permission in user databases.'
        }
        New-HTMLSection -Direction column -HeaderText "Certificate expiration (warning within $CertificateExpiryWarningDays days)" -CanCollapse {
            New-NodeAwareTable -Data $Certificates -Columns @('SqlNode', 'DatabaseName', 'CertificateName', 'Subject', 'ExpiryDate', 'DaysToExpiry', 'KeyEncryption', 'Status') -EmptyMessage 'No user certificates found.'
        }
        if (@($script:NodeLabels).Count -gt 1) {
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - server properties' -CanCollapse { New-SafeHTMLTable -Data $PropertyCompare -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - sp_configure (differences first)' -CanCollapse { New-SafeHTMLTable -Data ($ConfigCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - logins' -CanCollapse { New-SafeHTMLTable -Data ($LoginCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - SQL Server Agent jobs' -CanCollapse { New-SafeHTMLTable -Data ($JobCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - global trace flags and linked servers' -CanCollapse {
                New-SafeHTMLTable -Data $TraceFlagCompare -StatusColumns 'Comparison' -EmptyMessage 'No global trace flags enabled on any node.'
                New-SafeHTMLTable -Data $LinkedCompare -StatusColumns 'Comparison' -EmptyMessage 'No linked servers on any node.'
            }
        } else {
            New-HTMLSection -Direction column -HeaderText 'SQL Server configuration (sp_configure)' -CanCollapse { New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.Configurations -Columns @('Setting', 'Value')) -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Global trace flags and linked servers' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.TraceFlags -Columns @('TraceFlag', 'Global', 'Session')) -EmptyMessage 'No global trace flags enabled.'
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.LinkedServers -Columns @('LinkedServer', 'Product', 'Provider', 'DataSource')) -EmptyMessage 'No linked servers.'
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Windows user and administrator accounts' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Admins -Columns @('SqlNode', 'Member', 'MemberType', 'SID') -EmptyMessage 'Administrator membership could not be collected.'
            New-NodeAwareTable -Data $WinData.LocalUsers -Columns @('SqlNode', 'AccountName', 'Disabled', 'PasswordRequired', 'PasswordExpires', 'Status')
        }
    }

    # ---------------------------------------------------------------- 8. Investigation
    New-HTMLTab -Name 'Investigation' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Information requiring further administrative analysis' -CanCollapse {
            New-SafeHTMLTable -Data ($script:Findings | Sort-Object @{ Expression = { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } } }, Area, SqlNode) -PagingLength 25 -EmptyMessage 'No findings - every check is within its configured threshold.'
        }
        New-HTMLSection -Direction column -HeaderText "SQL Server Agent job duration trends (last $HistoryLookbackDays days)" -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-JobDurationTrendHtml -Trends @($JobTrends | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server Agent job statistics' -CanCollapse {
            New-NodeAwareTable -Data $JobTrends -Columns @('SqlNode', 'JobName', 'Runs', 'Failed', 'AvgDuration', 'MaxDuration', 'LastDuration', 'LastVsAvgPct', 'LastRun', 'LastOutcome', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Database integrity verification history - DBCC CHECKDB (threshold $CheckDbMaxAgeDays days)" -CanCollapse {
            New-NodeAwareTable -Data $CheckDbStatus -Columns @('SqlNode', 'DatabaseName', 'State', 'LastKnownGood', 'Age', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'DBCC CHECKDB evidence from the SQL Server error log' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.CheckDbLog -Columns @('SqlNode', 'LogDate', 'Status', 'Message') -EmptyMessage 'No DBCC CHECKDB entries in the current or previous error log.'
        }
        New-HTMLSection -Direction column -HeaderText "Long-running requests (> $LongRunningQuerySeconds s)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.LongRunning -Columns @('SqlNode', 'SessionId', 'LoginName', 'HostName', 'ProgramName', 'DatabaseName', 'Command', 'RequestStatus', 'StartTime', 'ElapsedSec', 'CpuMs', 'LogicalReads', 'WaitType', 'BlockedBy', 'SqlText') -EmptyMessage 'No long-running requests at collection time.'
        }
        New-HTMLSection -Direction column -HeaderText "Blocking (waits > $BlockingWaitSeconds s)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.Blocking -Columns @('SqlNode', 'BlockedSession', 'BlockingSession', 'DatabaseName', 'WaitType', 'WaitSec', 'WaitResource', 'BlockedLogin', 'BlockingLogin', 'BlockingHost', 'BlockingProgram', 'BlockedSql', 'BlockingSql') -EmptyMessage 'No blocking at collection time.'
        }
        New-HTMLSection -Direction column -HeaderText 'Collection errors and permission issues' -CanCollapse {
            New-SafeHTMLTable -Data $script:CollectionErrors -PagingLength 15 -EmptyMessage 'All sections were collected successfully.'
        }
        if ($EnableMetricHistory) {
            New-HTMLSection -Direction column -HeaderText "Monitoring history (retention $MetricHistoryRetentionDays days)" -CanCollapse {
                New-SafeHTMLTable -Data ($HistoryRows | Sort-Object Timestamp -Descending) -PagingLength 25
            }
        }
    }
}

Write-Host ''
Write-Host "Dashboard written to: $ReportPath" -ForegroundColor Green
Write-Host ("Overall status: {0} - {1} critical, {2} warning findings, {3} collection error(s)." -f $overall, $critTotal, $warnTotal, $script:CollectionErrors.Count)
if ($EnableMetricHistory) { Write-Host "Metric history: $(Join-Path $OutputFolder 'History\MetricHistory.csv')" }
#requires -Version 5.1
<#
.SYNOPSIS
    SQL Server Operations Dashboard - PowerShell monitoring and reporting engine.

.DESCRIPTION
    Collects read-only operational information from one or more SQL Server instances
    (standalone or Always On Availability Group nodes) and the Windows hosts they run on,
    then produces a single standalone HTML dashboard built with the PSWriteHTML module.

    The dashboard has eight tabs:
      1. Overview                  - KPIs, health cards, instance and database inventory, OS, services
      2. Always On                 - AG topology, replica connectivity/sync, database sync, queues, listeners
      3. Backups                   - Freshness, compliance (across AG replicas), history, restores, gaps
      4. Capacity & Performance    - CPU history, memory, storage, DB sizes, resource donuts, waits, I/O, log/VLF
      5. Cluster & Windows         - Failover Cluster health, groups, resources, OS, event logs, services, accounts
      6. SQL Errors                - Error log, grouped messages, suspect pages, memory dumps, Windows critical events
      7. Security & Configuration  - DB security, orphaned users, privileged roles, compliance, certs, HA config diff
      8. Investigation             - Agent job duration trends, CHECKDB evidence, long-running / blocking, findings,
                                     collection and permission errors, optional metric history

    READ-ONLY: every SQL statement is a SELECT against DMVs / catalog views / msdb history, or a
    read-only system procedure (xp_readerrorlog, DBCC TRACESTATUS). Nothing is created, altered or
    deleted in any user or system database, and no SQL Server configuration is changed.

.NOTES
    Script location : any folder, on any drive (e.g. C:\Apps\SQLServer-Operations\SQLServer-Operations.ps1)
    Report output   : <script folder>\Output\SQLServer-Operations.html   (override with -OutputFolder)
    History (opt.)  : <script folder>\Output\History\MetricHistory.csv

    Requirements
      - Windows PowerShell 5.1
      - PSWriteHTML module            (Install-Module PSWriteHTML -Scope CurrentUser)
      - Windows-integrated authentication to every SQL Server endpoint
      - SQL permissions: VIEW SERVER STATE, VIEW ANY DEFINITION, read on msdb history tables;
        sysadmin is simplest (error log reading and some security checks need it)
      - Windows permissions for remote CIM/WMI, Event Log and service queries on each host
      - FailoverClusters PowerShell module (RSAT) for cluster sections
      - SQL Server 2016 SP2 or later for every section (VLF counts and CHECKDB properties
        need it); older builds still report, failed sections appear under Investigation.

EXECUTIION

Install-Module PSWriteHTML -Scope CurrentUser   # one time



cd A:\SQLServer-Operations\output
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\SQLServer-Operations.ps1 -SqlInstances 'hlmm', 'abcd' #list all servers with comma


.EXAMPLE
    .\SQLServer-Operations.ps1 -SqlInstances 'SQLNODE1','SQLNODE2'

.EXAMPLE
    .\SQLServer-Operations.ps1 -SqlInstances 'SQLPROD01\INST1' -OutputFolder 'D:\Reports' -EnableMetricHistory $true -TimestampedFileName
#>
[CmdletBinding()]
param(
    # ---------------- SQL Server endpoints and instance identification ----------------
    [string[]] $SqlInstances = @('SQLNODE1', 'SQLNODE2'),
    [string]   $EnvironmentName = 'SQL Server Environment',

    # ---------------- SQL connection and query timeouts (seconds) ----------------
    [int] $ConnectionTimeoutSeconds = 15,
    [int] $QueryTimeoutSeconds = 120,

    # ---------------- Backup age thresholds ----------------
    [int] $FullBackupMaxAgeHours = 168,
    [int] $DiffBackupMaxAgeHours = 26,
    [int] $LogBackupMaxAgeMinutes = 60,

    # ---------------- Disk space thresholds (percent free) ----------------
    [int] $DiskFreeWarningPct = 20,
    [int] $DiskFreeCriticalPct = 10,

    # ---------------- Long-running query and blocking thresholds (seconds) ----------------
    [int] $LongRunningQuerySeconds = 300,
    [int] $BlockingWaitSeconds = 30,

    # ---------------- I/O latency and transaction log thresholds ----------------
    [int] $IoLatencyWarningMs = 20,
    [int] $IoLatencyCriticalMs = 50,
    [int] $VlfWarningCount = 1000,
    [int] $LogUsedWarningPct = 75,

    # ---------------- DBCC CHECKDB history age ----------------
    [int] $CheckDbMaxAgeDays = 7,

    # ---------------- Certificate expiration warning ----------------
    [int] $CertificateExpiryWarningDays = 30,

    # ---------------- Availability Group queue thresholds (KB) ----------------
    [long] $AgSendQueueWarningKB = 10240,
    [long] $AgRedoQueueWarningKB = 10240,

    # ---------------- Monitoring lookback periods ----------------
    [int] $LookbackHours = 24,          # error log, Windows events
    [int] $HistoryLookbackDays = 14,    # backup/restore history, Agent job history
    [int] $CpuHistoryMinutes = 240,     # CPU ring buffer samples (1 per minute)
    [int] $MaxEventsPerLog = 300,

    # ---------------- Windows Event Viewer / Failover Cluster collection ----------------
    [bool] $CollectEventViewer = $true,
    [bool] $CollectClusterInfo = $true,

    # ---------------- Optional metric history retention ----------------
    [bool] $EnableMetricHistory = $false,
    [int]  $MetricHistoryRetentionDays = 30,

    # ---------------- HTML report output location and filename ----------------
    [string] $OutputFolder = '',          # default: an Output folder next to this script
    [string] $ReportFileName = 'SQLServer-Operations.html',
    [switch] $TimestampedFileName
)

# =====================================================================================
#  Initialisation
# =====================================================================================
$ErrorActionPreference = 'Continue'
$script:ReportStart      = Get-Date
$script:CollectionErrors = New-Object System.Collections.Generic.List[object]
$script:Findings         = New-Object System.Collections.Generic.List[object]
$script:NodeMap          = @{}     # SqlInstance -> display label
$script:NodeOrder        = @()     # reachable SqlInstances, in the order given
$script:NodeLabels       = @()     # display labels of reachable nodes
$script:Inv              = [Globalization.CultureInfo]::InvariantCulture

if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
    throw 'The PSWriteHTML module is required. Install it with: Install-Module PSWriteHTML -Scope CurrentUser'
}
Import-Module PSWriteHTML -ErrorAction Stop

# Default output: an "Output" folder beside the .ps1 (works on any drive / any server).
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $scriptDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptDir) -and $MyInvocation.MyCommand.Path) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
    if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = (Get-Location).ProviderPath }
    $OutputFolder = Join-Path $scriptDir 'Output'
}
$qualifier = Split-Path -Path $OutputFolder -Qualifier -ErrorAction SilentlyContinue
if ($qualifier -and $qualifier -match '^[A-Za-z]:$' -and -not (Test-Path -LiteralPath "$qualifier\")) {
    throw "Drive $qualifier is not available on $env:COMPUTERNAME. Pass -OutputFolder with a valid path, or omit it to write next to the script."
}
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

if ($TimestampedFileName) {
    $base = [IO.Path]::GetFileNameWithoutExtension($ReportFileName)
    $ReportFileName = '{0}_{1}.html' -f $base, (Get-Date -Format 'yyyyMMdd_HHmm')
}
$ReportPath = Join-Path $OutputFolder $ReportFileName

# =====================================================================================
#  Generic helpers
# =====================================================================================
function ConvertTo-HtmlSafe {
    param([object] $Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function F {
    # Invariant-culture number for SVG/CSS (avoids "12,5" in comma-decimal locales)
    param([double] $Value, [string] $Format = '0.##')
    return $Value.ToString($Format, $script:Inv)
}

function Format-SizeMB {
    param([double] $MB)
    if ($MB -ge 1048576) { return ('{0:N2} TB' -f ($MB / 1048576)) }
    if ($MB -ge 1024)    { return ('{0:N2} GB' -f ($MB / 1024)) }
    return ('{0:N0} MB' -f $MB)
}

function Format-Duration {
    param([double] $Seconds)
    $ts = [TimeSpan]::FromSeconds([math]::Round($Seconds))
    if ($ts.TotalHours -ge 1)   { return ('{0}h {1:00}m {2:00}s' -f [math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds) }
    if ($ts.TotalMinutes -ge 1) { return ('{0}m {1:00}s' -f $ts.Minutes, $ts.Seconds) }
    return ('{0}s' -f $ts.Seconds)
}

function Format-Age {
    param([object] $Date)
    if ($null -eq $Date) { return 'Never' }
    $span = (Get-Date) - [datetime]$Date
    if ($span.TotalDays -ge 1)  { return ('{0:N1} days' -f $span.TotalDays) }
    if ($span.TotalHours -ge 1) { return ('{0:N1} hours' -f $span.TotalHours) }
    return ('{0:N0} min' -f $span.TotalMinutes)
}

function Get-WorstStatus {
    param([string[]] $Status)
    if ($Status -contains 'Critical') { return 'Critical' }
    if ($Status -contains 'Warning')  { return 'Warning' }
    if ($Status -contains 'OK')       { return 'OK' }
    return 'Info'
}

function Get-StatusClass {
    param([string] $Status)
    switch ($Status) { 'OK' { 'st-ok' } 'Warning' { 'st-warning' } 'Critical' { 'st-critical' } default { 'st-info' } }
}

function Get-StatusColor {
    param([string] $Status)
    switch ($Status) { 'OK' { '#2e9d5b' } 'Warning' { '#e0a100' } 'Critical' { '#d64545' } default { '#3a7bd5' } }
}

function Add-CollectionError {
    param([string] $Node, [string] $Section, [string] $Message)
    $category = 'Query / collection'
    if ($Message -match 'permission|denied|not authorized|login failed|VIEW SERVER STATE|does not have') { $category = 'Permission' }
    elseif ($Message -match 'timeout|network|not found or was not accessible|RPC server|WinRM|unavailable') { $category = 'Connectivity' }
    $script:CollectionErrors.Add([pscustomobject]@{
        Time     = Get-Date
        SqlNode  = $Node
        Section  = $Section
        Category = $category
        Message  = $Message
    })
}

function Add-Finding {
    param(
        [ValidateSet('Critical', 'Warning', 'Info')][string] $Status,
        [string] $Area, [string] $Node, [string] $Item, [string] $Detail
    )
    $script:Findings.Add([pscustomobject]@{
        Status  = $Status
        Area    = $Area
        SqlNode = $Node
        Item    = $Item
        Detail  = $Detail
    })
}

# =====================================================================================
#  SQL Server data collection functions
# =====================================================================================
function Get-ReportNodeLabel {
    <# Returns the display label used for a SQL Server node, e.g. "SQLNODE1 (Primary)". #>
    param([string] $SqlInstance)
    if ($script:NodeMap.ContainsKey($SqlInstance)) { return $script:NodeMap[$SqlInstance] }
    return $SqlInstance
}

function Invoke-DbQuery {
    <# Executes a read-only monitoring query and returns PSCustomObjects tagged with SqlNode. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $SqlInstance,
        [Parameter(Mandatory)][string] $Query,
        [string]    $Database = 'master',
        [hashtable] $Parameters,
        [string]    $Section = 'General',
        [int]       $Timeout = $QueryTimeoutSeconds
    )
    $connString = "Server=$SqlInstance;Database=$Database;Integrated Security=SSPI;" +
                  "Application Name=SQLServer-Operations-Dashboard;Connect Timeout=$ConnectionTimeoutSeconds"
    $conn = New-Object System.Data.SqlClient.SqlConnection $connString
    $table = New-Object System.Data.DataTable
    try {
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = $Timeout
        if ($Parameters) {
            foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue("@$k", $Parameters[$k]) }
        }
        $reader = $cmd.ExecuteReader()
        $table.Load($reader)
        $reader.Close()
    }
    catch {
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
        Add-CollectionError -Node (Get-ReportNodeLabel $SqlInstance) -Section $Section -Message $msg
        return
    }
    finally { $conn.Dispose() }

    $label   = Get-ReportNodeLabel $SqlInstance
    $columns = @($table.Columns | ForEach-Object { $_.ColumnName })
    foreach ($row in $table.Rows) {
        $o = [ordered]@{ SqlNode = $label }
        foreach ($c in $columns) {
            $v = $row[$c]
            if ($v -is [DBNull]) { $v = $null }
            $o[$c] = $v
        }
        [pscustomobject]$o
    }
}

function Get-DatabaseCategory {
    <# Classifies databases for structured reporting. #>
    param([string] $DatabaseName, [string] $AgName, [bool] $IsDistributor)
    if ($DatabaseName -in @('master', 'model', 'msdb', 'tempdb')) { return 'System databases' }
    if ($IsDistributor)                     { return 'Replication distribution' }
    if ($AgName)                            { return "Availability Group: $AgName" }
    if ($DatabaseName -eq 'SSISDB')         { return 'Integration Services catalog' }
    if ($DatabaseName -like 'ReportServer*'){ return 'Reporting Services' }
    return 'User databases (standalone)'
}

function Select-ReportColumns {
    <# Selects the relevant columns for report presentation. #>
    param([object[]] $Data, [object[]] $Columns, [switch] $ExcludeNode)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return @() }
    $cols = $Columns
    if ($ExcludeNode) { $cols = @($Columns | Where-Object { $_ -ne 'SqlNode' }) }
    return @($rows | Select-Object -Property $cols)
}

function Get-PerDatabaseQuery {
    <# Builds a read-only batch that runs $InnerQuery in every accessible online database.
       Placeholders: {DB} = QUOTENAME(db), {DBNAME} = db name for use inside N'...' literals. #>
    param([string] $TableDefinition, [string] $InnerQuery, [string] $OrderBy = '')
    $tpl = $InnerQuery.Replace("'", "''")
    return @"
SET NOCOUNT ON;
DECLARE @r TABLE ($TableDefinition);
DECLARE @db sysname, @sql nvarchar(max);
DECLARE @tpl nvarchar(max) = N'$tpl';
DECLARE dbs CURSOR LOCAL FAST_FORWARD FOR
    SELECT d.name FROM sys.databases d
    WHERE d.state = 0 AND d.name <> N'tempdb' AND HAS_DBACCESS(d.name) = 1;
OPEN dbs; FETCH NEXT FROM dbs INTO @db;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = REPLACE(REPLACE(@tpl, N'{DB}', QUOTENAME(@db)), N'{DBNAME}', REPLACE(@db, N'''', N''''''));
    BEGIN TRY INSERT INTO @r EXEC sys.sp_executesql @sql; END TRY BEGIN CATCH END CATCH;
    FETCH NEXT FROM dbs INTO @db;
END
CLOSE dbs; DEALLOCATE dbs;
SELECT * FROM @r $OrderBy;
"@
}

# =====================================================================================
#  HTML reporting functions
# =====================================================================================
$script:StatusPalette = @(
    @{ Values = @('OK', 'Healthy', 'HEALTHY', 'CONNECTED', 'SYNCHRONIZED', 'ONLINE', 'Online', 'Running', 'Up', 'Succeeded', 'Match', 'Present', 'Enabled'); Back = '#d8f3e3'; Fore = '#1e6b3e' }
    @{ Values = @('Warning', 'PARTIALLY_HEALTHY', 'SYNCHRONIZING', 'Retry', 'Canceled', 'Paused', 'Degraded', 'Overdue', 'NOT_HEALTHY_SECONDARY'); Back = '#fff1cc'; Fore = '#7a5b00' }
    @{ Values = @('Critical', 'NOT_HEALTHY', 'DISCONNECTED', 'NOT SYNCHRONIZING', 'NOT_SYNCHRONIZING', 'Stopped', 'Down', 'Failed', 'Missing', 'Mismatch', 'SUSPECT', 'OFFLINE', 'Offline', 'RECOVERY_PENDING', 'EMERGENCY', 'Never'); Back = '#f9d6d6'; Fore = '#8a1f1f' }
    @{ Values = @('Info', 'N/A', 'Unknown'); Back = '#e3edfb'; Fore = '#1f4f8f' }
)

function New-SafeHTMLTable {
    <# Generates a report table; renders a friendly message when there is no data. #>
    param(
        [object[]] $Data,
        [string]   $EmptyMessage = 'No data was returned for this section.',
        [int]      $PagingLength = 10,
        [string[]] $StatusColumns
    )
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>$(ConvertTo-HtmlSafe $EmptyMessage)</div>" }

    if (-not $StatusColumns) {
        $StatusColumns = @($rows[0].PSObject.Properties.Name | Where-Object { $_ -match '(Status|Health|State|Outcome|Compliance|Comparison|Role)$' })
    }
    $present = @{}
    foreach ($c in $StatusColumns) { $present[$c] = @($rows | ForEach-Object { [string]$_.$c } | Sort-Object -Unique) }
    $palette = $script:StatusPalette

    New-HTMLTable -DataTable $rows -Filtering -PagingLength $PagingLength -ScrollX -Buttons @('copyHtml5', 'excelHtml5', 'csvHtml5') {
        foreach ($c in $StatusColumns) {
            foreach ($p in $palette) {
                foreach ($v in $p.Values) {
                    if ($present[$c] -ccontains $v) {
                        New-TableCondition -Name $c -ComparisonType string -Operator eq -Value $v -BackgroundColor $p.Back -Color $p.Fore
                    }
                }
            }
        }
    }
}

function New-NodeAwarePanels {
    <# Renders the same content block side by side for every SQL Server node. #>
    param([scriptblock] $ContentForNode)
    $labels = @($script:NodeLabels)
    New-HTMLSection -Invisible {
        foreach ($lbl in $labels) {
            New-HTMLPanel {
                "<div class='ops-node-title'>$(ConvertTo-HtmlSafe $lbl)</div>"
                & $ContentForNode $lbl
            }
        }
    }
}

function New-NodeAwareTable {
    <# Organizes information per SQL Server node (side by side in HA environments). #>
    param(
        [object[]] $Data,
        [object[]] $Columns,
        [int]      $PagingLength = 10,
        [string]   $EmptyMessage = 'No data was returned for this node.'
    )
    $rows = @($Data | Where-Object { $null -ne $_ })
    if (@($script:NodeLabels).Count -le 1) {
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $rows -Columns $Columns -ExcludeNode) -PagingLength $PagingLength -EmptyMessage $EmptyMessage
        return
    }
    New-NodeAwarePanels {
        param($lbl)
        $nodeRows = @($rows | Where-Object { $_.SqlNode -eq $lbl })
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $nodeRows -Columns $Columns -ExcludeNode) -PagingLength $PagingLength -EmptyMessage $EmptyMessage
    }
}

function New-DatabaseGroupedTable {
    <# Groups database-related information by database category. #>
    param([object[]] $Data, [object[]] $Columns, [int] $PagingLength = 10)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No database information was collected.</div>" }
    $order = @{ 'System databases' = 0 }
    $groups = @($rows | Group-Object Category | Sort-Object @{ Expression = { if ($order.ContainsKey($_.Name)) { 0 } else { 1 } } }, Name)
    foreach ($g in $groups) {
        $crit = @($g.Group | Where-Object Status -eq 'Critical').Count
        $warn = @($g.Group | Where-Object Status -eq 'Warning').Count
        $st   = Get-WorstStatus @($g.Group | ForEach-Object Status)
        "<div class='ops-group-head'><span class='ops-pill $(Get-StatusClass $st)'>$st</span> $(ConvertTo-HtmlSafe $g.Name) <span class='ops-muted'>($($g.Count) rows, $crit critical, $warn warning)</span></div>"
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $g.Group -Columns $Columns) -PagingLength $PagingLength
    }
}

function New-BackupComplianceGroupedTable {
    <# Presents backup compliance results grouped by scope. #>
    param([object[]] $Data, [object[]] $Columns)
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No backup compliance information was collected.</div>" }
    foreach ($g in @($rows | Group-Object Scope | Sort-Object Name)) {
        $ok   = @($g.Group | Where-Object Status -eq 'OK').Count
        $st   = Get-WorstStatus @($g.Group | ForEach-Object Status)
        "<div class='ops-group-head'><span class='ops-pill $(Get-StatusClass $st)'>$st</span> $(ConvertTo-HtmlSafe $g.Name) <span class='ops-muted'>($ok of $($g.Count) compliant)</span></div>"
        New-SafeHTMLTable -Data (Select-ReportColumns -Data $g.Group -Columns $Columns) -PagingLength 15
    }
}

# =====================================================================================
#  Graphical monitoring functions (pure SVG / HTML, no external chart libraries)
# =====================================================================================
$script:ChartPalette = @('#3a7bd5', '#2e9d5b', '#e0a100', '#d64545', '#8e5bd6', '#17a2b8', '#ff7f50', '#6c757d', '#c2185b', '#00897b')

function New-DonutSvgInternal {
    param(
        [object[]] $Items, [string] $LabelProperty, [string] $ValueProperty,
        [string] $Title, [string] $Unit = '', [int] $Top = 8, [scriptblock] $ValueFormatter
    )
    if (-not $ValueFormatter) { $ValueFormatter = { param($v) '{0:N1}' -f $v } }
    $data = @($Items | Where-Object { $null -ne $_ -and [double]$_.$ValueProperty -gt 0 } |
              Sort-Object -Property @{ Expression = { [double]$_.$ValueProperty } } -Descending)
    if ($data.Count -eq 0) { return "<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-empty'>No data available.</div></div>" }
    if ($data.Count -gt $Top) {
        $others = ($data | Select-Object -Skip $Top | Measure-Object -Property $ValueProperty -Sum).Sum
        $data = @($data | Select-Object -First $Top) + @([pscustomobject]@{ $LabelProperty = 'Others'; $ValueProperty = $others })
    }
    $total = [double](($data | Measure-Object -Property $ValueProperty -Sum).Sum)
    $r = 60; $circ = 2 * [math]::PI * $r; $offset = 0.0; $i = 0
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-donut-body'>")
    [void]$sb.Append("<svg viewBox='0 0 160 160' width='160' height='160' role='img'><g transform='rotate(-90 80 80)'>")
    [void]$sb.Append("<circle cx='80' cy='80' r='$r' fill='none' stroke='#eef1f5' stroke-width='22'/>")
    $legend = New-Object System.Text.StringBuilder
    foreach ($d in $data) {
        $v   = [double]$d.$ValueProperty
        $len = $circ * $v / $total
        $col = $script:ChartPalette[$i % $script:ChartPalette.Count]
        $pct = 100 * $v / $total
        $lbl = ConvertTo-HtmlSafe $d.$LabelProperty
        $val = & $ValueFormatter $v
        [void]$sb.Append("<circle cx='80' cy='80' r='$r' fill='none' stroke='$col' stroke-width='22' stroke-dasharray='$(F $len) $(F ($circ - $len))' stroke-dashoffset='$(F (-$offset))'><title>$lbl : $val $Unit ($('{0:N1}' -f $pct)%)</title></circle>")
        [void]$legend.Append("<li><span class='ops-swatch' style='background:$col'></span><span class='ops-legend-label' title='$lbl'>$lbl</span><span class='ops-legend-val'>$val $Unit &middot; $('{0:N1}' -f $pct)%</span></li>")
        $offset += $len; $i++
    }
    [void]$sb.Append("</g><text x='80' y='78' text-anchor='middle' class='ops-donut-total'>$(& $ValueFormatter $total)</text>")
    [void]$sb.Append("<text x='80' y='96' text-anchor='middle' class='ops-donut-unit'>$(ConvertTo-HtmlSafe $Unit) total</text></svg>")
    [void]$sb.Append("<ul class='ops-legend'>$($legend.ToString())</ul></div></div>")
    return $sb.ToString()
}

function ConvertTo-CpuSvg {
    <# Generates the SQL Server / other process CPU utilization chart. #>
    param([object[]] $Samples, [string] $Title = 'CPU utilization history')
    $s = @($Samples | Where-Object { $null -ne $_.EventTime } | Sort-Object EventTime)
    if ($s.Count -lt 2) { return "<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-empty'>Not enough CPU samples available.</div></div>" }
    $w = 760; $h = 230; $l = 40; $rgt = 12; $t = 14; $b = 36
    $pw = $w - $l - $rgt; $ph = $h - $t - $b; $n = $s.Count
    $sqlPts = New-Object System.Collections.Generic.List[string]
    $othPts = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $n; $i++) {
        $x  = $l + ($pw * $i / ($n - 1))
        $ys = $t + $ph * (1 - [double]$s[$i].SQLProcessUtilization / 100)
        $yo = $t + $ph * (1 - [double]$s[$i].OtherProcessUtilization / 100)
        $sqlPts.Add("$(F $x),$(F $ys)"); $othPts.Add("$(F $x),$(F $yo)")
    }
    $avgSql = ($s | Measure-Object -Property SQLProcessUtilization -Average).Average
    $maxSql = ($s | Measure-Object -Property SQLProcessUtilization -Maximum).Maximum
    $avgOth = ($s | Measure-Object -Property OtherProcessUtilization -Average).Average
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>$(ConvertTo-HtmlSafe $Title)</div>")
    [void]$sb.Append("<svg viewBox='0 0 $w $h' class='ops-svg-wide' role='img'>")
    foreach ($g in 0, 25, 50, 75, 100) {
        $y = $t + $ph * (1 - $g / 100)
        [void]$sb.Append("<line x1='$l' x2='$($w - $rgt)' y1='$(F $y)' y2='$(F $y)' stroke='#e5e7eb' stroke-width='1'/>")
        [void]$sb.Append("<text x='$($l - 6)' y='$(F ($y + 4))' text-anchor='end' class='ops-axis'>$g%</text>")
    }
    $area = "$(F $l),$(F ($t + $ph)) " + ($sqlPts -join ' ') + " $(F ($l + $pw)),$(F ($t + $ph))"
    [void]$sb.Append("<polygon points='$area' fill='#3a7bd5' fill-opacity='0.15'/>")
    [void]$sb.Append("<polyline points='$($sqlPts -join ' ')' fill='none' stroke='#3a7bd5' stroke-width='2'/>")
    [void]$sb.Append("<polyline points='$($othPts -join ' ')' fill='none' stroke='#e0a100' stroke-width='1.5' stroke-dasharray='4 3'/>")
    foreach ($idx in 0, [int](($n - 1) / 2), ($n - 1)) {
        $x = $l + ($pw * $idx / ($n - 1))
        $anchor = if ($idx -eq 0) { 'start' } elseif ($idx -eq $n - 1) { 'end' } else { 'middle' }
        [void]$sb.Append("<text x='$(F $x)' y='$($h - 14)' text-anchor='$anchor' class='ops-axis'>$(([datetime]$s[$idx].EventTime).ToString('dd-MMM HH:mm'))</text>")
    }
    [void]$sb.Append("</svg><div class='ops-chart-legend'><span><span class='ops-swatch' style='background:#3a7bd5'></span>SQL Server</span>")
    [void]$sb.Append("<span><span class='ops-swatch' style='background:#e0a100'></span>Other processes</span>")
    [void]$sb.Append("<span class='ops-muted'>SQL avg $('{0:N0}' -f $avgSql)% &middot; SQL max $maxSql% &middot; other avg $('{0:N0}' -f $avgOth)% &middot; $n samples</span></div></div>")
    return $sb.ToString()
}

function ConvertTo-WaitDonutSvg {
    <# Generates the SQL Server wait statistics donut chart. #>
    param([object[]] $Waits, [int] $Top = 8)
    New-DonutSvgInternal -Items $Waits -LabelProperty 'WaitType' -ValueProperty 'WaitSeconds' -Title 'Top waits (share of wait time since startup)' -Unit 's' -Top $Top -ValueFormatter { param($v) '{0:N0}' -f $v }
}

function ConvertTo-DbResourcePie {
    <# Generates a per-database resource distribution donut chart. #>
    param([object[]] $Items, [string] $ValueProperty, [string] $Title, [string] $Unit = 'MB', [scriptblock] $ValueFormatter)
    New-DonutSvgInternal -Items $Items -LabelProperty 'DatabaseName' -ValueProperty $ValueProperty -Title $Title -Unit $Unit -Top 7 -ValueFormatter $ValueFormatter
}

function New-BarRowsHtml {
    param([object[]] $Rows)   # each: Label, Segments(@{Value;Color;Title}), Text, Max
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $Rows) {
        [void]$sb.Append("<div class='ops-bar-row'><div class='ops-bar-label' title='$(ConvertTo-HtmlSafe $r.Label)'>$(ConvertTo-HtmlSafe $r.Label)</div><div class='ops-bar-track'>")
        foreach ($seg in $r.Segments) {
            $pct = 0; if ($r.Max -gt 0) { $pct = [math]::Min(100, 100 * [double]$seg.Value / [double]$r.Max) }
            [void]$sb.Append("<div class='ops-bar-fill' style='width:$(F $pct)%;background:$($seg.Color)' title='$(ConvertTo-HtmlSafe $seg.Title)'></div>")
        }
        [void]$sb.Append("</div><div class='ops-bar-value'>$($r.Text)</div></div>")
    }
    return $sb.ToString()
}

function ConvertTo-DatabaseSizesHtml {
    <# Displays database size information (data vs log) as horizontal bars. #>
    param([object[]] $Databases, [int] $Top = 20)
    $rows = @($Databases | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [double]$_.DataMB + [double]$_.LogMB } } -Descending | Select-Object -First $Top)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No database size information.</div>" }
    $max = ($rows | ForEach-Object { [double]$_.DataMB + [double]$_.LogMB } | Measure-Object -Maximum).Maximum
    $bars = foreach ($d in $rows) {
        [pscustomobject]@{
            Label    = $d.DatabaseName
            Max      = $max
            Text     = Format-SizeMB ([double]$d.DataMB + [double]$d.LogMB)
            Segments = @(
                @{ Value = [double]$d.DataMB; Color = '#3a7bd5'; Title = "Data: $(Format-SizeMB $d.DataMB)" }
                @{ Value = [double]$d.LogMB;  Color = '#e0a100'; Title = "Log: $(Format-SizeMB $d.LogMB)" }
            )
        }
    }
    "<div class='ops-chart'><div class='ops-chart-title'>Database sizes (top $Top)</div>" +
    "<div class='ops-chart-legend'><span><span class='ops-swatch' style='background:#3a7bd5'></span>Data</span><span><span class='ops-swatch' style='background:#e0a100'></span>Log</span></div>" +
    (New-BarRowsHtml $bars) + '</div>'
}

function ConvertTo-FileIoLatencyHtml {
    <# Presents database file I/O latency with threshold colouring. #>
    param([object[]] $Files, [int] $Top = 20)
    $rows = @($Files | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [math]::Max([double]$_.AvgReadMs, [double]$_.AvgWriteMs) } } -Descending | Select-Object -First $Top)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No file I/O statistics.</div>" }
    $max = [math]::Max([double]$IoLatencyCriticalMs, (($rows | ForEach-Object { [math]::Max([double]$_.AvgReadMs, [double]$_.AvgWriteMs) }) | Measure-Object -Maximum).Maximum)
    $colorFor = { param($ms) if ($ms -ge $IoLatencyCriticalMs) { '#d64545' } elseif ($ms -ge $IoLatencyWarningMs) { '#e0a100' } else { '#2e9d5b' } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-chart'><div class='ops-chart-title'>File I/O latency (avg ms since startup; warning $IoLatencyWarningMs ms, critical $IoLatencyCriticalMs ms)</div>")
    [void]$sb.Append("<table class='ops-mini-table'><tr><th>Database / file</th><th>Type</th><th>Read latency</th><th>Write latency</th></tr>")
    foreach ($f in $rows) {
        $rd = [double]$f.AvgReadMs; $wr = [double]$f.AvgWriteMs
        $rw = F ([math]::Min(100, 100 * $rd / $max)); $ww = F ([math]::Min(100, 100 * $wr / $max))
        [void]$sb.Append("<tr><td title='$(ConvertTo-HtmlSafe $f.PhysicalName)'>$(ConvertTo-HtmlSafe $f.DatabaseName) / $(ConvertTo-HtmlSafe $f.LogicalName)</td><td>$(ConvertTo-HtmlSafe $f.FileType)</td>")
        [void]$sb.Append("<td><div class='ops-bar-track small'><div class='ops-bar-fill' style='width:$rw%;background:$(& $colorFor $rd)'></div></div><span class='ops-num'>$('{0:N1}' -f $rd) ms</span></td>")
        [void]$sb.Append("<td><div class='ops-bar-track small'><div class='ops-bar-fill' style='width:$ww%;background:$(& $colorFor $wr)'></div></div><span class='ops-num'>$('{0:N1}' -f $wr) ms</span></td></tr>")
    }
    [void]$sb.Append('</table></div>')
    return $sb.ToString()
}

function ConvertTo-MemoryClerkBarsHtml {
    <# Visualizes SQL Server memory clerk usage. #>
    param([object[]] $Clerks)
    $rows = @($Clerks | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [double]$_.MemoryMB } } -Descending)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No memory clerk information.</div>" }
    $max = ($rows | Measure-Object -Property MemoryMB -Maximum).Maximum
    $i = 0
    $bars = foreach ($c in $rows) {
        $col = $script:ChartPalette[$i % $script:ChartPalette.Count]; $i++
        [pscustomobject]@{ Label = $c.ClerkType; Max = $max; Text = Format-SizeMB $c.MemoryMB; Segments = @(@{ Value = [double]$c.MemoryMB; Color = $col; Title = "$($c.ClerkType): $(Format-SizeMB $c.MemoryMB)" }) }
    }
    "<div class='ops-chart'><div class='ops-chart-title'>Memory clerks (top $($rows.Count))</div>" + (New-BarRowsHtml $bars) + '</div>'
}

function ConvertTo-JobDurationSparkline {
    param([double[]] $Durations, [int] $Width = 170, [int] $Height = 34)
    $d = @($Durations)
    if ($d.Count -lt 2) { return "<span class='ops-muted'>1 run</span>" }
    $max = ($d | Measure-Object -Maximum).Maximum; if ($max -le 0) { $max = 1 }
    $pts = for ($i = 0; $i -lt $d.Count; $i++) {
        $x = 2 + ($Width - 4) * $i / ($d.Count - 1)
        $y = 2 + ($Height - 4) * (1 - $d[$i] / $max)
        "$(F $x),$(F $y)"
    }
    $lastX = 2 + ($Width - 4); $lastY = 2 + ($Height - 4) * (1 - $d[-1] / $max)
    return "<svg width='$Width' height='$Height' viewBox='0 0 $Width $Height'><polyline points='$($pts -join ' ')' fill='none' stroke='#3a7bd5' stroke-width='1.5'/><circle cx='$(F $lastX)' cy='$(F $lastY)' r='2.5' fill='#d64545'/></svg>"
}

function ConvertTo-JobDurationTrendHtml {
    param([object[]] $Trends)
    $rows = @($Trends | Where-Object { $null -ne $_ } | Sort-Object JobName)
    if ($rows.Count -eq 0) { return "<div class='ops-empty'>No SQL Server Agent job history in the lookback period.</div>" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table class='ops-mini-table'><tr><th>Job</th><th>Runs</th><th>Avg</th><th>Max</th><th>Last</th><th>Duration trend</th><th>Last outcome</th></tr>")
    foreach ($j in $rows) {
        [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $j.JobName)</td><td class='ops-num'>$($j.Runs)</td><td class='ops-num'>$($j.AvgDuration)</td><td class='ops-num'>$($j.MaxDuration)</td><td class='ops-num'>$($j.LastDuration)</td>")
        [void]$sb.Append("<td>$(ConvertTo-JobDurationSparkline -Durations $j.Durations)</td><td><span class='ops-pill $(Get-StatusClass $j.Status)'>$(ConvertTo-HtmlSafe $j.LastOutcome)</span></td></tr>")
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

# =====================================================================================
#  Operational health functions (KPI tiles and health cards)
# =====================================================================================
function ConvertTo-KpiHtmlText {
    <# Formats an operational indicator tile. #>
    param([string] $Title, [string] $Value, [string] $Subtitle = '', [string] $Status = 'Info')
    "<div class='ops-kpi $(Get-StatusClass $Status)'><div class='ops-kpi-title'>$(ConvertTo-HtmlSafe $Title)</div><div class='ops-kpi-value'>$(ConvertTo-HtmlSafe $Value)</div><div class='ops-kpi-sub'>$(ConvertTo-HtmlSafe $Subtitle)</div></div>"
}

function New-OpsCardHtml {
    param([string] $Title, [string] $Status, [System.Collections.IDictionary] $Rows, [string] $Footer = '')
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='ops-card $(Get-StatusClass $Status)'><div class='ops-card-head'><span>$(ConvertTo-HtmlSafe $Title)</span><span class='ops-pill $(Get-StatusClass $Status)'>$Status</span></div><table class='ops-card-table'>")
    foreach ($k in $Rows.Keys) { [void]$sb.Append("<tr><td>$(ConvertTo-HtmlSafe $k)</td><td>$(ConvertTo-HtmlSafe $Rows[$k])</td></tr>") }
    [void]$sb.Append('</table>')
    if ($Footer) { [void]$sb.Append("<div class='ops-card-foot'>$(ConvertTo-HtmlSafe $Footer)</div>") }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function ConvertTo-AgHealthCard {
    <# Presents Availability Group health. #>
    param([object[]] $Replicas, [object[]] $Databases)
    $rep = @($Replicas | Where-Object { $null -ne $_ })
    if ($rep.Count -eq 0) { return New-OpsCardHtml -Title 'Always On Availability Groups' -Status 'Info' -Rows ([ordered]@{ 'Configuration' = 'No Availability Groups detected' }) }
    $dbs = @($Databases | Where-Object { $null -ne $_ })
    $agCount   = @($rep | Select-Object -ExpandProperty AgName -Unique).Count
    $repOk     = @($rep | Where-Object SynchronizationHealth -eq 'HEALTHY').Count
    $repConn   = @($rep | Where-Object { $_.ConnectedState -eq 'CONNECTED' -or $_.Role -eq 'PRIMARY' }).Count
    $dbOk      = @($dbs | Where-Object SynchronizationHealth -eq 'HEALTHY').Count
    $maxSend   = [double](($dbs | Measure-Object -Property LogSendQueueKB -Maximum).Maximum)
    $maxRedo   = [double](($dbs | Measure-Object -Property RedoQueueKB -Maximum).Maximum)
    $status    = Get-WorstStatus @(@($rep | ForEach-Object Status) + @($dbs | ForEach-Object Status))
    New-OpsCardHtml -Title 'Always On Availability Groups' -Status $status -Rows ([ordered]@{
        'Availability Groups'        = $agCount
        'Replicas healthy'           = "$repOk of $($rep.Count)"
        'Replicas connected'         = "$repConn of $($rep.Count)"
        'Database replicas healthy'  = "$dbOk of $($dbs.Count)"
        'Max log send queue'         = Format-SizeMB ($maxSend / 1024)
        'Max redo queue'             = Format-SizeMB ($maxRedo / 1024)
    })
}

function ConvertTo-BackupHealthCard {
    <# Presents backup health. #>
    param([object[]] $Compliance)
    $c = @($Compliance | Where-Object { $null -ne $_ -and $_.Status -ne 'Info' })
    if ($c.Count -eq 0) { return New-OpsCardHtml -Title 'Backup compliance' -Status 'Info' -Rows ([ordered]@{ 'Databases evaluated' = 0 }) }
    $status = Get-WorstStatus @($c | ForEach-Object Status)
    New-OpsCardHtml -Title 'Backup compliance' -Status $status -Rows ([ordered]@{
        'Databases evaluated'     = $c.Count
        'Compliant'               = @($c | Where-Object Status -eq 'OK').Count
        'Warning (overdue)'       = @($c | Where-Object Status -eq 'Warning').Count
        'Critical (missing)'      = @($c | Where-Object Status -eq 'Critical').Count
        'Thresholds (full/diff/log)' = "$FullBackupMaxAgeHours h / $DiffBackupMaxAgeHours h / $LogBackupMaxAgeMinutes min"
    })
}

function ConvertTo-HostMemoryCard {
    <# Displays Windows host memory information. #>
    param([object] $Memory, [string] $Node)
    if ($null -eq $Memory) { return New-OpsCardHtml -Title "Host memory - $Node" -Status 'Info' -Rows ([ordered]@{ 'Status' = 'Not collected' }) }
    $total = [double]$Memory.HostTotalMB; $avail = [double]$Memory.HostAvailableMB
    $usedPct = 0; if ($total -gt 0) { $usedPct = 100 * ($total - $avail) / $total }
    $status = 'OK'
    if ($Memory.SystemMemoryState -match 'low' -or $usedPct -ge 97) { $status = 'Critical' } elseif ($usedPct -ge 92) { $status = 'Warning' }
    New-OpsCardHtml -Title "Host memory - $Node" -Status $status -Rows ([ordered]@{
        'Physical memory'   = Format-SizeMB $total
        'Available'         = Format-SizeMB $avail
        'Used'              = ('{0:N1}%' -f $usedPct)
        'Page file free'    = "$(Format-SizeMB $Memory.PageFileAvailableMB) of $(Format-SizeMB $Memory.PageFileTotalMB)"
        'OS memory state'   = $Memory.SystemMemoryState
    })
}

function ConvertTo-SqlProcessMemoryCard {
    <# Displays SQL Server process memory information. #>
    param([object] $Memory, [string] $Node)
    if ($null -eq $Memory) { return New-OpsCardHtml -Title "SQL Server memory - $Node" -Status 'Info' -Rows ([ordered]@{ 'Status' = 'Not collected' }) }
    $status = 'OK'
    if ($Memory.ProcessPhysicalMemoryLow -eq $true -or $Memory.ProcessVirtualMemoryLow -eq $true) { $status = 'Critical' }
    elseif ([double]$Memory.PageLifeExpectancy -lt 300) { $status = 'Warning' }
    $maxMem = [double]$Memory.MaxServerMemoryMB
    $maxText = if ($maxMem -ge 2147483647) { 'Unlimited (default)' } else { Format-SizeMB $maxMem }
    if ($maxMem -ge 2147483647 -and $status -eq 'OK') { $status = 'Warning' }
    New-OpsCardHtml -Title "SQL Server memory - $Node" -Status $status -Rows ([ordered]@{
        'Physical memory in use' = Format-SizeMB $Memory.SqlPhysicalMemoryMB
        'Total / target server'  = "$(Format-SizeMB $Memory.TotalServerMemoryMB) / $(Format-SizeMB $Memory.TargetServerMemoryMB)"
        'Max server memory'      = $maxText
        'Locked pages'           = Format-SizeMB $Memory.LockedPagesMB
        'Page life expectancy'   = ('{0:N0} s' -f [double]$Memory.PageLifeExpectancy)
        'Memory utilization'     = "$($Memory.MemoryUtilizationPct)%"
    })
}

# =====================================================================================
#  Windows / cluster collection
# =====================================================================================
function Get-WindowsHostData {
    param([string] $ComputerName, [string] $Label)
    $r = @{ OS = @(); Volumes = @(); Services = @(); LocalUsers = @(); Admins = @(); Events = @() }
    $session = $null
    try { $session = New-CimSession -ComputerName $ComputerName -OperationTimeoutSec 30 -ErrorAction Stop }
    catch {
        try { $session = New-CimSession -ComputerName $ComputerName -SessionOption (New-CimSessionOption -Protocol Dcom) -OperationTimeoutSec 30 -ErrorAction Stop }
        catch { Add-CollectionError -Node $Label -Section 'Windows (CIM session)' -Message $_.Exception.Message }
    }
    if ($session) {
        try {
            $os  = Get-CimInstance -CimSession $session -ClassName Win32_OperatingSystem -ErrorAction Stop
            $cs  = Get-CimInstance -CimSession $session -ClassName Win32_ComputerSystem -ErrorAction Stop
            $cpu = @(Get-CimInstance -CimSession $session -ClassName Win32_Processor -ErrorAction Stop)[0]
            $r.OS = @([pscustomobject]@{
                SqlNode         = $Label
                ComputerName    = $cs.Name
                Domain          = $cs.Domain
                OperatingSystem = $os.Caption
                Version         = $os.Version
                Build           = $os.BuildNumber
                Architecture    = $os.OSArchitecture
                Manufacturer    = $cs.Manufacturer
                Model           = $cs.Model
                Processor       = $cpu.Name
                Sockets         = $cs.NumberOfProcessors
                LogicalCPUs     = $cs.NumberOfLogicalProcessors
                MemoryGB        = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
                FreeMemoryGB    = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
                LastBoot        = $os.LastBootUpTime
                UptimeDays      = [math]::Round(((Get-Date) - $os.LastBootUpTime).TotalDays, 1)
                InstallDate     = $os.InstallDate
            })
        } catch { Add-CollectionError -Node $Label -Section 'Windows operating system' -Message $_.Exception.Message }

        try {
            $r.Volumes = @(Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType=3' -ErrorAction Stop |
                Where-Object { $_.Capacity -gt 0 -and $_.Name -notlike '\\?\*' } | ForEach-Object {
                    $freePct = [math]::Round(100 * $_.FreeSpace / $_.Capacity, 1)
                    $st = if ($freePct -lt $DiskFreeCriticalPct) { 'Critical' } elseif ($freePct -lt $DiskFreeWarningPct) { 'Warning' } else { 'OK' }
                    [pscustomobject]@{
                        SqlNode = $Label; Volume = $_.Name; Label = $_.Label; FileSystem = $_.FileSystem
                        CapacityGB = [math]::Round($_.Capacity / 1GB, 2); FreeGB = [math]::Round($_.FreeSpace / 1GB, 2)
                        FreePct = $freePct; BlockSizeKB = [math]::Round($_.BlockSize / 1KB, 0); Status = $st
                    }
                })
        } catch { Add-CollectionError -Node $Label -Section 'Windows volumes' -Message $_.Exception.Message }

        try {
            $pattern = '^(MSSQL|SQLAgent|SQLSERVERAGENT|SQLBrowser|SQLWriter|MSSQLFDLauncher|SQLTELEMETRY|SSISTELEMETRY|MsDtsServer|ReportServer|SQLServerReportingServices|PowerBIReportServer|MSSQLLaunchpad|MSOLAP|ClusSvc|MSDTC)'
            $r.Services = @(Get-CimInstance -CimSession $session -ClassName Win32_Service -ErrorAction Stop |
                Where-Object { $_.Name -match $pattern } | ForEach-Object {
                    $st = 'OK'
                    if ($_.State -ne 'Running') { $st = if ($_.StartMode -eq 'Auto') { 'Critical' } else { 'Info' } }
                    [pscustomobject]@{
                        SqlNode = $Label; ServiceName = $_.Name; DisplayName = $_.DisplayName
                        State = $_.State; StartMode = $_.StartMode; ServiceAccount = $_.StartName; Status = $st
                    }
                })
        } catch { Add-CollectionError -Node $Label -Section 'Windows services' -Message $_.Exception.Message }

        try {
            $r.LocalUsers = @(Get-CimInstance -CimSession $session -ClassName Win32_UserAccount -Filter 'LocalAccount=True' -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    SqlNode = $Label; AccountName = $_.Name; FullName = $_.FullName; Disabled = $_.Disabled
                    LockedOut = $_.Lockout; PasswordRequired = $_.PasswordRequired; PasswordExpires = $_.PasswordExpires; SID = $_.SID
                    Status = if (-not $_.Disabled -and $_.SID -like '*-500') { 'Warning' } elseif (-not $_.Disabled -and -not $_.PasswordRequired) { 'Warning' } else { 'OK' }
                }
            })
        } catch { Add-CollectionError -Node $Label -Section 'Windows local accounts' -Message $_.Exception.Message }

        try {
            $grp = Get-CimInstance -CimSession $session -ClassName Win32_Group -Filter "LocalAccount=True AND SID='S-1-5-32-544'" -ErrorAction Stop
            if ($grp) {
                $r.Admins = @(Get-CimAssociatedInstance -CimSession $session -InputObject $grp -Association Win32_GroupUser -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{
                        SqlNode = $Label; Group = $grp.Name; Member = "$($_.Domain)\$($_.Name)"
                        MemberType = $_.CimClass.CimClassName -replace '^Win32_', ''; SID = $_.SID
                    }
                })
            }
        } catch { Add-CollectionError -Node $Label -Section 'Windows administrators' -Message $_.Exception.Message }

        Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
    }

    if ($CollectEventViewer) {
        $start = (Get-Date).AddHours(-$LookbackHours)
        foreach ($log in 'System', 'Application') {
            try {
                $r.Events += @(Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{ LogName = $log; Level = 1, 2, 3; StartTime = $start } -MaxEvents $MaxEventsPerLog -ErrorAction Stop | ForEach-Object {
                    $msg = [string]$_.Message
                    if ($msg.Length -gt 600) { $msg = $msg.Substring(0, 600) + '...' }
                    [pscustomobject]@{
                        SqlNode = $Label; TimeCreated = $_.TimeCreated; LogName = $_.LogName
                        Level = switch ($_.Level) { 1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } default { [string]$_.LevelDisplayName } }
                        Provider = $_.ProviderName; EventId = $_.Id; Message = ($msg -replace '\s+', ' ')
                    }
                })
            } catch {
                if ($_.Exception.Message -notmatch 'No events were found') {
                    Add-CollectionError -Node $Label -Section "Event Viewer ($log)" -Message $_.Exception.Message
                }
            }
        }
    }
    return $r
}

function Get-ClusterData {
    param([string[]] $ComputerNames)
    $out = @{ Clusters = @(); Nodes = @(); Groups = @(); Resources = @() }
    if (-not $CollectClusterInfo) { return $out }
    if (-not (Get-Command Get-ClusterNode -ErrorAction SilentlyContinue)) {
        try { Import-Module FailoverClusters -ErrorAction Stop } catch { }
    }
    if (-not (Get-Command Get-ClusterNode -ErrorAction SilentlyContinue)) {
        if ($script:IsHA) { Add-CollectionError -Node 'All' -Section 'Failover Cluster' -Message 'FailoverClusters PowerShell cmdlets are not available on this machine (install RSAT Failover Clustering Tools).' }
        return $out
    }
    $seen = @{}
    foreach ($cn in $ComputerNames) {
        try { $cl = Get-Cluster -Name $cn -ErrorAction Stop } catch { continue }
        if ($seen.ContainsKey($cl.Name)) { continue }
        $seen[$cl.Name] = $true
        try {
            $q = Get-ClusterQuorum -Cluster $cl.Name -ErrorAction SilentlyContinue
            $out.Clusters += [pscustomobject]@{
                ClusterName = $cl.Name; Domain = $cl.Domain
                QuorumType = [string]$q.QuorumType; QuorumResource = [string]$q.QuorumResource
                SameSubnetThreshold = $cl.SameSubnetThreshold; CrossSubnetThreshold = $cl.CrossSubnetThreshold
            }
            $out.Nodes += @(Get-ClusterNode -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; NodeName = $_.Name; State = [string]$_.State; NodeWeight = $_.NodeWeight
                    DynamicWeight = $_.DynamicWeight; Status = if ([string]$_.State -eq 'Up') { 'OK' } elseif ([string]$_.State -eq 'Paused') { 'Warning' } else { 'Critical' }
                }
            })
            $out.Groups += @(Get-ClusterGroup -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; GroupName = $_.Name; OwnerNode = [string]$_.OwnerNode; State = [string]$_.State
                    GroupType = [string]$_.GroupType; Status = if ([string]$_.State -eq 'Online') { 'OK' } elseif ([string]$_.State -eq 'PartialOnline') { 'Warning' } else { 'Critical' }
                }
            })
            $out.Resources += @(Get-ClusterResource -Cluster $cl.Name -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    ClusterName = $cl.Name; ResourceName = $_.Name; ResourceType = [string]$_.ResourceType
                    OwnerGroup = [string]$_.OwnerGroup; OwnerNode = [string]$_.OwnerNode; State = [string]$_.State
                    Status = if ([string]$_.State -eq 'Online') { 'OK' } elseif ([string]$_.State -eq 'Offline') { 'Warning' } else { 'Critical' }
                }
            })
        } catch { Add-CollectionError -Node $cl.Name -Section 'Failover Cluster' -Message $_.Exception.Message }
    }
    return $out
}

# =====================================================================================
#  Monitoring queries (all read-only)
# =====================================================================================
$Queries = @{}

$Queries.Discovery = @'
SELECT
    CAST(SERVERPROPERTY('ServerName') AS nvarchar(256))                  AS ServerName,
    CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS nvarchar(256)) AS PhysicalHost,
    CAST(SERVERPROPERTY('Edition') AS nvarchar(256))                     AS Edition,
    CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64))               AS ProductVersion,
    CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(64))                 AS ProductLevel,
    CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(64))           AS ProductUpdateLevel,
    CAST(SERVERPROPERTY('Collation') AS nvarchar(128))                   AS Collation,
    CAST(ISNULL(SERVERPROPERTY('IsClustered'), 0) AS int)                AS IsClustered,
    CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS int)              AS IsHadrEnabled,
    CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int)              AS WindowsAuthOnly,
    si.sqlserver_start_time                                              AS StartTime,
    si.cpu_count                                                         AS CpuCount,
    si.physical_memory_kb / 1024                                         AS HostMemoryMB,
    (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'max server memory (MB)')        AS MaxServerMemoryMB,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'max degree of parallelism')        AS MaxDop,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'cost threshold for parallelism')   AS CostThreshold,
    (SELECT COUNT(*) FROM sys.availability_groups)                                                            AS AgCount,
    (SELECT COUNT(*) FROM sys.dm_hadr_availability_replica_states WHERE is_local = 1 AND role = 1)            AS LocalPrimaryCount,
    (SELECT COUNT(*) FROM sys.dm_hadr_availability_replica_states WHERE is_local = 1 AND role = 2)            AS LocalSecondaryCount
FROM sys.dm_os_sys_info si;
'@

$Queries.Databases = @'
SELECT d.name AS DatabaseName, d.database_id AS DatabaseId, d.state_desc AS State, d.recovery_model_desc AS RecoveryModel,
       d.compatibility_level AS CompatLevel, SUSER_SNAME(d.owner_sid) AS Owner, d.user_access_desc AS UserAccess,
       d.is_read_only AS IsReadOnly, d.page_verify_option_desc AS PageVerify, d.is_auto_close_on AS AutoClose,
       d.is_auto_shrink_on AS AutoShrink, d.create_date AS Created, d.collation_name AS Collation,
       d.is_distributor AS IsDistributor, d.log_reuse_wait_desc AS LogReuseWait, ag.name AS AgName,
       CAST(sz.DataMB AS decimal(18,2)) AS DataMB, CAST(sz.LogMB AS decimal(18,2)) AS LogMB
FROM sys.databases d
LEFT JOIN sys.availability_replicas ar ON ar.replica_id = d.replica_id
LEFT JOIN sys.availability_groups ag  ON ag.group_id = ar.group_id
OUTER APPLY (SELECT SUM(CASE WHEN mf.type = 0 THEN CAST(mf.size AS bigint) ELSE 0 END) * 8 / 1024.0 AS DataMB,
                    SUM(CASE WHEN mf.type = 1 THEN CAST(mf.size AS bigint) ELSE 0 END) * 8 / 1024.0 AS LogMB
             FROM sys.master_files mf WHERE mf.database_id = d.database_id) sz
ORDER BY d.name;
'@

$Queries.AgReplicas = @'
SELECT ag.name AS AgName, ar.replica_server_name AS ReplicaServer, ISNULL(ars.role_desc, 'UNKNOWN') AS Role,
       ar.availability_mode_desc AS AvailabilityMode, ar.failover_mode_desc AS FailoverMode, ar.seeding_mode_desc AS SeedingMode,
       ar.secondary_role_allow_connections_desc AS ReadableSecondary, ar.endpoint_url AS EndpointUrl,
       ar.backup_priority AS BackupPriority, ag.automated_backup_preference_desc AS BackupPreference, ar.session_timeout AS SessionTimeout,
       ars.connected_state_desc AS ConnectedState, ars.operational_state_desc AS OperationalState,
       ars.recovery_health_desc AS RecoveryHealth, ars.synchronization_health_desc AS SynchronizationHealth,
       ars.last_connect_error_number AS LastConnectError, ars.last_connect_error_description AS LastConnectErrorDescription,
       ars.last_connect_error_timestamp AS LastConnectErrorTime
FROM sys.availability_groups ag
JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states ars ON ars.replica_id = ar.replica_id
ORDER BY ag.name, ar.replica_server_name;
'@

$Queries.AgDatabases = @'
SELECT ag.name AS AgName, ar.replica_server_name AS ReplicaServer, DB_NAME(drs.database_id) AS DatabaseName,
       drs.is_local AS IsLocal, drs.is_primary_replica AS IsPrimaryReplica,
       drs.synchronization_state_desc AS SynchronizationState, drs.synchronization_health_desc AS SynchronizationHealth,
       drs.database_state_desc AS DatabaseState, drs.is_suspended AS IsSuspended, drs.suspend_reason_desc AS SuspendReason,
       drs.log_send_queue_size AS LogSendQueueKB, drs.log_send_rate AS LogSendRateKBs,
       drs.redo_queue_size AS RedoQueueKB, drs.redo_rate AS RedoRateKBs,
       drs.last_commit_time AS LastCommitTime, drs.last_hardened_time AS LastHardenedTime, drs.last_redone_time AS LastRedoneTime
FROM sys.dm_hadr_database_replica_states drs
JOIN sys.availability_replicas ar ON ar.replica_id = drs.replica_id
JOIN sys.availability_groups ag  ON ag.group_id = drs.group_id
ORDER BY ag.name, DatabaseName, ar.replica_server_name;
'@

$Queries.AgListeners = @'
SELECT ag.name AS AgName, l.dns_name AS ListenerName, l.port AS Port, lip.ip_address AS IpAddress,
       lip.ip_subnet_mask AS SubnetMask, lip.is_dhcp AS IsDhcp, lip.state_desc AS State
FROM sys.availability_group_listeners l
JOIN sys.availability_groups ag ON ag.group_id = l.group_id
LEFT JOIN sys.availability_group_listener_ip_addresses lip ON lip.listener_id = l.listener_id
ORDER BY ag.name, l.dns_name;
'@

$Queries.BackupStatus = @'
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel, d.state_desc AS State, ag.name AS AgName,
       CASE WHEN d.replica_id IS NULL THEN NULL WHEN sys.fn_hadr_is_primary_replica(d.name) = 1 THEN 'PRIMARY' ELSE 'SECONDARY' END AS LocalRole,
       b.LastFull, b.LastDiff, b.LastLog
FROM sys.databases d
LEFT JOIN sys.availability_replicas ar ON ar.replica_id = d.replica_id
LEFT JOIN sys.availability_groups ag  ON ag.group_id = ar.group_id
OUTER APPLY (SELECT MAX(CASE WHEN bs.type = 'D' THEN bs.backup_finish_date END) AS LastFull,
                    MAX(CASE WHEN bs.type = 'I' THEN bs.backup_finish_date END) AS LastDiff,
                    MAX(CASE WHEN bs.type = 'L' THEN bs.backup_finish_date END) AS LastLog
             FROM msdb.dbo.backupset bs WHERE bs.database_name = d.name) b
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.BackupHistory = @'
SELECT TOP (500) bs.database_name AS DatabaseName,
       CASE bs.type WHEN 'D' THEN 'Full' WHEN 'I' THEN 'Differential' WHEN 'L' THEN 'Log' WHEN 'F' THEN 'File' ELSE bs.type END AS BackupType,
       bs.is_copy_only AS CopyOnly, bs.backup_start_date AS Started, bs.backup_finish_date AS Finished,
       DATEDIFF(second, bs.backup_start_date, bs.backup_finish_date) AS DurationSec,
       CAST(bs.backup_size / 1048576.0 AS decimal(18,2)) AS SizeMB,
       CAST(bs.compressed_backup_size / 1048576.0 AS decimal(18,2)) AS CompressedMB,
       bmf.physical_device_name AS Device, bs.user_name AS ExecutedBy
FROM msdb.dbo.backupset bs
JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE bs.backup_finish_date >= DATEADD(day, -@days, GETDATE())
ORDER BY bs.backup_finish_date DESC;
'@

$Queries.RestoreHistory = @'
SELECT TOP (200) rh.destination_database_name AS DatabaseName, rh.restore_date AS RestoreDate,
       CASE rh.restore_type WHEN 'D' THEN 'Database' WHEN 'I' THEN 'Differential' WHEN 'L' THEN 'Log' WHEN 'F' THEN 'File'
                            WHEN 'G' THEN 'Filegroup' WHEN 'V' THEN 'Verify only' ELSE rh.restore_type END AS RestoreType,
       rh.user_name AS RestoredBy, rh.[replace] AS WithReplace, rh.[recovery] AS WithRecovery,
       bs.database_name AS SourceDatabase, bs.server_name AS SourceServer, bs.backup_finish_date AS SourceBackupDate,
       bmf.physical_device_name AS Device
FROM msdb.dbo.restorehistory rh
LEFT JOIN msdb.dbo.backupset bs ON bs.backup_set_id = rh.backup_set_id
LEFT JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE rh.restore_date >= DATEADD(day, -@days, GETDATE())
ORDER BY rh.restore_date DESC;
'@

$Queries.CpuHistory = @'
DECLARE @ts_now bigint = (SELECT cpu_ticks / (cpu_ticks / ms_ticks) FROM sys.dm_os_sys_info);
SELECT TOP (@minutes)
       DATEADD(ms, -1 * (@ts_now - y.[timestamp]), GETDATE()) AS EventTime,
       y.SQLProcessUtilization,
       y.SystemIdle,
       CASE WHEN 100 - y.SystemIdle - y.SQLProcessUtilization < 0 THEN 0 ELSE 100 - y.SystemIdle - y.SQLProcessUtilization END AS OtherProcessUtilization
FROM (SELECT x.[timestamp],
             x.record.value('(./Record/@id)[1]', 'int') AS record_id,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS SystemIdle,
             x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS SQLProcessUtilization
      FROM (SELECT [timestamp], CONVERT(xml, record) AS record
            FROM sys.dm_os_ring_buffers
            WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR' AND record LIKE N'%<SystemHealth>%') x) y
ORDER BY y.record_id DESC;
'@

$Queries.Memory = @'
SELECT pm.physical_memory_in_use_kb / 1024 AS SqlPhysicalMemoryMB, pm.locked_page_allocations_kb / 1024 AS LockedPagesMB,
       pm.memory_utilization_percentage AS MemoryUtilizationPct,
       CAST(pm.process_physical_memory_low AS bit) AS ProcessPhysicalMemoryLow, CAST(pm.process_virtual_memory_low AS bit) AS ProcessVirtualMemoryLow,
       sm.total_physical_memory_kb / 1024 AS HostTotalMB, sm.available_physical_memory_kb / 1024 AS HostAvailableMB,
       sm.total_page_file_kb / 1024 AS PageFileTotalMB, sm.available_page_file_kb / 1024 AS PageFileAvailableMB,
       sm.system_memory_state_desc AS SystemMemoryState,
       (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'max server memory (MB)') AS MaxServerMemoryMB,
       (SELECT CAST(value_in_use AS bigint) FROM sys.configurations WHERE name = 'min server memory (MB)') AS MinServerMemoryMB,
       (SELECT TOP 1 cntr_value / 1024 FROM sys.dm_os_performance_counters WHERE counter_name = 'Target Server Memory (KB)') AS TargetServerMemoryMB,
       (SELECT TOP 1 cntr_value / 1024 FROM sys.dm_os_performance_counters WHERE counter_name = 'Total Server Memory (KB)')  AS TotalServerMemoryMB,
       (SELECT TOP 1 cntr_value FROM sys.dm_os_performance_counters WHERE counter_name = 'Page life expectancy' AND object_name LIKE '%Buffer Manager%') AS PageLifeExpectancy
FROM sys.dm_os_process_memory pm CROSS JOIN sys.dm_os_sys_memory sm;
'@

$Queries.DatabaseFiles = @'
SELECT DB_NAME(mf.database_id) AS DatabaseName, mf.name AS LogicalName, mf.type_desc AS FileType, mf.physical_name AS PhysicalName,
       CAST(CAST(mf.size AS bigint) * 8 / 1024.0 AS decimal(18,2)) AS SizeMB,
       CASE mf.max_size WHEN -1 THEN 'Unlimited' WHEN 0 THEN 'No growth' ELSE CAST(CAST(mf.max_size AS bigint) * 8 / 1024 AS varchar(20)) + ' MB' END AS MaxSize,
       CASE mf.is_percent_growth WHEN 1 THEN CAST(mf.growth AS varchar(10)) + ' %' ELSE CAST(CAST(mf.growth AS bigint) * 8 / 1024 AS varchar(20)) + ' MB' END AS Growth,
       mf.state_desc AS FileState
FROM sys.master_files mf
ORDER BY DatabaseName, mf.type, mf.file_id;
'@

$Queries.SqlVolumes = @'
SELECT DISTINCT vs.volume_mount_point AS VolumePath, vs.logical_volume_name AS VolumeLabel, vs.file_system_type AS FileSystem,
       CAST(vs.total_bytes / 1073741824.0 AS decimal(18,2)) AS TotalGB,
       CAST(vs.available_bytes / 1073741824.0 AS decimal(18,2)) AS FreeGB
FROM sys.master_files mf
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs;
'@

$Queries.BufferCache = @'
SELECT CASE database_id WHEN 32767 THEN 'ResourceDb' ELSE DB_NAME(database_id) END AS DatabaseName,
       CAST(COUNT_BIG(*) * 8 / 1024.0 AS decimal(18,2)) AS BufferCacheMB
FROM sys.dm_os_buffer_descriptors
GROUP BY database_id
ORDER BY BufferCacheMB DESC;
'@

$Queries.DbIo = @'
SELECT DB_NAME(vfs.database_id) AS DatabaseName,
       CAST(SUM(vfs.num_of_bytes_read + vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS TotalIoMB,
       CAST(SUM(vfs.num_of_bytes_read) / 1048576.0 AS decimal(18,2)) AS ReadMB,
       CAST(SUM(vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS WrittenMB,
       SUM(vfs.num_of_reads) AS Reads, SUM(vfs.num_of_writes) AS Writes, SUM(vfs.io_stall) AS IoStallMs
FROM sys.dm_io_virtual_file_stats(NULL, NULL) vfs
GROUP BY vfs.database_id
ORDER BY TotalIoMB DESC;
'@

$Queries.DbCpu = @'
WITH DbCpu AS (
    SELECT pa.DatabaseID, SUM(qs.total_worker_time / 1000) AS CpuMs
    FROM sys.dm_exec_query_stats qs
    CROSS APPLY (SELECT CONVERT(int, value) AS DatabaseID FROM sys.dm_exec_plan_attributes(qs.plan_handle) WHERE attribute = N'dbid') pa
    GROUP BY pa.DatabaseID)
SELECT CASE WHEN DatabaseID = 32767 THEN 'ResourceDb' ELSE DB_NAME(DatabaseID) END AS DatabaseName, CpuMs,
       CAST(CpuMs * 100.0 / NULLIF(SUM(CpuMs) OVER (), 0) AS decimal(5,2)) AS CpuPct
FROM DbCpu
WHERE DatabaseID IS NOT NULL AND (DatabaseID = 32767 OR DB_NAME(DatabaseID) IS NOT NULL)
ORDER BY CpuMs DESC;
'@

$Queries.WaitStats = @'
WITH w AS (
    SELECT wait_type, wait_time_ms / 1000.0 AS WaitS, (wait_time_ms - signal_wait_time_ms) / 1000.0 AS ResourceS,
           signal_wait_time_ms / 1000.0 AS SignalS, waiting_tasks_count AS WaitCount
    FROM sys.dm_os_wait_stats
    WHERE waiting_tasks_count > 0
      AND wait_type NOT LIKE N'SLEEP%' AND wait_type NOT LIKE N'PREEMPTIVE_XE%' AND wait_type NOT LIKE N'QDS%'
      AND wait_type NOT IN (
        N'BROKER_EVENTHANDLER', N'BROKER_RECEIVE_WAITFOR', N'BROKER_TASK_STOP', N'BROKER_TO_FLUSH', N'BROKER_TRANSMITTER',
        N'CHECKPOINT_QUEUE', N'CHKPT', N'CLR_AUTO_EVENT', N'CLR_MANUAL_EVENT', N'CLR_SEMAPHORE', N'CXCONSUMER',
        N'DBMIRROR_DBM_EVENT', N'DBMIRROR_EVENTS_QUEUE', N'DBMIRROR_WORKER_QUEUE', N'DBMIRRORING_CMD', N'DIRTY_PAGE_POLL',
        N'DISPATCHER_QUEUE_SEMAPHORE', N'EXECSYNC', N'FSAGENT', N'FT_IFTS_SCHEDULER_IDLE_WAIT', N'FT_IFTSHC_MUTEX',
        N'HADR_CLUSAPI_CALL', N'HADR_FILESTREAM_IOMGR_IOCOMPLETION', N'HADR_LOGCAPTURE_WAIT', N'HADR_NOTIFICATION_DEQUEUE',
        N'HADR_TIMER_TASK', N'HADR_WORK_QUEUE', N'KSOURCE_WAKEUP', N'LAZYWRITER_SLEEP', N'LOGMGR_QUEUE',
        N'MEMORY_ALLOCATION_EXT', N'ONDEMAND_TASK_QUEUE', N'PARALLEL_REDO_DRAIN_WORKER', N'PARALLEL_REDO_LOG_CACHE',
        N'PARALLEL_REDO_TRAN_LIST', N'PARALLEL_REDO_WORKER_SYNC', N'PARALLEL_REDO_WORKER_WAIT_WORK',
        N'PREEMPTIVE_OS_FLUSHFILEBUFFERS', N'PVS_PREALLOCATE', N'PWAIT_ALL_COMPONENTS_INITIALIZED',
        N'PWAIT_DIRECTLOGCONSUMER_GETNEXT', N'PWAIT_EXTENSIBILITY_CLEANUP_TASK', N'REDO_THREAD_PENDING_WORK',
        N'REQUEST_FOR_DEADLOCK_SEARCH', N'RESOURCE_QUEUE', N'SERVER_IDLE_CHECK', N'SNI_HTTP_ACCEPT', N'SOS_WORK_DISPATCHER',
        N'SP_SERVER_DIAGNOSTICS_SLEEP', N'SQLTRACE_BUFFER_FLUSH', N'SQLTRACE_INCREMENTAL_FLUSH_SLEEP', N'SQLTRACE_WAIT_ENTRIES',
        N'STARTUP_DEPENDENCY_MANAGER', N'UCS_SESSION_REGISTRATION', N'VDI_CLIENT_OTHER', N'WAIT_FOR_RESULTS', N'WAITFOR',
        N'WAITFOR_TASKSHUTDOWN', N'WAIT_XTP_RECOVERY', N'WAIT_XTP_HOST_WAIT', N'WAIT_XTP_OFFLINE_CKPT_NEW_LOG',
        N'WAIT_XTP_CKPT_CLOSE', N'XE_DISPATCHER_JOIN', N'XE_DISPATCHER_WAIT', N'XE_TIMER_EVENT', N'XE_LIVE_TARGET_TVF'))
SELECT TOP (15) wait_type AS WaitType, CAST(WaitS AS decimal(18,2)) AS WaitSeconds, CAST(ResourceS AS decimal(18,2)) AS ResourceSeconds,
       CAST(SignalS AS decimal(18,2)) AS SignalSeconds, WaitCount,
       CAST(100.0 * WaitS / SUM(WaitS) OVER () AS decimal(5,2)) AS Percentage,
       CAST(WaitS * 1000 / WaitCount AS decimal(18,2)) AS AvgWaitMs
FROM w
ORDER BY WaitS DESC;
'@

$Queries.FileLatency = @'
SELECT DB_NAME(vfs.database_id) AS DatabaseName, mf.name AS LogicalName, mf.type_desc AS FileType, mf.physical_name AS PhysicalName,
       vfs.num_of_reads AS Reads, vfs.num_of_writes AS Writes,
       CAST(CASE WHEN vfs.num_of_reads = 0 THEN 0 ELSE vfs.io_stall_read_ms * 1.0 / vfs.num_of_reads END AS decimal(18,2)) AS AvgReadMs,
       CAST(CASE WHEN vfs.num_of_writes = 0 THEN 0 ELSE vfs.io_stall_write_ms * 1.0 / vfs.num_of_writes END AS decimal(18,2)) AS AvgWriteMs,
       CAST((vfs.num_of_bytes_read + vfs.num_of_bytes_written) / 1048576.0 AS decimal(18,2)) AS TotalIoMB
FROM sys.dm_io_virtual_file_stats(NULL, NULL) vfs
JOIN sys.master_files mf ON mf.database_id = vfs.database_id AND mf.file_id = vfs.file_id
ORDER BY (vfs.io_stall_read_ms + vfs.io_stall_write_ms) DESC;
'@

$Queries.LogUsage = @'
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel, d.log_reuse_wait_desc AS LogReuseWait,
       CAST(ls.cntr_value / 1024.0 AS decimal(18,2)) AS LogSizeMB, CAST(lu.cntr_value / 1024.0 AS decimal(18,2)) AS LogUsedMB,
       CAST(CASE WHEN ISNULL(ls.cntr_value, 0) = 0 THEN 0 ELSE lu.cntr_value * 100.0 / ls.cntr_value END AS decimal(5,2)) AS LogUsedPct
FROM sys.databases d
LEFT JOIN sys.dm_os_performance_counters ls
       ON RTRIM(ls.instance_name) COLLATE DATABASE_DEFAULT = d.name COLLATE DATABASE_DEFAULT
      AND ls.counter_name = 'Log File(s) Size (KB)' AND ls.object_name LIKE '%:Databases%'
LEFT JOIN sys.dm_os_performance_counters lu
       ON RTRIM(lu.instance_name) COLLATE DATABASE_DEFAULT = d.name COLLATE DATABASE_DEFAULT
      AND lu.counter_name = 'Log File(s) Used Size (KB)' AND lu.object_name LIKE '%:Databases%'
WHERE d.state = 0
ORDER BY d.name;
'@

$Queries.VlfCounts = @'
SELECT d.name AS DatabaseName, v.VlfCount, v.ActiveVlfCount
FROM sys.databases d
CROSS APPLY (SELECT COUNT(*) AS VlfCount, SUM(CASE WHEN li.vlf_active = 1 THEN 1 ELSE 0 END) AS ActiveVlfCount
             FROM sys.dm_db_log_info(d.database_id) li) v
WHERE d.state = 0 AND HAS_DBACCESS(d.name) = 1
  AND d.database_id NOT IN (SELECT drs.database_id FROM sys.dm_hadr_database_replica_states drs
                            WHERE drs.is_local = 1 AND drs.is_primary_replica = 0);
'@

$Queries.MemoryClerks = @'
SELECT TOP (15) type AS ClerkType, CAST(SUM(pages_kb) / 1024.0 AS decimal(18,2)) AS MemoryMB
FROM sys.dm_os_memory_clerks
GROUP BY type
ORDER BY SUM(pages_kb) DESC;
'@

$Queries.ServerServices = @'
SELECT servicename AS ServiceName, startup_type_desc AS StartMode, status_desc AS State, service_account AS ServiceAccount,
       last_startup_time AS LastStartup, is_clustered AS IsClustered, instant_file_initialization_enabled AS InstantFileInit
FROM sys.dm_server_services;
'@

$Queries.ErrorLog = @'
SET NOCOUNT ON;
DECLARE @s datetime = DATEADD(hour, -@hours, GETDATE()), @e datetime = GETDATE();
DECLARE @log TABLE (LogDate datetime, ProcessInfo nvarchar(100), [Text] nvarchar(max));
INSERT INTO @log EXEC master.dbo.xp_readerrorlog 0, 1, NULL, NULL, @s, @e, N'desc';
BEGIN TRY INSERT INTO @log EXEC master.dbo.xp_readerrorlog 1, 1, NULL, NULL, @s, @e, N'desc'; END TRY BEGIN CATCH END CATCH;
SELECT LogDate, ProcessInfo,
       CASE WHEN [Text] LIKE '%Severity: 1[7-9]%' OR [Text] LIKE '%Severity: 2[0-5]%' OR [Text] LIKE '%stack dump%'
                 OR [Text] LIKE '%corrupt%' OR [Text] LIKE '%I/O requests taking longer%' THEN 'Critical'
            WHEN [Text] LIKE '%error%' OR [Text] LIKE '%fail%' THEN 'Error'
            ELSE 'Warning' END AS Level,
       [Text] AS Message
FROM @log
WHERE ([Text] LIKE '%error%' OR [Text] LIKE '%fail%' OR [Text] LIKE '%warning%' OR [Text] LIKE '%severity%'
       OR [Text] LIKE '%deadlock%' OR [Text] LIKE '%I/O requests taking longer%' OR [Text] LIKE '%stack dump%'
       OR [Text] LIKE '%corrupt%' OR [Text] LIKE '%terminat%' OR [Text] LIKE '%non-yielding%')
  AND [Text] NOT LIKE '%found 0 errors%' AND [Text] NOT LIKE '%\ERRORLOG%' AND [Text] NOT LIKE '%without errors%'
  AND [Text] NOT LIKE 'Logging SQL Server messages%' AND [Text] NOT LIKE '%informational message%'
ORDER BY LogDate DESC;
'@

$Queries.CheckDbLog = @'
SET NOCOUNT ON;
DECLARE @log TABLE (LogDate datetime, ProcessInfo nvarchar(100), [Text] nvarchar(max));
INSERT INTO @log EXEC master.dbo.xp_readerrorlog 0, 1, N'DBCC CHECKDB';
BEGIN TRY INSERT INTO @log EXEC master.dbo.xp_readerrorlog 1, 1, N'DBCC CHECKDB'; END TRY BEGIN CATCH END CATCH;
SELECT TOP (200) LogDate, [Text] AS Message,
       CASE WHEN [Text] LIKE '%found 0 errors and repaired 0 errors%' THEN 'OK' ELSE 'Critical' END AS Status
FROM @log
ORDER BY LogDate DESC;
'@

$Queries.SuspectPages = @'
SELECT DB_NAME(sp.database_id) AS DatabaseName, sp.file_id AS FileId, sp.page_id AS PageId,
       CASE sp.event_type WHEN 1 THEN '823/824 error' WHEN 2 THEN 'Bad checksum' WHEN 3 THEN 'Torn page'
                          WHEN 4 THEN 'Restored' WHEN 5 THEN 'Repaired' WHEN 7 THEN 'Deallocated by DBCC' ELSE CAST(sp.event_type AS varchar(10)) END AS EventType,
       sp.error_count AS ErrorCount, sp.last_update_date AS LastUpdate
FROM msdb.dbo.suspect_pages sp
ORDER BY sp.last_update_date DESC;
'@

$Queries.MemoryDumps = @'
SELECT filename AS DumpFile, creation_time AS Created, CAST(size_in_bytes / 1048576.0 AS decimal(18,2)) AS SizeMB
FROM sys.dm_server_memory_dumps
ORDER BY creation_time DESC;
'@

$Queries.DbSecurity = @'
SELECT d.name AS DatabaseName, SUSER_SNAME(d.owner_sid) AS Owner, d.is_trustworthy_on AS Trustworthy,
       d.is_db_chaining_on AS DbChaining, d.is_encrypted AS TdeEnabled, dek.encryption_state AS EncryptionState,
       d.containment_desc AS Containment, d.is_broker_enabled AS BrokerEnabled, d.state_desc AS State
FROM sys.databases d
LEFT JOIN sys.dm_database_encryption_keys dek ON dek.database_id = d.database_id
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.OrphanedUsers = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, UserName sysname, UserType nvarchar(60), AuthenticationType nvarchar(60), CreateDate datetime' -InnerQuery @'
SELECT N'{DBNAME}', dp.name, dp.type_desc, dp.authentication_type_desc, dp.create_date
FROM {DB}.sys.database_principals dp
LEFT JOIN sys.server_principals sp ON sp.sid = dp.sid
WHERE sp.sid IS NULL AND dp.type IN ('S','U','G') AND dp.authentication_type IN (1,3)
  AND dp.principal_id > 4 AND dp.sid IS NOT NULL AND dp.name NOT LIKE '##%'
'@ -OrderBy 'ORDER BY DatabaseName, UserName'

$Queries.DbRoleMembers = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, RoleName sysname, MemberName sysname, MemberType nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', r.name, m.name, m.type_desc
FROM {DB}.sys.database_role_members rm
JOIN {DB}.sys.database_principals r ON r.principal_id = rm.role_principal_id
JOIN {DB}.sys.database_principals m ON m.principal_id = rm.member_principal_id
WHERE r.name IN ('db_owner','db_securityadmin','db_accessadmin','db_ddladmin','db_backupoperator') AND m.name <> 'dbo'
'@ -OrderBy 'ORDER BY DatabaseName, RoleName, MemberName'

$Queries.GuestAccess = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, PermissionState nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', p.state_desc
FROM {DB}.sys.database_permissions p
WHERE p.grantee_principal_id = 2 AND p.permission_name = 'CONNECT' AND p.state IN ('G','W')
  AND N'{DBNAME}' NOT IN (N'master', N'msdb', N'tempdb')
'@ -OrderBy 'ORDER BY DatabaseName'

$Queries.Certificates = Get-PerDatabaseQuery -TableDefinition 'DatabaseName sysname, CertificateName sysname, Subject nvarchar(4000), ExpiryDate datetime, KeyEncryption nvarchar(60)' -InnerQuery @'
SELECT N'{DBNAME}', c.name, c.subject, c.expiry_date, c.pvt_key_encryption_type_desc
FROM {DB}.sys.certificates c
WHERE c.name NOT LIKE '##%'
'@ -OrderBy 'ORDER BY ExpiryDate'

$Queries.ServerRoleMembers = @'
SELECT r.name AS ServerRole, m.name AS MemberName, m.type_desc AS MemberType, m.is_disabled AS IsDisabled
FROM sys.server_role_members rm
JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id
JOIN sys.server_principals m ON m.principal_id = rm.member_principal_id
WHERE r.name IN ('sysadmin','securityadmin','serveradmin','setupadmin','processadmin','diskadmin','dbcreator','bulkadmin')
ORDER BY r.name, m.name;
'@

$Queries.InstanceSecurity = @'
SELECT
    (SELECT CAST(is_disabled AS int) FROM sys.server_principals WHERE sid = 0x01) AS SaDisabled,
    (SELECT name FROM sys.server_principals WHERE sid = 0x01)                     AS SaLoginName,
    CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int)                         AS WindowsAuthOnly,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'xp_cmdshell')                  AS XpCmdShell,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'clr enabled')                  AS ClrEnabled,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'Ole Automation Procedures')    AS OleAutomation,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'Ad Hoc Distributed Queries')   AS AdHocDistributedQueries,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'cross db ownership chaining')  AS CrossDbOwnershipChaining,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'remote admin connections')     AS RemoteAdminConnections,
    (SELECT CAST(value_in_use AS int) FROM sys.configurations WHERE name = 'contained database authentication') AS ContainedDbAuth,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE is_disabled = 0 AND is_policy_checked = 0)     AS SqlLoginsWithoutPolicy,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE is_disabled = 0 AND is_expiration_checked = 0) AS SqlLoginsWithoutExpiration,
    (SELECT COUNT(*) FROM sys.sql_logins WHERE PWDCOMPARE('', password_hash) = 1)            AS BlankPasswordLogins,
    (SELECT COUNT(*) FROM sys.server_role_members rm JOIN sys.server_principals r ON r.principal_id = rm.role_principal_id WHERE r.name = 'sysadmin') AS SysadminMembers;
'@

$Queries.Configurations = @'
SELECT name AS Setting, CAST(value_in_use AS nvarchar(50)) AS Value FROM sys.configurations ORDER BY name;
'@

$Queries.Logins = @'
SELECT name AS LoginName, type_desc AS LoginType, is_disabled AS IsDisabled, default_database_name AS DefaultDatabase
FROM sys.server_principals
WHERE type IN ('S','U','G') AND name NOT LIKE '##%' AND name NOT LIKE 'NT SERVICE\%' AND name NOT LIKE 'NT AUTHORITY\%'
ORDER BY name;
'@

$Queries.AgentJobs = @'
SELECT j.name AS JobName, j.enabled AS IsEnabled, SUSER_SNAME(j.owner_sid) AS Owner, c.name AS Category
FROM msdb.dbo.sysjobs j
LEFT JOIN msdb.dbo.syscategories c ON c.category_id = j.category_id
ORDER BY j.name;
'@

$Queries.LinkedServers = @'
SELECT name AS LinkedServer, product AS Product, provider AS Provider, data_source AS DataSource
FROM sys.servers WHERE is_linked = 1 ORDER BY name;
'@

$Queries.TraceFlags = @'
SET NOCOUNT ON;
DECLARE @tf TABLE (TraceFlag int, Status int, Global int, Session int);
INSERT INTO @tf EXEC ('DBCC TRACESTATUS(-1) WITH NO_INFOMSGS');
SELECT TraceFlag, Status, Global, Session FROM @tf ORDER BY TraceFlag;
'@

$Queries.JobHistory = @'
SELECT j.name AS JobName, j.enabled AS IsEnabled,
       msdb.dbo.agent_datetime(h.run_date, h.run_time) AS RunStart,
       (h.run_duration / 10000) * 3600 + ((h.run_duration / 100) % 100) * 60 + (h.run_duration % 100) AS DurationSec,
       CASE h.run_status WHEN 0 THEN 'Failed' WHEN 1 THEN 'Succeeded' WHEN 2 THEN 'Retry' WHEN 3 THEN 'Canceled' WHEN 4 THEN 'In Progress' ELSE 'Unknown' END AS RunStatus
FROM msdb.dbo.sysjobhistory h
JOIN msdb.dbo.sysjobs j ON j.job_id = h.job_id
WHERE h.step_id = 0
  AND h.run_date >= CONVERT(int, CONVERT(char(8), DATEADD(day, -@days, GETDATE()), 112))
ORDER BY j.name, RunStart;
'@

$Queries.CheckDb = @'
SELECT d.name AS DatabaseName, d.state_desc AS State,
       CAST(DATABASEPROPERTYEX(d.name, 'LastGoodCheckDbTime') AS datetime) AS LastGoodCheckDb
FROM sys.databases d
WHERE d.name <> 'tempdb'
ORDER BY d.name;
'@

$Queries.LongRunning = @'
SELECT r.session_id AS SessionId, s.login_name AS LoginName, s.host_name AS HostName, s.program_name AS ProgramName,
       DB_NAME(r.database_id) AS DatabaseName, r.status AS RequestStatus, r.command AS Command, r.start_time AS StartTime,
       CAST(r.total_elapsed_time / 1000.0 AS decimal(18,1)) AS ElapsedSec, r.cpu_time AS CpuMs, r.logical_reads AS LogicalReads,
       r.wait_type AS WaitType, r.blocking_session_id AS BlockedBy, LEFT(t.text, 2000) AS SqlText
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE s.is_user_process = 1 AND r.session_id <> @@SPID AND r.total_elapsed_time >= @ms
  AND r.wait_type NOT IN (N'BROKER_RECEIVE_WAITFOR', N'WAITFOR')
ORDER BY r.total_elapsed_time DESC;
'@

$Queries.Blocking = @'
SELECT r.session_id AS BlockedSession, r.blocking_session_id AS BlockingSession, DB_NAME(r.database_id) AS DatabaseName,
       r.wait_type AS WaitType, CAST(r.wait_time / 1000.0 AS decimal(18,1)) AS WaitSec, r.wait_resource AS WaitResource,
       s.login_name AS BlockedLogin, bs.login_name AS BlockingLogin, bs.host_name AS BlockingHost, bs.program_name AS BlockingProgram,
       LEFT(t.text, 1000) AS BlockedSql, LEFT(bt.text, 1000) AS BlockingSql
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
LEFT JOIN sys.dm_exec_sessions bs ON bs.session_id = r.blocking_session_id
LEFT JOIN sys.dm_exec_connections bc ON bc.session_id = r.blocking_session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
OUTER APPLY sys.dm_exec_sql_text(bc.most_recent_sql_handle) bt
WHERE r.blocking_session_id <> 0 AND r.wait_time >= @ms
ORDER BY r.wait_time DESC;
'@

# =====================================================================================
#  Phase 1 - Discovery (reachability, HA role, node labels)
# =====================================================================================
Write-Host "SQL Server Operations Dashboard - collection started $($script:ReportStart)" -ForegroundColor Cyan
$Discovery = @{}
$Unreachable = @()
foreach ($inst in $SqlInstances) {
    Write-Host "  Discovering $inst ..."
    $row = @(Invoke-DbQuery -SqlInstance $inst -Query $Queries.Discovery -Section 'Discovery / connectivity') | Select-Object -First 1
    if ($null -eq $row) {
        $script:NodeMap[$inst] = "$inst (Unreachable)"
        $Unreachable += $inst
        Add-Finding -Status Critical -Area 'Databases' -Node $inst -Item 'Instance unreachable' -Detail 'Connection or discovery query failed; see collection errors.'
        continue
    }
    $role = 'Standalone'
    if ([int]$row.AgCount -gt 0) {
        if ([int]$row.LocalPrimaryCount -gt 0 -and [int]$row.LocalSecondaryCount -gt 0) { $role = 'Primary/Secondary' }
        elseif ([int]$row.LocalPrimaryCount -gt 0) { $role = 'Primary' }
        elseif ([int]$row.LocalSecondaryCount -gt 0) { $role = 'Secondary' }
        else { $role = 'AG Replica' }
    } elseif ([int]$row.IsClustered -eq 1) { $role = 'FCI' }
    $script:NodeMap[$inst] = "$inst ($role)"
    $row.SqlNode = $script:NodeMap[$inst]
    $row | Add-Member -NotePropertyName Role -NotePropertyValue $role
    $Discovery[$inst] = $row
    $script:NodeOrder += $inst
}
$script:NodeLabels = @($script:NodeOrder | ForEach-Object { Get-ReportNodeLabel $_ })
$script:IsHA = (@($Discovery.Values | Where-Object { [int]$_.AgCount -gt 0 -or [int]$_.IsClustered -eq 1 }).Count -gt 0)
$HasAG = (@($Discovery.Values | Where-Object { [int]$_.AgCount -gt 0 }).Count -gt 0)
$EnvironmentMode = if ($HasAG) { 'High Availability - Always On Availability Groups' } elseif ($script:IsHA) { 'High Availability - Failover Cluster Instance' } else { 'Standalone' }

# =====================================================================================
#  Phase 2 - SQL Server collection
# =====================================================================================
$collectionPlan = @(
    @{ Key = 'Databases';         Section = 'Database inventory';          Query = $Queries.Databases }
    @{ Key = 'AgReplicas';        Section = 'AG replicas';                 Query = $Queries.AgReplicas;     AgOnly = $true }
    @{ Key = 'AgDatabases';       Section = 'AG database states';          Query = $Queries.AgDatabases;    AgOnly = $true }
    @{ Key = 'AgListeners';       Section = 'AG listeners';                Query = $Queries.AgListeners;    AgOnly = $true }
    @{ Key = 'BackupStatus';      Section = 'Backup freshness';            Query = $Queries.BackupStatus }
    @{ Key = 'BackupHistory';     Section = 'Backup history';              Query = $Queries.BackupHistory;  Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'RestoreHistory';    Section = 'Restore history';             Query = $Queries.RestoreHistory; Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'CpuHistory';        Section = 'CPU history';                 Query = $Queries.CpuHistory;     Params = @{ minutes = $CpuHistoryMinutes } }
    @{ Key = 'Memory';            Section = 'Memory';                      Query = $Queries.Memory }
    @{ Key = 'DatabaseFiles';     Section = 'Database files';              Query = $Queries.DatabaseFiles }
    @{ Key = 'SqlVolumes';        Section = 'Storage volumes';             Query = $Queries.SqlVolumes }
    @{ Key = 'BufferCache';       Section = 'Buffer cache per database';   Query = $Queries.BufferCache }
    @{ Key = 'DbIo';              Section = 'I/O per database';            Query = $Queries.DbIo }
    @{ Key = 'DbCpu';             Section = 'CPU per database';            Query = $Queries.DbCpu }
    @{ Key = 'WaitStats';         Section = 'Wait statistics';             Query = $Queries.WaitStats }
    @{ Key = 'FileLatency';       Section = 'File I/O latency';            Query = $Queries.FileLatency }
    @{ Key = 'LogUsage';          Section = 'Transaction log usage';       Query = $Queries.LogUsage }
    @{ Key = 'VlfCounts';         Section = 'VLF statistics';              Query = $Queries.VlfCounts }
    @{ Key = 'MemoryClerks';      Section = 'Memory clerks';               Query = $Queries.MemoryClerks }
    @{ Key = 'ServerServices';    Section = 'SQL Server services (DMV)';   Query = $Queries.ServerServices }
    @{ Key = 'ErrorLog';          Section = 'SQL Server error log';        Query = $Queries.ErrorLog;       Params = @{ hours = $LookbackHours } }
    @{ Key = 'CheckDbLog';        Section = 'CHECKDB error log evidence';  Query = $Queries.CheckDbLog }
    @{ Key = 'SuspectPages';      Section = 'Suspect pages';               Query = $Queries.SuspectPages }
    @{ Key = 'MemoryDumps';       Section = 'Memory dumps';                Query = $Queries.MemoryDumps }
    @{ Key = 'DbSecurity';        Section = 'Database security';           Query = $Queries.DbSecurity }
    @{ Key = 'OrphanedUsers';     Section = 'Orphaned users';              Query = $Queries.OrphanedUsers }
    @{ Key = 'DbRoleMembers';     Section = 'Privileged database roles';   Query = $Queries.DbRoleMembers }
    @{ Key = 'GuestAccess';       Section = 'Guest access';                Query = $Queries.GuestAccess }
    @{ Key = 'Certificates';      Section = 'Certificates';                Query = $Queries.Certificates }
    @{ Key = 'ServerRoleMembers'; Section = 'Server role membership';      Query = $Queries.ServerRoleMembers }
    @{ Key = 'InstanceSecurity';  Section = 'Instance security checks';    Query = $Queries.InstanceSecurity }
    @{ Key = 'Configurations';    Section = 'sp_configure values';         Query = $Queries.Configurations }
    @{ Key = 'Logins';            Section = 'Logins';                      Query = $Queries.Logins }
    @{ Key = 'AgentJobs';         Section = 'Agent jobs';                  Query = $Queries.AgentJobs }
    @{ Key = 'LinkedServers';     Section = 'Linked servers';              Query = $Queries.LinkedServers }
    @{ Key = 'TraceFlags';        Section = 'Trace flags';                 Query = $Queries.TraceFlags }
    @{ Key = 'JobHistory';        Section = 'Agent job history';           Query = $Queries.JobHistory;     Params = @{ days = $HistoryLookbackDays } }
    @{ Key = 'CheckDb';           Section = 'DBCC CHECKDB history';        Query = $Queries.CheckDb }
    @{ Key = 'LongRunning';       Section = 'Long-running requests';       Query = $Queries.LongRunning;    Params = @{ ms = $LongRunningQuerySeconds * 1000 } }
    @{ Key = 'Blocking';          Section = 'Blocking';                    Query = $Queries.Blocking;       Params = @{ ms = $BlockingWaitSeconds * 1000 } }
)

$SqlData = @{}
foreach ($p in $collectionPlan) { $SqlData[$p.Key] = @() }
foreach ($inst in $script:NodeOrder) {
    Write-Host "  Collecting SQL Server data from $(Get-ReportNodeLabel $inst) ..."
    foreach ($p in $collectionPlan) {
        if ($p.AgOnly -and [int]$Discovery[$inst].AgCount -eq 0) { continue }
        $qArgs = @{ SqlInstance = $inst; Query = $p.Query; Section = $p.Section }
        if ($p.Params) { $qArgs.Parameters = $p.Params }
        $SqlData[$p.Key] += @(Invoke-DbQuery @qArgs)
    }
}

# =====================================================================================
#  Phase 3 - Windows and cluster collection
# =====================================================================================
$WinData = @{ OS = @(); Volumes = @(); Services = @(); LocalUsers = @(); Admins = @(); Events = @() }
$hostSeen = @{}
foreach ($inst in $script:NodeOrder) {
    $hostName = [string]$Discovery[$inst].PhysicalHost
    if (-not $hostName -or $hostSeen.ContainsKey($hostName)) { continue }
    $hostSeen[$hostName] = $true
    Write-Host "  Collecting Windows data from $hostName ..."
    $h = Get-WindowsHostData -ComputerName $hostName -Label (Get-ReportNodeLabel $inst)
    foreach ($k in @($WinData.Keys)) { $WinData[$k] += @($h[$k]) }
}
Write-Host '  Collecting Failover Cluster data ...'
$ClusterData = Get-ClusterData -ComputerNames @($hostSeen.Keys)

# Fall back to sys.dm_server_services if remote service enumeration failed
if (@($WinData.Services).Count -eq 0 -and @($SqlData.ServerServices).Count -gt 0) {
    $WinData.Services = @($SqlData.ServerServices | ForEach-Object {
        [pscustomobject]@{
            SqlNode = $_.SqlNode; ServiceName = $_.ServiceName; DisplayName = $_.ServiceName; State = $_.State
            StartMode = $_.StartMode; ServiceAccount = $_.ServiceAccount
            Status = if ($_.State -eq 'Running') { 'OK' } elseif ($_.StartMode -eq 'Automatic') { 'Critical' } else { 'Info' }
        }
    })
}

# =====================================================================================
#  Phase 4 - Analysis and findings
# =====================================================================================
Write-Host '  Analysing collected data ...'
$now = Get-Date

# ---- Instance inventory ----
$versionNames = @{ '11' = '2012'; '12' = '2014'; '13' = '2016'; '14' = '2017'; '15' = '2019'; '16' = '2022'; '17' = '2025' }
$InstanceInventory = @(
    foreach ($inst in $script:NodeOrder) {
        $d = $Discovery[$inst]; $major = ([string]$d.ProductVersion).Split('.')[0]
        [pscustomobject]@{
            SqlNode = $d.SqlNode; ServerName = $d.ServerName; PhysicalHost = $d.PhysicalHost; Role = $d.Role
            SqlVersion = "SQL Server $($versionNames[$major])"; ProductVersion = $d.ProductVersion
            Level = (@($d.ProductLevel, $d.ProductUpdateLevel) | Where-Object { $_ }) -join ' '
            Edition = $d.Edition; Collation = $d.Collation; HadrEnabled = [bool]$d.IsHadrEnabled; Clustered = [bool]$d.IsClustered
            WindowsAuthOnly = [bool]$d.WindowsAuthOnly; StartTime = $d.StartTime
            UptimeDays = [math]::Round(($now - [datetime]$d.StartTime).TotalDays, 1)
            CPUs = $d.CpuCount; HostMemory = Format-SizeMB $d.HostMemoryMB
            MaxServerMemory = if ([double]$d.MaxServerMemoryMB -ge 2147483647) { 'Unlimited' } else { Format-SizeMB $d.MaxServerMemoryMB }
            MaxDop = $d.MaxDop; CostThreshold = $d.CostThreshold; Status = 'OK'
        }
    }
    foreach ($u in $Unreachable) {
        [pscustomobject]@{ SqlNode = Get-ReportNodeLabel $u; ServerName = $u; Role = 'Unknown'; Status = 'Critical' }
    }
)

# ---- Databases ----
foreach ($db in $SqlData.Databases) {
    $cat = Get-DatabaseCategory -DatabaseName $db.DatabaseName -AgName $db.AgName -IsDistributor ([bool]$db.IsDistributor)
    $notes = @(); $st = 'OK'
    switch -Regex ($db.State) {
        '^ONLINE$'                                  { }
        '^(OFFLINE|RESTORING)$'                     { $st = 'Warning';  $notes += "State $($db.State)" }
        default                                     { $st = 'Critical'; $notes += "State $($db.State)" }
    }
    if ($db.AutoClose -eq $true)  { if ($st -eq 'OK') { $st = 'Warning' }; $notes += 'AUTO_CLOSE on' }
    if ($db.AutoShrink -eq $true) { if ($st -eq 'OK') { $st = 'Warning' }; $notes += 'AUTO_SHRINK on' }
    if ($db.PageVerify -and $db.PageVerify -ne 'CHECKSUM' -and $db.DatabaseName -ne 'tempdb') { if ($st -eq 'OK') { $st = 'Warning' }; $notes += "PAGE_VERIFY $($db.PageVerify)" }
    if (-not $db.Owner) { $notes += 'Owner SID has no login' }
    $db | Add-Member -NotePropertyName Category -NotePropertyValue $cat -Force
    $db | Add-Member -NotePropertyName TotalSize -NotePropertyValue (Format-SizeMB ([double]$db.DataMB + [double]$db.LogMB)) -Force
    $db | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $db | Add-Member -NotePropertyName Notes -NotePropertyValue ($notes -join '; ') -Force
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Databases' -Node $db.SqlNode -Item $db.DatabaseName -Detail ($notes -join '; ') }
}

# ---- Always On ----
$roleOf = @{}; foreach ($inst in $script:NodeOrder) { $roleOf[(Get-ReportNodeLabel $inst)] = $Discovery[$inst].Role }
$AgTopology = @(
    $SqlData.AgReplicas | Group-Object AgName, ReplicaServer | ForEach-Object {
        $best = @($_.Group | Sort-Object @{ Expression = { if ($roleOf[$_.SqlNode] -like 'Primary*') { 0 } elseif ($_.ConnectedState) { 1 } else { 2 } } })[0]
        $st = 'OK'
        if ($best.ConnectedState -eq 'DISCONNECTED' -or $best.SynchronizationHealth -eq 'NOT_HEALTHY') { $st = 'Critical' }
        elseif ($best.SynchronizationHealth -eq 'PARTIALLY_HEALTHY' -or ($best.OperationalState -and $best.OperationalState -ne 'ONLINE' -and $best.Role -eq 'PRIMARY')) { $st = 'Warning' }
        elseif (-not $best.SynchronizationHealth) { $st = 'Info' }
        $best | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
        $best | Add-Member -NotePropertyName ReportedBy -NotePropertyValue $best.SqlNode -Force
        if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Always On' -Node $best.SqlNode -Item "$($best.AgName) / $($best.ReplicaServer)" -Detail "Connected: $($best.ConnectedState); sync health: $($best.SynchronizationHealth); last error: $($best.LastConnectErrorDescription)" }
        $best
    }
)

$AgDbStates = @(
    $SqlData.AgDatabases | Group-Object AgName, ReplicaServer, DatabaseName | ForEach-Object {
        @($_.Group | Sort-Object @{ Expression = { if ($roleOf[$_.SqlNode] -like 'Primary*') { 0 } else { 1 } } })[0]
    }
)
foreach ($g in ($AgDbStates | Group-Object AgName, DatabaseName)) {
    $prim = @($g.Group | Where-Object { $_.IsPrimaryReplica -eq $true }) | Select-Object -First 1
    foreach ($r in $g.Group) {
        $loss = $null
        if ($prim -and $r.IsPrimaryReplica -ne $true -and $prim.LastCommitTime -and $r.LastCommitTime) {
            $loss = [math]::Max(0, [math]::Round(([datetime]$prim.LastCommitTime - [datetime]$r.LastCommitTime).TotalSeconds, 0))
        }
        $rec = $null
        if ([double]$r.RedoRateKBs -gt 0) { $rec = [math]::Round([double]$r.RedoQueueKB / [double]$r.RedoRateKBs, 0) }
        $st = 'OK'; $why = @()
        if ($r.SynchronizationHealth -eq 'NOT_HEALTHY' -or $r.IsSuspended -eq $true) { $st = 'Critical'; $why += "health $($r.SynchronizationHealth); suspended $($r.IsSuspended) $($r.SuspendReason)" }
        elseif ($r.SynchronizationHealth -eq 'PARTIALLY_HEALTHY') { $st = 'Warning'; $why += 'partially healthy' }
        if ([double]$r.LogSendQueueKB -gt $AgSendQueueWarningKB) { if ($st -eq 'OK') { $st = 'Warning' }; $why += "log send queue $([double]$r.LogSendQueueKB) KB" }
        if ([double]$r.RedoQueueKB -gt $AgRedoQueueWarningKB)    { if ($st -eq 'OK') { $st = 'Warning' }; $why += "redo queue $([double]$r.RedoQueueKB) KB" }
        $r | Add-Member -NotePropertyName EstDataLossSec -NotePropertyValue $loss -Force
        $r | Add-Member -NotePropertyName EstRecoverySec -NotePropertyValue $rec -Force
        $r | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
        if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Always On' -Node $r.SqlNode -Item "$($r.AgName) / $($r.DatabaseName) @ $($r.ReplicaServer)" -Detail ($why -join '; ') }
    }
}
$AgListeners = @($SqlData.AgListeners | Group-Object AgName, ListenerName, IpAddress | ForEach-Object {
    $l = $_.Group[0]
    $l | Add-Member -NotePropertyName Status -NotePropertyValue $(if ($l.State -eq 'ONLINE') { 'OK' } elseif ($l.State) { 'Warning' } else { 'Info' }) -Force
    $l })

# ---- Backups (compliance across AG replicas) ----
$BackupCompliance = @(
    $SqlData.BackupStatus | Group-Object { if ($_.AgName) { "AG|$($_.AgName)|$($_.DatabaseName)" } else { "SA|$($_.SqlNode)|$($_.DatabaseName)" } } | ForEach-Object {
        $rows  = @($_.Group); $first = $rows[0]
        $isSys = $first.DatabaseName -in @('master', 'model', 'msdb')
        $scope = if ($isSys) { 'System databases' } elseif ($first.AgName) { 'Availability Group databases (evaluated across queried replicas)' } else { 'Standalone databases' }
        $lastFull = ($rows | Where-Object LastFull | Sort-Object LastFull -Descending | Select-Object -First 1)
        $lastDiff = ($rows | Where-Object LastDiff | Sort-Object LastDiff -Descending | Select-Object -First 1)
        $lastLog  = ($rows | Where-Object LastLog  | Sort-Object LastLog  -Descending | Select-Object -First 1)
        $recModel = @($rows | Where-Object RecoveryModel | Select-Object -ExpandProperty RecoveryModel)[0]
        $states   = @($rows | ForEach-Object State)

        $fullSt = 'OK'; $diffSt = 'N/A'; $logSt = 'N/A'
        if (-not $lastFull) { $fullSt = 'Critical' } elseif (($now - [datetime]$lastFull.LastFull).TotalHours -gt $FullBackupMaxAgeHours) { $fullSt = 'Warning' }
        if ($lastDiff) {
            $diffAge = ($now - [datetime]$lastDiff.LastDiff).TotalHours
            $fullAge = if ($lastFull) { ($now - [datetime]$lastFull.LastFull).TotalHours } else { [double]::MaxValue }
            $diffSt = if ($diffAge -gt $DiffBackupMaxAgeHours -and $fullAge -gt $DiffBackupMaxAgeHours) { 'Warning' } else { 'OK' }
        }
        if (-not $isSys -and $recModel -and $recModel -ne 'SIMPLE') {
            if (-not $lastLog) { $logSt = 'Critical' } elseif (($now - [datetime]$lastLog.LastLog).TotalMinutes -gt $LogBackupMaxAgeMinutes) { $logSt = 'Warning' } else { $logSt = 'OK' }
        }
        $overall = Get-WorstStatus @($fullSt, $diffSt, $logSt)
        if ($states -notcontains 'ONLINE') { $overall = 'Info' }

        $item = [pscustomobject]@{
            Scope = $scope; DatabaseName = $first.DatabaseName; AgName = $first.AgName; RecoveryModel = $recModel
            ReplicasQueried = (@($rows | ForEach-Object SqlNode | Sort-Object -Unique) -join ', ')
            LastFull = if ($lastFull) { $lastFull.LastFull } else { $null }; FullAge = Format-Age $(if ($lastFull) { $lastFull.LastFull }); FullTakenOn = if ($lastFull) { $lastFull.SqlNode } else { '' }; FullStatus = $fullSt
            LastDiff = if ($lastDiff) { $lastDiff.LastDiff } else { $null }; DiffAge = if ($lastDiff) { Format-Age $lastDiff.LastDiff } else { 'Not used' }; DiffStatus = $diffSt
            LastLog = if ($lastLog) { $lastLog.LastLog } else { $null }; LogAge = if ($lastLog) { Format-Age $lastLog.LastLog } else { $(if ($logSt -eq 'N/A') { 'N/A' } else { 'Never' }) }; LogTakenOn = if ($lastLog) { $lastLog.SqlNode } else { '' }; LogStatus = $logSt
            Status = $overall
            SqlNode = if ($first.AgName) { 'AG (all replicas)' } else { $first.SqlNode }
        }
        if ($overall -in 'Critical', 'Warning') {
            $why = @(); if ($fullSt -ne 'OK') { $why += "full: $($item.FullAge)" }; if ($diffSt -eq 'Warning') { $why += "diff: $($item.DiffAge)" }; if ($logSt -in 'Critical', 'Warning') { $why += "log: $($item.LogAge)" }
            Add-Finding -Status $overall -Area 'Backups' -Node $item.SqlNode -Item $item.DatabaseName -Detail ($why -join '; ')
        }
        $item
    }
)
$BackupGaps = @($BackupCompliance | Where-Object { $_.Status -in 'Critical', 'Warning' })

# ---- Capacity ----
$SqlVolumes = @($SqlData.SqlVolumes | ForEach-Object {
    $pct = 0; if ([double]$_.TotalGB -gt 0) { $pct = [math]::Round(100 * [double]$_.FreeGB / [double]$_.TotalGB, 1) }
    $st = if ($pct -lt $DiskFreeCriticalPct) { 'Critical' } elseif ($pct -lt $DiskFreeWarningPct) { 'Warning' } else { 'OK' }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $_.SqlNode -Item "Volume $($_.VolumePath)" -Detail "$pct% free ($($_.FreeGB) GB of $($_.TotalGB) GB)" }
    $_ | Add-Member -NotePropertyName FreePct -NotePropertyValue $pct -Force
    $_ | Add-Member -NotePropertyName UsedGB -NotePropertyValue ([math]::Round([double]$_.TotalGB - [double]$_.FreeGB, 2)) -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

foreach ($f in $SqlData.FileLatency) {
    $worst = [math]::Max([double]$f.AvgReadMs, [double]$f.AvgWriteMs)
    $st = if ($worst -ge $IoLatencyCriticalMs) { 'Critical' } elseif ($worst -ge $IoLatencyWarningMs) { 'Warning' } else { 'OK' }
    $f | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    if ($st -ne 'OK' -and ([long]$f.Reads + [long]$f.Writes) -gt 1000) { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $f.SqlNode -Item "$($f.DatabaseName) / $($f.LogicalName)" -Detail "Avg read $($f.AvgReadMs) ms, avg write $($f.AvgWriteMs) ms" }
}

$vlfIndex = @{}; foreach ($v in $SqlData.VlfCounts) { $vlfIndex["$($v.SqlNode)|$($v.DatabaseName)"] = $v }
$LogUsage = @($SqlData.LogUsage | ForEach-Object {
    $v = $vlfIndex["$($_.SqlNode)|$($_.DatabaseName)"]
    $vlf = if ($v) { [int]$v.VlfCount } else { $null }
    $st = 'OK'; $why = @()
    if ([double]$_.LogUsedPct -ge $LogUsedWarningPct -and $_.LogReuseWait -ne 'NOTHING') { $st = 'Warning'; $why += "log $($_.LogUsedPct)% used, reuse wait $($_.LogReuseWait)" }
    if ($vlf -and $vlf -ge $VlfWarningCount) { $st = 'Warning'; $why += "$vlf VLFs" }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Capacity & Performance' -Node $_.SqlNode -Item "$($_.DatabaseName) transaction log" -Detail ($why -join '; ') }
    $_ | Add-Member -NotePropertyName VlfCount -NotePropertyValue $vlf -Force
    $_ | Add-Member -NotePropertyName ActiveVlfs -NotePropertyValue $(if ($v) { $v.ActiveVlfCount }) -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

foreach ($m in $SqlData.Memory) {
    if ($m.SystemMemoryState -match 'low') { Add-Finding -Status Critical -Area 'Capacity & Performance' -Node $m.SqlNode -Item 'Host memory' -Detail $m.SystemMemoryState }
    if ([double]$m.MaxServerMemoryMB -ge 2147483647) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $m.SqlNode -Item 'max server memory' -Detail 'Left at default (unlimited).' }
    if ([double]$m.PageLifeExpectancy -lt 300) { Add-Finding -Status Warning -Area 'Capacity & Performance' -Node $m.SqlNode -Item 'Page life expectancy' -Detail "$($m.PageLifeExpectancy) seconds" }
}
foreach ($node in $script:NodeLabels) {
    $cpu = @($SqlData.CpuHistory | Where-Object SqlNode -eq $node)
    if ($cpu.Count -gt 0) {
        $avg = ($cpu | Measure-Object -Property SQLProcessUtilization -Average).Average
        if ($avg -ge 80) { Add-Finding -Status Critical -Area 'Capacity & Performance' -Node $node -Item 'SQL Server CPU' -Detail ('Average {0:N0}% over {1} minutes' -f $avg, $cpu.Count) }
        elseif ($avg -ge 60) { Add-Finding -Status Warning -Area 'Capacity & Performance' -Node $node -Item 'SQL Server CPU' -Detail ('Average {0:N0}% over {1} minutes' -f $avg, $cpu.Count) }
    }
}

# ---- Cluster & Windows ----
foreach ($n in $ClusterData.Nodes)     { if ($n.Status -ne 'OK') { Add-Finding -Status $n.Status -Area 'Cluster & Windows' -Node $n.ClusterName -Item "Cluster node $($n.NodeName)" -Detail "State $($n.State)" } }
foreach ($g in $ClusterData.Groups)    { if ($g.Status -ne 'OK') { Add-Finding -Status $g.Status -Area 'Cluster & Windows' -Node $g.ClusterName -Item "Cluster group $($g.GroupName)" -Detail "State $($g.State) on $($g.OwnerNode)" } }
foreach ($r in $ClusterData.Resources) { if ($r.Status -eq 'Critical') { Add-Finding -Status Critical -Area 'Cluster & Windows' -Node $r.ClusterName -Item "Cluster resource $($r.ResourceName)" -Detail "State $($r.State)" } }
foreach ($s in $WinData.Services)  { if ($s.Status -eq 'Critical') { Add-Finding -Status Critical -Area 'Cluster & Windows' -Node $s.SqlNode -Item "Service $($s.ServiceName)" -Detail "Automatic service is $($s.State)" } }
foreach ($v in $WinData.Volumes)   { if ($v.Status -ne 'OK') { Add-Finding -Status $v.Status -Area 'Cluster & Windows' -Node $v.SqlNode -Item "Windows volume $($v.Volume)" -Detail "$($v.FreePct)% free" } }

$clusterProviders = 'Microsoft-Windows-FailoverClustering|Microsoft-Windows-DNS-Client|DNS Client Events|NETLOGON|Microsoft-Windows-Security-Kerberos|LsaSrv|Microsoft-Windows-GroupPolicy|Microsoft-Windows-Time-Service'
$dnsIds = @(1196, 1257, 1578, 1579)
$adIds  = @(1193, 1194, 1206, 1207, 1222, 1223, 1226, 1228, 1683, 1685, 1688)
$ClusterDnsAdEvents = @($WinData.Events | Where-Object { $_.LogName -eq 'System' -and $_.Provider -match $clusterProviders } | ForEach-Object {
    $cat = 'Cluster'
    if ($_.Provider -match 'DNS' -or ($_.Provider -match 'FailoverClustering' -and $dnsIds -contains $_.EventId)) { $cat = 'DNS' }
    elseif ($_.Provider -match 'NETLOGON|Kerberos|LsaSrv|GroupPolicy|Time-Service' -or ($_.Provider -match 'FailoverClustering' -and $adIds -contains $_.EventId)) { $cat = 'Active Directory' }
    $_ | Add-Member -NotePropertyName Category -NotePropertyValue $cat -Force
    $_ })
$WindowsCritical = @($WinData.Events | Where-Object { $_.Level -eq 'Critical' -or ($_.Level -eq 'Error' -and $_.Message -match 'fail|unexpected|terminat|crash|corrupt') })
foreach ($g in ($WindowsCritical | Where-Object Level -eq 'Critical' | Group-Object SqlNode)) {
    Add-Finding -Status Critical -Area 'SQL Errors' -Node $g.Name -Item 'Windows critical events' -Detail "$($g.Count) critical event(s) in the last $LookbackHours h"
}
foreach ($g in ($ClusterDnsAdEvents | Where-Object Level -in 'Critical', 'Error' | Group-Object SqlNode)) {
    Add-Finding -Status Warning -Area 'Cluster & Windows' -Node $g.Name -Item 'Cluster / DNS / AD events' -Detail "$($g.Count) error event(s) in the last $LookbackHours h"
}

# ---- SQL errors ----
$ErrorGroups = @($SqlData.ErrorLog | Group-Object SqlNode, @{ Expression = {
        $t = [string]$_.Message
        $t = $t -replace '\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b', '<guid>'
        $t = $t -replace "'[^']*'", "'<x>'" -replace '\[CLIENT: [^\]]*\]', '[CLIENT: <x>]' -replace '\d+', '#'
        $t.Trim() } } | ForEach-Object {
    $items = @($_.Group | Sort-Object LogDate)
    [pscustomobject]@{
        SqlNode = $items[0].SqlNode; Occurrences = $items.Count; Level = Get-WorstStatus @($items | ForEach-Object { if ($_.Level -eq 'Critical') { 'Critical' } elseif ($_.Level -eq 'Error') { 'Warning' } else { 'Info' } })
        FirstSeen = $items[0].LogDate; LastSeen = $items[-1].LogDate; SampleMessage = $items[-1].Message
    } } | Sort-Object Occurrences -Descending)
foreach ($g in ($SqlData.ErrorLog | Where-Object Level -eq 'Critical' | Group-Object SqlNode)) {
    Add-Finding -Status Critical -Area 'SQL Errors' -Node $g.Name -Item 'Critical error log entries' -Detail "$($g.Count) entries (severity 17+, stack dumps, corruption or slow I/O) in $LookbackHours h"
}
foreach ($g in ($SqlData.ErrorLog | Where-Object Message -like '*Login failed*' | Group-Object SqlNode)) {
    if ($g.Count -ge 10) { Add-Finding -Status Warning -Area 'SQL Errors' -Node $g.Name -Item 'Failed logins' -Detail "$($g.Count) failed login messages in $LookbackHours h" }
}
foreach ($s in $SqlData.SuspectPages) {
    if ($s.EventType -notin 'Restored', 'Repaired', 'Deallocated by DBCC') { Add-Finding -Status Critical -Area 'SQL Errors' -Node $s.SqlNode -Item "Suspect page in $($s.DatabaseName)" -Detail "File $($s.FileId) page $($s.PageId): $($s.EventType), $($s.ErrorCount) error(s)" }
}
foreach ($g in ($SqlData.MemoryDumps | Where-Object { $_.Created -ge $now.AddDays(-$HistoryLookbackDays) } | Group-Object SqlNode)) {
    Add-Finding -Status Warning -Area 'SQL Errors' -Node $g.Name -Item 'SQL Server memory dumps' -Detail "$($g.Count) dump(s) in the last $HistoryLookbackDays days"
}

# ---- Security & configuration ----
$DbSecurity = @($SqlData.DbSecurity | ForEach-Object {
    $isSys = $_.DatabaseName -in @('master', 'model', 'msdb')
    $why = @()
    if ($_.Trustworthy -eq $true -and $_.DatabaseName -ne 'msdb') { $why += 'TRUSTWORTHY on' }
    if ($_.DbChaining -eq $true -and -not $isSys) { $why += 'DB chaining on' }
    if (-not $_.Owner) { $why += 'Owner has no matching login' }
    $_ | Add-Member -NotePropertyName Encryption -NotePropertyValue $(switch ([int]$_.EncryptionState) { 0 { 'None' } 1 { 'Unencrypted' } 2 { 'Encrypting' } 3 { 'Encrypted' } 4 { 'Key change' } 5 { 'Decrypting' } 6 { 'Protection change' } default { 'None' } }) -Force
    $_ | Add-Member -NotePropertyName Issues -NotePropertyValue ($why -join '; ') -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $(if ($why.Count) { 'Warning' } else { 'OK' }) -Force
    if ($why.Count) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $_.SqlNode -Item $_.DatabaseName -Detail ($why -join '; ') }
    $_ })
foreach ($g in ($SqlData.OrphanedUsers | Group-Object SqlNode)) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $g.Name -Item 'Orphaned database users' -Detail "$($g.Count) user(s) without a matching login" }
foreach ($g in $SqlData.GuestAccess) { Add-Finding -Status Warning -Area 'Security & Configuration' -Node $g.SqlNode -Item $g.DatabaseName -Detail 'guest user has CONNECT permission' }

$SecurityChecks = @(foreach ($s in $SqlData.InstanceSecurity) {
    $chk = @(
        @{ Check = 'sa login disabled';                   Value = [bool]$s.SaDisabled;               Expected = $true;  Sev = 'Warning' }
        @{ Check = 'sa login renamed';                    Value = ($s.SaLoginName -ne 'sa');          Expected = $true;  Sev = 'Info' }
        @{ Check = 'Windows authentication only';         Value = [bool]$s.WindowsAuthOnly;          Expected = $true;  Sev = 'Info' }
        @{ Check = 'xp_cmdshell disabled';                Value = ([int]$s.XpCmdShell -eq 0);         Expected = $true;  Sev = 'Warning' }
        @{ Check = 'CLR disabled';                        Value = ([int]$s.ClrEnabled -eq 0);         Expected = $true;  Sev = 'Info' }
        @{ Check = 'OLE Automation disabled';             Value = ([int]$s.OleAutomation -eq 0);      Expected = $true;  Sev = 'Warning' }
        @{ Check = 'Ad Hoc Distributed Queries disabled'; Value = ([int]$s.AdHocDistributedQueries -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'Cross-database ownership chaining off'; Value = ([int]$s.CrossDbOwnershipChaining -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'Remote admin connections (DAC) enabled'; Value = ([int]$s.RemoteAdminConnections -eq 1); Expected = $true; Sev = 'Info' }
        @{ Check = 'No logins with blank password';       Value = ([int]$s.BlankPasswordLogins -eq 0); Expected = $true; Sev = 'Critical' }
        @{ Check = 'All SQL logins enforce password policy'; Value = ([int]$s.SqlLoginsWithoutPolicy -eq 0); Expected = $true; Sev = 'Warning' }
        @{ Check = 'All SQL logins enforce expiration';   Value = ([int]$s.SqlLoginsWithoutExpiration -eq 0); Expected = $true; Sev = 'Info' }
        @{ Check = 'sysadmin members (5 or fewer)';       Value = ([int]$s.SysadminMembers -le 5);    Expected = $true;  Sev = 'Warning' }
    )
    foreach ($c in $chk) {
        $pass = ($c.Value -eq $c.Expected)
        $st = if ($pass) { 'OK' } else { $c.Sev }
        if (-not $pass -and $c.Sev -ne 'Info') { Add-Finding -Status $c.Sev -Area 'Security & Configuration' -Node $s.SqlNode -Item $c.Check -Detail 'Check not met' }
        [pscustomobject]@{ SqlNode = $s.SqlNode; Check = $c.Check; Result = if ($pass) { 'Pass' } else { 'Not met' }; Status = $st }
    }
})
$Certificates = @($SqlData.Certificates | ForEach-Object {
    $days = if ($_.ExpiryDate) { [math]::Floor(([datetime]$_.ExpiryDate - $now).TotalDays) } else { $null }
    $st = if ($null -eq $days) { 'Info' } elseif ($days -lt 0) { 'Critical' } elseif ($days -le $CertificateExpiryWarningDays) { 'Warning' } else { 'OK' }
    if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Security & Configuration' -Node $_.SqlNode -Item "Certificate $($_.CertificateName) ($($_.DatabaseName))" -Detail "Expires $($_.ExpiryDate) ($days days)" }
    $_ | Add-Member -NotePropertyName DaysToExpiry -NotePropertyValue $days -Force
    $_ | Add-Member -NotePropertyName Status -NotePropertyValue $st -Force
    $_ })

function Compare-AcrossNodes {
    param([object[]] $Data, [string] $KeyProperty, [scriptblock] $ValueOf, [string] $Area, [string] $What, [switch] $Presence)
    $labels = @($script:NodeLabels)
    $index = @{}
    foreach ($r in $Data) { $index["$($r.SqlNode)|$($r.$KeyProperty)"] = $r }
    $keys = @($Data | ForEach-Object { [string]$_.$KeyProperty } | Sort-Object -Unique)
    foreach ($k in $keys) {
        $o = [ordered]@{ Item = $k }
        $vals = @()
        foreach ($l in $labels) {
            $r = $index["$l|$k"]
            $v = if ($null -eq $r) { 'Missing' } else { [string](& $ValueOf $r) }
            $o[$l] = $v; $vals += $v
        }
        $same = (@($vals | Sort-Object -Unique).Count -le 1)
        $o['Comparison'] = if ($same) { 'Match' } else { 'Mismatch' }
        if (-not $same -and $labels.Count -gt 1) { Add-Finding -Status Warning -Area $Area -Node 'HA nodes' -Item "$What $k" -Detail (($labels | ForEach-Object { "$_ = $($o[$_])" }) -join ' | ') }
        [pscustomobject]$o
    }
}
$ConfigCompare = @(); $PropertyCompare = @(); $LoginCompare = @(); $JobCompare = @(); $TraceFlagCompare = @(); $LinkedCompare = @()
if (@($script:NodeLabels).Count -gt 1) {
    $ConfigCompare   = @(Compare-AcrossNodes -Data $SqlData.Configurations -KeyProperty Setting -ValueOf { param($r) $r.Value } -Area 'Security & Configuration' -What 'sp_configure')
    $propRows = foreach ($inst in $script:NodeOrder) {
        $d = $Discovery[$inst]
        foreach ($p in 'Edition', 'ProductVersion', 'ProductUpdateLevel', 'Collation', 'MaxServerMemoryMB', 'MaxDop', 'CostThreshold', 'CpuCount', 'HostMemoryMB', 'IsHadrEnabled', 'WindowsAuthOnly') {
            [pscustomobject]@{ SqlNode = $d.SqlNode; Property = $p; Value = $d.$p }
        }
    }
    $PropertyCompare = @(Compare-AcrossNodes -Data $propRows -KeyProperty Property -ValueOf { param($r) $r.Value } -Area 'Security & Configuration' -What 'Server property')
    $LoginCompare    = @(Compare-AcrossNodes -Data $SqlData.Logins -KeyProperty LoginName -ValueOf { param($r) if ($r.IsDisabled) { 'Present (disabled)' } else { 'Present' } } -Area 'Security & Configuration' -What 'Login')
    $JobCompare      = @(Compare-AcrossNodes -Data $SqlData.AgentJobs -KeyProperty JobName -ValueOf { param($r) if ($r.IsEnabled) { 'Enabled' } else { 'Present (disabled)' } } -Area 'Security & Configuration' -What 'Agent job')
    $TraceFlagCompare = @(Compare-AcrossNodes -Data ($SqlData.TraceFlags | Where-Object { $_.Global -eq 1 }) -KeyProperty TraceFlag -ValueOf { param($r) 'Enabled' } -Area 'Security & Configuration' -What 'Trace flag')
    $LinkedCompare   = @(Compare-AcrossNodes -Data $SqlData.LinkedServers -KeyProperty LinkedServer -ValueOf { param($r) 'Present' } -Area 'Security & Configuration' -What 'Linked server')
}

# ---- Investigation ----
$JobTrends = @($SqlData.JobHistory | Group-Object SqlNode, JobName | ForEach-Object {
    $runs = @($_.Group | Sort-Object RunStart)
    $dur  = [double[]]@($runs | ForEach-Object { [double]$_.DurationSec })
    $avg  = ($dur | Measure-Object -Average).Average
    $max  = ($dur | Measure-Object -Maximum).Maximum
    $last = $runs[-1]
    $failed = @($runs | Where-Object RunStatus -eq 'Failed').Count
    $trendPct = if ($avg -gt 0) { [math]::Round(100 * ([double]$last.DurationSec - $avg) / $avg, 0) } else { 0 }
    $st = 'OK'
    if ($last.RunStatus -eq 'Failed') { $st = 'Critical' }
    elseif ($failed -gt 0 -or ($avg -ge 60 -and [double]$last.DurationSec -gt 2 * $avg)) { $st = 'Warning' }
    if ($st -ne 'OK') { Add-Finding -Status $st -Area 'Investigation' -Node $last.SqlNode -Item "Agent job $($last.JobName)" -Detail "Last outcome $($last.RunStatus); $failed failure(s) in $HistoryLookbackDays days; last run $(Format-Duration $last.DurationSec) vs avg $(Format-Duration $avg)" }
    [pscustomobject]@{
        SqlNode = $last.SqlNode; JobName = $last.JobName; Runs = $runs.Count; Failed = $failed
        AvgDuration = Format-Duration $avg; MaxDuration = Format-Duration $max; LastDuration = Format-Duration $last.DurationSec
        LastRun = $last.RunStart; LastOutcome = $last.RunStatus; LastVsAvgPct = $trendPct; Status = $st; Durations = $dur
    } })

$CheckDbStatus = @($SqlData.CheckDb | ForEach-Object {
    $last = $_.LastGoodCheckDb
    $st = 'OK'; $age = $null
    if ($null -eq $last) { $st = 'Info'; $age = 'Unknown (not readable on this replica)' }
    elseif (([datetime]$last).Year -le 1900) { $st = 'Critical'; $age = 'Never'; $last = $null }
    else {
        $days = ($now - [datetime]$last).TotalDays; $age = '{0:N1} days' -f $days
        if ($days -gt $CheckDbMaxAgeDays) { $st = 'Warning' }
    }
    if ($_.State -ne 'ONLINE') { $st = 'Info' }
    if ($st -in 'Critical', 'Warning') { Add-Finding -Status $st -Area 'Investigation' -Node $_.SqlNode -Item "DBCC CHECKDB $($_.DatabaseName)" -Detail "Last known good: $age (threshold $CheckDbMaxAgeDays days)" }
    [pscustomobject]@{ SqlNode = $_.SqlNode; DatabaseName = $_.DatabaseName; State = $_.State; LastKnownGood = $last; Age = $age; Status = $st }
})
foreach ($e in ($SqlData.CheckDbLog | Where-Object Status -eq 'Critical')) { Add-Finding -Status Critical -Area 'Investigation' -Node $e.SqlNode -Item 'DBCC CHECKDB reported errors' -Detail $e.Message }
foreach ($g in ($SqlData.LongRunning | Group-Object SqlNode)) { Add-Finding -Status Warning -Area 'Investigation' -Node $g.Name -Item 'Long-running requests' -Detail "$($g.Count) request(s) running longer than $LongRunningQuerySeconds s" }
foreach ($g in ($SqlData.Blocking | Group-Object SqlNode))    { Add-Finding -Status Warning -Area 'Investigation' -Node $g.Name -Item 'Blocking' -Detail "$($g.Count) blocked request(s) waiting longer than $BlockingWaitSeconds s" }
if ($script:CollectionErrors.Count -gt 0) {
    Add-Finding -Status Info -Area 'Investigation' -Node 'All' -Item 'Collection errors' -Detail "$($script:CollectionErrors.Count) section(s) could not be collected; see 'Collection errors and permission issues'."
}

# ---- Optional metric history ----
$HistoryRows = @()
if ($EnableMetricHistory) {
    $histFolder = Join-Path $OutputFolder 'History'
    if (-not (Test-Path -LiteralPath $histFolder)) { New-Item -ItemType Directory -Path $histFolder -Force | Out-Null }
    $histFile = Join-Path $histFolder 'MetricHistory.csv'
    $stamp = $now.ToString('yyyy-MM-dd HH:mm:ss', $script:Inv)
    $newRows = foreach ($node in $script:NodeLabels) {
        $cpu = @($SqlData.CpuHistory | Where-Object SqlNode -eq $node)
        $mem = @($SqlData.Memory | Where-Object SqlNode -eq $node) | Select-Object -First 1
        $metrics = [ordered]@{
            'SQL CPU avg %'            = if ($cpu) { [math]::Round(($cpu | Measure-Object -Property SQLProcessUtilization -Average).Average, 1) } else { $null }
            'Page life expectancy'     = if ($mem) { $mem.PageLifeExpectancy } else { $null }
            'SQL memory in use MB'     = if ($mem) { $mem.SqlPhysicalMemoryMB } else { $null }
            'Databases'                = @($SqlData.Databases | Where-Object SqlNode -eq $node).Count
            'Total database size GB'   = [math]::Round((($SqlData.Databases | Where-Object SqlNode -eq $node | ForEach-Object { [double]$_.DataMB + [double]$_.LogMB }) | Measure-Object -Sum).Sum / 1024, 2)
            'Critical findings'        = @($script:Findings | Where-Object { $_.SqlNode -eq $node -and $_.Status -eq 'Critical' }).Count
            'Warning findings'         = @($script:Findings | Where-Object { $_.SqlNode -eq $node -and $_.Status -eq 'Warning' }).Count
            'Error log entries'        = @($SqlData.ErrorLog | Where-Object SqlNode -eq $node).Count
        }
        foreach ($k in $metrics.Keys) { [pscustomobject]@{ Timestamp = $stamp; SqlNode = $node; Metric = $k; Value = $metrics[$k] } }
    }
    $existing = @()
    if (Test-Path -LiteralPath $histFile) { $existing = @(Import-Csv -LiteralPath $histFile) }
    $cutoff = $now.AddDays(-$MetricHistoryRetentionDays)
    $kept = @($existing | Where-Object {
        $t = [datetime]::MinValue
        [datetime]::TryParseExact($_.Timestamp, 'yyyy-MM-dd HH:mm:ss', $script:Inv, [Globalization.DateTimeStyles]::None, [ref]$t) -and $t -ge $cutoff })
    $HistoryRows = @($kept) + @($newRows)
    $HistoryRows | Export-Csv -LiteralPath $histFile -NoTypeInformation -Encoding UTF8
}

# =====================================================================================
#  Phase 5 - HTML dashboard
# =====================================================================================
Write-Host '  Building HTML dashboard ...'
$AreaNames = @('Databases', 'Always On', 'Backups', 'Capacity & Performance', 'Cluster & Windows', 'SQL Errors', 'Security & Configuration', 'Investigation')
$AreaSummary = @(foreach ($a in $AreaNames) {
    $f = @($script:Findings | Where-Object Area -eq $a)
    $crit = @($f | Where-Object Status -eq 'Critical').Count; $warn = @($f | Where-Object Status -eq 'Warning').Count
    [pscustomobject]@{ Area = $a; Critical = $crit; Warning = $warn; Status = if ($crit) { 'Critical' } elseif ($warn) { 'Warning' } else { 'OK' } }
})
$critTotal = @($script:Findings | Where-Object Status -eq 'Critical').Count
$warnTotal = @($script:Findings | Where-Object Status -eq 'Warning').Count
$overall   = if ($critTotal) { 'Critical' } elseif ($warnTotal) { 'Warning' } else { 'OK' }
$dbTotal   = @($SqlData.Databases).Count
$dbOnline  = @($SqlData.Databases | Where-Object State -eq 'ONLINE').Count
$duration  = Format-Duration ((Get-Date) - $script:ReportStart).TotalSeconds

$Css = @'
<style>
.ops-banner{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px;padding:14px 18px;border-radius:10px;background:linear-gradient(90deg,#1f3b64,#2f5d9a);color:#fff;margin-bottom:10px}
.ops-banner h1{font-size:20px;margin:0;font-weight:600;color:#fff}
.ops-banner .ops-meta{font-size:12px;opacity:.9}
.ops-kpi-grid,.ops-card-grid{display:flex;flex-wrap:wrap;gap:12px;margin:4px 0 8px}
.ops-kpi{flex:1 1 150px;border-radius:8px;padding:12px 14px;background:#fff;border:1px solid #e3e7ee;border-left:5px solid #3a7bd5}
.ops-kpi.st-ok,.ops-card.st-ok{border-left-color:#2e9d5b}.ops-kpi.st-warning,.ops-card.st-warning{border-left-color:#e0a100}
.ops-kpi.st-critical,.ops-card.st-critical{border-left-color:#d64545}.ops-kpi.st-info,.ops-card.st-info{border-left-color:#3a7bd5}
.ops-kpi-title{font-size:11px;text-transform:uppercase;color:#6b7280;letter-spacing:.05em}
.ops-kpi-value{font-size:26px;font-weight:600;color:#1f2937;margin-top:4px}
.ops-kpi-sub{font-size:12px;color:#6b7280;margin-top:2px}
.ops-card{flex:1 1 260px;border-radius:8px;padding:10px 14px;background:#fff;border:1px solid #e3e7ee;border-left:5px solid #3a7bd5}
.ops-card-head{display:flex;justify-content:space-between;align-items:center;font-weight:600;color:#1f2937;margin-bottom:6px}
.ops-card-table{width:100%;font-size:12px;border-collapse:collapse}
.ops-card-table td{padding:3px 0;border-bottom:1px solid #f1f3f6}.ops-card-table td:last-child{text-align:right;font-weight:500}
.ops-card-foot{font-size:11px;color:#6b7280;margin-top:6px}
.ops-pill{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:600}
.ops-pill.st-ok{background:#d8f3e3;color:#1e6b3e}.ops-pill.st-warning{background:#fff1cc;color:#7a5b00}
.ops-pill.st-critical{background:#f9d6d6;color:#8a1f1f}.ops-pill.st-info{background:#e3edfb;color:#1f4f8f}
.ops-node-title{font-weight:600;font-size:14px;margin:2px 0 8px;color:#1f3b64;border-bottom:2px solid #e3e7ee;padding-bottom:4px}
.ops-group-head{font-weight:600;font-size:13px;margin:14px 0 6px;color:#1f2937}
.ops-muted{color:#6b7280;font-weight:400;font-size:12px}
.ops-empty{color:#6b7280;font-style:italic;padding:8px 2px}
.ops-chart{margin:6px 0 12px}
.ops-chart-title{font-weight:600;font-size:13px;color:#1f2937;margin-bottom:6px}
.ops-chart-legend{display:flex;flex-wrap:wrap;gap:14px;font-size:12px;margin:4px 0}
.ops-svg-wide{width:100%;max-width:900px;height:auto}
.ops-axis{font-size:10px;fill:#6b7280}
.ops-donut-body{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.ops-donut-total{font-size:15px;font-weight:600;fill:#1f2937}.ops-donut-unit{font-size:10px;fill:#6b7280}
.ops-legend{list-style:none;margin:0;padding:0;font-size:12px;min-width:200px}
.ops-legend li{display:flex;align-items:center;gap:6px;margin:3px 0}
.ops-legend-label{max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ops-legend-val{color:#6b7280;margin-left:auto;padding-left:8px;white-space:nowrap}
.ops-swatch{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:4px;vertical-align:middle}
.ops-bar-row{display:flex;align-items:center;gap:8px;margin:3px 0;font-size:12px}
.ops-bar-label{flex:0 0 200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ops-bar-track{flex:1;background:#eef1f5;border-radius:4px;height:14px;overflow:hidden;display:flex}
.ops-bar-track.small{height:8px;display:inline-flex;width:110px;vertical-align:middle;margin-right:6px}
.ops-bar-fill{height:100%}
.ops-bar-value{flex:0 0 90px;text-align:right;color:#374151}
.ops-mini-table{width:100%;border-collapse:collapse;font-size:12px}
.ops-mini-table th{text-align:left;background:#f4f6f9;padding:5px 6px;font-weight:600;color:#374151}
.ops-mini-table td{padding:4px 6px;border-bottom:1px solid #eef1f5;vertical-align:middle}
.ops-num{font-variant-numeric:tabular-nums;white-space:nowrap}
.ops-donut-grid{display:flex;flex-wrap:wrap;gap:18px}
.ops-donut-grid .ops-chart{flex:1 1 360px}
</style>
'@

$Banner = "<div class='ops-banner'><div><h1>SQL Server Operations Dashboard</h1><div class='ops-meta'>$(ConvertTo-HtmlSafe $EnvironmentName) &middot; $(ConvertTo-HtmlSafe $EnvironmentMode)</div></div>" +
          "<div class='ops-meta'>Generated $($now.ToString('yyyy-MM-dd HH:mm')) by $(ConvertTo-HtmlSafe "$env:USERDOMAIN\$env:USERNAME") on $(ConvertTo-HtmlSafe $env:COMPUTERNAME) &middot; collection time $duration<br/>" +
          "Nodes: $(ConvertTo-HtmlSafe (($script:NodeLabels + @($Unreachable | ForEach-Object { Get-ReportNodeLabel $_ })) -join ', ')) &middot; read-only collection</div></div>"

$Kpis = "<div class='ops-kpi-grid'>" +
    (ConvertTo-KpiHtmlText -Title 'Overall status' -Value $overall -Subtitle "$critTotal critical, $warnTotal warning findings" -Status $overall) +
    (ConvertTo-KpiHtmlText -Title 'SQL Server nodes' -Value "$(@($script:NodeOrder).Count) / $(@($SqlInstances).Count)" -Subtitle 'reachable / configured' -Status $(if ($Unreachable.Count) { 'Critical' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Databases online' -Value "$dbOnline / $dbTotal" -Subtitle 'across all nodes' -Status $(if ($dbOnline -lt $dbTotal) { 'Warning' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Availability Groups' -Value "$(@($AgTopology | Select-Object -ExpandProperty AgName -Unique).Count)" -Subtitle "$(@($AgTopology | Where-Object Status -eq 'OK').Count) of $(@($AgTopology).Count) replicas healthy" -Status (Get-WorstStatus @($AgTopology | ForEach-Object Status))) +
    (ConvertTo-KpiHtmlText -Title 'Backup compliance' -Value "$(@($BackupCompliance | Where-Object Status -eq 'OK').Count) / $(@($BackupCompliance | Where-Object Status -ne 'Info').Count)" -Subtitle 'databases compliant' -Status (Get-WorstStatus @($BackupCompliance | ForEach-Object Status))) +
    (ConvertTo-KpiHtmlText -Title 'SQL error log' -Value "$(@($SqlData.ErrorLog).Count)" -Subtitle "entries in last $LookbackHours h" -Status $(if (@($SqlData.ErrorLog | Where-Object Level -eq 'Critical').Count) { 'Critical' } elseif (@($SqlData.ErrorLog).Count) { 'Warning' } else { 'OK' })) +
    (ConvertTo-KpiHtmlText -Title 'Collection errors' -Value "$($script:CollectionErrors.Count)" -Subtitle 'sections not collected' -Status $(if ($script:CollectionErrors.Count) { 'Warning' } else { 'OK' })) +
    '</div>'

$HealthCards = "<div class='ops-card-grid'>" + (ConvertTo-AgHealthCard -Replicas $AgTopology -Databases $AgDbStates) + (ConvertTo-BackupHealthCard -Compliance $BackupCompliance)
foreach ($node in $script:NodeLabels) {
    $m = @($SqlData.Memory | Where-Object SqlNode -eq $node) | Select-Object -First 1
    $HealthCards += (ConvertTo-HostMemoryCard -Memory $m -Node $node) + (ConvertTo-SqlProcessMemoryCard -Memory $m -Node $node)
}
$HealthCards += '</div>'

$dbColumns     = @('SqlNode', 'DatabaseName', 'State', 'Status', 'RecoveryModel', 'CompatLevel', 'Owner', 'TotalSize', 'DataMB', 'LogMB', 'UserAccess', 'IsReadOnly', 'PageVerify', 'LogReuseWait', 'AgName', 'Created', 'Notes')
$backupColumns = @('DatabaseName', 'AgName', 'RecoveryModel', 'Status', 'LastFull', 'FullAge', 'FullTakenOn', 'FullStatus', 'LastDiff', 'DiffAge', 'DiffStatus', 'LastLog', 'LogAge', 'LogTakenOn', 'LogStatus', 'ReplicasQueried')

New-HTML -TitleText "SQL Server Operations Dashboard - $EnvironmentName" -FilePath $ReportPath {

    # ---------------------------------------------------------------- 1. Overview
    New-HTMLTab -Name 'Overview' {
        New-HTMLSection -Invisible -Direction column { $Css; $Banner; $Kpis }
        New-HTMLSection -Direction column -HeaderText 'Operational health' -CanCollapse { $HealthCards }
        New-HTMLSection -Direction column -HeaderText 'Status summary by area' -CanCollapse { New-SafeHTMLTable -Data $AreaSummary -PagingLength 10 }
        New-HTMLSection -Direction column -HeaderText 'SQL Server instance inventory' -CanCollapse { New-SafeHTMLTable -Data $InstanceInventory }
        New-HTMLSection -Direction column -HeaderText 'Database inventory and health' -CanCollapse { New-DatabaseGroupedTable -Data $SqlData.Databases -Columns $dbColumns }
        New-HTMLSection -Direction column -HeaderText 'Operating system per SQL Server node' -CanCollapse { New-SafeHTMLTable -Data $WinData.OS -EmptyMessage 'Operating system information could not be collected (see Investigation).' }
        New-HTMLSection -Direction column -HeaderText 'SQL Server-related Windows services' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Services -Columns @('SqlNode', 'ServiceName', 'DisplayName', 'State', 'StartMode', 'ServiceAccount', 'Status')
        }
    }

    # ---------------------------------------------------------------- 2. Always On
    New-HTMLTab -Name 'Always On' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        if (-not $HasAG) {
            New-HTMLSection -Direction column -HeaderText 'Always On Availability Groups' { "<div class='ops-empty'>No Availability Groups were detected on the queried instances. The environment is reported as $(ConvertTo-HtmlSafe $EnvironmentMode).</div>" }
        } else {
            New-HTMLSection -Direction column -HeaderText 'Availability Group health' -CanCollapse { "<div class='ops-card-grid'>$(ConvertTo-AgHealthCard -Replicas $AgTopology -Databases $AgDbStates)</div>" }
            New-HTMLSection -Direction column -HeaderText 'Availability Group topology and replica roles' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgTopology -Columns @('AgName', 'ReplicaServer', 'Role', 'AvailabilityMode', 'FailoverMode', 'SeedingMode', 'ReadableSecondary', 'BackupPriority', 'BackupPreference', 'SessionTimeout', 'EndpointUrl', 'ReportedBy'))
            }
            New-HTMLSection -Direction column -HeaderText 'Replica connectivity, synchronization state and health' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgTopology -Columns @('AgName', 'ReplicaServer', 'Role', 'ConnectedState', 'OperationalState', 'RecoveryHealth', 'SynchronizationHealth', 'Status', 'LastConnectError', 'LastConnectErrorDescription', 'LastConnectErrorTime'))
            }
            New-HTMLSection -Direction column -HeaderText 'Database synchronization' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgDbStates -Columns @('AgName', 'DatabaseName', 'ReplicaServer', 'IsPrimaryReplica', 'SynchronizationState', 'SynchronizationHealth', 'DatabaseState', 'IsSuspended', 'SuspendReason', 'Status', 'LastCommitTime', 'LastHardenedTime', 'LastRedoneTime')) -PagingLength 15
            }
            New-HTMLSection -Direction column -HeaderText "Log send and redo queues (warning: send > $AgSendQueueWarningKB KB, redo > $AgRedoQueueWarningKB KB)" -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data ($AgDbStates | Where-Object { $_.IsPrimaryReplica -ne $true }) -Columns @('AgName', 'DatabaseName', 'ReplicaServer', 'LogSendQueueKB', 'LogSendRateKBs', 'RedoQueueKB', 'RedoRateKBs', 'EstDataLossSec', 'EstRecoverySec', 'Status')) -PagingLength 15 -EmptyMessage 'No secondary database replica statistics available.'
            }
            New-HTMLSection -Direction column -HeaderText 'Availability Group listener configuration' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $AgListeners -Columns @('AgName', 'ListenerName', 'Port', 'IpAddress', 'SubnetMask', 'IsDhcp', 'State', 'Status')) -EmptyMessage 'No listeners are configured.'
            }
            New-HTMLSection -Direction column -HeaderText 'Independent database inventory per HA node' -CanCollapse {
                New-NodeAwareTable -Data $SqlData.Databases -Columns @('SqlNode', 'DatabaseName', 'State', 'RecoveryModel', 'AgName', 'TotalSize', 'Status') -PagingLength 15
            }
        }
    }

    # ---------------------------------------------------------------- 3. Backups
    New-HTMLTab -Name 'Backups' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Backup health' -CanCollapse { "<div class='ops-card-grid'>$(ConvertTo-BackupHealthCard -Compliance $BackupCompliance)</div>" }
        New-HTMLSection -Direction column -HeaderText "Backup freshness and compliance (full $FullBackupMaxAgeHours h, differential $DiffBackupMaxAgeHours h, log $LogBackupMaxAgeMinutes min)" -CanCollapse {
            New-BackupComplianceGroupedTable -Data $BackupCompliance -Columns $backupColumns
        }
        New-HTMLSection -Direction column -HeaderText 'Missing or outdated backup information' -CanCollapse {
            New-SafeHTMLTable -Data (Select-ReportColumns -Data $BackupGaps -Columns (@('Scope') + $backupColumns)) -EmptyMessage 'Every evaluated database meets the configured backup thresholds.'
        }
        New-HTMLSection -Direction column -HeaderText 'Backup status across queried Availability Group replicas' -CanCollapse {
            New-NodeAwareTable -Data ($SqlData.BackupStatus | Where-Object AgName) -Columns @('SqlNode', 'DatabaseName', 'AgName', 'LocalRole', 'LastFull', 'LastDiff', 'LastLog') -EmptyMessage 'No Availability Group databases on this node.'
        }
        New-HTMLSection -Direction column -HeaderText "Local backup history (last $HistoryLookbackDays days, diagnostic)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.BackupHistory -Columns @('SqlNode', 'DatabaseName', 'BackupType', 'CopyOnly', 'Started', 'Finished', 'DurationSec', 'SizeMB', 'CompressedMB', 'Device', 'ExecutedBy') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Recent database restore operations (last $HistoryLookbackDays days)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.RestoreHistory -Columns @('SqlNode', 'DatabaseName', 'RestoreDate', 'RestoreType', 'RestoredBy', 'WithReplace', 'WithRecovery', 'SourceDatabase', 'SourceServer', 'SourceBackupDate', 'Device') -EmptyMessage 'No restores recorded in the lookback period.'
        }
    }

    # ---------------------------------------------------------------- 4. Capacity & Performance
    New-HTMLTab -Name 'Capacity & Performance' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText "SQL Server CPU utilization history (last $CpuHistoryMinutes minutes)" -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-CpuSvg -Samples @($SqlData.CpuHistory | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server and operating system memory' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $m = @($SqlData.Memory | Where-Object SqlNode -eq $n) | Select-Object -First 1
                "<div class='ops-card-grid'>$(ConvertTo-SqlProcessMemoryCard -Memory $m -Node $n)$(ConvertTo-HostMemoryCard -Memory $m -Node $n)</div>"
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Database size and file allocation' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-DatabaseSizesHtml -Databases @($SqlData.Databases | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'Database files' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.DatabaseFiles -Columns @('SqlNode', 'DatabaseName', 'LogicalName', 'FileType', 'SizeMB', 'Growth', 'MaxSize', 'FileState', 'PhysicalName') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Storage volume capacity (warning < $DiskFreeWarningPct% free, critical < $DiskFreeCriticalPct% free)" -CanCollapse {
            New-NodeAwareTable -Data $SqlVolumes -Columns @('SqlNode', 'VolumePath', 'VolumeLabel', 'FileSystem', 'TotalGB', 'UsedGB', 'FreeGB', 'FreePct', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Database resource distribution' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $alloc = @($SqlData.Databases | Where-Object SqlNode -eq $n | ForEach-Object { [pscustomobject]@{ DatabaseName = $_.DatabaseName; AllocatedMB = [double]$_.DataMB + [double]$_.LogMB } })
                "<div class='ops-donut-grid'>" +
                (ConvertTo-DbResourcePie -Items $alloc -ValueProperty 'AllocatedMB' -Title 'Allocated storage per database' -Unit 'GB' -ValueFormatter { param($v) '{0:N1}' -f ($v / 1024) }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.BufferCache | Where-Object SqlNode -eq $n) -ValueProperty 'BufferCacheMB' -Title 'Buffer cache utilization per database' -Unit 'MB' -ValueFormatter { param($v) '{0:N0}' -f $v }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.DbIo | Where-Object SqlNode -eq $n) -ValueProperty 'TotalIoMB' -Title 'I/O activity per database (since startup)' -Unit 'GB' -ValueFormatter { param($v) '{0:N1}' -f ($v / 1024) }) +
                (ConvertTo-DbResourcePie -Items @($SqlData.DbCpu | Where-Object SqlNode -eq $n) -ValueProperty 'CpuMs' -Title 'Estimated CPU distribution per database (cached plans)' -Unit 's' -ValueFormatter { param($v) '{0:N0}' -f ($v / 1000) }) +
                '</div>'
            }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server wait statistics' -CanCollapse {
            New-NodeAwarePanels { param($n)
                $w = @($SqlData.WaitStats | Where-Object SqlNode -eq $n)
                ConvertTo-WaitDonutSvg -Waits $w
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $w -Columns @('WaitType', 'Percentage', 'WaitSeconds', 'ResourceSeconds', 'SignalSeconds', 'WaitCount', 'AvgWaitMs')) -PagingLength 15
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Database file I/O latency' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-FileIoLatencyHtml -Files @($SqlData.FileLatency | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'File I/O statistics' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.FileLatency -Columns @('SqlNode', 'DatabaseName', 'LogicalName', 'FileType', 'Reads', 'Writes', 'AvgReadMs', 'AvgWriteMs', 'TotalIoMB', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText "Transaction log usage, recovery models and VLFs (warning: > $LogUsedWarningPct% used, >= $VlfWarningCount VLFs)" -CanCollapse {
            New-NodeAwareTable -Data $LogUsage -Columns @('SqlNode', 'DatabaseName', 'RecoveryModel', 'LogSizeMB', 'LogUsedMB', 'LogUsedPct', 'LogReuseWait', 'VlfCount', 'ActiveVlfs', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server memory clerk consumption' -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-MemoryClerkBarsHtml -Clerks @($SqlData.MemoryClerks | Where-Object SqlNode -eq $n) }
        }
    }

    # ---------------------------------------------------------------- 5. Cluster & Windows
    New-HTMLTab -Name 'Cluster & Windows' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Windows Server operating system information' -CanCollapse { New-SafeHTMLTable -Data $WinData.OS -EmptyMessage 'Operating system information could not be collected.' }
        New-HTMLSection -Direction column -HeaderText 'Windows Failover Cluster' -CanCollapse {
            New-SafeHTMLTable -Data $ClusterData.Clusters -EmptyMessage $(if ($CollectClusterInfo) { 'No Windows Failover Cluster was detected for the SQL Server hosts.' } else { 'Cluster collection is disabled (-CollectClusterInfo $false).' })
        }
        New-HTMLSection -Direction column -HeaderText 'Cluster node health' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Nodes -EmptyMessage 'No cluster nodes.' }
        New-HTMLSection -Direction column -HeaderText 'Cluster resource groups and ownership' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Groups -EmptyMessage 'No cluster groups.' }
        New-HTMLSection -Direction column -HeaderText 'Cluster resources' -CanCollapse { New-SafeHTMLTable -Data $ClusterData.Resources -EmptyMessage 'No cluster resources.' -PagingLength 15 }
        New-HTMLSection -Direction column -HeaderText "Cluster-related DNS and Active Directory events (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $ClusterDnsAdEvents -Columns @('SqlNode', 'TimeCreated', 'Category', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage 'No cluster, DNS or Active Directory warnings or errors.'
        }
        New-HTMLSection -Direction column -HeaderText "Windows System event log - critical, error, warning (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data ($WinData.Events | Where-Object LogName -eq 'System') -Columns @('SqlNode', 'TimeCreated', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage $(if ($CollectEventViewer) { 'No matching events.' } else { 'Event Viewer collection is disabled.' })
        }
        New-HTMLSection -Direction column -HeaderText "Windows Application event log - critical, error, warning (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data ($WinData.Events | Where-Object LogName -eq 'Application') -Columns @('SqlNode', 'TimeCreated', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage $(if ($CollectEventViewer) { 'No matching events.' } else { 'Event Viewer collection is disabled.' })
        }
        New-HTMLSection -Direction column -HeaderText 'Windows service availability' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Services -Columns @('SqlNode', 'ServiceName', 'DisplayName', 'State', 'StartMode', 'ServiceAccount', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Windows storage volumes' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Volumes -Columns @('SqlNode', 'Volume', 'Label', 'FileSystem', 'CapacityGB', 'FreeGB', 'FreePct', 'BlockSizeKB', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Windows local accounts' -CanCollapse {
            New-NodeAwareTable -Data $WinData.LocalUsers -Columns @('SqlNode', 'AccountName', 'FullName', 'Disabled', 'LockedOut', 'PasswordRequired', 'PasswordExpires', 'Status')
        }
        New-HTMLSection -Direction column -HeaderText 'Local Administrators group membership' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Admins -Columns @('SqlNode', 'Member', 'MemberType', 'SID')
        }
    }

    # ---------------------------------------------------------------- 6. SQL Errors
    New-HTMLTab -Name 'SQL Errors' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText "Grouped SQL Server error messages (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $ErrorGroups -Columns @('SqlNode', 'Occurrences', 'Level', 'FirstSeen', 'LastSeen', 'SampleMessage') -EmptyMessage 'No errors or warnings in the SQL Server error log.'
        }
        New-HTMLSection -Direction column -HeaderText 'Recent operational errors and warnings' -CanCollapse {
            New-NodeAwareTable -Data ($SqlData.ErrorLog | Where-Object Level -in 'Critical', 'Error' | Sort-Object LogDate -Descending | Select-Object -First 200) -Columns @('SqlNode', 'LogDate', 'Level', 'ProcessInfo', 'Message') -EmptyMessage 'No errors in the lookback period.'
        }
        New-HTMLSection -Direction column -HeaderText "SQL Server error log analysis - all matched entries (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.ErrorLog -Columns @('SqlNode', 'LogDate', 'Level', 'ProcessInfo', 'Message') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Suspect pages (msdb.dbo.suspect_pages)' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.SuspectPages -Columns @('SqlNode', 'DatabaseName', 'FileId', 'PageId', 'EventType', 'ErrorCount', 'LastUpdate') -EmptyMessage 'No suspect pages recorded.'
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server memory dump indicators' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.MemoryDumps -Columns @('SqlNode', 'Created', 'SizeMB', 'DumpFile') -EmptyMessage 'No memory dumps recorded.'
        }
        New-HTMLSection -Direction column -HeaderText "Windows critical events and operational failures (last $LookbackHours h)" -CanCollapse {
            New-NodeAwareTable -Data $WindowsCritical -Columns @('SqlNode', 'TimeCreated', 'LogName', 'Level', 'Provider', 'EventId', 'Message') -EmptyMessage 'No Windows critical events or failures.'
        }
    }

    # ---------------------------------------------------------------- 7. Security & Configuration
    New-HTMLTab -Name 'Security & Configuration' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Database security compliance' -CanCollapse {
            New-NodeAwareTable -Data $SecurityChecks -Columns @('SqlNode', 'Check', 'Result', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Database security configuration' -CanCollapse {
            New-NodeAwareTable -Data $DbSecurity -Columns @('SqlNode', 'DatabaseName', 'Owner', 'Trustworthy', 'DbChaining', 'TdeEnabled', 'Encryption', 'Containment', 'BrokerEnabled', 'Issues', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Orphaned database users' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.OrphanedUsers -Columns @('SqlNode', 'DatabaseName', 'UserName', 'UserType', 'AuthenticationType', 'CreateDate') -EmptyMessage 'No orphaned users found.'
        }
        New-HTMLSection -Direction column -HeaderText 'Privileged database role membership' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.DbRoleMembers -Columns @('SqlNode', 'DatabaseName', 'RoleName', 'MemberName', 'MemberType') -PagingLength 15 -EmptyMessage 'No members in privileged database roles (other than dbo).'
        }
        New-HTMLSection -Direction column -HeaderText 'Privileged server role membership' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.ServerRoleMembers -Columns @('SqlNode', 'ServerRole', 'MemberName', 'MemberType', 'IsDisabled') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'Guest access in user databases' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.GuestAccess -Columns @('SqlNode', 'DatabaseName', 'PermissionState') -EmptyMessage 'guest has no CONNECT permission in user databases.'
        }
        New-HTMLSection -Direction column -HeaderText "Certificate expiration (warning within $CertificateExpiryWarningDays days)" -CanCollapse {
            New-NodeAwareTable -Data $Certificates -Columns @('SqlNode', 'DatabaseName', 'CertificateName', 'Subject', 'ExpiryDate', 'DaysToExpiry', 'KeyEncryption', 'Status') -EmptyMessage 'No user certificates found.'
        }
        if (@($script:NodeLabels).Count -gt 1) {
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - server properties' -CanCollapse { New-SafeHTMLTable -Data $PropertyCompare -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - sp_configure (differences first)' -CanCollapse { New-SafeHTMLTable -Data ($ConfigCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - logins' -CanCollapse { New-SafeHTMLTable -Data ($LoginCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - SQL Server Agent jobs' -CanCollapse { New-SafeHTMLTable -Data ($JobCompare | Sort-Object { $_.Comparison -eq 'Match' }, Item) -StatusColumns 'Comparison' -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Configuration comparison between HA nodes - global trace flags and linked servers' -CanCollapse {
                New-SafeHTMLTable -Data $TraceFlagCompare -StatusColumns 'Comparison' -EmptyMessage 'No global trace flags enabled on any node.'
                New-SafeHTMLTable -Data $LinkedCompare -StatusColumns 'Comparison' -EmptyMessage 'No linked servers on any node.'
            }
        } else {
            New-HTMLSection -Direction column -HeaderText 'SQL Server configuration (sp_configure)' -CanCollapse { New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.Configurations -Columns @('Setting', 'Value')) -PagingLength 15 }
            New-HTMLSection -Direction column -HeaderText 'Global trace flags and linked servers' -CanCollapse {
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.TraceFlags -Columns @('TraceFlag', 'Global', 'Session')) -EmptyMessage 'No global trace flags enabled.'
                New-SafeHTMLTable -Data (Select-ReportColumns -Data $SqlData.LinkedServers -Columns @('LinkedServer', 'Product', 'Provider', 'DataSource')) -EmptyMessage 'No linked servers.'
            }
        }
        New-HTMLSection -Direction column -HeaderText 'Windows user and administrator accounts' -CanCollapse {
            New-NodeAwareTable -Data $WinData.Admins -Columns @('SqlNode', 'Member', 'MemberType', 'SID') -EmptyMessage 'Administrator membership could not be collected.'
            New-NodeAwareTable -Data $WinData.LocalUsers -Columns @('SqlNode', 'AccountName', 'Disabled', 'PasswordRequired', 'PasswordExpires', 'Status')
        }
    }

    # ---------------------------------------------------------------- 8. Investigation
    New-HTMLTab -Name 'Investigation' {
        New-HTMLSection -Invisible -Direction column { $Banner }
        New-HTMLSection -Direction column -HeaderText 'Information requiring further administrative analysis' -CanCollapse {
            New-SafeHTMLTable -Data ($script:Findings | Sort-Object @{ Expression = { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } } }, Area, SqlNode) -PagingLength 25 -EmptyMessage 'No findings - every check is within its configured threshold.'
        }
        New-HTMLSection -Direction column -HeaderText "SQL Server Agent job duration trends (last $HistoryLookbackDays days)" -CanCollapse {
            New-NodeAwarePanels { param($n) ConvertTo-JobDurationTrendHtml -Trends @($JobTrends | Where-Object SqlNode -eq $n) }
        }
        New-HTMLSection -Direction column -HeaderText 'SQL Server Agent job statistics' -CanCollapse {
            New-NodeAwareTable -Data $JobTrends -Columns @('SqlNode', 'JobName', 'Runs', 'Failed', 'AvgDuration', 'MaxDuration', 'LastDuration', 'LastVsAvgPct', 'LastRun', 'LastOutcome', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText "Database integrity verification history - DBCC CHECKDB (threshold $CheckDbMaxAgeDays days)" -CanCollapse {
            New-NodeAwareTable -Data $CheckDbStatus -Columns @('SqlNode', 'DatabaseName', 'State', 'LastKnownGood', 'Age', 'Status') -PagingLength 15
        }
        New-HTMLSection -Direction column -HeaderText 'DBCC CHECKDB evidence from the SQL Server error log' -CanCollapse {
            New-NodeAwareTable -Data $SqlData.CheckDbLog -Columns @('SqlNode', 'LogDate', 'Status', 'Message') -EmptyMessage 'No DBCC CHECKDB entries in the current or previous error log.'
        }
        New-HTMLSection -Direction column -HeaderText "Long-running requests (> $LongRunningQuerySeconds s)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.LongRunning -Columns @('SqlNode', 'SessionId', 'LoginName', 'HostName', 'ProgramName', 'DatabaseName', 'Command', 'RequestStatus', 'StartTime', 'ElapsedSec', 'CpuMs', 'LogicalReads', 'WaitType', 'BlockedBy', 'SqlText') -EmptyMessage 'No long-running requests at collection time.'
        }
        New-HTMLSection -Direction column -HeaderText "Blocking (waits > $BlockingWaitSeconds s)" -CanCollapse {
            New-NodeAwareTable -Data $SqlData.Blocking -Columns @('SqlNode', 'BlockedSession', 'BlockingSession', 'DatabaseName', 'WaitType', 'WaitSec', 'WaitResource', 'BlockedLogin', 'BlockingLogin', 'BlockingHost', 'BlockingProgram', 'BlockedSql', 'BlockingSql') -EmptyMessage 'No blocking at collection time.'
        }
        New-HTMLSection -Direction column -HeaderText 'Collection errors and permission issues' -CanCollapse {
            New-SafeHTMLTable -Data $script:CollectionErrors -PagingLength 15 -EmptyMessage 'All sections were collected successfully.'
        }
        if ($EnableMetricHistory) {
            New-HTMLSection -Direction column -HeaderText "Monitoring history (retention $MetricHistoryRetentionDays days)" -CanCollapse {
                New-SafeHTMLTable -Data ($HistoryRows | Sort-Object Timestamp -Descending) -PagingLength 25
            }
        }
    }
}

Write-Host ''
Write-Host "Dashboard written to: $ReportPath" -ForegroundColor Green
Write-Host ("Overall status: {0} - {1} critical, {2} warning findings, {3} collection error(s)." -f $overall, $critTotal, $warnTotal, $script:CollectionErrors.Count)
if ($EnableMetricHistory) { Write-Host "Metric history: $(Join-Path $OutputFolder 'History\MetricHistory.csv')" }

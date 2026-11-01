<#
.SYNOPSIS
    MigrationReadiness - read-only readiness check for SQL Server migrations and HA/DR partners.

.DESCRIPTION
    Compares a REFERENCE server (the current production server, or the AG primary) with one or
    more TARGETS (the new build, or the AG secondaries / DR replicas) across every facet that
    usually breaks a cut-over or a failover:

      Platform & OS          OS build, domain, time zone, locale, power plan, cluster membership
      Hardware & disk layout cores, RAM, every volume (letter/mount, size, 64K AU), free space per drive for
                             the databases (honours -PathMap), tempdb file count/size/growth/location,
                             master/model/msdb file locations
      Installed software     VC++ runtimes, Java, client tools, backup/AV/monitoring agents, ...
      Drivers & providers    ODBC drivers, system DSNs, OLE DB providers (x64/x86) + provider options
                             (AllowInProcess ...), SQL client aliases, .NET 3.5 / 4.x, SCHANNEL TLS settings
      Windows features & services   Failover Clustering, .NET 3.5, MPIO, SSIS/SSRS/SSAS/Full-Text services,
                             start modes, service accounts, third-party services, Windows Firewall inbound
                             rules for the SQL port / HADR endpoint / listener / Browser
      SQL instance           version/CU, edition (incl. persisted Enterprise features), collation, HADR,
                             FILESTREAM, Full-Text, PolyBase, IFI, locked pages, TCP port, force encryption
      Certificates           SQL TLS certificate (present in LocalMachine\My, expiry, SAN covers host/FQDN/
                             listener), machine certificates with private keys, master certs (TDE/backup)
      Configuration          sp_configure parity, startup trace flags (obsolete ones recognised)
      CLR & assemblies       user assemblies, permission sets, clr enabled / clr strict security,
                             sys.trusted_assemblies, TRUSTWORTHY reset on restore
      Linked servers         definition, options (rpc out, data access, collation, DTC, timeouts), provider
                             registered + provider options on the target, stored logins, code references
      Agent & automation     jobs, job step definitions (signature), job SCHEDULES, categories, proxies +
                             subsystems + principals, operators, alerts, fail-safe operator, Agent mail
                             profile, job history settings, Database Mail profiles / ACCOUNTS (relay, port,
                             SSL, auth) / public default profile / settings, maintenance plans, and every
                             file/exe/script/PowerShell module the job steps reference (checked on target)
      SSIS / SSRS / SSAS     SSISDB projects/environments, msdb packages, services
      Security objects       credentials, TDE certificates (restore blocker), EKM providers, audits,
                             server triggers, endpoints (HADR port/encryption)
      Server objects         startup procedures, custom extended procs + DLLs, user objects in master/msdb,
                             custom messages, XE sessions, Resource Governor, backup devices
      Databases              compat level, owners, CDC/replication/broker notes, cross-database
                             dependencies, code using xp_cmdshell / OLE automation / ad hoc queries
      Files & paths          database folders, backup/audit/XE paths, job-step binaries on the target
      Source health          last full/log backup, last good CHECKDB, suspect pages, DMKs encrypted by the SMK,
                             backup-encryption certificates, log shipping, mirroring, distributor role
      Host policy            OS + SQL support lifecycle, pending reboot, patch level (UBR / last update), FIPS,
                             MSDTC network access, Defender exclusions, machine env vars + PATH, scheduled
                             tasks, local Administrators, disk health, NTFS compression/EFS on SQL folders
      Identity & DNS         Kerberos SPNs (read-only LDAP search from the runner), DNS forward/reverse
      Hidden dependencies    hard-coded server names/IPs in job steps and in code, jobs already failing,
                             other SQL instances on the host, file versions of referenced binaries
    Every report ends with a coverage map of the team migration checklist (A-K): automated / partial / manual.

    Each target gets a readiness verdict and score, a "needs installation / needs action" list and
    a preparation T-SQL script. Logins are deliberately out of scope - use LoginSync.ps1 for those.

    Everything is in one file:
      * No parameters (or -Gui) opens a WinForms window.
      * With parameters it runs headless and writes: interactive HTML, Word (.docx), plain text
        (Notepad), CSV, JSON and prep_<target>.sql. Exit code for schedulers.

    PRODUCTION GUARDRAILS (not optional)
      * SQL: only static SELECTs against catalog views/DMVs/msdb system tables. A guard refuses
        anything else before it reaches SQL Server. LOCK_TIMEOUT, DEADLOCK_PRIORITY LOW,
        READ UNCOMMITTED, pooling off, one connection per server, closed immediately.
      * Windows: CIM (WMI) queries and the read-only StdRegProv methods (EnumKey/EnumValues/Get*Value)
        only. A guard refuses any other registry method. Nothing is installed, started, stopped or
        changed. No remote code execution (no Invoke-Command), no admin shares.
        Certificates are parsed in memory from the registry copy of LocalMachine\My (no store
        access; private keys are never read).
      * prep_<target>.sql is for review only: aborts unless @@SERVERNAME is the target, every change
        is IF NOT EXISTS / commented out, secrets are placeholders.

.EXAMPLE
    .\MigrationReadiness.ps1                                     # GUI
.EXAMPLE
    .\MigrationReadiness.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01
.EXAMPLE
    .\MigrationReadiness.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01 -PathMap 'D:\=E:\','L:\=F:\'
.EXAMPLE
    .\MigrationReadiness.ps1 -DiscoverAgFrom SQLAGLISTENER01      # primary = reference, real modes
.EXAMPLE
    .\MigrationReadiness.ps1 -ReferenceServer SQL01 -TargetServer SQL02 -TargetMode Synchronous -AsyncTarget SQLDR01
.EXAMPLE
    .\MigrationReadiness.ps1 -Demo                               # sample report, no server needed

.NOTES
    Exit codes: 0 ready / below -FailOnSeverity, 2 findings at or above -FailOnSeverity, 1 run failed.
    Permissions: SQL - VIEW SERVER STATE + VIEW ANY DEFINITION (sysadmin gives the fullest picture,
    msdb/SSISDB reads need db_datareader or sysadmin). Windows - local admin or WMI/DCOM read rights
    (root\cimv2, root\default StdRegProv) on each host, WinRM (WSMan) or DCOM reachable.
#>
[CmdletBinding(DefaultParameterSetName = 'Gui')]
param(
    [Parameter(ParameterSetName = 'Manual', Mandatory = $true)]
    [string]$ReferenceServer,

    [Parameter(ParameterSetName = 'Manual')]
    [string[]]$TargetServer = @(),

    [Parameter(ParameterSetName = 'Manual')]
    [ValidateSet('Migration', 'Synchronous', 'Asynchronous')]
    [string]$TargetMode = 'Migration',

    [Parameter(ParameterSetName = 'Manual')]
    [string[]]$AsyncTarget = @(),

    [Parameter(ParameterSetName = 'Ag', Mandatory = $true)]
    [string]$DiscoverAgFrom,

    [Parameter(ParameterSetName = 'Ag')]
    [string]$AvailabilityGroup,

    [Parameter(ParameterSetName = 'Ag')]
    [switch]$SyncReplicasOnly,

    [Parameter(ParameterSetName = 'Demo', Mandatory = $true)]
    [switch]$Demo,

    [Parameter(ParameterSetName = 'Gui')]
    [switch]$Gui,

    # Windows host to use for OS checks when it differs from the SQL name (listener/alias/port).
    # Format 'SQLNAME=HOST'. By default the physical node reported by SQL Server is used.
    [string[]]$ComputerMap = @(),

    # Drive/folder mapping old -> new for migrations, e.g. 'D:\=E:\','L:\SQLLogs=F:\Logs'
    [string[]]$PathMap = @(),

    [ValidateSet('Sql', 'Os', 'Paths', 'Software', 'CodeScan', 'Spn')]
    [string[]]$Skip = @(),

    [pscredential]$SqlCredential,
    [pscredential]$WindowsCredential,
    [switch]$Encrypt,
    [switch]$TrustServerCertificate,
    [ValidateSet('Wsman', 'Dcom')]
    [string]$CimProtocol = 'Wsman',

    [int]$LockTimeoutMs = 3000,
    [int]$QueryTimeoutSec = 30,
    [int]$MaxDatabases = 200,

    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ('MigrationReadiness_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [ValidateSet('Html', 'Docx', 'Txt', 'Csv', 'Json', 'Sql')]
    [string[]]$Format = @('Html', 'Docx', 'Txt', 'Csv', 'Json', 'Sql'),

    [ValidateSet('Blocker', 'High', 'Medium', 'Low', 'Never')]
    [string]$FailOnSeverity = 'High',

    [switch]$Open
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:ToolName    = 'MigrationReadiness'
$script:ToolVersion = '1.2.0'
$script:AppName     = 'MigrationReadiness (read-only)'
$script:SevOrder    = @{ 'Blocker' = 5; 'High' = 4; 'Medium' = 3; 'Low' = 2; 'Info' = 1; 'OK' = 0 }
$script:SevPenalty  = @{ 'Blocker' = 25; 'High' = 8; 'Medium' = 3; 'Low' = 1; 'Info' = 0 }
$script:Categories  = @(
    'Platform & OS', 'Hardware & storage', 'Installed software', 'Drivers & providers',
    'Windows features & services', 'SQL instance', 'Configuration', 'CLR & assemblies',
    'Linked servers & external', 'Agent & automation', 'SSIS / SSRS / SSAS', 'Security objects',
    'Server objects', 'Databases', 'Files & paths', 'Collection gaps')
$script:LogSink = $null   # GUI replaces this with a scriptblock

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    if ($script:LogSink) { & $script:LogSink $line } else {
        switch ($Level) {
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
            default { Write-Host $line }
        }
    }
}

function New-List { return New-Object System.Collections.ArrayList }
function Add-To { param($List, $Item) [void]$List.Add($Item) }
function ConvertTo-Arr { param($x) if ($null -eq $x) { return @() } return @($x) }
function Get-Prop {
    param($Obj, [string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    if ($Obj -is [System.Collections.IDictionary]) { if ($Obj.Contains($Name)) { return $Obj[$Name] } return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value } return $Default
}
function Format-MB {
    param($Mb)
    if ($null -eq $Mb -or "$Mb" -eq '') { return '-' }
    $v = [double]$Mb
    if ($v -ge 1048576) { return ('{0:N1} TB' -f ($v / 1048576)) }
    if ($v -ge 1024) { return ('{0:N1} GB' -f ($v / 1024)) }
    return ('{0:N0} MB' -f $v)
}
function ConvertTo-VersionSafe {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $m = [regex]::Match($Text, '\d+(\.\d+){0,3}')
    if (-not $m.Success) { return $null }
    $t = $m.Value; if ($t -notmatch '\.') { $t = "$t.0" }
    try { return [version]$t } catch { return $null }
}
function Get-SqlMajorName {
    param([int]$Major)
    switch ($Major) { 9 { '2005' } 10 { '2008/R2' } 11 { '2012' } 12 { '2014' } 13 { '2016' } 14 { '2017' } 15 { '2019' } 16 { '2022' } 17 { '2025' } default { "v$Major" } }
}
function Get-Sum {
    # null/empty-safe sum of a numeric property (Measure-Object returns nothing for empty input)
    param($Items, [string]$Property)
    $t = [double]0
    foreach ($i in (ConvertTo-Arr $Items)) { $v = Get-Prop $i $Property; if ($null -ne $v -and "$v" -ne '') { $t += [double]$v } }
    return $t
}
function Get-WinParent {
    # string-only parent of a Windows path (never touches the local file system)
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    $p = $Path.TrimEnd('\'); $i = $p.LastIndexOf('\')
    if ($i -lt 0) { return '' }
    if ($i -eq 2 -and $p[1] -eq ':') { return $p.Substring(0, 3) }
    return $p.Substring(0, $i)
}
function Get-WinLeaf { param([string]$Path) $p = ([string]$Path).TrimEnd('\'); return $p.Substring($p.LastIndexOf('\') + 1) }
function Protect-Secret {
    # masks password-like fragments in provider strings / commands before they reach a report
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $t = [regex]::Replace($Text, '(?i)(pwd|password|passwd|secret|token|apikey|api_key)\s*[=:]\s*("[^"]*"|''[^'']*''|[^;\s]+)', '$1=****')
    $t = [regex]::Replace($t, '(?i)(-p|/p|-password|/password)\s+("[^"]*"|\S+)', '$1 ****')
    return $t
}

#region ---------------------------------------------------------------- SQL (read-only) access
$script:SqlSessionSetup = 'SET LOCK_TIMEOUT {0}; SET DEADLOCK_PRIORITY LOW; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED; SET NOCOUNT ON;'

function Assert-ReadOnlySql {
    # Refuses anything that is not a single static SELECT/WITH. String literals and [identifiers]
    # are removed before the keyword scan, so a LIKE '%xp_cmdshell%' pattern or a database name
    # can't trip it, and nothing inside them can execute anyway.
    param([string]$Sql)
    $s = $Sql -replace "N?'([^']|'')*'", "''"
    $s = $s -replace '\[([^\]]|\]\])*\]', '[x]'
    if ($s -match '--|/\*') { throw 'Read-only guard: comments are not allowed in collector queries.' }
    $trim = $s.Trim().TrimEnd(';')
    if ($trim -notmatch '^(?is)\s*(SELECT|WITH)\b') { throw 'Read-only guard: query must start with SELECT or WITH.' }
    if ($trim.Contains(';')) { throw 'Read-only guard: multiple statements are not allowed.' }
    $bad = '(?i)\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|TRUNCATE|EXEC|EXECUTE|GRANT|REVOKE|DENY|DBCC|BACKUP|RESTORE|KILL|SHUTDOWN|RECONFIGURE|WAITFOR|OPENROWSET|OPENQUERY|OPENDATASOURCE|OPENXML|INTO|USE|SET|DECLARE|BULK|GO)\b|\bxp_|\bsp_'
    $m = [regex]::Match($trim, $bad)
    if ($m.Success) { throw ("Read-only guard: forbidden token '{0}'." -f $m.Value) }
}

function New-SqlConn {
    param([string]$Server)
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source'] = $Server
    $b['Initial Catalog'] = 'master'
    $b['Application Name'] = $script:AppName
    $b['Pooling'] = $false
    $b['Connect Timeout'] = 15
    if ($Encrypt) { $b['Encrypt'] = $true }
    if ($TrustServerCertificate) { $b['TrustServerCertificate'] = $true }
    if ($SqlCredential) {
        $b['User ID'] = $SqlCredential.UserName
        $b['Password'] = $SqlCredential.GetNetworkCredential().Password
    } else { $b['Integrated Security'] = $true }
    $c = New-Object System.Data.SqlClient.SqlConnection($b.ConnectionString)
    $c.Open()
    $cmd = $c.CreateCommand()
    $cmd.CommandText = ($script:SqlSessionSetup -f [int]$LockTimeoutMs)   # fixed constant, not user input
    $cmd.CommandTimeout = 10
    [void]$cmd.ExecuteNonQuery()
    return $c
}

function Invoke-RoSql {
    param($Conn, [string]$Sql)
    Assert-ReadOnlySql -Sql $Sql
    $cmd = $Conn.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = $QueryTimeoutSec
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    $rows = New-List
    foreach ($r in $dt.Rows) {
        $o = [ordered]@{}
        foreach ($col in $dt.Columns) {
            $v = $r[$col.ColumnName]
            if ($v -is [System.DBNull]) { $v = $null }
            elseif ($v -is [string]) { $v = $v.TrimEnd([char]0) }
            $o[$col.ColumnName] = $v
        }
        Add-To $rows ([pscustomobject]$o)
    }
    return $rows.ToArray()
}

function Invoke-Section {
    # runs one collector section; a failure is recorded, never fatal
    param($Snap, [string]$Area, [string]$Name, [scriptblock]$Body)
    try {
        $r = & $Body
        $Snap[$Area].Sections[$Name] = 'ok'
        if ($null -ne $r) { return $r }
        return
    } catch {
        $msg = $_.Exception.Message -replace '\s+', ' '
        if ($msg.Length -gt 220) { $msg = $msg.Substring(0, 220) + '...' }
        $Snap[$Area].Sections[$Name] = "error: $msg"
        Write-Log ("  {0}/{1}: {2}" -f $Area, $Name, $msg) 'WARN'
        return
    }
}

function Get-Q { param([string]$Db, [string]$Tail) return ('[' + $Db.Replace(']', ']]') + '].' + $Tail) }

function Get-PathRefs {
    # extracts local absolute paths / UNC paths and module names from a job-step command
    param([string]$Text)
    $paths = New-List; $uncs = New-List; $mods = New-List
    if ([string]::IsNullOrWhiteSpace($Text)) { return @{ Paths = @(); Unc = @(); Modules = @() } }
    $ext = 'exe|bat|cmd|ps1|psm1|psd1|vbs|js|wsf|dtsx|ispac|py|jar|dll|sql|config|xml|json|ini|txt|csv|fmt|xsd|rdl'
    foreach ($m in [regex]::Matches($Text, "(?i)(?<![\\\w])([a-z]:\\(?:[^""'\r\n|<>*?:;]+?\\)*[^""'\r\n|<>*?:;\\]+?\.($ext))(?=$|[\s""',;)])")) {
        if (-not $paths.Contains($m.Groups[1].Value)) { Add-To $paths $m.Groups[1].Value }
    }
    foreach ($m in [regex]::Matches($Text, "(?i)(\\\\[a-z0-9_.$-]+\\[^""'\r\n|<>*?:;]+)")) {
        $u = $m.Groups[1].Value.Trim().TrimEnd('\')
        if (-not $uncs.Contains($u)) { Add-To $uncs $u }
    }
    foreach ($m in [regex]::Matches($Text, '(?i)(?:Import-Module|#Requires\s+-Modules?|using\s+module)\s+[''"]?([A-Za-z][\w.\-]+)')) {
        if (-not $mods.Contains($m.Groups[1].Value)) { Add-To $mods $m.Groups[1].Value }
    }
    if ($Text -match '(?i)\b(Invoke-Sqlcmd|Get-SqlInstance|Backup-SqlDatabase|Restore-SqlDatabase)\b' -and -not $mods.Contains('SqlServer')) { Add-To $mods 'SqlServer' }
    if ($Text -match '(?i)\b\w+-Dba\w+\b' -and -not $mods.Contains('dbatools')) { Add-To $mods 'dbatools' }
    return @{ Paths = $paths.ToArray(); Unc = $uncs.ToArray(); Modules = $mods.ToArray() }
}

function Get-ScheduleText {
    # human readable msdb schedule (sysschedules semantics)
    param($x)
    $t = { param($v) $v = [int]$v; '{0:00}:{1:00}' -f [int][Math]::Floor($v / 10000), [int]([Math]::Floor($v / 100) % 100) }
    $days = { param($m) $n = @(); $map = @('Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'); for ($i = 0; $i -lt 7; $i++) { if ([int]$m -band (1 -shl $i)) { $n += $map[$i] } }; $n -join ',' }
    $f = switch ([int]$x.FreqType) {
        1 { 'once' } 4 { if ([int]$x.FreqInterval -gt 1) { "every $($x.FreqInterval) days" } else { 'daily' } }
        8 { "weekly ($(& $days $x.FreqInterval))" + $(if ([int]$x.Recurrence -gt 1) { " every $($x.Recurrence) weeks" } else { '' }) }
        16 { "monthly day $($x.FreqInterval)" + $(if ([int]$x.Recurrence -gt 1) { " every $($x.Recurrence) months" } else { '' }) }
        32 { "monthly relative ($($x.RelInterval)/$($x.FreqInterval))" } 64 { 'when Agent starts' } 128 { 'when CPU idle' } default { "type $($x.FreqType)" }
    }
    $sub = switch ([int]$x.SubdayType) { 2 { " every $($x.SubdayInterval) sec" } 4 { " every $($x.SubdayInterval) min" } 8 { " every $($x.SubdayInterval) h" } default { '' } }
    $win = if ([int]$x.FreqType -in @(64, 128)) { '' } elseif ($sub) { " $(& $t $x.StartTime)-$(& $t $x.EndTime)" } else { " at $(& $t $x.StartTime)" }
    $end = if ([int]$x.EndDate -lt 99991231) { " until $($x.EndDate)" } else { '' }
    return ($f + $sub + $win + $end)
}

function Get-SqlSnapshot {
    param($Snap)
    $server = $Snap.Name
    $S = $Snap.Sql
    $conn = $null
    try {
        $conn = New-SqlConn -Server $server
    } catch {
        $S.Sections['connect'] = 'error: ' + ($_.Exception.Message -replace '\s+', ' ')
        Write-Log "  SQL connect to $server failed: $($_.Exception.Message)" 'ERROR'
        return
    }
    $Snap.SqlOk = $true
    try {
        # ---- instance
        $inst = Invoke-Section $Snap 'Sql' 'instance' {
            $q = @"
SELECT CAST(SERVERPROPERTY('ServerName') AS nvarchar(256)) AS ServerName,
 CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS nvarchar(256)) AS PhysicalNode,
 CAST(SERVERPROPERTY('MachineName') AS nvarchar(256)) AS MachineName,
 CAST(SERVERPROPERTY('InstanceName') AS nvarchar(256)) AS InstanceName,
 CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(64)) AS ProductVersion,
 CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(64)) AS ProductLevel,
 CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(64)) AS ProductUpdateLevel,
 CAST(SERVERPROPERTY('Edition') AS nvarchar(128)) AS Edition,
 CAST(SERVERPROPERTY('EngineEdition') AS int) AS EngineEdition,
 CAST(SERVERPROPERTY('Collation') AS nvarchar(128)) AS Collation,
 CAST(SERVERPROPERTY('IsHadrEnabled') AS int) AS IsHadrEnabled,
 CAST(SERVERPROPERTY('IsClustered') AS int) AS IsClustered,
 CAST(SERVERPROPERTY('IsFullTextInstalled') AS int) AS IsFullTextInstalled,
 CAST(SERVERPROPERTY('IsPolyBaseInstalled') AS int) AS IsPolyBaseInstalled,
 CAST(SERVERPROPERTY('IsIntegratedSecurityOnly') AS int) AS IsIntegratedSecurityOnly,
 CAST(SERVERPROPERTY('FilestreamConfiguredLevel') AS int) AS FilestreamConfiguredLevel,
 CAST(SERVERPROPERTY('InstanceDefaultDataPath') AS nvarchar(512)) AS DefaultDataPath,
 CAST(SERVERPROPERTY('InstanceDefaultLogPath') AS nvarchar(512)) AS DefaultLogPath,
 CAST(SERVERPROPERTY('InstanceDefaultBackupPath') AS nvarchar(512)) AS DefaultBackupPath,
 CAST(SERVERPROPERTY('ResourceVersion') AS nvarchar(64)) AS ResourceVersion
"@
            @(Invoke-RoSql $conn $q)[0]
        }
        if ($inst) {
            foreach ($p in $inst.PSObject.Properties) { $S.Instance[$p.Name] = $p.Value }
            $v = ConvertTo-VersionSafe $inst.ProductVersion
            $S.Instance['Major'] = if ($v) { $v.Major } else { 0 }
            $S.Instance['Build'] = if ($v) { $v.ToString() } else { '' }
            $S.Instance['VersionName'] = Get-SqlMajorName $S.Instance['Major']
        }
        $major = [int](Get-Prop $S.Instance 'Major' 0)
        $own = @(@((Get-Prop $S.Instance 'ServerName'), (Get-Prop $S.Instance 'PhysicalNode'), (Get-Prop $S.Instance 'MachineName')) | Where-Object { $_ } | ForEach-Object { ([string]$_ -split '\\')[0] } | Sort-Object -Unique)
        $S.Instance['OwnNames'] = $own

        Invoke-Section $Snap 'Sql' 'sysinfo' {
            if ($major -ge 13) {
                $q = 'SELECT cpu_count AS CpuCount, physical_memory_kb/1024 AS MemoryMB, sql_memory_model_desc AS MemoryModel, virtual_machine_type_desc AS VmType FROM sys.dm_os_sys_info'
            } elseif ($major -ge 11) {
                $q = 'SELECT cpu_count AS CpuCount, physical_memory_kb/1024 AS MemoryMB, NULL AS MemoryModel, virtual_machine_type_desc AS VmType FROM sys.dm_os_sys_info'
            } else {
                $q = 'SELECT cpu_count AS CpuCount, physical_memory_in_bytes/1048576 AS MemoryMB, NULL AS MemoryModel, NULL AS VmType FROM sys.dm_os_sys_info'
            }
            $r = @(Invoke-RoSql $conn $q)[0]
            foreach ($p in $r.PSObject.Properties) { $S.Instance[$p.Name] = $p.Value }
        } | Out-Null

        $S.Services = @(Invoke-Section $Snap 'Sql' 'services' {
            try {
                Invoke-RoSql $conn 'SELECT servicename AS ServiceName, service_account AS Account, startup_type_desc AS StartMode, status_desc AS Status, CAST(instant_file_initialization_enabled AS nvarchar(5)) AS Ifi, filename AS Binary FROM sys.dm_server_services'
            } catch {
                Invoke-RoSql $conn 'SELECT servicename AS ServiceName, service_account AS Account, startup_type_desc AS StartMode, status_desc AS Status, NULL AS Ifi, filename AS Binary FROM sys.dm_server_services'
            }
        })
        $eng = $S.Services | Where-Object { $_.ServiceName -like 'SQL Server (*' } | Select-Object -First 1
        if ($eng) {
            $S.Instance['Ifi'] = $eng.Ifi
            $S.Instance['ServiceAccount'] = $eng.Account
            $bin = [regex]::Match([string]$eng.Binary, '(?i)"?([a-z]:\\[^"]+\\)sqlservr\.exe')
            if ($bin.Success) { $S.Instance['BinnPath'] = $bin.Groups[1].Value }
        }

        # ---- configuration + startup parameters + network
        $cfg = Invoke-Section $Snap 'Sql' 'config' { Invoke-RoSql $conn 'SELECT name, CAST(value_in_use AS nvarchar(64)) AS value, CAST(value AS nvarchar(64)) AS configured FROM sys.configurations' }
        foreach ($r in (ConvertTo-Arr $cfg)) { $S.Config[[string](Get-Prop $r 'name')] = Get-Prop $r 'value'; $cv = Get-Prop $r 'configured'; if ($null -ne $cv -and [string]$cv -ne [string](Get-Prop $r 'value')) { $S.ConfigPending[[string](Get-Prop $r 'name')] = [string]$cv } }

        $reg = Invoke-Section $Snap 'Sql' 'registry' {
            Invoke-RoSql $conn "SELECT registry_key, value_name, CAST(value_data AS nvarchar(1024)) AS value_data FROM sys.dm_server_registry WHERE value_name LIKE N'SQLArg%' OR value_name IN (N'TcpPort', N'TcpDynamicPorts', N'ForceEncryption', N'Certificate', N'BackupDirectory', N'Enabled', N'HideInstance', N'NumErrorLogs') OR registry_key LIKE N'%\SQLServerAgent'"
        }
        $flags = New-List
        foreach ($r in (ConvertTo-Arr $reg)) {
            $k = [string]$r.registry_key; $n = [string]$r.value_name; $d = ([string]$r.value_data).Trim()
            if (-not (Get-Prop $S.Instance 'RegRoot') -and $k -match '^(?i)HKLM\\(Software\\Microsoft\\Microsoft SQL Server\\MSSQL[^\\]+)') { $S.Instance['RegRoot'] = $Matches[1] }
            if ($k -match '(?i)\\Microsoft SQL Server\\[^\\]+\\SQLServerAgent$') { $S.AgentProps[$n] = $d; continue }
            if ($n -like 'SQLArg*') {
                if ($d -match '^-T(\d+)') { Add-To $flags ([int]$Matches[1]) }
                elseif ($d -match '^-e(.+)') { $S.Instance['ErrorLogPath'] = $Matches[1] }
                elseif ($d -match '^-d(.+)') { $S.Instance['MasterData'] = $Matches[1] }
            }
            elseif ($k -match '\\Tcp\\IPAll$' -and $n -eq 'TcpPort') { $S.Net['TcpPort'] = $d }
            elseif ($k -match '\\Tcp\\IPAll$' -and $n -eq 'TcpDynamicPorts') { $S.Net['TcpDynamicPorts'] = $d }
            elseif ($k -match '\\SuperSocketNetLib\\Tcp$' -and $n -eq 'Enabled') { $S.Net['TcpEnabled'] = $d }
            elseif ($k -match '\\SuperSocketNetLib\\Np$' -and $n -eq 'Enabled') { $S.Net['NamedPipes'] = $d }
            elseif ($k -match '\\SuperSocketNetLib$' -and $n -eq 'ForceEncryption') { $S.Net['ForceEncryption'] = $d }
            elseif ($k -match '\\SuperSocketNetLib$' -and $n -eq 'Certificate') { $S.Net['Certificate'] = $d }
            elseif ($k -match '\\SuperSocketNetLib$' -and $n -eq 'HideInstance') { $S.Net['HideInstance'] = $d }
            elseif ($n -eq 'NumErrorLogs') { $S.Instance['NumErrorLogs'] = $d }
            elseif ($n -eq 'BackupDirectory' -and -not (Get-Prop $S.Instance 'DefaultBackupPath')) { $S.Instance['DefaultBackupPath'] = $d }
        }
        $S.TraceFlags = @($flags | Sort-Object -Unique)

        # ---- databases + files
        $dbq = if ($major -ge 11) {
            "SELECT d.name AS Name, d.database_id AS Id, d.compatibility_level AS Compat, d.collation_name AS Collation, d.recovery_model_desc AS Recovery, d.state_desc AS State, CAST(d.is_read_only AS int) AS ReadOnly, CAST(d.is_trustworthy_on AS int) AS Trustworthy, CAST(d.is_db_chaining_on AS int) AS DbChaining, CAST(d.is_encrypted AS int) AS Encrypted, CAST(d.is_broker_enabled AS int) AS Broker, CAST(d.is_cdc_enabled AS int) AS Cdc, CAST(d.is_published AS int) + CAST(d.is_merge_published AS int) AS Published, CAST(d.is_subscribed AS int) AS Subscribed, CAST(d.is_distributor AS int) AS Distributor, d.containment_desc AS Containment, SUSER_SNAME(d.owner_sid) AS Owner, CAST(CASE WHEN d.replica_id IS NULL THEN 0 ELSE 1 END AS int) AS InAg, (SELECT CAST(SUM(CAST(mf.size AS bigint)) * 8 / 1024 AS bigint) FROM sys.master_files mf WHERE mf.database_id = d.database_id) AS SizeMB FROM sys.databases d"
        } else {
            "SELECT d.name AS Name, d.database_id AS Id, d.compatibility_level AS Compat, d.collation_name AS Collation, d.recovery_model_desc AS Recovery, d.state_desc AS State, CAST(d.is_read_only AS int) AS ReadOnly, CAST(d.is_trustworthy_on AS int) AS Trustworthy, CAST(d.is_db_chaining_on AS int) AS DbChaining, CAST(d.is_encrypted AS int) AS Encrypted, CAST(d.is_broker_enabled AS int) AS Broker, CAST(d.is_cdc_enabled AS int) AS Cdc, CAST(d.is_published AS int) + CAST(d.is_merge_published AS int) AS Published, CAST(d.is_subscribed AS int) AS Subscribed, CAST(d.is_distributor AS int) AS Distributor, 'NONE' AS Containment, SUSER_SNAME(d.owner_sid) AS Owner, 0 AS InAg, (SELECT CAST(SUM(CAST(mf.size AS bigint)) * 8 / 1024 AS bigint) FROM sys.master_files mf WHERE mf.database_id = d.database_id) AS SizeMB FROM sys.databases d"
        }
        $S.Databases = @(Invoke-Section $Snap 'Sql' 'databases' { Invoke-RoSql $conn $dbq })
        $S.DbHealth = @(Invoke-Section $Snap 'Sql' 'db_health' {
            $chk = if ($major -ge 13) { "CAST(DATABASEPROPERTYEX(d.name, 'LastGoodCheckDbTime') AS datetime)" } else { 'CAST(NULL AS datetime)' }
            Invoke-RoSql $conn ("SELECT d.name AS Db, CAST(d.is_master_key_encrypted_by_server AS int) AS DmkBySmk, d.recovery_model_desc AS Recovery, (SELECT MAX(b.backup_finish_date) FROM msdb.dbo.backupset b WHERE b.database_name = d.name AND b.type = 'D') AS LastFull, (SELECT MAX(b.backup_finish_date) FROM msdb.dbo.backupset b WHERE b.database_name = d.name AND b.type = 'L') AS LastLog, {0} AS LastCheckDb, (SELECT COUNT(*) FROM msdb.dbo.suspect_pages sp WHERE sp.database_id = d.database_id) AS SuspectPages, (SELECT COUNT(*) FROM sys.database_mirroring m WHERE m.database_id = d.database_id AND m.mirroring_guid IS NOT NULL) AS Mirrored FROM sys.databases d WHERE d.database_id > 4 AND d.source_database_id IS NULL" -f $chk)
        })
        Invoke-Section $Snap 'Sql' 'server_misc' {
            $m = @(Invoke-RoSql $conn "SELECT (SELECT COUNT(*) FROM master.sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##') AS MasterDmk, (SELECT COUNT(*) FROM msdb.dbo.syspolicy_policies WHERE is_system = 0) AS Policies, (SELECT COUNT(*) FROM msdb.dbo.sysmanagement_shared_registered_servers) AS CmsServers, (SELECT COUNT(*) FROM sys.dm_os_nodes WHERE node_state_desc NOT LIKE N'%DAC%') AS NumaNodes")[0]
            if ($m) { foreach ($p in $m.PSObject.Properties) { $S.Instance[$p.Name] = $p.Value } }
        } | Out-Null
        $S.LogShipping = @(Invoke-Section $Snap 'Sql' 'log_shipping' { Invoke-RoSql $conn "SELECT N'primary' AS Role, primary_database AS Db FROM msdb.dbo.log_shipping_primary_databases UNION ALL SELECT N'secondary', secondary_database FROM msdb.dbo.log_shipping_secondary_databases" })
        if ($major -ge 12) {
            $S.BackupCerts = @(Invoke-Section $Snap 'Sql' 'backup_encryption' { Invoke-RoSql $conn 'SELECT CONVERT(varchar(100), b.encryptor_thumbprint, 2) AS Thumbprint, MAX(b.backup_finish_date) AS LastUsed, COUNT(*) AS Backups FROM msdb.dbo.backupset b WHERE b.encryptor_thumbprint IS NOT NULL AND b.backup_finish_date > DATEADD(day, -60, GETDATE()) GROUP BY b.encryptor_thumbprint' })
        }
        $S.Files = @(Invoke-Section $Snap 'Sql' 'files' {
            Invoke-RoSql $conn 'SELECT DB_NAME(database_id) AS Db, database_id AS DbId, type_desc AS Type, name AS LogicalName, physical_name AS Path, CAST(size AS bigint) * 8 / 1024 AS SizeMB, CASE WHEN is_percent_growth = 1 THEN CAST(growth AS varchar(20)) + ''%'' WHEN growth = 0 THEN ''none'' ELSE CAST(CAST(growth AS bigint) * 8 / 1024 AS varchar(20)) + '' MB'' END AS Growth, max_size AS MaxSize FROM sys.master_files'
        })

        # ---- per-database scan (user databases, ONLINE only, capped)
        $userDbs = @($S.Databases | Where-Object { $_.Id -gt 4 -and $_.State -eq 'ONLINE' -and $_.Name -ne 'SSISDB' -and $_.Distributor -eq 0 } | Sort-Object Name | Select-Object -First $MaxDatabases)
        if (@($S.Databases | Where-Object { $_.Id -gt 4 }).Count -gt $MaxDatabases) { $S.Sections['dbscan-cap'] = "error: only the first $MaxDatabases user databases were scanned (-MaxDatabases)" }
        $selfLike = if ($own.Count) { (@($own | ForEach-Object { "definition LIKE N'%" + (([string]$_).Replace("'", "''") -replace '([\[\]%_])', '[$1]') + "%'" }) -join ' OR ') } else { '1 = 0' }
        $asm = New-List; $feat = New-List; $lref = New-List; $xref = New-List; $code = New-List; $sku = New-List
        foreach ($d in $userDbs) {
            $db = [string]$d.Name
            try {
                $r = @(Invoke-RoSql $conn ("SELECT (SELECT COUNT(*) FROM {0}) AS FtCatalogs, (SELECT COUNT(*) FROM {1} WHERE type = 'FD') AS FsFilegroups, (SELECT COUNT(*) FROM {1} WHERE type = 'FX') AS XtpFilegroups, (SELECT COUNT(*) FROM {2} WHERE is_ms_shipped = 0) AS BrokerQueues, (SELECT COUNT(*) FROM {3} WHERE is_user_defined = 1) AS Assemblies" -f (Get-Q $db 'sys.fulltext_catalogs'), (Get-Q $db 'sys.filegroups'), (Get-Q $db 'sys.service_queues'), (Get-Q $db 'sys.assemblies')))[0]
                Add-To $feat ([pscustomobject]@{ Db = $db; FtCatalogs = $r.FtCatalogs; FsFilegroups = $r.FsFilegroups; XtpFilegroups = $r.XtpFilegroups; BrokerQueues = $r.BrokerQueues; Assemblies = $r.Assemblies })
                if ($r.Assemblies -gt 0) {
                    foreach ($a in @(Invoke-RoSql $conn ("SELECT name AS Name, clr_name AS ClrName, permission_set_desc AS PermissionSet, CONVERT(varchar(20), create_date, 23) AS Created FROM {0} WHERE is_user_defined = 1" -f (Get-Q $db 'sys.assemblies')))) {
                        Add-To $asm ([pscustomobject]@{ Db = $db; Name = $a.Name; ClrName = $a.ClrName; PermissionSet = $a.PermissionSet; Created = $a.Created; Trustworthy = $d.Trustworthy })
                    }
                }
            } catch { $S.Sections["db:$db"] = 'error: ' + ($_.Exception.Message -replace '\s+', ' ') }
            try {
                foreach ($f in @(Invoke-RoSql $conn ('SELECT feature_name AS Feature FROM ' + (Get-Q $db 'sys.dm_db_persisted_sku_features')))) {
                    Add-To $sku ([pscustomobject]@{ Db = $db; Feature = $f.Feature })
                }
            } catch { }
            try {
                foreach ($x in @(Invoke-RoSql $conn ('SELECT DISTINCT referenced_server_name AS RefServer, referenced_database_name AS RefDb FROM ' + (Get-Q $db 'sys.sql_expression_dependencies') + ' WHERE referenced_server_name IS NOT NULL OR (referenced_database_name IS NOT NULL AND referenced_database_name <> DB_NAME())'))) {
                    if ($x.RefServer) { Add-To $lref ([pscustomobject]@{ Db = $db; LinkedServer = $x.RefServer; RefDb = $x.RefDb }) }
                    elseif ($x.RefDb -and $x.RefDb -ne $db) { Add-To $xref ([pscustomobject]@{ Db = $db; RefDb = $x.RefDb }) }
                }
            } catch { $S.Sections["deps:$db"] = 'error: ' + ($_.Exception.Message -replace '\s+', ' ') }
            if ($Skip -notcontains 'CodeScan') {
                try {
                    $c = @(Invoke-RoSql $conn ("SELECT SUM(CASE WHEN definition LIKE N'%xp[_]cmdshell%' THEN 1 ELSE 0 END) AS XpCmdshell, SUM(CASE WHEN definition LIKE N'%sp[_]OACreate%' THEN 1 ELSE 0 END) AS OleAutomation, SUM(CASE WHEN definition LIKE N'%OPENROWSET%' OR definition LIKE N'%OPENDATASOURCE%' THEN 1 ELSE 0 END) AS AdHocQueries, SUM(CASE WHEN definition LIKE N'%sp[_]send[_]dbmail%' THEN 1 ELSE 0 END) AS DbMail, SUM(CASE WHEN definition LIKE N'%BULK INSERT%' THEN 1 ELSE 0 END) AS BulkInsert, SUM(CASE WHEN definition LIKE N'%BEGIN DISTRIBUTED%' THEN 1 ELSE 0 END) AS DistTx, SUM(CASE WHEN {1} THEN 1 ELSE 0 END) AS SelfRef FROM {0}" -f (Get-Q $db 'sys.sql_modules'), $selfLike))[0]
                    Add-To $code ([pscustomobject]@{ Db = $db; XpCmdshell = [int]("0$($c.XpCmdshell)"); OleAutomation = [int]("0$($c.OleAutomation)"); AdHocQueries = [int]("0$($c.AdHocQueries)"); DbMail = [int]("0$($c.DbMail)"); BulkInsert = [int]("0$($c.BulkInsert)"); DistTx = [int]("0$($c.DistTx)"); SelfRef = [int]("0$($c.SelfRef)") })
                } catch { }
            }
        }
        $S.DbFeatures = $feat.ToArray(); $S.Assemblies = $asm.ToArray(); $S.SkuFeatures = $sku.ToArray()
        $S.LinkedRefs = $lref.ToArray(); $S.CrossDbRefs = $xref.ToArray(); $S.CodeUse = $code.ToArray()
        $S.Sections['dbscan'] = "ok ($($userDbs.Count) databases)"

        if ($major -ge 14) {
            $S.TrustedAssemblies = @(Invoke-Section $Snap 'Sql' 'trusted_assemblies' { Invoke-RoSql $conn 'SELECT CONVERT(varchar(200), hash, 2) AS Hash, description AS Description FROM sys.trusted_assemblies' })
        }

        # ---- linked servers / credentials
        $S.LinkedServers = @(Invoke-Section $Snap 'Sql' 'linked_servers' {
            $rows = Invoke-RoSql $conn 'SELECT s.name AS Name, s.product AS Product, s.provider AS Provider, s.data_source AS DataSource, s.provider_string AS ProviderString, s.catalog AS Catalog, CAST(s.is_rpc_out_enabled AS int) AS RpcOut, CAST(s.is_data_access_enabled AS int) AS DataAccess, CAST(s.is_remote_login_enabled AS int) AS Rpc, CAST(s.is_collation_compatible AS int) AS CollationCompatible, CAST(s.uses_remote_collation AS int) AS RemoteCollation, s.collation_name AS RemoteCollationName, CAST(s.lazy_schema_validation AS int) AS LazySchema, CAST(s.is_remote_proc_transaction_promotion_enabled AS int) AS DtcPromotion, s.connect_timeout AS ConnectTimeout, s.query_timeout AS QueryTimeout, (SELECT COUNT(*) FROM sys.linked_logins ll WHERE ll.server_id = s.server_id AND ll.uses_self_credential = 0 AND ll.remote_name IS NOT NULL) AS StoredLogins FROM sys.servers s WHERE s.is_linked = 1'
            foreach ($r in $rows) { $r.ProviderString = Protect-Secret $r.ProviderString }
            $rows
        })
        $S.Credentials = @(Invoke-Section $Snap 'Sql' 'credentials' { Invoke-RoSql $conn 'SELECT name AS Name, credential_identity AS CredIdentity FROM sys.credentials' })
        $S.Logins = @(Invoke-Section $Snap 'Sql' 'login_names' { @(Invoke-RoSql $conn "SELECT name AS Name FROM sys.server_principals WHERE type IN ('S','U','G') AND name NOT LIKE N'##%'") | ForEach-Object { $_.Name } })

        # ---- Agent / msdb
        $S.Jobs = @(Invoke-Section $Snap 'Sql' 'jobs' { Invoke-RoSql $conn 'SELECT j.name AS Name, CAST(j.enabled AS int) AS Enabled, SUSER_SNAME(j.owner_sid) AS Owner, c.name AS Category, (SELECT COUNT(*) FROM msdb.dbo.sysjobsteps s WHERE s.job_id = j.job_id) AS Steps, js.last_run_outcome AS LastOutcome, js.last_run_date AS LastRunDate FROM msdb.dbo.sysjobs j LEFT JOIN msdb.dbo.syscategories c ON c.category_id = j.category_id LEFT JOIN msdb.dbo.sysjobservers js ON js.job_id = j.job_id' })
        $steps = Invoke-Section $Snap 'Sql' 'job_steps' { Invoke-RoSql $conn 'SELECT j.name AS Job, s.step_id AS StepId, s.step_name AS Step, s.subsystem AS Subsystem, s.command AS Command, s.output_file_name AS OutputFile, s.database_name AS StepDb, s.on_success_action AS OnSuccess, s.on_fail_action AS OnFail, p.name AS Proxy FROM msdb.dbo.sysjobsteps s JOIN msdb.dbo.sysjobs j ON j.job_id = s.job_id LEFT JOIN msdb.dbo.sysproxies p ON p.proxy_id = s.proxy_id' }
        $refs = New-List
        foreach ($st in (ConvertTo-Arr $steps)) {
            $pr = Get-PathRefs -Text ([string]$st.Command)
            $outDir = $null
            if ($st.OutputFile -and ([string]$st.OutputFile) -match '^(?i)([a-z]:\\.*)\\[^\\]+$') { $outDir = $Matches[1] }
            $catalog = $null
            if ($st.Subsystem -eq 'SSIS' -and ([string]$st.Command) -match '(?i)/ISSERVER\s+"?\\?"?\\SSISDB\\([^\\]+)\\([^\\]+)\\') { $catalog = "$($Matches[1])\$($Matches[2])" }
            $selfHits = @($own | Where-Object { ([string]$st.Command) -match ('(?i)(?<![\w-])' + [regex]::Escape([string]$_) + '(?![\w-])') })
            $ipHits = @([regex]::Matches([string]$st.Command, '\b(?:\d{1,3}\.){3}\d{1,3}\b') | ForEach-Object { $_.Value } | Where-Object { $_ -notmatch '^(127\.|0\.)' } | Sort-Object -Unique)
            Add-To $refs ([pscustomobject]@{ Job = $st.Job; StepId = $st.StepId; Step = $st.Step; Subsystem = $st.Subsystem; Proxy = $st.Proxy; Paths = $pr.Paths; Unc = $pr.Unc; Modules = $pr.Modules; OutputDir = $outDir; SsisProject = $catalog; SelfRefs = $selfHits; IpRefs = $ipHits })
        }
        $S.JobSteps = $refs.ToArray()
        Invoke-Section $Snap 'Sql' 'job_signatures' {
            $sig = @{}
            foreach ($st in (ConvertTo-Arr $steps | Sort-Object Job, StepId)) {
                if (-not $sig.ContainsKey([string]$st.Job)) { $sig[[string]$st.Job] = New-Object System.Text.StringBuilder }
                [void]$sig[[string]$st.Job].Append(("{0}|{1}|{2}|{3}|{4}|{5}|{6}`n" -f (Get-Prop $st 'StepId'), (Get-Prop $st 'Subsystem'), (Get-Prop $st 'StepDb'), (Get-Prop $st 'Proxy'), (Get-Prop $st 'OnSuccess'), (Get-Prop $st 'OnFail'), ([string](Get-Prop $st 'Command') -replace '\s+', ' ').Trim()))
            }
            $sha = [System.Security.Cryptography.SHA1]::Create()
            foreach ($j in $sig.Keys) { $S.JobSignatures[$j] = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($sig[$j].ToString())))).Replace('-', '').Substring(0, 12) }
        } | Out-Null
        $S.Schedules = @(Invoke-Section $Snap 'Sql' 'schedules' { Invoke-RoSql $conn 'SELECT j.name AS Job, s.name AS Schedule, CAST(s.enabled AS int) AS Enabled, s.freq_type AS FreqType, s.freq_interval AS FreqInterval, s.freq_subday_type AS SubdayType, s.freq_subday_interval AS SubdayInterval, s.freq_relative_interval AS RelInterval, s.freq_recurrence_factor AS Recurrence, s.active_start_time AS StartTime, s.active_end_time AS EndTime, s.active_end_date AS EndDate FROM msdb.dbo.sysjobschedules js JOIN msdb.dbo.sysschedules s ON s.schedule_id = js.schedule_id JOIN msdb.dbo.sysjobs j ON j.job_id = js.job_id' })
        foreach ($x in $S.Schedules) { $txt = try { Get-ScheduleText $x } catch { 'unparsed schedule' }; $x | Add-Member -NotePropertyName Text -NotePropertyValue $txt -Force }
        $S.ProxySubsystems = @(Invoke-Section $Snap 'Sql' 'proxy_subsystems' { Invoke-RoSql $conn 'SELECT p.name AS Proxy, ss.subsystem AS Subsystem FROM msdb.dbo.sysproxies p JOIN msdb.dbo.sysproxysubsystem ps ON ps.proxy_id = p.proxy_id JOIN msdb.dbo.syssubsystems ss ON ss.subsystem_id = ps.subsystem_id' })
        $S.ProxyLogins = @(Invoke-Section $Snap 'Sql' 'proxy_logins' { Invoke-RoSql $conn 'SELECT p.name AS Proxy, COALESCE(SUSER_SNAME(pl.sid), USER_NAME(pl.sid), CONVERT(varchar(100), pl.sid, 1)) AS Principal FROM msdb.dbo.sysproxies p JOIN msdb.dbo.sysproxylogin pl ON pl.proxy_id = p.proxy_id' })
        $S.MailAccounts = @(Invoke-Section $Snap 'Sql' 'dbmail_accounts' { Invoke-RoSql $conn 'SELECT a.name AS Account, a.email_address AS Email, a.display_name AS DisplayName, a.replyto_address AS ReplyTo, s.servername AS SmtpServer, s.port AS Port, CAST(s.enable_ssl AS int) AS Ssl, CAST(s.use_default_credentials AS int) AS DefaultCreds, s.username AS SmtpUser FROM msdb.dbo.sysmail_account a LEFT JOIN msdb.dbo.sysmail_server s ON s.account_id = a.account_id' })
        $S.MailPrincipals = @(Invoke-Section $Snap 'Sql' 'dbmail_principals' { Invoke-RoSql $conn "SELECT p.name AS Profile, CASE WHEN pp.principal_sid = 0x00 THEN N'public' ELSE COALESCE(SUSER_SNAME(pp.principal_sid), CONVERT(varchar(100), pp.principal_sid, 1)) END AS Principal, CAST(pp.is_default AS int) AS IsDefault FROM msdb.dbo.sysmail_principalprofile pp JOIN msdb.dbo.sysmail_profile p ON p.profile_id = pp.profile_id" })
        $S.MailConfig = @(Invoke-Section $Snap 'Sql' 'dbmail_config' { Invoke-RoSql $conn 'SELECT paramname AS Name, paramvalue AS Value FROM msdb.dbo.sysmail_configuration' })
        $S.BackupDevices = @(Invoke-Section $Snap 'Sql' 'backup_devices' { Invoke-RoSql $conn 'SELECT name AS Name, type_desc AS Type, physical_name AS Path FROM sys.backup_devices' })
        $S.MaintPlans = @(Invoke-Section $Snap 'Sql' 'maint_plans' { Invoke-RoSql $conn 'SELECT name AS Name, owner AS Owner FROM msdb.dbo.sysmaintplan_plans' })
        $S.JobCategories = @(Invoke-Section $Snap 'Sql' 'job_categories' { Invoke-RoSql $conn "SELECT name AS Name, category_class AS Class FROM msdb.dbo.syscategories WHERE category_id >= 100" })
        $S.Proxies = @(Invoke-Section $Snap 'Sql' 'proxies' { Invoke-RoSql $conn 'SELECT p.name AS Name, c.name AS Credential, CAST(p.enabled AS int) AS Enabled FROM msdb.dbo.sysproxies p LEFT JOIN sys.credentials c ON c.credential_id = p.credential_id' })
        $S.Operators = @(Invoke-Section $Snap 'Sql' 'operators' { Invoke-RoSql $conn 'SELECT name AS Name, CAST(enabled AS int) AS Enabled, email_address AS Email FROM msdb.dbo.sysoperators' })
        $S.Alerts = @(Invoke-Section $Snap 'Sql' 'alerts' { Invoke-RoSql $conn 'SELECT name AS Name, message_id AS MessageId, severity AS Severity, CAST(enabled AS int) AS Enabled FROM msdb.dbo.sysalerts' })
        $S.Mail = @(Invoke-Section $Snap 'Sql' 'dbmail' { Invoke-RoSql $conn 'SELECT p.name AS Profile, a.name AS Account, a.email_address AS Email, s.servername AS SmtpServer, s.port AS Port, CAST(s.enable_ssl AS int) AS Ssl FROM msdb.dbo.sysmail_profile p LEFT JOIN msdb.dbo.sysmail_profileaccount pa ON pa.profile_id = p.profile_id LEFT JOIN msdb.dbo.sysmail_account a ON a.account_id = pa.account_id LEFT JOIN msdb.dbo.sysmail_server s ON s.account_id = a.account_id' })

        # ---- SSIS
        if (@($S.Databases | Where-Object { $_.Name -eq 'SSISDB' }).Count -gt 0) {
            $S.SsisProjects = @(Invoke-Section $Snap 'Sql' 'ssisdb' { Invoke-RoSql $conn 'SELECT f.name AS Folder, p.name AS Project, CONVERT(varchar(20), p.last_deployed_time, 23) AS Deployed FROM SSISDB.catalog.projects p JOIN SSISDB.catalog.folders f ON f.folder_id = p.folder_id' })
            $S.SsisEnvironments = @(Invoke-Section $Snap 'Sql' 'ssisdb_env' { Invoke-RoSql $conn 'SELECT f.name AS Folder, e.name AS Environment FROM SSISDB.catalog.environments e JOIN SSISDB.catalog.folders f ON f.folder_id = e.folder_id' })
        }
        $S.SsisMsdb = @(Invoke-Section $Snap 'Sql' 'msdb_packages' { Invoke-RoSql $conn "SELECT f.foldername AS Folder, p.name AS Package FROM msdb.dbo.sysssispackages p LEFT JOIN msdb.dbo.sysssispackagefolders f ON f.folderid = p.folderid WHERE ISNULL(f.foldername, N'') NOT IN (N'Data Collector', N'Maintenance Plans') AND p.packagetype <> 6" })

        # ---- server objects / security objects
        $S.ServerTriggers = @(Invoke-Section $Snap 'Sql' 'server_triggers' { Invoke-RoSql $conn 'SELECT name AS Name, CAST(is_disabled AS int) AS Disabled FROM sys.server_triggers' })
        $S.Endpoints = @(Invoke-Section $Snap 'Sql' 'endpoints' { Invoke-RoSql $conn 'SELECT e.name AS Name, e.type_desc AS Type, e.state_desc AS State, t.port AS Port, m.role_desc AS Role, m.encryption_algorithm_desc AS Encryption, SUSER_NAME(e.principal_id) AS Owner FROM sys.endpoints e LEFT JOIN sys.tcp_endpoints t ON t.endpoint_id = e.endpoint_id LEFT JOIN sys.database_mirroring_endpoints m ON m.endpoint_id = e.endpoint_id WHERE e.endpoint_id > 5' })
        $S.StartupProcs = @(Invoke-Section $Snap 'Sql' 'startup_procs' { @(Invoke-RoSql $conn "SELECT name AS Name FROM master.sys.procedures WHERE OBJECTPROPERTY(object_id, 'ExecIsStartup') = 1") | ForEach-Object { $_.Name } })
        $S.SysDbObjects = @(Invoke-Section $Snap 'Sql' 'sysdb_objects' { Invoke-RoSql $conn "SELECT N'master' AS Db, SCHEMA_NAME(o.schema_id) + N'.' + o.name AS Name, o.type_desc AS Type FROM master.sys.objects o WHERE o.is_ms_shipped = 0 AND o.type IN ('U','P','V','FN','IF','TF','PC','FS','FT','SN') UNION ALL SELECT N'msdb', SCHEMA_NAME(o.schema_id) + N'.' + o.name, o.type_desc FROM msdb.sys.objects o WHERE o.is_ms_shipped = 0 AND o.type IN ('U','P','V','FN','IF','TF','PC','FS','FT','SN') AND o.name NOT LIKE N'sysutility%' AND o.name NOT LIKE N'syspolicy%'" })
        $S.XProcs = @(Invoke-Section $Snap 'Sql' 'extended_procs' { Invoke-RoSql $conn 'SELECT name AS Name, dll_name AS Dll FROM master.sys.extended_procedures WHERE is_ms_shipped = 0' })
        $S.Audits = @(Invoke-Section $Snap 'Sql' 'audits' { Invoke-RoSql $conn 'SELECT a.name AS Name, a.type_desc AS Type, a.on_failure_desc AS OnFailure, CAST(a.is_state_enabled AS int) AS Enabled, f.log_file_path AS Path FROM sys.server_audits a LEFT JOIN sys.server_file_audits f ON f.audit_id = a.audit_id' })
        $S.AuditSpecs = @(Invoke-Section $Snap 'Sql' 'audit_specs' { Invoke-RoSql $conn 'SELECT name AS Name, CAST(is_state_enabled AS int) AS Enabled FROM sys.server_audit_specifications' })
        $S.XeSessions = @(Invoke-Section $Snap 'Sql' 'xe_sessions' { Invoke-RoSql $conn "SELECT s.name AS Name, CAST(s.startup_state AS int) AS AutoStart, (SELECT TOP (1) CAST(f.value AS nvarchar(512)) FROM sys.server_event_session_fields f WHERE f.event_session_id = s.event_session_id AND f.name = N'filename') AS FilePath FROM sys.server_event_sessions s WHERE s.name NOT IN (N'system_health', N'AlwaysOn_health', N'telemetry_xevents', N'sp_server_diagnostics session', N'hkenginexesession', N'QuickSessionStandard', N'QuickSessionTSQL')" })
        Invoke-Section $Snap 'Sql' 'resource_governor' {
            $c = @(Invoke-RoSql $conn "SELECT CAST(is_enabled AS int) AS Enabled, OBJECT_SCHEMA_NAME(classifier_function_id, 1) + N'.' + OBJECT_NAME(classifier_function_id, 1) AS Classifier FROM sys.resource_governor_configuration")[0]
            $S.ResourceGovernor = [ordered]@{
                Enabled    = $c.Enabled
                Classifier = $c.Classifier
                Pools      = @(Invoke-RoSql $conn 'SELECT name AS Name, min_cpu_percent AS MinCpu, max_cpu_percent AS MaxCpu, min_memory_percent AS MinMem, max_memory_percent AS MaxMem FROM sys.resource_governor_resource_pools WHERE pool_id > 2')
                Groups     = @(Invoke-RoSql $conn 'SELECT g.name AS Name, p.name AS Pool FROM sys.resource_governor_workload_groups g JOIN sys.resource_governor_resource_pools p ON p.pool_id = g.pool_id WHERE g.group_id > 2')
            }
        } | Out-Null
        $S.CryptoProviders = @(Invoke-Section $Snap 'Sql' 'crypto_providers' { Invoke-RoSql $conn 'SELECT name AS Name, dll_path AS Dll, CAST(is_enabled AS int) AS Enabled FROM sys.cryptographic_providers' })
        $S.Certificates = @(Invoke-Section $Snap 'Sql' 'certificates' { Invoke-RoSql $conn "SELECT name AS Name, subject AS Subject, CONVERT(varchar(100), thumbprint, 2) AS Thumbprint, CONVERT(varchar(20), expiry_date, 23) AS Expiry, pvt_key_encryption_type_desc AS PrivateKey FROM master.sys.certificates WHERE name NOT LIKE N'##%'" })
        $S.AsymKeys = @(Invoke-Section $Snap 'Sql' 'asym_keys' { Invoke-RoSql $conn "SELECT name AS Name, CONVERT(varchar(100), thumbprint, 2) AS Thumbprint, provider_type AS ProviderType FROM master.sys.asymmetric_keys WHERE name NOT LIKE N'##%'" })
        $S.Tde = @(Invoke-Section $Snap 'Sql' 'tde' { Invoke-RoSql $conn 'SELECT DB_NAME(k.database_id) AS Db, CONVERT(varchar(100), k.encryptor_thumbprint, 2) AS Thumbprint, k.encryption_state AS State FROM sys.dm_database_encryption_keys k WHERE k.database_id <> 2' })
        $S.Messages = @(Invoke-Section $Snap 'Sql' 'messages' { Invoke-RoSql $conn 'SELECT message_id AS Id, language_id AS Lang, severity AS Severity, CAST(is_event_logged AS int) AS Logged, text AS Text FROM sys.messages WHERE message_id > 50000' })
        if ($major -ge 11) {
            if ([int](Get-Prop $S.Instance 'IsHadrEnabled' 0) -eq 1) {
                $S.AgStates = @(Invoke-Section $Snap 'Sql' 'ag_states' { Invoke-RoSql $conn 'SELECT ar.replica_server_name AS Replica, DB_NAME(rs.database_id) AS Db, rs.synchronization_state_desc AS SyncState, rs.synchronization_health_desc AS Health, rs.log_send_queue_size AS SendQueueKB, rs.redo_queue_size AS RedoQueueKB, CAST(rs.is_suspended AS int) AS Suspended FROM sys.dm_hadr_database_replica_states rs JOIN sys.availability_replicas ar ON ar.replica_id = rs.replica_id' })
            }
            $S.Ag = @(Invoke-Section $Snap 'Sql' 'ag' { Invoke-RoSql $conn 'SELECT ag.name AS AgName, ar.replica_server_name AS Replica, ar.availability_mode_desc AS Mode, ar.failover_mode_desc AS Failover, ar.endpoint_url AS Endpoint, (SELECT TOP (1) l.dns_name FROM sys.availability_group_listeners l WHERE l.group_id = ag.group_id) AS Listener, (SELECT TOP (1) l.port FROM sys.availability_group_listeners l WHERE l.group_id = ag.group_id) AS ListenerPort FROM sys.availability_groups ag JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id' })
        }
    } finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

function Get-AgTopology {
    # Discover replicas from any replica or listener. Returns @{ Primary; Replicas=@(name,mode) }
    param([string]$Server, [string]$Ag)
    $conn = New-SqlConn -Server $Server
    try {
        $rows = Invoke-RoSql $conn 'SELECT ag.name AS AgName, ar.replica_server_name AS Replica, ar.availability_mode_desc AS Mode, ar.failover_mode_desc AS Failover, ags.primary_replica AS PrimaryReplica FROM sys.availability_groups ag JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id LEFT JOIN sys.dm_hadr_availability_group_states ags ON ags.group_id = ag.group_id'
    } finally { $conn.Close(); $conn.Dispose() }
    $rows = @($rows)
    if ($rows.Count -eq 0) { throw "No availability groups found on $Server." }
    $names = @($rows | Select-Object -ExpandProperty AgName -Unique)
    if ($Ag) { $rows = @($rows | Where-Object { $_.AgName -eq $Ag }); if ($rows.Count -eq 0) { throw "AG '$Ag' not found on $Server." } }
    elseif ($names.Count -gt 1) { Write-Log ("Several AGs found ({0}); using '{1}'. Use -AvailabilityGroup to choose." -f ($names -join ', '), $names[0]) 'WARN'; $rows = @($rows | Where-Object { $_.AgName -eq $names[0] }) }
    $primary = ($rows | Where-Object { $_.PrimaryReplica } | Select-Object -First 1).PrimaryReplica
    if (-not $primary) { throw 'Primary replica could not be determined (VIEW SERVER STATE needed).' }
    return @{ AgName = $rows[0].AgName; Primary = $primary; Replicas = @($rows | Where-Object { $_.Replica -ne $primary } | ForEach-Object { @{ Name = $_.Replica; Mode = $(if ($_.Mode -like 'SYNCHRONOUS*') { 'Synchronous' } else { 'Asynchronous' }) } }) }
}
#endregion

#region ---------------------------------------------------------------- Windows (read-only) access
$script:HKLM = [uint32]2147483650
$script:RegReadMethods = @('EnumKey', 'EnumValues', 'GetStringValue', 'GetExpandedStringValue', 'GetDWORDValue', 'GetMultiStringValue', 'GetQWORDValue', 'GetBinaryValue')

function New-OsSession {
    param([string]$Computer)
    $short = ($Computer -split '\.')[0]
    if ($short -eq $env:COMPUTERNAME -or $Computer -in @('.', 'localhost', '127.0.0.1')) {
        return New-CimSession -SessionOption (New-CimSessionOption -Protocol Dcom) -ErrorAction Stop
    }
    $opt = New-CimSessionOption -Protocol $CimProtocol
    $p = @{ ComputerName = $Computer; SessionOption = $opt; OperationTimeoutSec = 60; ErrorAction = 'Stop' }
    if ($WindowsCredential) { $p['Credential'] = $WindowsCredential }
    return New-CimSession @p
}

function Invoke-RegRead {
    # StdRegProv, read methods only. Anything else is refused here.
    param($Cs, [string]$Method, [hashtable]$Arguments)
    if ($script:RegReadMethods -notcontains $Method) { throw "Read-only guard: registry method '$Method' is not allowed." }
    $a = @{ hDefKey = $script:HKLM }
    foreach ($k in $Arguments.Keys) { $a[$k] = $Arguments[$k] }
    $r = Invoke-CimMethod -CimSession $Cs -Namespace 'root/default' -ClassName 'StdRegProv' -MethodName $Method -Arguments $a -ErrorAction Stop
    if ($r.ReturnValue -ne 0) { return $null }
    return $r
}
function Get-RegSubKeys { param($Cs, [string]$Key) $r = Invoke-RegRead $Cs 'EnumKey' @{ sSubKeyName = $Key }; if ($r -and $r.sNames) { return @($r.sNames) } return @() }
function Get-RegValueNames { param($Cs, [string]$Key) $r = Invoke-RegRead $Cs 'EnumValues' @{ sSubKeyName = $Key }; if ($r -and $r.sNames) { return @($r.sNames) } return @() }
function Get-RegString {
    param($Cs, [string]$Key, [string]$Name)
    $r = Invoke-RegRead $Cs 'GetStringValue' @{ sSubKeyName = $Key; sValueName = $Name }
    if ($r) { return $r.sValue }
    $r = Invoke-RegRead $Cs 'GetExpandedStringValue' @{ sSubKeyName = $Key; sValueName = $Name }
    if ($r) { return $r.sValue }
    return $null
}
function Get-RegDword { param($Cs, [string]$Key, [string]$Name) $r = Invoke-RegRead $Cs 'GetDWORDValue' @{ sSubKeyName = $Key; sValueName = $Name }; if ($r) { return [int64]$r.uValue } return $null }

function ConvertFrom-CertBlob {
    # Parses a SystemCertificates registry "Blob" (serialized cert element) without touching any store.
    param([byte[]]$Blob, [string]$Thumb)
    if (-not $Blob -or $Blob.Length -lt 12) { return $null }
    $i = 0; $der = $null; $hasKey = $false; $friendly = $null
    while ($i + 12 -le $Blob.Length) {
        $id = [BitConverter]::ToUInt32($Blob, $i); $len = [BitConverter]::ToUInt32($Blob, $i + 8); $i += 12
        if ($len -gt ($Blob.Length - $i)) { break }
        if ($id -eq 32) { $der = New-Object byte[] $len; [Array]::Copy($Blob, $i, $der, 0, $len) }
        elseif ($id -eq 2) { $hasKey = $true }
        elseif ($id -eq 11 -and $len -gt 2) { $friendly = [Text.Encoding]::Unicode.GetString($Blob, $i, $len - 2) }
        $i += $len
    }
    if (-not $der) { return $null }
    $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 (, $der)
    $san = @()
    foreach ($e in $c.Extensions) { if ($e.Oid.Value -eq '2.5.29.17') { $san = @(([string]$e.Format($false)) -split ',\s*' | ForEach-Object { ($_ -replace '^(DNS Name|DNS|IP Address)[=:]\s*', '').Trim() } | Where-Object { $_ }) } }
    $eku = @()
    foreach ($e in $c.Extensions) { if ($e.Oid.Value -eq '2.5.29.37') { $eku = @($e.EnhancedKeyUsages | ForEach-Object { $_.FriendlyName }) } }
    return [pscustomobject]@{ Thumbprint = $c.Thumbprint; Subject = $c.Subject; Issuer = $c.Issuer; NotAfter = $c.NotAfter.ToString('yyyy-MM-dd'); NotAfterDate = $c.NotAfter; SelfSigned = ($c.Subject -eq $c.Issuer); San = $san; Eku = ($eku -join ', '); HasPrivateKey = $hasKey; FriendlyName = $friendly }
}

function Get-DotNetName {
    param($Release)
    if ($null -eq $Release) { return $null }
    $map = @(@(533320, '4.8.1'), @(528040, '4.8'), @(461808, '4.7.2'), @(461308, '4.7.1'), @(460798, '4.7'), @(394802, '4.6.2'), @(394254, '4.6.1'), @(393295, '4.6'), @(379893, '4.5.2'), @(378675, '4.5.1'), @(378389, '4.5'))
    foreach ($m in $map) { if ($Release -ge $m[0]) { return $m[1] } }
    return '4.0'
}

$script:KnownOleDb = @('MSOLEDBSQL', 'MSOLEDBSQL19', 'SQLNCLI11', 'SQLNCLI10', 'SQLOLEDB', 'MSDASQL', 'Microsoft.ACE.OLEDB.12.0', 'Microsoft.ACE.OLEDB.16.0', 'Microsoft.Jet.OLEDB.4.0', 'OraOLEDB.Oracle', 'MSDAORA', 'IBMDADB2', 'DB2OLEDB', 'MSOLAP', 'ADsDSOObject', 'Search.CollatorDSO')

function Get-OsSnapshot {
    param($Snap, [string[]]$ExtraProviders = @())
    $O = $Snap.Os
    $computer = $Snap.Computer
    $cs = $null
    try { $cs = New-OsSession -Computer $computer } catch {
        $O.Sections['connect'] = 'error: ' + ($_.Exception.Message -replace '\s+', ' ')
        Write-Log "  Windows (CIM) connect to $computer failed: $($_.Exception.Message)" 'ERROR'
        return
    }
    $Snap.OsOk = $true
    $Snap.CimSession = $cs
    Invoke-Section $Snap 'Os' 'os' {
        $os = Get-CimInstance -CimSession $cs -ClassName Win32_OperatingSystem -Property Caption, Version, BuildNumber, OSArchitecture, Locale, OSLanguage, TotalVisibleMemorySize, LastBootUpTime
        $c = Get-CimInstance -CimSession $cs -ClassName Win32_ComputerSystem -Property Domain, PartOfDomain, Manufacturer, Model, NumberOfLogicalProcessors, TotalPhysicalMemory
        $tz = Get-CimInstance -CimSession $cs -ClassName Win32_TimeZone -Property Caption, StandardName, Bias
        $cpu = @(Get-CimInstance -CimSession $cs -ClassName Win32_Processor -Property Name, NumberOfCores, NumberOfLogicalProcessors)
        $O.Info = [ordered]@{
            Caption = $os.Caption; Version = $os.Version; Build = $os.BuildNumber; Arch = $os.OSArchitecture
            Locale = $os.Locale; Language = $os.OSLanguage; Domain = $c.Domain; PartOfDomain = [bool]$c.PartOfDomain
            Manufacturer = $c.Manufacturer; Model = $c.Model
            IsVirtual = [bool]($c.Model -match '(?i)virtual|vmware|kvm|hvm|xen|qemu|google|amazon|ec2|azure')
            Cores = [int](Get-Sum $cpu NumberOfCores); Logical = [int]$c.NumberOfLogicalProcessors
            Sockets = $cpu.Count; CpuName = ($cpu | Select-Object -First 1).Name
            MemoryMB = [int64]([int64]$c.TotalPhysicalMemory / 1MB)
            TimeZone = $tz.Caption; TimeZoneStd = $tz.StandardName; TzBias = $tz.Bias
            LastBoot = if ($os.LastBootUpTime) { ([datetime]$os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm') } else { $null }
        }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'volumes' {
        $O.Volumes = @(Get-CimInstance -CimSession $cs -ClassName Win32_Volume -Filter 'DriveType=3' -Property Name, DriveLetter, Label, Capacity, FreeSpace, BlockSize, FileSystem |
            Where-Object { $_.Name -notmatch '^\\\\\?\\' } |
            ForEach-Object { [pscustomobject]@{ Name = $_.Name; Letter = $_.DriveLetter; Label = $_.Label; SizeMB = [int64]($_.Capacity / 1MB); FreeMB = [int64]($_.FreeSpace / 1MB); BlockSize = $_.BlockSize; FileSystem = $_.FileSystem } })
    } | Out-Null
    Invoke-Section $Snap 'Os' 'power' {
        $p = Get-CimInstance -CimSession $cs -Namespace 'root/cimv2/power' -ClassName Win32_PowerPlan -Filter 'IsActive=TRUE' -ErrorAction Stop | Select-Object -First 1
        $O.PowerPlan = if ($p) { $p.ElementName } else { $null }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'services' {
        $O.Services = @(Get-CimInstance -CimSession $cs -ClassName Win32_Service -Property Name, DisplayName, StartMode, State, StartName, PathName |
            ForEach-Object { [pscustomobject]@{ Name = $_.Name; DisplayName = $_.DisplayName; StartMode = $_.StartMode; State = $_.State; Account = $_.StartName; Path = $_.PathName } })
    } | Out-Null
    Invoke-Section $Snap 'Os' 'features' {
        $O.Features = @(Get-CimInstance -CimSession $cs -ClassName Win32_ServerFeature -Property ID, Name -ErrorAction Stop | ForEach-Object { $_.Name } | Sort-Object -Unique)
    } | Out-Null

    # registry-based facets
    if ($Skip -notcontains 'Software') {
        Invoke-Section $Snap 'Os' 'software' {
            $list = New-List
            foreach ($view in @(@('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'x64'), @('SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall', 'x86'))) {
                foreach ($k in (Get-RegSubKeys $cs $view[0])) {
                    $path = $view[0] + '\' + $k
                    # one EnumValues per key, then only the values that exist (keeps CIM round-trips low)
                    $ev = Invoke-RegRead $cs 'EnumValues' @{ sSubKeyName = $path }
                    if (-not $ev -or -not $ev.sNames) { continue }
                    $names = @($ev.sNames)
                    if ($names -notcontains 'DisplayName' -or $names -contains 'ParentKeyName') { continue }
                    if ($names -contains 'SystemComponent' -and (Get-RegDword $cs $path 'SystemComponent') -eq 1) { continue }
                    $dn = Get-RegString $cs $path 'DisplayName'
                    if ([string]::IsNullOrWhiteSpace($dn) -or $dn -match '\bKB\d{6,}\b') { continue }
                    $ver = if ($names -contains 'DisplayVersion') { Get-RegString $cs $path 'DisplayVersion' } else { $null }
                    $pub = if ($names -contains 'Publisher') { Get-RegString $cs $path 'Publisher' } else { $null }
                    Add-To $list ([pscustomobject]@{ Name = $dn.Trim(); Version = $ver; Publisher = $pub; View = $view[1] })
                }
            }
            $O.Software = $list.ToArray()
        } | Out-Null
    }
    Invoke-Section $Snap 'Os' 'dotnet' {
        $rel = Get-RegDword $cs 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' 'Release'
        $n35 = Get-RegDword $cs 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5' 'Install'
        $strong = Get-RegDword $cs 'SOFTWARE\Microsoft\.NETFramework\v4.0.30319' 'SchUseStrongCrypto'
        $O.DotNet = [ordered]@{ Release = $rel; Version = (Get-DotNetName $rel); Net35 = [bool]($n35 -eq 1); StrongCrypto = $strong }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'odbc' {
        $drv = New-List; $dsn = New-List
        foreach ($v in @(@('SOFTWARE\ODBC', 'x64'), @('SOFTWARE\WOW6432Node\ODBC', 'x86'))) {
            foreach ($n in (Get-RegValueNames $cs ($v[0] + '\ODBCINST.INI\ODBC Drivers'))) { Add-To $drv ([pscustomobject]@{ Name = $n; Bitness = $v[1] }) }
            foreach ($n in (Get-RegValueNames $cs ($v[0] + '\ODBC.INI\ODBC Data Sources'))) {
                $d = Get-RegString $cs ($v[0] + '\ODBC.INI\ODBC Data Sources') $n
                $srv = Get-RegString $cs ($v[0] + '\ODBC.INI\' + $n) 'Server'
                Add-To $dsn ([pscustomobject]@{ Name = $n; Driver = $d; Server = $srv; Bitness = $v[1] })
            }
        }
        $O.OdbcDrivers = $drv.ToArray(); $O.Dsns = $dsn.ToArray()
    } | Out-Null
    Invoke-Section $Snap 'Os' 'oledb' {
        $res = [ordered]@{}
        foreach ($p in (@($script:KnownOleDb) + @($ExtraProviders) | Where-Object { $_ } | Sort-Object -Unique)) {
            $clsid = Get-RegString $cs ("SOFTWARE\Classes\$p\CLSID") ''
            $x64 = $false; $x86 = $false
            if ($clsid) {
                $x64 = [bool](Get-RegString $cs ("SOFTWARE\Classes\CLSID\$clsid\InprocServer32") '')
                $x86 = [bool](Get-RegString $cs ("SOFTWARE\Classes\WOW6432Node\CLSID\$clsid\InprocServer32") '')
            }
            $res[$p] = [ordered]@{ x64 = $x64; x86 = $x86 }
        }
        $O.OleDb = $res
    } | Out-Null
    Invoke-Section $Snap 'Os' 'aliases' {
        $al = New-List
        foreach ($v in @(@('SOFTWARE\Microsoft\MSSQLServer\Client\ConnectTo', 'x64'), @('SOFTWARE\WOW6432Node\Microsoft\MSSQLServer\Client\ConnectTo', 'x86'))) {
            foreach ($n in (Get-RegValueNames $cs $v[0])) { Add-To $al ([pscustomobject]@{ Name = $n; Target = (Get-RegString $cs $v[0] $n); Bitness = $v[1] }) }
        }
        $O.Aliases = $al.ToArray()
    } | Out-Null
    Invoke-Section $Snap 'Os' 'tls' {
        $t = [ordered]@{}
        foreach ($proto in @('SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')) {
            foreach ($side in @('Server', 'Client')) {
                $e = Get-RegDword $cs ("SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\$side") 'Enabled'
                $t["$proto $side"] = if ($null -eq $e) { 'OS default' } elseif ($e -eq 0) { 'Disabled' } else { 'Enabled' }
            }
        }
        $O.Tls = $t
    } | Out-Null
    $regRoot = [string](Get-Prop $Snap.Sql.Instance 'RegRoot' '')
    if ($regRoot) {
        Invoke-Section $Snap 'Os' 'agent_reg' {
            $a = [ordered]@{}
            foreach ($n in @('DatabaseMailProfile', 'UseDatabaseMail', 'AlertFailSafeOperator', 'AlertNotificationMethod', 'JobHistoryMaxRows', 'JobHistoryMaxRowsPerJob', 'ErrorLogFile', 'RestartSQLServer', 'MonitorAutoStart')) {
                $v = Get-RegString $cs ($regRoot + '\SQLServerAgent') $n
                if ($null -eq $v) { $v = Get-RegDword $cs ($regRoot + '\SQLServerAgent') $n }
                $a[$n] = $v
            }
            $O.AgentReg = $a
        } | Out-Null
        Invoke-Section $Snap 'Os' 'provider_options' {
            $po = [ordered]@{}
            foreach ($p in (Get-RegSubKeys $cs ($regRoot + '\Providers'))) {
                $o = [ordered]@{}
                foreach ($n in @('AllowInProcess', 'DynamicParameters', 'NestedQueries', 'LevelZeroOnly', 'NonTransactedUpdates', 'IndexAsAccessPath', 'DisallowAdHocAccess', 'SqlServerLIKE')) {
                    $v = Get-RegDword $cs ($regRoot + '\Providers\' + $p) $n
                    if ($null -ne $v) { $o[$n] = [int]$v }
                }
                $po[$p] = $o
            }
            $O.ProviderOptions = $po
        } | Out-Null
    }
    Invoke-Section $Snap 'Os' 'certificates' {
        $list = New-List
        $base = 'SOFTWARE\Microsoft\SystemCertificates\MY\Certificates'
        foreach ($th in (Get-RegSubKeys $cs $base)) {
            $r = Invoke-RegRead $cs 'GetBinaryValue' @{ sSubKeyName = ($base + '\' + $th); sValueName = 'Blob' }
            if (-not $r -or -not $r.uValue) { continue }
            try { $c = ConvertFrom-CertBlob -Blob ([byte[]]$r.uValue) -Thumb $th; if ($c) { Add-To $list $c } } catch { }
        }
        $O.Certificates = $list.ToArray()
        $sqlTh = ([string](Get-Prop $Snap.Sql.Net 'Certificate' '')).Trim().ToUpperInvariant()
        if ($sqlTh) { $O.SqlCert = $O.Certificates | Where-Object { $_.Thumbprint -eq $sqlTh } | Select-Object -First 1; if (-not $O.SqlCert) { $O.SqlCert = [pscustomobject]@{ Thumbprint = $sqlTh; Missing = $true } } }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'firewall' {
        $prof = @(Get-CimInstance -CimSession $cs -Namespace 'root/StandardCimv2' -ClassName MSFT_NetFirewallProfile -ErrorAction Stop)
        $on = @($prof | Where-Object { [int]$_.Enabled -eq 1 }).Count
        $rules = @{}
        foreach ($r in @(Get-CimInstance -CimSession $cs -Namespace 'root/StandardCimv2' -ClassName MSFT_NetFirewallRule -Filter 'Enabled = 1 AND Direction = 1 AND Action = 2' -ErrorAction Stop)) { $rules[[string]$r.InstanceID] = $r.DisplayName }
        $ports = New-List; $apps = New-List
        foreach ($f in @(Get-CimInstance -CimSession $cs -Namespace 'root/StandardCimv2' -ClassName MSFT_NetFirewallPortFilter -ErrorAction Stop)) {
            if (-not $rules.ContainsKey([string]$f.InstanceID)) { continue }
            foreach ($lp in @($f.LocalPort)) { if ($lp -and $lp -ne 'Any') { Add-To $ports ("{0}:{1}" -f $f.Protocol, $lp) } }
        }
        foreach ($f in @(Get-CimInstance -CimSession $cs -Namespace 'root/StandardCimv2' -ClassName MSFT_NetFirewallApplicationFilter -ErrorAction Stop)) {
            if ($rules.ContainsKey([string]$f.InstanceID) -and [string]$f.Program -match '(?i)sqlservr\.exe$') { Add-To $apps ([string]$f.Program) }
        }
        $O.Firewall = [ordered]@{ ProfilesOn = $on; Ports = @($ports | Sort-Object -Unique); SqlProgramRules = @($apps | Sort-Object -Unique) }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'os_detail' {
        $cv = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $O.Patch = [ordered]@{
            InstallationType = Get-RegString $cs $cv 'InstallationType'
            DisplayVersion   = Get-RegString $cs $cv 'DisplayVersion'
            Ubr              = Get-RegDword $cs $cv 'UBR'
            LastHotfix       = $null
        }
        try {
            $q = @(Get-CimInstance -CimSession $cs -ClassName Win32_QuickFixEngineering -Property HotFixID, InstalledOn -ErrorAction Stop | Where-Object { $_.InstalledOn })
            if ($q.Count) { $O.Patch.LastHotfix = (@($q | ForEach-Object { [datetime]$_.InstalledOn }) | Sort-Object -Descending | Select-Object -First 1).ToString('yyyy-MM-dd'); $O.Patch.LastHotfixId = ($q | Sort-Object { [datetime]$_.InstalledOn } -Descending | Select-Object -First 1).HotFixID }
        } catch { }
        $pr = New-List
        if ((Get-RegSubKeys $cs 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing') -contains 'RebootPending') { Add-To $pr 'Component Based Servicing' }
        if ((Get-RegSubKeys $cs 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update') -contains 'RebootRequired') { Add-To $pr 'Windows Update' }
        $pf = Invoke-RegRead $cs 'GetMultiStringValue' @{ sSubKeyName = 'SYSTEM\CurrentControlSet\Control\Session Manager'; sValueName = 'PendingFileRenameOperations' }
        if ($pf -and $pf.sValue -and @($pf.sValue | Where-Object { $_ }).Count) { Add-To $pr 'Pending file rename operations' }
        $O.PendingReboot = $pr.ToArray()
        $O.Fips = Get-RegDword $cs 'SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy' 'Enabled'
        $dtc = [ordered]@{}
        foreach ($n in @('NetworkDtcAccess', 'NetworkDtcAccessTransactions', 'NetworkDtcAccessInbound', 'NetworkDtcAccessOutbound', 'XaTransactions')) { $dtc[$n] = Get-RegDword $cs 'SOFTWARE\Microsoft\MSDTC\Security' $n }
        $O.Msdtc = $dtc
        $inst = New-List
        foreach ($kind in @(@('SQL', 'Engine'), @('RS', 'Reporting'), @('OLAP', 'Analysis'))) {
            foreach ($n in (Get-RegValueNames $cs ('SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\' + $kind[0]))) { Add-To $inst ("{0}: {1} ({2})" -f $kind[1], $n, (Get-RegString $cs ('SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\' + $kind[0]) $n)) }
        }
        $O.SqlInstances = $inst.ToArray()
    } | Out-Null
    Invoke-Section $Snap 'Os' 'environment' {
        $O.EnvVars = [ordered]@{}
        foreach ($e in @(Get-CimInstance -CimSession $cs -ClassName Win32_Environment -Filter 'SystemVariable = TRUE' -Property Name, VariableValue -ErrorAction Stop)) { $O.EnvVars[[string]$e.Name] = [string]$e.VariableValue }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'scheduled_tasks' {
        $t = New-List
        foreach ($x in @(Get-CimInstance -CimSession $cs -Namespace 'Root/Microsoft/Windows/TaskScheduler' -ClassName MSFT_ScheduledTask -ErrorAction Stop | Where-Object { [string]$_.TaskPath -notlike '\Microsoft\*' })) {
            $acts = @($x.Actions | ForEach-Object { Protect-Secret ((([string](Get-Prop $_ 'Execute' '')) + ' ' + ([string](Get-Prop $_ 'Arguments' ''))).Trim()) } | Where-Object { $_ })
            Add-To $t ([pscustomobject]@{ Name = ([string]$x.TaskPath + [string]$x.TaskName); State = [string]$x.State; RunAs = [string](Get-Prop $x.Principal 'UserId' ''); Actions = ($acts -join ' ; ') })
        }
        $O.ScheduledTasks = $t.ToArray()
    } | Out-Null
    Invoke-Section $Snap 'Os' 'defender' {
        $mp = Get-CimInstance -CimSession $cs -Namespace 'root/Microsoft/Windows/Defender' -ClassName MSFT_MpPreference -ErrorAction Stop | Select-Object -First 1
        $st = $null; try { $st = Get-CimInstance -CimSession $cs -Namespace 'root/Microsoft/Windows/Defender' -ClassName MSFT_MpComputerStatus -ErrorAction Stop | Select-Object -First 1 } catch { }
        $O.Defender = [ordered]@{
            RealTime   = if ($st) { [bool](Get-Prop $st 'RealTimeProtectionEnabled' $false) } else { $null }
            Paths      = @(Get-Prop $mp 'ExclusionPath' @() | Where-Object { $_ -and $_ -notmatch '^N/A' })
            Processes  = @(Get-Prop $mp 'ExclusionProcess' @() | Where-Object { $_ -and $_ -notmatch '^N/A' })
            Extensions = @(Get-Prop $mp 'ExclusionExtension' @() | Where-Object { $_ -and $_ -notmatch '^N/A' })
            Hidden     = (@(Get-Prop $mp 'ExclusionPath' @()) -match '^N/A').Count -gt 0
        }
    } | Out-Null
    Invoke-Section $Snap 'Os' 'local_admins' {
        $g = Get-CimInstance -CimSession $cs -ClassName Win32_Group -Filter "LocalAccount = TRUE AND SID = 'S-1-5-32-544'" -ErrorAction Stop | Select-Object -First 1
        $O.LocalAdmins = @(Get-CimAssociatedInstance -CimSession $cs -InputObject $g -Association Win32_GroupUser -ErrorAction Stop | ForEach-Object { "$($_.Domain)\$($_.Name)" } | Sort-Object -Unique)
    } | Out-Null
    Invoke-Section $Snap 'Os' 'disks' {
        $O.Disks = @(Get-CimInstance -CimSession $cs -Namespace 'root/Microsoft/Windows/Storage' -ClassName MSFT_Disk -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ Number = $_.Number; Name = $_.FriendlyName; SizeMB = [int64]($_.Size / 1MB); Offline = [bool]$_.IsOffline; ReadOnly = [bool]$_.IsReadOnly; Health = $(switch ([int]$_.HealthStatus) { 0 { 'Healthy' } 1 { 'Warning' } 2 { 'Unhealthy' } default { 'Unknown' } }); Partition = [string]$_.PartitionStyle } })
        $O.PageFiles = @(Get-CimInstance -CimSession $cs -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue | ForEach-Object { "$($_.Name) ($($_.AllocatedBaseSize) MB)" })
    } | Out-Null
    Invoke-Section $Snap 'Os' 'cluster' {
        $O.ClusterName = Get-RegString $cs 'Cluster' 'ClusterName'
    } | Out-Null
    Invoke-Section $Snap 'Os' 'ps_modules' {
        $mods = New-List
        foreach ($dir in @('\program files\windowspowershell\modules\', '\program files\powershell\modules\')) {
            $q = "Drive='C:' AND Path='" + $dir.Replace('\', '\\') + "'"
            foreach ($d in @(Get-CimInstance -CimSession $cs -ClassName Win32_Directory -Filter $q -Property Name -ErrorAction SilentlyContinue)) {
                $n = Get-WinLeaf $d.Name
                if (-not $mods.Contains($n)) { Add-To $mods $n }
            }
        }
        $O.PsModules = @($mods | Sort-Object)
    } | Out-Null
}

function Get-SpnInfo {
    # Read-only LDAP search (DirectorySearcher) for MSSQLSvc SPNs of a host, its FQDN and listener names.
    param([string[]]$Names)
    $names2 = @($Names | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object -Unique)
    if (-not $names2.Count) { return @() }
    $esc = { param($v) ($v -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29') }
    $filter = '(|' + (($names2 | ForEach-Object { '(servicePrincipalName=MSSQLSvc/' + (& $esc $_) + '*)' }) -join '') + ')'
    $ds = New-Object System.DirectoryServices.DirectorySearcher
    $ds.Filter = $filter; $ds.PageSize = 200; $ds.ClientTimeout = [TimeSpan]::FromSeconds(20)
    [void]$ds.PropertiesToLoad.Add('samaccountname'); [void]$ds.PropertiesToLoad.Add('serviceprincipalname')
    $out = New-List
    foreach ($r in $ds.FindAll()) {
        $acct = [string]($r.Properties['samaccountname'] | Select-Object -First 1)
        foreach ($spn in $r.Properties['serviceprincipalname']) {
            $sp = [string]$spn
            if ($sp -notmatch '^(?i)MSSQLSvc/([^:/]+)') { continue }
            if ($names2 -contains $Matches[1].ToLowerInvariant()) { Add-To $out ([pscustomobject]@{ Spn = $sp; Account = $acct }) }
        }
    }
    return $out.ToArray()
}

function Get-DnsInfo {
    param([string]$Name)
    $o = [ordered]@{ Name = $Name; Ips = @(); Reverse = @(); Error = $null }
    try {
        $o.Ips = @([System.Net.Dns]::GetHostAddresses($Name) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.ToString() })
        foreach ($ip in $o.Ips) { try { $o.Reverse += [System.Net.Dns]::GetHostEntry($ip).HostName } catch { $o.Reverse += "(no PTR for $ip)" } }
    } catch { $o.Error = $_.Exception.Message }
    return $o
}

function Test-OsPaths {
    # existence check through CIM_LogicalFile (local paths of that host only). $null = could not check.
    param($Snap, [string[]]$Paths)
    $res = [ordered]@{}
    if (-not $Snap.OsOk -or -not $Snap.CimSession) { foreach ($p in $Paths) { $res[$p] = $null }; return $res }
    foreach ($p in ($Paths | Where-Object { $_ } | Sort-Object -Unique)) {
        if ($p -notmatch '^[A-Za-z]:\\') { $res[$p] = $null; continue }
        $norm = $p.TrimEnd('\'); if ($norm.Length -eq 2) { $norm += '\' }
        try {
            $f = "Name='" + $norm.Replace('\', '\\').Replace("'", "\'") + "'"
            $hit = Get-CimInstance -CimSession $Snap.CimSession -ClassName CIM_LogicalFile -Filter $f -ErrorAction Stop | Select-Object -First 1
            $res[$p] = [bool]$hit
            if ($hit) { $Snap.PathInfo[$p] = [pscustomobject]@{ Compressed = [bool](Get-Prop $hit 'Compressed' $false); Encrypted = [bool](Get-Prop $hit 'Encrypted' $false); Version = [string](Get-Prop $hit 'Version' ''); IsDir = ([string]$hit.CimClass.CimClassName -eq 'Win32_Directory') } }
        } catch { $res[$p] = $null }
    }
    return $res
}
#endregion

#region ---------------------------------------------------------------- snapshots, paths, orchestration
function New-Snapshot {
    param([string]$Name, [string]$Role)
    return [ordered]@{
        Name = $Name; Role = $Role; Computer = $null; SqlOk = $false; OsOk = $false; CimSession = $null
        Sql  = [ordered]@{ Sections = [ordered]@{}; Instance = [ordered]@{}; Config = [ordered]@{}; Net = [ordered]@{}; Services = @(); TraceFlags = @()
            Databases = @(); Files = @(); DbFeatures = @(); Assemblies = @(); SkuFeatures = @(); LinkedRefs = @(); CrossDbRefs = @(); CodeUse = @(); TrustedAssemblies = @()
            LinkedServers = @(); Credentials = @(); Logins = @(); Jobs = @(); JobSteps = @(); Proxies = @(); Operators = @(); Alerts = @(); Mail = @()
            SsisProjects = @(); SsisEnvironments = @(); SsisMsdb = @(); ServerTriggers = @(); Endpoints = @(); StartupProcs = @(); SysDbObjects = @(); XProcs = @()
            Audits = @(); AuditSpecs = @(); XeSessions = @(); ResourceGovernor = $null; CryptoProviders = @(); Certificates = @(); AsymKeys = @(); Tde = @(); Messages = @(); Ag = @()
            AgentProps = [ordered]@{}; JobSignatures = @{}; ConfigPending = [ordered]@{}; DbHealth = @(); LogShipping = @(); BackupCerts = @(); AgStates = @(); Schedules = @(); ProxySubsystems = @(); ProxyLogins = @(); MailAccounts = @(); MailPrincipals = @(); MailConfig = @(); BackupDevices = @(); MaintPlans = @(); JobCategories = @() }
        Os   = [ordered]@{ Sections = [ordered]@{}; Info = $null; Volumes = @(); PowerPlan = $null; Services = @(); Features = $null; Software = $null; DotNet = $null
            OdbcDrivers = @(); Dsns = @(); OleDb = [ordered]@{}; Aliases = @(); Tls = $null; ClusterName = $null; PsModules = $null
            AgentReg = $null; ProviderOptions = $null; Certificates = $null; SqlCert = $null; Firewall = $null
            Patch = $null; PendingReboot = $null; Fips = $null; Msdtc = $null; SqlInstances = $null; EnvVars = $null; ScheduledTasks = $null; Defender = $null; LocalAdmins = $null; Disks = $null; PageFiles = $null }
        PathChecks = [ordered]@{}; PathInfo = @{}; Spns = $null; Dns = $null
    }
}

function Get-HostFromSqlName {
    param([string]$Name)
    $h = $Name -replace '^(?i)tcp:', ''
    $h = ($h -split '[\\,]')[0]
    if ($h -in @('.', '(local)', 'localhost')) { return $env:COMPUTERNAME }
    return $h
}

function Convert-MappedPath {
    param([string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return $Path }
    foreach ($m in $PathMap) {
        $pair = $m -split '=', 2
        if ($pair.Count -ne 2) { continue }
        if ($Path.StartsWith($pair[0], [System.StringComparison]::OrdinalIgnoreCase)) { return $pair[1] + $Path.Substring($pair[0].Length) }
    }
    return $Path
}

function Get-RequiredPaths {
    # everything the reference depends on that must exist on a target's local disks
    param($R, $T)
    $req = New-List
    $tBinn = Get-Prop $T.Sql.Instance 'BinnPath'
    $userDbIds = @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Id -gt 4 } | ForEach-Object { $_.Id })
    $seen = @{}
    foreach ($f in $R.Sql.Files) {
        if ($userDbIds -notcontains $f.DbId) { continue }
        if ($f.Type -eq 'FILESTREAM') { $dir = [string]$f.Path } else { $dir = Get-WinParent ([string]$f.Path) }
        if (-not $dir -or $seen.ContainsKey($dir)) { continue }
        $seen[$dir] = 1
        Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath $dir); RefPath = $dir; Kind = 'Database folder'; Context = (@((ConvertTo-Arr $R.Sql.Files) | Where-Object { $_.Path -like "$dir\*" -or $_.Path -eq $dir } | Select-Object -ExpandProperty Db -Unique) -join ', ') })
    }
    foreach ($s in $R.Sql.JobSteps) {
        foreach ($p in $s.Paths) { Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath $p); RefPath = $p; Kind = 'Job step file'; Context = "$($s.Job) / step $($s.StepId) ($($s.Subsystem))" }) }
        if ($s.OutputDir) { Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath $s.OutputDir); RefPath = $s.OutputDir; Kind = 'Job output folder'; Context = "$($s.Job) / step $($s.StepId)" }) }
    }
    foreach ($a in $R.Sql.Audits) { if ($a.Path) { Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath ([string]$a.Path).TrimEnd('\')); RefPath = $a.Path; Kind = 'Audit folder'; Context = $a.Name }) } }
    foreach ($x in $R.Sql.XeSessions) {
        if ($x.FilePath -and ([string]$x.FilePath) -match '^[A-Za-z]:\\') { $d = Get-WinParent ([string]$x.FilePath); Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath $d); RefPath = $d; Kind = 'XE target folder'; Context = $x.Name }) }
    }
    foreach ($x in $R.Sql.XProcs) {
        $dll = [string]$x.Dll
        if ($dll -match '^[A-Za-z]:\\') { $p = Convert-MappedPath $dll } elseif ($tBinn) { $p = ([string]$tBinn).TrimEnd('\') + '\' + $dll } else { $p = $null }
        if ($p) { Add-To $req ([pscustomobject]@{ Path = $p; RefPath = $dll; Kind = 'Extended proc DLL'; Context = $x.Name }) }
    }
    foreach ($c in $R.Sql.CryptoProviders) { if ($c.Dll) { Add-To $req ([pscustomobject]@{ Path = (Convert-MappedPath $c.Dll); RefPath = $c.Dll; Kind = 'EKM provider DLL'; Context = $c.Name }) } }
    foreach ($k in @('DefaultDataPath', 'DefaultLogPath', 'DefaultBackupPath')) {
        $v = Get-Prop $T.Sql.Instance $k
        if ($v) { Add-To $req ([pscustomobject]@{ Path = ([string]$v).TrimEnd('\'); RefPath = ''; Kind = 'Target default folder'; Context = $k }) }
    }
    $seenReq = @{}; $out = New-List
    foreach ($q in $req) { $k = "$($q.Kind)|$($q.Path)".ToLowerInvariant(); if (-not $seenReq.ContainsKey($k)) { $seenReq[$k] = 1; Add-To $out $q } }
    return $out.ToArray()
}

function Invoke-Collection {
    param($Snapshots)
    $R = $Snapshots[0]
    $total = $Snapshots.Count; $i = 0
    foreach ($s in $Snapshots) {
        $i++
        Write-Log ("[{0}/{1}] {2} ({3})" -f $i, $total, $s.Name, $s.Role)
        if ($Skip -notcontains 'Sql') { Write-Log '  SQL Server catalog...'; Get-SqlSnapshot -Snap $s }
        $map = $ComputerMap | Where-Object { ($_ -split '=', 2)[0] -eq $s.Name } | Select-Object -First 1
        if ($map) { $s.Computer = ($map -split '=', 2)[1] }
        elseif ($s.SqlOk -and (Get-Prop $s.Sql.Instance 'PhysicalNode')) { $s.Computer = [string]$s.Sql.Instance.PhysicalNode }
        else { $s.Computer = Get-HostFromSqlName $s.Name }
        if ($Skip -notcontains 'Os') {
            Write-Log "  Windows facets on $($s.Computer)..."
            $prov = @()
            if ($R.SqlOk) { $prov = @((ConvertTo-Arr $R.Sql.LinkedServers) | ForEach-Object { $_.Provider } | Where-Object { $_ }) }
            Get-OsSnapshot -Snap $s -ExtraProviders $prov
        }
    }
    foreach ($s in $Snapshots) {
        $s.Dns = Get-DnsInfo -Name ([string]$s.Computer)
        if ($Skip -notcontains 'Spn') {
            $fq = if ($s.OsOk -and $s.Os.Info -and $s.Os.Info.Domain) { "$(([string]$s.Computer -split '\.')[0]).$($s.Os.Info.Domain)" } else { $null }
            $names = @(([string]$s.Computer -split '\.')[0], $fq) + @(ConvertTo-Arr $s.Sql.Ag | Where-Object { $_.Listener } | ForEach-Object { $_.Listener; if ($s.Os.Info -and $s.Os.Info.Domain) { "$($_.Listener).$($s.Os.Info.Domain)" } })
            try { $s.Spns = @(Get-SpnInfo -Names $names); $s.Os.Sections['spn'] = 'ok' } catch { $s.Os.Sections['spn'] = 'error: ' + ($_.Exception.Message -replace '\s+', ' '); Write-Log "  SPN lookup for $($s.Computer) failed (domain-joined runner + AD read needed): $($_.Exception.Message)" 'WARN' }
        }
    }
    if ($Skip -notcontains 'Paths' -and $Skip -notcontains 'Os') {
        foreach ($s in $Snapshots) {
            if (-not $s.OsOk) { continue }
            $req = if ($s -eq $R) { Get-RequiredPaths -R $R -T $R | Where-Object { $_.Kind -ne 'Target default folder' } } else { Get-RequiredPaths -R $R -T $s }
            if ($s -eq $R) { foreach ($q in $req) { $q.Path = $q.RefPath } }
            try {
                Write-Log ("  Checking {0} referenced paths on {1}..." -f @($req).Count, $s.Computer)
                $chk = Test-OsPaths -Snap $s -Paths @($req | ForEach-Object { $_.Path })
                foreach ($k in $chk.Keys) { $s.PathChecks[$k] = $chk[$k] }
            } catch { $s.Os.Sections['paths'] = 'error: ' + $_.Exception.Message; Write-Log "  path checks on $($s.Computer) failed: $($_.Exception.Message)" 'WARN' }
        }
    }
    foreach ($s in $Snapshots) { if ($s.CimSession) { try { Remove-CimSession $s.CimSession } catch { } ; $s.CimSession = $null } }
}
#endregion

#region ---------------------------------------------------------------- comparison engine
$script:Findings = New-List
$script:Checks = @{}

function Add-Check { param($T, [string]$Category, [int]$N = 1) $k = "$($T.Name)|$Category"; if (-not $script:Checks.ContainsKey($k)) { $script:Checks[$k] = 0 }; $script:Checks[$k] += $N }

function Add-Finding {
    param($T, [string]$Category, [string]$Item, $Ref, $Tgt, [string]$Sev, [switch]$Hard, [string]$Action, [string]$Fix = '', [switch]$Install, [string]$Detail = '')
    $eff = $Sev
    if ($T.Role -eq 'Asynchronous' -and -not $Hard) {
        switch ($Sev) { 'High' { $eff = 'Medium' } 'Medium' { $eff = 'Low' } 'Low' { $eff = 'Info' } }
    }
    Add-Check $T $Category
    Add-To $script:Findings ([pscustomobject]@{
        Target = $T.Name; Role = $T.Role; Category = $Category; Severity = $eff; BaseSeverity = $Sev; Item = $Item
        Reference = [string]$Ref; TargetValue = [string]$Tgt; Action = $Action; Fix = $Fix; Install = [bool]$Install; Detail = $Detail
    })
}

function Test-Sec {
    # true when the section was collected on both sides; otherwise records the gap once
    param($R, $T, [string]$Area, [string]$Name, [string]$Category, [string]$Label)
    $okR = ($Area -eq 'Sql' -and $R.SqlOk) -or ($Area -eq 'Os' -and $R.OsOk)
    $okT = ($Area -eq 'Sql' -and $T.SqlOk) -or ($Area -eq 'Os' -and $T.OsOk)
    if (-not $okR -or -not $okT) { return $false }   # whole-area gaps are reported once elsewhere
    $sr = [string]$R[$Area].Sections[$Name]; $st = [string]$T[$Area].Sections[$Name]
    if ($sr -like 'ok*' -and $st -like 'ok*') { return $true }
    if ($sr -notlike 'ok*') { Add-Finding $T 'Collection gaps' "$Label (reference)" $sr '' 'Info' -Action "Not collected on the reference; $Label was not compared. Grant the permission in the message and re-run." }
    if ($st -notlike 'ok*') { Add-Finding $T 'Collection gaps' $Label '' $st 'Medium' -Action "Could not verify $Label on this target. Grant the permission in the message and re-run." }
    return $false
}

function Compare-NamedSet {
    param($T, [string]$Category, [string]$Label, $RefItems, $TgtItems, [scriptblock]$Key, [string]$Sev, [switch]$Hard, [string]$Action, [string]$Fix = '', [scriptblock]$Show = $null, [scriptblock]$Diff = $null, [string]$DiffSev = 'Medium', [switch]$ReportExtra, [int]$GroupAbove = 0)
    $tIdx = @{}
    foreach ($ti in (ConvertTo-Arr $TgtItems)) { $tIdx[[string](& $Key $ti)] = $ti }
    $rKeys = @{}
    $missing = @(ConvertTo-Arr $RefItems | Where-Object { -not $tIdx.ContainsKey([string](& $Key $_)) })
    $grouped = ($GroupAbove -gt 0 -and $missing.Count -gt $GroupAbove)
    if ($grouped) {
        $names = @($missing | ForEach-Object { [string](& $Key $_) })
        Add-Finding $T $Category "$Label`s missing: $($missing.Count)" ("$($missing.Count) on reference") 'missing' $Sev -Hard:$Hard -Action $Action -Fix $Fix -Detail (($names | Select-Object -First 40) -join '; ')
        Add-Check $T $Category ($missing.Count - 1)
    }
    foreach ($ri in (ConvertTo-Arr $RefItems)) {
        $k = [string](& $Key $ri); $rKeys[$k] = 1
        $disp = if ($Show) { & $Show $ri } else { $k }
        if (-not $tIdx.ContainsKey($k)) {
            if (-not $grouped) { Add-Finding $T $Category "$Label`: $k" $disp 'missing' $Sev -Hard:$Hard -Action $Action -Fix $Fix }
        } elseif ($Diff) {
            $d = & $Diff $ri $tIdx[$k]
            if ($d) { Add-Finding $T $Category "$Label`: $k" $d[0] $d[1] $DiffSev -Action $(if ($d.Count -gt 2) { $d[2] } else { "Align $Label settings." }) -Fix $Fix }
            else { Add-Check $T $Category }
        } else { Add-Check $T $Category }
    }
    if ($ReportExtra) {
        foreach ($k in $tIdx.Keys) { if (-not $rKeys.ContainsKey($k)) { Add-Finding $T $Category "$Label`: $k" '' 'only on target' 'Info' -Action 'Exists only on the target. Review; nothing is removed by this tool.' } }
    }
}

# ---------- software classification
$script:SwNoise = '(?i)^(Update for|Security Update|Hotfix|Definition Update|Service Pack)|\bKB\d{6,}\b|Language Pack|Microsoft Edge|WebView2|Windows Driver Package|^Intel\(R\)|^Realtek|^NVIDIA|Update Health|PC Health Check|Teams Machine-Wide|OneDrive|Microsoft Update Health|Windows Subsystem'
$script:SwRules = @(
    @{ Cat = 'Drivers & providers'; Kind = 'Data driver'; Sev = 'High'; Rx = '(?i)ODBC Driver|OLE ?DB|Native Client|Oracle.*(Client|ODAC|Data Provider|Data Access|Instant|Provider)|IBM.*(DB2|Data Server)|Access Database Engine|Teradata|SAP.*(Connector|NCo|HANA)|MySQL Connector|psqlODBC|PostgreSQL.*ODBC|Npgsql|Snowflake|Simba|Informix|Sybase|DataDirect|CData|Devart|Microsoft Connector for' },
    @{ Cat = 'Installed software'; Kind = 'Runtime'; Sev = 'High'; Rx = '(?i)Visual C\+\+.*(Redistributable|Runtime)|\.NET (Core|Runtime|Desktop Runtime|Framework)|ASP\.NET Core|\bJava\b|\bJRE\b|\bJDK\b|OpenJDK|Temurin|Corretto|\bPython\b|R for Windows|Microsoft R (Open|Client)|Strawberry Perl|Node\.js' },
    @{ Cat = 'Installed software'; Kind = 'SQL component'; Sev = 'High'; Rx = '(?i)Integration Services|Reporting Services|Analysis Services|Power BI Report Server|Data Quality|Master Data Services|PolyBase|Machine Learning|Data-Tier|DacFx|SqlPackage|Command Line Utilities|\bsqlcmd\b|System CLR Types|Shared Management Objects|\bSMO\b|PowerShell Extensions|Microsoft Connector|Attunity|SSIS' },
    @{ Cat = 'Installed software'; Kind = 'SQL setup/tools'; Sev = 'Low'; Rx = '(?i)Microsoft SQL Server|SQL Server Management Studio|Azure Data Studio|SQL Server Browser|VSS Writer|Data Migration Assistant' },
    @{ Cat = 'Installed software'; Kind = 'Agent'; Sev = 'Medium'; Rx = '(?i)Commvault|Veeam|NetBackup|Veritas|Rubrik|Cohesity|Avamar|Networker|Data Protector|Spectrum Protect|Druva|Redgate|SQL Monitor|SQL Backup|Idera|SQL Diagnostic|Spotlight|LiteSpeed|SolarWinds|SentryOne|SQL Sentry|Datadog|Dynatrace|New Relic|AppDynamics|Splunk|Elastic Agent|Zabbix|Nagios|NSClient|Operations Manager|Monitoring Agent|Azure Arc|Connected Machine|Defender for|CrowdStrike|Falcon|SentinelOne|Symantec|McAfee|Trellix|Trend Micro|Sophos|Carbon Black|Cylance|Tanium|Qualys|Rapid7|Tenable|Nessus|BigFix|Configuration Manager|Ivanti|CyberArk|Imperva|Guardium|DbProtect|Varonis' },
    @{ Cat = 'Installed software'; Kind = 'Tooling'; Sev = 'Medium'; Rx = '(?i)Crystal Reports|BusinessObjects|Informatica|Tableau|Qlik|Autosys|Control-M|Tidal|ActiveBatch|JAMS|Automic|UC4|WinSCP|7-Zip|GnuPG|Gpg4win|\bPGP\b|FileZilla|Robo-FTP|MOVEit|Globalscape|\bcURL\b|\bGit\b|PowerShell 7|PowerShell \d' },
    @{ Cat = 'Installed software'; Kind = 'Platform agent'; Sev = 'Info'; Rx = '(?i)VMware Tools|Hyper-V|Citrix|Azure VM Agent|Windows Azure VM Agent|Amazon SSM|EC2Launch|AWS PV|Google Compute|Nutanix|XenServer' }
)
function Get-SwClass {
    param([string]$Name)
    foreach ($r in $script:SwRules) { if ($Name -match $r.Rx) { return $r } }
    return @{ Cat = 'Installed software'; Kind = 'Other'; Sev = 'Low' }
}
function Get-SwKey {
    param([string]$Name, [string]$View)
    $arch = $View
    if ($Name -match '(?i)\bx64\b|64-bit|amd64|\(64') { $arch = 'x64' } elseif ($Name -match '(?i)\bx86\b|32-bit|\(32') { $arch = 'x86' }
    $n = $Name -replace '(?i)\s*\((x64|x86|64-bit|32-bit|amd64)\)', '' -replace '(?i)\b(x64|x86|amd64)\b', ''
    $n = $n -replace '\s+[-\u2013]\s+\d+(\.\d+)+.*$', ''
    $n = $n -replace '(?i)\bUpdate \d+\b', ''
    $n = [regex]::Replace($n, '\b(\d+\.\d+)\.\d+(\.\d+)*\b', '$1')
    $n = ($n -replace '\s{2,}', ' ').Trim([char[]]@([char]32, [char]45, [char]0x2013))
    return @{ Key = ($n.ToLowerInvariant() + '|' + $arch); Clean = $n; Arch = $arch }
}

# ---------- service family mapping
function Get-SvcFamily {
    param([string]$Name)
    switch -Regex ($Name) {
        '^(?i)(MSSQLSERVER|MSSQL\$.+)$' { return 'Database Engine' }
        '^(?i)(SQLSERVERAGENT|SQLAgent\$.+)$' { return 'SQL Server Agent' }
        '^(?i)SQLBrowser$' { return 'SQL Browser' }
        '^(?i)MsDtsServer\d+$' { return 'Integration Services' }
        '^(?i)SSISScaleOut' { return 'SSIS Scale Out' }
        '^(?i)(ReportServer(\$.+)?|SQLServerReportingServices|PowerBIReportServer)$' { return 'Reporting Services' }
        '^(?i)(MSSQLServerOLAPService|MSOLAP\$.+)$' { return 'Analysis Services' }
        '^(?i)MSSQLFDLauncher' { return 'Full-Text Filter Daemon' }
        '^(?i)MSSQLLaunchpad' { return 'Launchpad (ML Services)' }
        '^(?i)SQLPBENGINE|^SQLPBDMS' { return 'PolyBase' }
        '^(?i)SQLWriter$' { return 'SQL VSS Writer' }
        '^(?i)MSDTC$' { return 'Distributed Transaction Coordinator' }
        '^(?i)ClusSvc$' { return 'Cluster Service' }
        '^(?i)SQLTELEMETRY' { return 'SQL CEIP' }
        default { return $null }
    }
}

function Get-VolumeFor {
    param($Volumes, [string]$Path)
    $best = $null
    foreach ($v in $Volumes) { if ($Path.StartsWith([string]$v.Name, [System.StringComparison]::OrdinalIgnoreCase) -and ($null -eq $best -or ([string]$v.Name).Length -gt ([string]$best.Name).Length)) { $best = $v } }
    return $best
}

# ---------- the rules
$script:ObsoleteTf = @{
    1117 = 'No effect on 2016+ (use ALTER DATABASE ... AUTOGROW_ALL_FILES)'; 1118 = 'No effect on 2016+ (mixed extents off by default)'
    2371 = 'Default behaviour at compatibility level 130+'; 8048 = 'Default behaviour on 2016+'; 1236 = 'Default behaviour on 2016+'
    4199 = 'Use QUERY_OPTIMIZER_HOTFIXES database-scoped option on 2016+'; 9481 = 'Use LEGACY_CARDINALITY_ESTIMATION database-scoped option on 2016+'
    8079 = 'Automatic soft-NUMA is default on 2016+'; 2453 = 'Table variable deferred compilation on compat 150+'
}
$script:ProviderNames = @{
    'MSOLEDBSQL' = 'Microsoft OLE DB Driver for SQL Server (x64)'; 'MSOLEDBSQL19' = 'Microsoft OLE DB Driver 19 for SQL Server (x64)'
    'SQLNCLI11' = 'SQL Server 2012 Native Client (deprecated; prefer MSOLEDBSQL)'; 'SQLNCLI10' = 'SQL Server 2008 Native Client (deprecated; prefer MSOLEDBSQL)'
    'SQLOLEDB' = 'Microsoft OLE DB Provider for SQL Server (Windows built-in, legacy)'; 'MSDASQL' = 'OLE DB Provider for ODBC (Windows built-in) + the ODBC driver/DSN it uses'
    'Microsoft.ACE.OLEDB.12.0' = 'Microsoft Access Database Engine 2010 Redistributable (x64)'; 'Microsoft.ACE.OLEDB.16.0' = 'Microsoft Access Database Engine 2016 Redistributable (x64)'
    'OraOLEDB.Oracle' = 'Oracle Provider for OLE DB (ODAC, x64)'; 'IBMDADB2' = 'IBM Data Server Driver / DB2 client (x64)'; 'DB2OLEDB' = 'Microsoft OLE DB Provider for DB2 (Host Integration Server)'
    'MSOLAP' = 'Microsoft Analysis Services OLE DB Provider (MSOLAP, x64)'; 'Microsoft.Jet.OLEDB.4.0' = 'Jet 4.0 (32-bit only - not usable by 64-bit SQL Server)'
}
function Get-MinCompat { param([int]$Major) if ($Major -ge 12) { return 100 } if ($Major -eq 11) { return 90 } if ($Major -eq 10) { return 80 } return 0 }

# Extended-support end dates (Microsoft lifecycle). Unknown/newer versions return $null.
$script:SqlSupportEnd = @{ 10 = '2019-07-09'; 11 = '2022-07-12'; 12 = '2024-07-09'; 13 = '2026-07-14'; 14 = '2027-10-12'; 15 = '2030-01-08'; 16 = '2033-01-11' }
$script:OsSupportEnd = @{ 7601 = @('Windows Server 2008 R2', '2020-01-14'); 9200 = @('Windows Server 2012', '2023-10-10'); 9600 = @('Windows Server 2012 R2', '2023-10-10'); 14393 = @('Windows Server 2016', '2027-01-12'); 17763 = @('Windows Server 2019', '2029-01-09'); 20348 = @('Windows Server 2022', '2031-10-14'); 26100 = @('Windows Server 2025', '2034-11-14') }

function Compare-Target {
    param($R, $T)
    $isMig = $T.Role -eq 'Migration'
    $isHa = -not $isMig
    $RI = $R.Sql.Instance; $TI = $T.Sql.Instance
    $rMaj = [int](Get-Prop $RI 'Major' 0); $tMaj = [int](Get-Prop $TI 'Major' 0)

    # ---------------- whole-area gaps
    if ($Skip -notcontains 'Sql' -and -not $T.SqlOk) { Add-Finding $T 'Collection gaps' 'SQL Server connection' '' ([string]$T.Sql.Sections['connect']) 'High' -Hard -Action 'Could not connect to SQL Server; no SQL facets were verified. Check name/port, firewall and permissions (VIEW SERVER STATE, VIEW ANY DEFINITION).' }
    if ($Skip -notcontains 'Os' -and -not $T.OsOk) { Add-Finding $T 'Collection gaps' "Windows (CIM) on $($T.Computer)" '' ([string]$T.Os.Sections['connect']) 'High' -Hard -Action 'Software, drivers, services, volumes and paths were NOT verified. Enable WinRM (or use -CimProtocol Dcom) and grant WMI read / local admin.' }

    # ================= Platform & OS
    if (Test-Sec $R $T 'Os' 'os' 'Platform & OS' 'Operating system') {
        $c = 'Platform & OS'; $ro = $R.Os.Info; $to = $T.Os.Info
        $rb = [int]("0" + $ro.Build); $tb = [int]("0" + $to.Build)
        if ($tb -lt $rb) { Add-Finding $T $c 'Windows version' "$($ro.Caption) (build $($ro.Build))" "$($to.Caption) (build $($to.Build))" 'Medium' -Action 'Target runs an older Windows build than the reference. Confirm the SQL Server version and every driver/agent are supported on it.' }
        elseif ($isHa -and $tb -ne $rb) { Add-Finding $T $c 'Windows version' "$($ro.Caption) (build $($ro.Build))" "$($to.Caption) (build $($to.Build))" 'Medium' -Action 'Cluster nodes should run the same Windows version (mixed only during a rolling OS upgrade).' }
        else { Add-Check $T $c }
        if ($ro.Arch -ne $to.Arch) { Add-Finding $T $c 'OS architecture' $ro.Arch $to.Arch 'High' -Action 'Architectures differ; 32-bit drivers/providers will not load in 64-bit SQL Server.' } else { Add-Check $T $c }
        if ([string]$ro.Domain -ne [string]$to.Domain) { Add-Finding $T $c 'Domain' $ro.Domain $to.Domain $(if ($isHa) { 'High' } else { 'Medium' }) -Action 'Different domain: Windows logins, service accounts, SPNs/Kerberos delegation and file-share access need a trust or new accounts.' } else { Add-Check $T $c }
        if ([string]$ro.TimeZone -ne [string]$to.TimeZone) { Add-Finding $T $c 'Time zone' $ro.TimeZone $to.TimeZone 'Medium' -Action 'Agent schedules run in local time and GETDATE()/SYSDATETIME() values shift. Align the time zone or review schedules and code.' } else { Add-Check $T $c }
        if ([string]$ro.Locale -ne [string]$to.Locale) { Add-Finding $T $c 'System locale' $ro.Locale $to.Locale 'Low' -Action 'Different system locale can change date/number formats in CmdExec/PowerShell job output and file imports.' } else { Add-Check $T $c }
        if ($T.Os.Sections['power'] -like 'ok*' -and $T.Os.PowerPlan -and $T.Os.PowerPlan -notmatch '(?i)high performance|ultimate') { Add-Finding $T $c 'Power plan (target baseline)' $R.Os.PowerPlan $T.Os.PowerPlan 'Low' -Action 'Use the High performance power plan on SQL Server hosts.' } else { Add-Check $T $c }
        if ($isHa) {
            if (-not $T.Os.ClusterName) { Add-Finding $T $c 'Windows cluster membership' $R.Os.ClusterName 'not a cluster node' 'High' -Action 'AG replicas (except clusterless/read-scale AGs) must be nodes of the same WSFC.' }
            elseif ($R.Os.ClusterName -and $R.Os.ClusterName -ne $T.Os.ClusterName) { Add-Finding $T $c 'Windows cluster membership' $R.Os.ClusterName $T.Os.ClusterName 'High' -Action 'Replica is in a different cluster (distributed AG?). Confirm the design.' }
            else { Add-Check $T $c }
        }
    }

    # ================= Hardware & storage
    if ($R.OsOk -and $T.OsOk -and $R.Os.Sections['os'] -like 'ok*' -and $T.Os.Sections['os'] -like 'ok*') {
        $c = 'Hardware & storage'; $ro = $R.Os.Info; $to = $T.Os.Info
        if ([int]$to.Logical -lt [int]$ro.Logical) { Add-Finding $T $c 'Logical processors' $ro.Logical $to.Logical $(if ($T.Role -eq 'Synchronous') { 'High' } else { 'Medium' }) -Action 'Target has fewer logical CPUs. Validate capacity (and MAXDOP) for the workload after cut-over / failover.' } else { Add-Check $T $c }
        if ([int64]$to.MemoryMB -lt [int64]$ro.MemoryMB * 0.95) { Add-Finding $T $c 'Physical memory' (Format-MB $ro.MemoryMB) (Format-MB $to.MemoryMB) $(if ($T.Role -eq 'Synchronous') { 'High' } else { 'Medium' }) -Action 'Target has less RAM. Re-size max server memory and validate the workload.' } else { Add-Check $T $c }
    }
    if ($T.SqlOk -and $T.Sql.Sections['config'] -like 'ok*') {
        $c = 'Hardware & storage'
        $msm = [int64]("0" + (Get-Prop $T.Sql.Config 'max server memory (MB)' '0'))
        $tmem = [int64](Get-Prop $TI 'MemoryMB' 0); if ($T.OsOk -and $T.Os.Info) { $tmem = [int64]$T.Os.Info.MemoryMB }
        if ($msm -ge 2147483647) { Add-Finding $T $c 'max server memory (target baseline)' (Get-Prop $R.Sql.Config 'max server memory (MB)') 'unlimited (default)' 'Medium' -Action ('Set max server memory; leave memory for the OS and other services (roughly {0} MB on this host).' -f [int]($tmem * 0.85)) }
        elseif ($tmem -gt 0 -and $msm -gt $tmem) { Add-Finding $T $c 'max server memory (target baseline)' '' "$msm MB > $tmem MB physical" 'High' -Action 'max server memory exceeds physical RAM on the target (copied from a bigger box?). Lower it.' }
        else { Add-Check $T $c }
    }
    if ($R.SqlOk -and $T.OsOk -and $T.Os.Sections['volumes'] -like 'ok*') {
        $c = 'Hardware & storage'
        $tDbs = @((ConvertTo-Arr $T.Sql.Databases) | ForEach-Object { $_.Name })
        $need = @{}
        foreach ($f in $R.Sql.Files) {
            if ($f.DbId -le 4 -or $tDbs -contains $f.Db) { continue }
            $p = Convert-MappedPath ([string]$f.Path)
            $v = Get-VolumeFor $T.Os.Volumes $p
            $key = if ($v) { [string]$v.Name } else { ($p.Substring(0, [Math]::Min(3, $p.Length))) + ' (no such volume)' }
            if (-not $need.ContainsKey($key)) { $need[$key] = @{ MB = [int64]0; Vol = $v; Dbs = New-List } }
            $need[$key].MB += [int64]$f.SizeMB
            if (-not $need[$key].Dbs.Contains($f.Db)) { Add-To $need[$key].Dbs $f.Db }
        }
        foreach ($k in $need.Keys) {
            $n = $need[$k]; $dbList = ($n.Dbs | Select-Object -First 8) -join ', '
            if (-not $n.Vol) { Add-Finding $T $c "Volume $k" (Format-MB $n.MB) 'volume missing' $(if ($isHa) { 'Blocker' } else { 'High' }) -Hard -Action "Databases on the reference use $k but the target has no such volume. Create it, or pass -PathMap and plan RESTORE ... WITH MOVE." -Detail $dbList }
            elseif ([int64]$n.Vol.FreeMB -lt $n.MB) { Add-Finding $T $c "Free space on $k" ("needs " + (Format-MB $n.MB)) ("free " + (Format-MB $n.Vol.FreeMB)) 'Blocker' -Hard -Action 'Not enough free space for the databases that will land on this volume (data + log at current allocated size).' -Detail $dbList }
            elseif ([int64]$n.Vol.FreeMB -lt [int64]($n.MB * 1.25)) { Add-Finding $T $c "Free space on $k" ("needs " + (Format-MB $n.MB)) ("free " + (Format-MB $n.Vol.FreeMB)) 'High' -Action 'Less than 25% headroom after the databases land. Plan growth.' -Detail $dbList }
            else { Add-Check $T $c }
            if ($n.Vol -and $n.Vol.BlockSize -and [int]$n.Vol.BlockSize -ne 65536) { Add-Finding $T $c "Allocation unit on $k" '64 KB recommended' ('{0} KB' -f ([int]$n.Vol.BlockSize / 1024)) 'Low' -Action 'Format SQL data/log volumes with 64 KB allocation units (requires reformat; do before the data lands).' }
        }
    }

    # ================= Installed software
    if ($Skip -notcontains 'Software' -and (Test-Sec $R $T 'Os' 'software' 'Installed software' 'Installed software')) {
        $tIdx = @{}
        foreach ($s in $T.Os.Software) { if ($s.Name -match $script:SwNoise) { continue }; $k = Get-SwKey $s.Name $s.View; $tIdx[$k.Key] = $s }
        $done = @{}
        foreach ($s in $R.Os.Software) {
            if ($s.Name -match $script:SwNoise) { continue }
            $k = Get-SwKey $s.Name $s.View
            if ($done.ContainsKey($k.Key)) { continue }; $done[$k.Key] = 1
            $cl = Get-SwClass $s.Name
            $sev = $cl.Sev
            if ($cl.Kind -in @('SQL setup/tools', 'SQL component') -and $s.Name -match '20\d\d' -and $rMaj -ne $tMaj) { $sev = 'Info' }
            if ($tIdx.ContainsKey($k.Key)) {
                $tv = ConvertTo-VersionSafe $tIdx[$k.Key].Version; $rv = ConvertTo-VersionSafe $s.Version
                if ($rv -and $tv -and $tv -lt $rv -and $sev -ne 'Info') { Add-Finding $T $cl.Cat ("$($cl.Kind): " + $k.Clean + " ($($k.Arch))") $s.Version $tIdx[$k.Key].Version $(if ($sev -eq 'High') { 'Medium' } else { 'Low' }) -Install -Action 'Older version on the target. Upgrade to at least the reference version.' -Detail $s.Publisher }
                else { Add-Check $T $cl.Cat }
            } else {
                Add-Finding $T $cl.Cat ("$($cl.Kind): " + $k.Clean + " ($($k.Arch))") ("$($s.Name) $($s.Version)") 'not installed' $sev -Install:($sev -ne 'Info') -Action $(if ($sev -eq 'Info') { 'Version-specific / platform component; usually replaced by the target equivalent. Confirm.' } else { 'Install on the target (same or newer version) or confirm it is no longer needed.' }) -Detail $s.Publisher
            }
        }
    }

    # ================= Drivers & providers
    $c = 'Drivers & providers'
    if (Test-Sec $R $T 'Os' 'odbc' $c 'ODBC drivers / DSNs') {
        Compare-NamedSet $T $c 'ODBC driver' $R.Os.OdbcDrivers $T.Os.OdbcDrivers { "$($args[0].Name) ($($args[0].Bitness))" } 'High' -Action 'Install this ODBC driver (same bitness) on the target - SSIS, linked servers (MSDASQL), jobs and apps may use it.' -ReportExtra:$false
        foreach ($f in @($script:Findings | Where-Object { $_.Target -eq $T.Name -and $_.Item -like 'ODBC driver:*' })) { $f.Install = $true }
        $usedDsn = @((ConvertTo-Arr $R.Sql.LinkedServers) | Where-Object { $_.Provider -eq 'MSDASQL' } | ForEach-Object { [string]$_.DataSource })
        $tD = @{}; foreach ($d in $T.Os.Dsns) { $tD["$($d.Name)|$($d.Bitness)"] = $d }
        foreach ($d in $R.Os.Dsns) {
            $k = "$($d.Name)|$($d.Bitness)"
            if (-not $tD.ContainsKey($k)) { Add-Finding $T $c "System DSN: $($d.Name) ($($d.Bitness))" "$($d.Driver) -> $($d.Server)" 'missing' $(if ($usedDsn -contains $d.Name) { 'High' } else { 'Medium' }) -Install -Action $(if ($usedDsn -contains $d.Name) { 'Used by a linked server. Create the system DSN (same bitness) on the target.' } else { 'Create the system DSN on the target if apps/jobs/SSIS use it.' }) }
            elseif ([string]$tD[$k].Driver -ne [string]$d.Driver) { Add-Finding $T $c "System DSN: $($d.Name) ($($d.Bitness))" $d.Driver $tD[$k].Driver 'Medium' -Action 'DSN uses a different driver on the target.' }
            else { Add-Check $T $c }
        }
    }
    if (Test-Sec $R $T 'Os' 'oledb' $c 'OLE DB providers') {
        $usedProv = @((ConvertTo-Arr $R.Sql.LinkedServers) | ForEach-Object { [string]$_.Provider } | Where-Object { $_ } | Sort-Object -Unique)
        foreach ($p in $R.Os.OleDb.Keys) {
            $rp = $R.Os.OleDb[$p]; $tp = $T.Os.OleDb[$p]
            if (-not $tp) { continue }
            if ($rp.x64 -and -not $tp.x64) {
                $nm = if ($script:ProviderNames.ContainsKey($p)) { $script:ProviderNames[$p] } else { $p }
                Add-Finding $T $c "OLE DB provider (x64): $p" 'registered' 'not registered' $(if ($usedProv -contains $p) { 'High' } else { 'Medium' }) -Install -Action ("Install: $nm." + $(if ($usedProv -contains $p) { ' A linked server uses it.' } else { '' }))
            } elseif ($rp.x86 -and -not $tp.x86) {
                Add-Finding $T $c "OLE DB provider (x86): $p" 'registered' 'not registered' 'Low' -Install -Action '32-bit provider - only matters for 32-bit SSIS/DTExec or apps on this host.'
            } else { Add-Check $T $c }
        }
    }
    if (Test-Sec $R $T 'Os' 'dotnet' $c '.NET Framework') {
        $rr = [int64]("0" + $R.Os.DotNet.Release); $tr = [int64]("0" + $T.Os.DotNet.Release)
        if ($tr -lt $rr) { Add-Finding $T $c '.NET Framework 4.x' $R.Os.DotNet.Version $(if ($T.Os.DotNet.Version) { $T.Os.DotNet.Version } else { 'not installed' }) 'High' -Install -Action 'Install the same or newer .NET Framework 4.x (SQLCLR, SSIS script tasks, PowerShell and vendor tools depend on it).' } else { Add-Check $T $c }
        if ($R.Os.DotNet.Net35 -and -not $T.Os.DotNet.Net35) { Add-Finding $T $c '.NET Framework 3.5' 'installed' 'not installed' 'Medium' -Install -Action 'Enable the .NET Framework 3.5 feature (older SSIS packages/tools and some setup chains need it).' } else { Add-Check $T $c }
        if ([string]$R.Os.DotNet.StrongCrypto -ne [string]$T.Os.DotNet.StrongCrypto) { Add-Finding $T $c '.NET SchUseStrongCrypto' $R.Os.DotNet.StrongCrypto $T.Os.DotNet.StrongCrypto 'Low' -Action '.NET apps on the host negotiate TLS differently. Align with your TLS policy.' } else { Add-Check $T $c }
    }
    if (Test-Sec $R $T 'Os' 'aliases' $c 'SQL client aliases') {
        Compare-NamedSet $T $c 'SQL client alias' $R.Os.Aliases $T.Os.Aliases { "$($args[0].Name) ($($args[0].Bitness))" } 'Medium' -Action 'Create the alias (cliconfg.exe / SQL Server Configuration Manager) - linked servers, jobs or apps may connect through it.' -Show { "-> $($args[0].Target)" } -Diff { if ($args[0].Target -ne $args[1].Target) { , @($args[0].Target, $args[1].Target, 'Alias points elsewhere on the target. Confirm it is intentional.') } } -DiffSev 'Low'
    }
    if (Test-Sec $R $T 'Os' 'tls' $c 'TLS protocol settings') {
        foreach ($k in $R.Os.Tls.Keys) {
            $rv = $R.Os.Tls[$k]; $tv = $T.Os.Tls[$k]
            if ($rv -eq $tv) { Add-Check $T $c; continue }
            $sev = 'Low'; $act = 'Protocol setting differs. Align with your security baseline.'
            if ($k -like 'TLS 1.2*' -and $tv -eq 'Disabled') { $sev = 'High'; $act = 'TLS 1.2 is disabled on the target - modern drivers will fail to connect.' }
            elseif ($k -match 'TLS 1\.[01] (Server|Client)' -and $rv -ne 'Disabled' -and $tv -eq 'Disabled') { $sev = 'Medium'; $act = 'The target disables an old TLS version the reference still allows. Clients/linked servers using old drivers (SQLOLEDB, SQLNCLI10, old JDBC/ODBC) will fail; inventory them first.' }
            Add-Finding $T $c "SCHANNEL $k" $rv $tv $sev -Action $act
        }
    }

    # ================= Windows features & services
    $c = 'Windows features & services'
    if ($R.OsOk -and $T.OsOk -and $null -ne $R.Os.Features -and $null -ne $T.Os.Features) {
        foreach ($f in $R.Os.Features) {
            if ($T.Os.Features -contains $f) { Add-Check $T $c; continue }
            $sev = 'Low'; $hard = $false
            if ($f -match '(?i)^Failover Clustering$') { if ($isHa) { $sev = 'Blocker'; $hard = $true } else { $sev = 'Medium' } }
            elseif ($f -match '(?i)\.NET Framework 3\.5|Multipath|Failover Cluster Module|Failover Cluster Management') { $sev = 'Medium' }
            Add-Finding $T $c "Windows feature: $f" 'installed' 'not installed' $sev -Hard:$hard -Install -Action 'Add the Windows feature on the target (Install-WindowsFeature) if still needed.'
        }
    } elseif ($R.OsOk -and $T.OsOk -and ($null -eq $R.Os.Features -or $null -eq $T.Os.Features)) { Add-Check $T $c }
    if (Test-Sec $R $T 'Os' 'services' $c 'Windows services') {
        $fam = { param($list) $h = @{}; foreach ($s in $list) { $f = Get-SvcFamily $s.Name; if ($f) { if (-not $h.ContainsKey($f)) { $h[$f] = New-List }; Add-To $h[$f] $s } }; return $h }
        $rf = & $fam $R.Os.Services; $tf = & $fam $T.Os.Services
        foreach ($f in $rf.Keys) {
            $rs = $rf[$f] | Select-Object -First 1
            if ($rs.StartMode -eq 'Disabled') { Add-Check $T $c; continue }
            if (-not $tf.ContainsKey($f)) {
                $sev = switch ($f) { 'SQL CEIP' { 'Info' } 'SQL Browser' { 'Medium' } 'SQL VSS Writer' { 'Medium' } 'Distributed Transaction Coordinator' { 'Medium' } 'Cluster Service' { if ($isHa) { 'Blocker' } else { 'Low' } } default { 'High' } }
                Add-Finding $T $c "Service: $f" "$($rs.DisplayName) [$($rs.StartMode)]" 'not installed' $sev -Hard:($f -eq 'Cluster Service' -and $isHa) -Install:($sev -ne 'Info') -Action $(switch ($f) { 'Integration Services' { 'Install SSIS (SQL Server setup > Integration Services) - SSIS job steps/packages depend on it.' } 'Reporting Services' { 'Install SSRS/PBIRS and plan the ReportServer DB move + encryption key backup/restore.' } 'Analysis Services' { 'Install SSAS in the same mode (Tabular/Multidimensional) and migrate databases.' } 'Full-Text Filter Daemon' { 'Add the Full-Text and Semantic Extractions feature (SQL Server setup).' } 'Launchpad (ML Services)' { 'Add Machine Learning Services (R/Python) and the same runtimes/packages.' } 'PolyBase' { 'Add the PolyBase feature.' } 'SQL Browser' { 'Needed for named instances on dynamic ports. Install/enable or use fixed ports + aliases.' } default { 'Install/enable this component on the target.' } })
                continue
            }
            $ts = $tf[$f] | Select-Object -First 1
            if ($f -in @('Database Engine', 'SQL Server Agent', 'Integration Services', 'Reporting Services', 'Analysis Services') -and $rs.StartMode -eq 'Auto' -and $ts.StartMode -ne 'Auto') {
                Add-Finding $T $c "Service start mode: $f" $rs.StartMode $ts.StartMode $(if ($f -eq 'SQL Server Agent') { 'High' } else { 'Medium' }) -Action 'Set the service to Automatic on the target (jobs/alerts do not run while Agent is stopped).'
            } else { Add-Check $T $c }
            if ($isHa -and $f -eq 'Database Engine' -and [string]$rs.Account -ne [string]$ts.Account) { Add-Finding $T $c 'Engine service account' $rs.Account $ts.Account 'Low' -Action 'Different engine accounts across replicas: each needs a login with CONNECT on the HADR endpoint, and file-share/backup permissions.' }
        }
        $isThird = { param($s) $p = [string]$s.Path; ($p -and $p -notmatch '(?i)\\windows\\|\\Microsoft SQL Server\\|\\Microsoft\\|svchost' -and -not (Get-SvcFamily $s.Name)) }
        $tThird = @{}; foreach ($s in $T.Os.Services) { $tThird[$s.DisplayName] = $s }
        foreach ($s in $R.Os.Services) {
            if (-not (& $isThird $s) -or $s.StartMode -eq 'Disabled') { continue }
            if ($tThird.ContainsKey($s.DisplayName)) { Add-Check $T $c }
            else { Add-Finding $T $c "Third-party service: $($s.DisplayName)" "[$($s.StartMode)] $($s.Account)" 'not installed' 'Medium' -Install -Action 'Agent/service present on the reference only (backup, monitoring, AV, scheduler?). Install or confirm it is not needed.' -Detail (Protect-Secret ([string]$s.Path)) }
        }
    }

    # ================= SQL instance
    $c = 'SQL instance'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'instance' $c 'Instance properties')) {
        $rv = ConvertTo-VersionSafe $RI.ProductVersion; $tv = ConvertTo-VersionSafe $TI.ProductVersion
        $rTxt = "SQL $($RI.VersionName) $($RI.ProductLevel) $($RI.ProductUpdateLevel) ($($RI.ProductVersion))"; $tTxt = "SQL $($TI.VersionName) $($TI.ProductLevel) $($TI.ProductUpdateLevel) ($($TI.ProductVersion))"
        if ($tMaj -lt $rMaj) { Add-Finding $T $c 'Version' $rTxt $tTxt 'Blocker' -Hard -Action 'Target runs an OLDER major version. Databases cannot be restored/attached/seeded to an older version. Install the same or newer version.' }
        elseif ($isHa -and $tv -and $rv -and $tv -lt $rv) { Add-Finding $T $c 'Version / CU' $rTxt $tTxt 'High' -Action 'Secondary is on a lower build than the primary. Patch it to the same CU (patch secondaries first during rolling upgrades).' }
        elseif ($isHa -and $tv -ne $rv) { Add-Finding $T $c 'Version / CU' $rTxt $tTxt 'Medium' -Action 'Replica on a newer build than the primary. Fine only during a rolling upgrade; do not run mixed long-term.' }
        elseif ($isMig -and $tMaj -gt $rMaj) { Add-Finding $T $c 'Version (upgrade)' $rTxt $tTxt 'Info' -Action 'Upgrade migration: plan compatibility-level testing (Query Store baseline, new CE), check discontinued features (DMA/Upgrade Assessment), keep the old compat level at cut-over.' }
        elseif ($isMig -and $tv -and $rv -and $tv -lt $rv) { Add-Finding $T $c 'Version / CU' $rTxt $tTxt 'Medium' -Action 'Target is on a lower CU than the source. Patch the target to at least the source CU before restoring.' }
        else { Add-Check $T $c }

        $rEd = [string]$RI.Edition; $tEd = [string]$TI.Edition; $tEng = [int]$TI.EngineEdition; $rEng = [int]$RI.EngineEdition
        if ($tEd -match '(?i)Developer|Evaluation' -and $rEd -notmatch '(?i)Developer|Evaluation') { Add-Finding $T $c 'Edition licensing' $rEd $tEd 'High' -Action 'Developer/Evaluation edition is not licensed for production (Evaluation also expires). Install the licensed edition before cut-over.' }
        elseif ($rEd -ne $tEd) { Add-Finding $T $c 'Edition' $rEd $tEd $(if ($rEng -eq 3 -and $tEng -ne 3) { 'High' } else { 'Low' }) -Action $(if ($rEng -eq 3 -and $tEng -ne 3) { 'Enterprise -> lower edition: buffer pool/core caps, Basic AG only (1 DB, 2 replicas, no readable secondary), no online index ops/partition switching parity on old builds. See the persisted-feature checks under Databases.' } else { 'Edition differs - confirm licensing.' }) }
        else { Add-Check $T $c }
        if ($tEng -ne 3) {
            $isSp1 = ($tMaj -gt 13) -or ($tMaj -eq 13 -and $tv -and $tv -ge [version]'13.0.4001.0')
            foreach ($sf in ((ConvertTo-Arr $R.Sql.SkuFeatures) | Group-Object Feature)) {
                $dbs = ($sf.Group | Select-Object -ExpandProperty Db -Unique) -join ', '
                $bad = (-not $isSp1) -or ($sf.Name -eq 'TransparentDataEncryption' -and $tMaj -lt 15) -or ($tEng -eq 4 -and $sf.Name -in @('TransparentDataEncryption', 'Partitioning', 'ChangeCapture'))
                if ($bad) { Add-Finding $T 'Databases' "Edition feature in use: $($sf.Name)" "used by $dbs" "not supported by $tEd" 'Blocker' -Hard -Action 'These databases will not restore/start on this edition. Use Enterprise, a newer version (2016 SP1+/2019+ for TDE in Standard), or remove the feature first.' -Detail $dbs }
                else { Add-Finding $T 'Databases' "Edition feature in use: $($sf.Name)" "used by $dbs" "supported by $tEd" 'Info' -Action 'Feature available on the target edition (2016 SP1+ programmability parity); check scale limits.' -Detail $dbs }
            }
            if ($tEng -eq 4) {
                foreach ($d in @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Id -gt 4 -and [int64]$_.SizeMB -gt 10240 })) { Add-Finding $T 'Databases' "Express 10 GB limit: $($d.Name)" (Format-MB $d.SizeMB) $tEd 'Blocker' -Hard -Action 'Database exceeds the Express edition size limit.' }
                if (@($R.Sql.Jobs).Count -gt 0) { Add-Finding $T 'Agent & automation' 'SQL Server Agent (Express)' "$(@($R.Sql.Jobs).Count) jobs" 'no Agent in Express' 'High' -Action 'Express has no SQL Server Agent. Use Task Scheduler or another edition.' }
            }
        }
        if ([string]$RI.Collation -ne [string]$TI.Collation) { Add-Finding $T $c 'Server collation' $RI.Collation $TI.Collation 'High' -Action 'tempdb/master collation differs: temp-table joins raise "Cannot resolve the collation conflict", sorting/case-sensitivity of system objects changes. Rebuild system DBs with the reference collation or fix code (COLLATE DATABASE_DEFAULT).' } else { Add-Check $T $c }
        if ($isHa -and [int]$TI.IsHadrEnabled -ne 1) { Add-Finding $T $c 'Always On availability groups' 'enabled' 'disabled' 'Blocker' -Hard -Action 'Enable Always On in SQL Server Configuration Manager (needs a service restart).' }
        elseif ($isMig -and [int]$RI.IsHadrEnabled -eq 1 -and [int]$TI.IsHadrEnabled -ne 1) { Add-Finding $T $c 'Always On availability groups' 'enabled' 'disabled' 'Medium' -Action 'Reference uses HADR. Enable it on the target if the new build will host an AG.' }
        else { Add-Check $T $c }
        $fsDbs = @((ConvertTo-Arr $R.Sql.DbFeatures) | Where-Object { $_.FsFilegroups -gt 0 } | ForEach-Object { $_.Db })
        if ($fsDbs.Count -gt 0 -and [int]$TI.FilestreamConfiguredLevel -eq 0) { Add-Finding $T $c 'FILESTREAM' ("level $($RI.FilestreamConfiguredLevel); used by " + ($fsDbs -join ', ')) 'disabled' 'Blocker' -Hard -Action 'Enable FILESTREAM in Configuration Manager and set filestream access level before restoring these databases.' -Detail ($fsDbs -join ', ') }
        elseif ([int]$RI.FilestreamConfiguredLevel -ne [int]$TI.FilestreamConfiguredLevel) { Add-Finding $T $c 'FILESTREAM level' $RI.FilestreamConfiguredLevel $TI.FilestreamConfiguredLevel 'Low' -Action 'FILESTREAM configuration differs.' }
        else { Add-Check $T $c }
        $ftDbs = @((ConvertTo-Arr $R.Sql.DbFeatures) | Where-Object { $_.FtCatalogs -gt 0 } | ForEach-Object { $_.Db })
        if ($ftDbs.Count -gt 0 -and [int]$TI.IsFullTextInstalled -ne 1) { Add-Finding $T $c 'Full-Text Search' ('used by ' + ($ftDbs -join ', ')) 'not installed' 'High' -Install -Action 'Install the Full-Text and Semantic Extractions for Search feature; full-text queries fail without it. Repopulate catalogs after restore.' -Detail ($ftDbs -join ', ') } else { Add-Check $T $c }
        if ([int]$RI.IsPolyBaseInstalled -eq 1 -and [int]$TI.IsPolyBaseInstalled -ne 1) { Add-Finding $T $c 'PolyBase' 'installed' 'not installed' 'Medium' -Install -Action 'Install PolyBase if external tables/data sources are used.' } else { Add-Check $T $c }
        if ([int]$RI.IsIntegratedSecurityOnly -eq 0 -and [int]$TI.IsIntegratedSecurityOnly -eq 1) { Add-Finding $T $c 'Authentication mode' 'Mixed (SQL + Windows)' 'Windows only' 'High' -Action 'SQL logins cannot connect. Switch the target to Mixed Mode (restart required).' } else { Add-Check $T $c }
        if ((Get-Prop $TI 'Ifi') -eq 'N') { Add-Finding $T $c 'Instant file initialization (target baseline)' (Get-Prop $RI 'Ifi') 'N' 'Low' -Action 'Grant "Perform volume maintenance tasks" to the engine service account: restores, seeding and growth are much faster.' } else { Add-Check $T $c }
        if ((Get-Prop $RI 'MemoryModel') -eq 'LOCK_PAGES' -and (Get-Prop $TI 'MemoryModel') -and (Get-Prop $TI 'MemoryModel') -ne 'LOCK_PAGES') { Add-Finding $T $c 'Lock pages in memory' 'LOCK_PAGES' (Get-Prop $TI 'MemoryModel') 'Low' -Action 'Reference uses Lock Pages in Memory; grant it to the target service account if it is part of your standard.' } else { Add-Check $T $c }
        if ([string]$RI.InstanceName -ne [string]$TI.InstanceName -and $isMig) { Add-Finding $T $c 'Instance name' $(if ($RI.InstanceName) { $RI.InstanceName } else { 'default (MSSQLSERVER)' }) $(if ($TI.InstanceName) { $TI.InstanceName } else { 'default (MSSQLSERVER)' }) 'Info' -Action 'Connection strings change unless you use a DNS alias / CNAME + SPN, or a client alias.' }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'registry' $c 'Network configuration')) {
        $rn = $R.Sql.Net; $tn = $T.Sql.Net
        if ((Get-Prop $tn 'TcpEnabled') -eq '0') { Add-Finding $T $c 'TCP/IP protocol' 'enabled' 'disabled' 'High' -Action 'Enable TCP/IP in Configuration Manager (restart).' } else { Add-Check $T $c }
        $rp = "$(Get-Prop $rn 'TcpPort')"; $tp = "$(Get-Prop $tn 'TcpPort')"
        if ($rp -ne $tp) { Add-Finding $T $c 'TCP port' $(if ($rp) { $rp } else { 'dynamic ' + (Get-Prop $rn 'TcpDynamicPorts') }) $(if ($tp) { $tp } else { 'dynamic ' + (Get-Prop $tn 'TcpDynamicPorts') }) $(if ($isMig) { 'Medium' } else { 'Low' }) -Action 'Apps, firewalls and connection strings that use host,port break. Set the same static port or update them.' } else { Add-Check $T $c }
        if ((Get-Prop $rn 'ForceEncryption') -eq '1' -and (Get-Prop $tn 'ForceEncryption') -ne '1') { Add-Finding $T $c 'Force encryption' 'Yes' 'No' 'Medium' -Action 'Reference forces TLS for all connections. Install a certificate and enable Force Encryption on the target.' } else { Add-Check $T $c }
        if ((Get-Prop $rn 'Certificate') -and -not (Get-Prop $tn 'Certificate')) { Add-Finding $T $c 'TLS certificate for SQL Server' 'configured' 'none (self-signed)' 'Medium' -Action 'Issue/install a certificate for the target host (and listener names) and bind it in Configuration Manager.' } else { Add-Check $T $c }
    }

    # ================= Configuration
    $c = 'Configuration'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'config' $c 'sp_configure')) {
        $use = @{ xp = @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.XpCmdshell -gt 0 } | ForEach-Object { $_.Db }); ole = @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.OleAutomation -gt 0 } | ForEach-Object { $_.Db }); adhoc = @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.AdHocQueries -gt 0 } | ForEach-Object { $_.Db }); mail = @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.DbMail -gt 0 } | ForEach-Object { $_.Db }) }
        $hasAsm = @($R.Sql.Assemblies).Count -gt 0
        $hasContained = @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Containment -and $_.Containment -ne 'NONE' }).Count -gt 0
        foreach ($k in $R.Sql.Config.Keys) {
            $rv = [string]$R.Sql.Config[$k]
            if (-not $T.Sql.Config.Contains($k)) { continue }
            $tv = [string]$T.Sql.Config[$k]
            if ($rv -eq $tv) { Add-Check $T $c; continue }
            $sev = 'Low'; $hard = $false; $act = 'Align with the reference (Copy-DbaSpConfigure, or the prep script) unless the difference is intended.'
            switch ($k) {
                'clr enabled' { if ($hasAsm -and $tv -eq '0') { $sev = 'High'; $act = 'User CLR assemblies exist on the reference; CLR objects fail until clr enabled = 1.' } }
                'xp_cmdshell' { if ($use.xp.Count -gt 0) { $sev = 'High'; $act = 'Code uses xp_cmdshell (' + ($use.xp -join ', ') + '). Enable it or refactor before cut-over.' } else { $sev = 'Medium' } }
                'Ole Automation Procedures' { if ($use.ole.Count -gt 0) { $sev = 'High'; $act = 'Code uses sp_OA* (' + ($use.ole -join ', ') + ').' } }
                'Ad Hoc Distributed Queries' { if ($use.adhoc.Count -gt 0) { $sev = 'High'; $act = 'Code uses OPENROWSET/OPENDATASOURCE (' + ($use.adhoc -join ', ') + ').' } }
                'Database Mail XPs' { if (@($R.Sql.Mail).Count -gt 0 -or $use.mail.Count -gt 0) { $sev = 'High'; $act = 'Database Mail is used on the reference (profiles/sp_send_dbmail).' } }
                'contained database authentication' { if ($hasContained -and $tv -eq '0') { $sev = 'Blocker'; $hard = $true; $act = 'Contained databases exist; they will not open until contained database authentication = 1.' } }
                'filestream access level' { $sev = 'Medium' }
                'max server memory (MB)' { $sev = $(if ($isMig) { 'Info' } else { 'Medium' }); $act = $(if ($isMig) { 'Sized per host - review against target RAM (see Hardware).' } else { 'Replicas should be able to carry the primary workload after failover.' }) }
                'min server memory (MB)' { $sev = 'Info' }
                'max degree of parallelism' { $sev = 'Medium' }
                'cost threshold for parallelism' { $sev = 'Medium' }
                'optimize for ad hoc workloads' { $sev = 'Low' }
                'remote admin connections' { $sev = 'Low' }
                'backup compression default' { $sev = 'Low' }
                'backup checksum default' { $sev = 'Low' }
                'scan for startup procs' { if (@($R.Sql.StartupProcs).Count -gt 0) { $sev = 'High'; $act = 'Startup procedures exist; they only run when scan for startup procs = 1.' } }
                'cross db ownership chaining' { $sev = 'Medium' }
                'remote access' { $sev = 'Low' }
                'show advanced options' { $sev = 'Info' }
                'clr strict security' { $sev = $(if ($hasAsm) { 'Medium' } else { 'Info' }) }
                'external scripts enabled' { $sev = 'Medium' }
                'polybase enabled' { $sev = 'Medium' }
                'user connections' { $sev = 'Medium' }
                'default trace enabled' { $sev = 'Low' }
                'blocked process threshold (s)' { $sev = 'Low' }
                'common criteria compliance enabled' { $sev = 'Medium' }
                'c2 audit mode' { $sev = 'Medium' }
                'priority boost' { $sev = 'Low' }
                'lightweight pooling' { $sev = 'Low' }
                'affinity mask' { $sev = 'Low' }
                'affinity64 mask' { $sev = 'Low' }
                'fill factor (%)' { $sev = 'Medium' }
                'max text repl size (B)' { $sev = 'Medium' }
                'default language' { $sev = 'Medium'; $act = 'New logins get a different default language - date parsing (SET DATEFORMAT) changes.' }
                'nested triggers' { $sev = 'Medium' }
                'user options' { $sev = 'Medium'; $act = 'Default SET options for every connection (ANSI_NULLS, QUOTED_IDENTIFIER, ARITHABORT ...) differ - query results, indexed views and computed columns can behave differently.' }
                'two digit year cutoff' { $sev = 'Medium'; $act = 'Two-digit years are interpreted differently (e.g. 49 -> 2049 vs 1949).' }
                'remote query timeout (s)' { $sev = 'Medium'; $act = 'Linked-server/remote queries time out at a different point.' }
                'default full-text language' { $sev = 'Low' }
                'recovery interval (min)' { $sev = 'Low' }
                'disallow results from triggers' { $sev = 'Medium' }
                'max worker threads' { $sev = 'Low' }
            }
            Add-Finding $T $c "sp_configure: $k" $rv $tv $sev -Hard:$hard -Action $act -Fix 'Copy-DbaSpConfigure / prep script section 1'
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'registry' $c 'Startup trace flags')) {
        foreach ($tf in $R.Sql.TraceFlags) {
            if ($T.Sql.TraceFlags -contains $tf) { Add-Check $T $c; continue }
            if ($script:ObsoleteTf.ContainsKey([int]$tf) -and $tMaj -ge 13) { Add-Finding $T $c "Startup trace flag -T$tf" 'on' 'off' 'Info' -Action ('Not needed on the target: ' + $script:ObsoleteTf[[int]$tf] + '.') }
            else { Add-Finding $T $c "Startup trace flag -T$tf" 'on' 'off' 'Medium' -Action ('Add -T' + $tf + ' to the startup parameters (Configuration Manager; restart) if still required. Also check DBCC TRACEON in jobs/startup procs.') -Fix 'SQL Server Configuration Manager' }
        }
        foreach ($tf in $T.Sql.TraceFlags) { if ($R.Sql.TraceFlags -notcontains $tf) { Add-Finding $T $c "Startup trace flag -T$tf" 'off' 'on' 'Low' -Action 'Trace flag only on the target. Confirm it is intended.' } }
    }

    # ================= CLR & assemblies
    $c = 'CLR & assemblies'
    if ($R.SqlOk -and $T.SqlOk -and @($R.Sql.Assemblies).Count -gt 0) {
        $tStrict = [string](Get-Prop $T.Sql.Config 'clr strict security' '')
        foreach ($g in ((ConvertTo-Arr $R.Sql.Assemblies) | Group-Object Db)) {
            $names = ($g.Group | ForEach-Object { "$($_.Name) [$($_.PermissionSet -replace '_ASSEMBLY', '')]" }) -join ', '
            $unsafe = @($g.Group | Where-Object { $_.PermissionSet -notlike 'SAFE*' }).Count -gt 0
            $tw = [int]($g.Group[0].Trustworthy)
            if ($tMaj -ge 14 -and $rMaj -lt 14 -and $tStrict -ne '0') {
                Add-Finding $T $c "clr strict security: $($g.Name)" $names "SQL $($TI.VersionName) treats every assembly as UNSAFE" 'High' -Action 'Sign the assemblies (certificate/asymmetric key + login with UNSAFE ASSEMBLY), or whitelist them with sys.sp_add_trusted_assembly, before go-live. TRUSTWORTHY ON is the discouraged fallback.' -Detail $names
            } else { Add-Check $T $c }
            if ($isMig -and $tw -eq 1) { Add-Finding $T $c "TRUSTWORTHY: $($g.Name)" 'ON' 'reset to OFF by restore' $(if ($unsafe) { 'High' } else { 'Medium' }) -Action 'RESTORE/attach sets TRUSTWORTHY OFF. EXTERNAL_ACCESS/UNSAFE assemblies and cross-database impersonation stop working until it is re-enabled (or the assemblies are signed). Database owner must also be the same login.' -Detail $names }
        }
        if ($T.Sql.Sections['trusted_assemblies'] -like 'ok*' -and $R.Sql.Sections['trusted_assemblies'] -like 'ok*') {
            Compare-NamedSet $T $c 'Trusted assembly hash' $R.Sql.TrustedAssemblies $T.Sql.TrustedAssemblies { ([string]$args[0].Hash).Substring(0, [Math]::Min(24, ([string]$args[0].Hash).Length)) + '...' } 'High' -Action 'Whitelisted on the reference only. Add with sys.sp_add_trusted_assembly on the target.' -Show { $args[0].Description }
        }
    }

    # ================= Linked servers & external
    $c = 'Linked servers & external'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'linked_servers' $c 'Linked servers')) {
        $refBy = @{}; foreach ($l in $R.Sql.LinkedRefs) { if (-not $refBy.ContainsKey($l.LinkedServer)) { $refBy[$l.LinkedServer] = New-List }; if (-not $refBy[$l.LinkedServer].Contains($l.Db)) { Add-To $refBy[$l.LinkedServer] $l.Db } }
        $tLs = @{}; foreach ($l in $T.Sql.LinkedServers) { $tLs[$l.Name] = $l }
        foreach ($l in $R.Sql.LinkedServers) {
            $users = if ($refBy.ContainsKey($l.Name)) { 'referenced by ' + ($refBy[$l.Name] -join ', ') } else { '' }
            if (-not $tLs.ContainsKey($l.Name)) { Add-Finding $T $c "Linked server: $($l.Name)" "$($l.Provider) -> $($l.DataSource)" 'missing' 'High' -Action ('Create it on the target (Copy-DbaLinkedServer keeps stored remote passwords; otherwise re-enter them) and test. ' + $users).Trim() -Fix 'Copy-DbaLinkedServer' -Detail $users }
            else {
                $tl = $tLs[$l.Name]
                if ($tl.Provider -ne $l.Provider -or $tl.DataSource -ne $l.DataSource) { Add-Finding $T $c "Linked server: $($l.Name)" "$($l.Provider) -> $($l.DataSource)" "$($tl.Provider) -> $($tl.DataSource)" 'Medium' -Action 'Definition differs on the target. Confirm the target points where it should.' } else { Add-Check $T $c }
                if ([int]$l.StoredLogins -gt [int]$tl.StoredLogins) { Add-Finding $T $c "Linked server logins: $($l.Name)" "$($l.StoredLogins) stored remote logins" "$($tl.StoredLogins)" 'Medium' -Action 'Remote login mappings missing; passwords are not readable from catalog views - re-enter them or use Copy-DbaLinkedServer.' }
            }
            $prov = [string]$l.Provider
            if ($prov -and $T.OsOk -and $T.Os.OleDb.Contains($prov) -and -not $T.Os.OleDb[$prov].x64) { Add-Finding $T 'Drivers & providers' "Provider for linked server $($l.Name): $prov" 'installed on reference' 'not registered (x64)' 'High' -Install -Action ('Install ' + $(if ($script:ProviderNames.ContainsKey($prov)) { $script:ProviderNames[$prov] } else { "the 64-bit $prov provider" }) + ' on the target.') }
            if ($prov -like 'SQLNCLI*' -and $tMaj -ge 16) { Add-Finding $T $c "Deprecated provider: $($l.Name)" $prov "not shipped with SQL $($TI.VersionName)" 'High' -Action 'SQL Server Native Client is removed from SQL 2022+. Recreate the linked server with MSOLEDBSQL (watch the Encrypt default in OLE DB Driver 19).' }
            if ($prov -eq 'MSDASQL' -and $T.OsOk -and $T.Os.Sections['odbc'] -like 'ok*' -and -not @((ConvertTo-Arr $T.Os.Dsns) | Where-Object { $_.Name -eq $l.DataSource -and $_.Bitness -eq 'x64' }).Count) { Add-Finding $T $c "DSN for linked server $($l.Name)" $l.DataSource 'missing (x64)' 'High' -Install -Action 'Create the 64-bit system DSN (and its ODBC driver) on the target.' }
        }
        foreach ($k in $refBy.Keys) { if (-not @((ConvertTo-Arr $R.Sql.LinkedServers) | Where-Object { $_.Name -eq $k }).Count -and -not @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Name -eq $k }).Count) { Add-Finding $T $c "Code references unknown server: $k" ('in ' + ($refBy[$k] -join ', ')) '' 'Info' -Action 'Code uses 4-part names for a server that is not a linked server on the reference either (stale code or remote alias). Review.' } }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'credentials' $c 'Credentials')) {
        Compare-NamedSet $T 'Security objects' 'Credential' $R.Sql.Credentials $T.Sql.Credentials { $args[0].Name } 'High' -Action 'Create the credential on the target; the secret cannot be read from the reference - re-enter it (Copy-DbaCredential can decrypt with sysadmin + DAC).' -Fix 'Copy-DbaCredential' -Show { "identity: $($args[0].CredIdentity)" } -Diff { if ($args[0].CredIdentity -ne $args[1].CredIdentity) { , @($args[0].CredIdentity, $args[1].CredIdentity, 'Credential identity differs.') } }
    }

    # ================= Agent & automation
    $c = 'Agent & automation'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'jobs' $c 'Agent jobs')) {
        $skipJob = { param($j) ($j.Name -eq 'syspolicy_purge_history') -or ($j.Category -eq 'Report Server') }
        $rJobs = @((ConvertTo-Arr $R.Sql.Jobs) | Where-Object { -not (& $skipJob $_) }); $tJobs = @((ConvertTo-Arr $T.Sql.Jobs) | Where-Object { -not (& $skipJob $_) })
        $ssrsJobs = @((ConvertTo-Arr $R.Sql.Jobs) | Where-Object { $_.Category -eq 'Report Server' }).Count
        Compare-NamedSet $T $c 'Agent job' $rJobs $tJobs { $args[0].Name } 'High' -GroupAbove 25 -Action $(if ($isHa) { 'Jobs are not replicated by AGs. Create it on this replica too (usually with a "is primary?" check in step 1).' } else { 'Script the job to the target (Copy-DbaAgentJob), then review steps, schedules, owners and proxies.' }) -Fix 'Copy-DbaAgentJob' -Show { "$(if ($args[0].Enabled) { 'enabled' } else { 'disabled' }), owner $($args[0].Owner), $($args[0].Steps) steps" } -Diff { if ([int]$args[0].Enabled -ne [int]$args[1].Enabled) { , @($(if ($args[0].Enabled) { 'enabled' } else { 'disabled' }), $(if ($args[1].Enabled) { 'enabled' } else { 'disabled' }), 'Enabled state differs (expected on secondaries; at cut-over enable on the new server).') } } -DiffSev 'Info'
        if ($ssrsJobs -gt 0) { Add-Finding $T 'SSIS / SSRS / SSAS' 'SSRS subscription jobs' "$ssrsJobs jobs" 'recreated by SSRS' 'Info' -Action 'Report Server subscription jobs are recreated automatically when the ReportServer DB is attached to the new SSRS instance.' }
        if ($T.Sql.Sections['login_names'] -like 'ok*') {
            foreach ($j in $rJobs) { if ($j.Owner -and $T.Sql.Logins -notcontains $j.Owner) { Add-Finding $T $c "Job owner login: $($j.Owner)" $j.Name 'login missing' 'Medium' -Action 'Job owner does not exist on the target. Create the login (see LoginSync) or change the owner (sa / service account).' } }
        }
        foreach ($s in $R.Sql.JobSteps) {
            if ($s.Subsystem -eq 'SSIS' -and $T.OsOk -and $T.Os.Sections['services'] -like 'ok*' -and -not @((ConvertTo-Arr $T.Os.Services) | Where-Object { (Get-SvcFamily $_.Name) -eq 'Integration Services' }).Count) { Add-Finding $T $c "SSIS job step: $($s.Job) / $($s.StepId)" 'SSIS subsystem' 'Integration Services not installed' 'High' -Install -Action 'Install Integration Services on the target (DTExec + catalog).' }
            if ($s.Subsystem -like 'ANALYSIS*' -and $T.OsOk -and -not @((ConvertTo-Arr $T.Os.Services) | Where-Object { (Get-SvcFamily $_.Name) -eq 'Analysis Services' }).Count) { Add-Finding $T $c "SSAS job step: $($s.Job) / $($s.StepId)" $s.Subsystem 'no local SSAS' 'Low' -Action 'Step talks to Analysis Services; confirm the SSAS server name is reachable from the target.' }
            if ($s.Subsystem -in @('Distribution', 'LogReader', 'Snapshot', 'Merge', 'QueueReader')) { Add-Finding $T $c "Replication agent job: $($s.Job)" $s.Subsystem '' 'Info' -Action 'Recreated when replication is configured on the target - do not copy by hand.' }
            foreach ($u in $s.Unc) { Add-Finding $T 'Files & paths' "UNC path in job: $u" "$($s.Job) / step $($s.StepId)" 'not checked' 'Low' -Action "Verify the target's Agent/proxy account can read/write this share (not testable read-only)." }
            foreach ($m in $s.Modules) {
                if ($T.OsOk -and $null -ne $T.Os.PsModules -and $T.Os.PsModules -notcontains $m -and $m -notin @('SQLPS', 'Microsoft.PowerShell.Management', 'Microsoft.PowerShell.Utility')) { Add-Finding $T $c "PowerShell module for job: $m" "$($s.Job) / step $($s.StepId)" 'not installed (Program Files)' 'High' -Install -Action "Install-Module $m -Scope AllUsers on the target (or copy it into Program Files\WindowsPowerShell\Modules)." }
            }
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'proxies' $c 'Agent proxies')) { Compare-NamedSet $T $c 'Agent proxy' $R.Sql.Proxies $T.Sql.Proxies { $args[0].Name } 'High' -Action 'Create the proxy (and its credential first); job steps that run as this proxy fail without it.' -Fix 'Copy-DbaAgentProxy' -Show { "credential $($args[0].Credential)" } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'operators' $c 'Agent operators')) { Compare-NamedSet $T $c 'Operator' $R.Sql.Operators $T.Sql.Operators { $args[0].Name } 'Medium' -Action 'Create the operator so job/alert notifications still go out.' -Fix 'Copy-DbaAgentOperator' -Show { $args[0].Email } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'alerts' $c 'Agent alerts')) { Compare-NamedSet $T $c 'Alert' $R.Sql.Alerts $T.Sql.Alerts { $args[0].Name } 'Medium' -GroupAbove 3 -Action 'Create the alert (severity 16-25 and error 823/824/825 alerts are typical).' -Fix 'Copy-DbaAgentAlert' -Show { if ($args[0].MessageId) { "error $($args[0].MessageId)" } else { "severity $($args[0].Severity)" } } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'dbmail' $c 'Database Mail')) {
        $rp = @((ConvertTo-Arr $R.Sql.Mail) | Group-Object Profile); $tp = @((ConvertTo-Arr $T.Sql.Mail) | Group-Object Profile)
        Compare-NamedSet $T $c 'Database Mail profile' $rp $tp { $args[0].Name } $(if (@($R.Sql.Operators).Count -gt 0 -or @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.DbMail -gt 0 }).Count -gt 0) { 'High' } else { 'Medium' }) -Action 'Create the profile/account (SMTP passwords must be re-entered). Also set it as the Agent mail profile and restart Agent.' -Fix 'Copy-DbaDbMail' -Show { (($args[0].Group | ForEach-Object { "$($_.Account) via $($_.SmtpServer):$($_.Port)" }) -join '; ') }
    }
    if ($T.OsOk -and $R.OsOk -and $null -ne $R.Os.PsModules -and $null -ne $T.Os.PsModules) {
        foreach ($m in $R.Os.PsModules) { if ($T.Os.PsModules -contains $m) { Add-Check $T $c } elseif (-not @($script:Findings | Where-Object { $_.Target -eq $T.Name -and $_.Item -eq "PowerShell module for job: $m" }).Count) { Add-Finding $T $c "PowerShell module: $m" 'installed (all users)' 'not installed' 'Low' -Install -Action 'Installed for all users on the reference only. Install if scripts on this host use it.' } }
    }

    # ================= SSIS / SSRS / SSAS
    $c = 'SSIS / SSRS / SSAS'
    if ($R.SqlOk -and $T.SqlOk) {
        if (@($R.Sql.SsisProjects).Count -gt 0) {
            if (-not @((ConvertTo-Arr $T.Sql.Databases) | Where-Object { $_.Name -eq 'SSISDB' }).Count) { Add-Finding $T $c 'SSIS catalog (SSISDB)' "$(@($R.Sql.SsisProjects).Count) projects" 'no catalog' 'High' -Install -Action $(if ($isHa) { 'Add SSISDB to the AG (needs the catalog master key password) and enable SSIS support for Always On.' } else { 'Create the SSIS catalog (needs CLR + a master key password) and redeploy the projects (.ispac / ISDeploymentWizard), then recreate environments.' }) }
            elseif ($T.Sql.Sections['ssisdb'] -like 'ok*') {
                Compare-NamedSet $T $c 'SSIS project' $R.Sql.SsisProjects $T.Sql.SsisProjects { "$($args[0].Folder)\$($args[0].Project)" } 'High' -Action 'Deploy the project to the target catalog (same folder) and re-map environments/parameters.' -Show { "deployed $($args[0].Deployed)" }
                Compare-NamedSet $T $c 'SSIS environment' $R.Sql.SsisEnvironments $T.Sql.SsisEnvironments { "$($args[0].Folder)\$($args[0].Environment)" } 'Medium' -Action 'Recreate the environment and its variables (sensitive values must be re-entered).'
            }
        }
        if ((Test-Sec $R $T 'Sql' 'msdb_packages' $c 'SSIS packages in msdb')) { Compare-NamedSet $T $c 'SSIS package (msdb)' $R.Sql.SsisMsdb $T.Sql.SsisMsdb { "$($args[0].Folder)\$($args[0].Package)" } 'High' -Action 'Move the package to the target msdb (dtutil /SQL ... /COPY) or convert to the project model.' }
        foreach ($d in @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Name -match '^ReportServer' -and $_.Name -notmatch 'TempDB$' })) {
            if ($isMig) { Add-Finding $T $c "SSRS database: $($d.Name)" 'present' '' 'High' -Action 'Back up the SSRS encryption key on the old server (rskeymgmt -e / Report Server Configuration Manager), restore ReportServer + ReportServerTempDB, restore the key, then fix the scale-out list. Without the key all data sources/subscriptions with stored credentials are lost.' }
        }
    }

    # ================= Security objects
    $c = 'Security objects'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'certificates' $c 'Certificates in master')) {
        $tThumbs = @((ConvertTo-Arr $T.Sql.Certificates) | ForEach-Object { $_.Thumbprint }) + @((ConvertTo-Arr $T.Sql.AsymKeys) | ForEach-Object { $_.Thumbprint })
        $tdeThumbs = @{}
        foreach ($k in $R.Sql.Tde) {
            if (-not $k.Thumbprint) { continue }
            $cert = (ConvertTo-Arr $R.Sql.Certificates) | Where-Object { $_.Thumbprint -eq $k.Thumbprint } | Select-Object -First 1
            $cn = if ($cert) { $cert.Name } else { "asymmetric key $($k.Thumbprint.Substring(0, 12))..." }
            if (-not $tdeThumbs.ContainsKey($k.Thumbprint)) { $tdeThumbs[$k.Thumbprint] = New-List }
            Add-To $tdeThumbs[$k.Thumbprint] ([pscustomobject]@{ Db = $k.Db; Cert = $cn })
        }
        foreach ($cert in @(ConvertTo-Arr $R.Sql.Certificates | Where-Object { $tdeThumbs.ContainsKey($_.Thumbprint) -and $_.Expiry -and $_.Expiry -lt (Get-Date -Format 'yyyy-MM-dd') })) {
            Add-Finding $T $c "Expired TDE certificate: $($cert.Name)" "expired $($cert.Expiry)" '' 'Info' -Action 'Expired certificates still decrypt TDE databases and can be backed up/restored, but plan a rotation (ALTER DATABASE ENCRYPTION KEY ... ENCRYPTION BY SERVER CERTIFICATE <new>) during the migration window.'
        }
        foreach ($th in $tdeThumbs.Keys) {
            $dbs = ($tdeThumbs[$th] | ForEach-Object { $_.Db }) -join ', '; $cn = $tdeThumbs[$th][0].Cert
            if ($tThumbs -contains $th) { Add-Check $T $c }
            else { Add-Finding $T $c "TDE certificate: $cn" "protects $dbs" 'missing on target' 'Blocker' -Hard -Action 'BACKUP CERTIFICATE ... WITH PRIVATE KEY on the reference and CREATE CERTIFICATE ... FROM FILE on the target (or Copy-DbaDbCertificate) BEFORE restoring/seeding. Without it the databases cannot be restored or joined.' -Fix 'Copy-DbaDbCertificate' -Detail $dbs }
        }
        foreach ($cert in $R.Sql.Certificates) {
            if ($tdeThumbs.ContainsKey($cert.Thumbprint)) { continue }
            if ($tThumbs -contains $cert.Thumbprint) { Add-Check $T $c }
            else { Add-Finding $T $c "Certificate in master: $($cert.Name)" "$($cert.Subject), expires $($cert.Expiry)" 'missing' 'Medium' -Action 'Used for backup encryption, module signing, endpoint auth or Service Broker? Copy it with its private key if needed (restoring encrypted backups needs it).' -Fix 'Copy-DbaDbCertificate' }
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'crypto_providers' $c 'EKM providers')) { Compare-NamedSet $T $c 'EKM provider' $R.Sql.CryptoProviders $T.Sql.CryptoProviders { $args[0].Name } 'High' -Hard -Action 'Install the vendor EKM/HSM client + DLL, CREATE CRYPTOGRAPHIC PROVIDER and the credential before restoring EKM-protected databases.' -Show { $args[0].Dll } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'audits' $c 'Server audits')) { Compare-NamedSet $T $c 'Server audit' $R.Sql.Audits $T.Sql.Audits { $args[0].Name } 'High' -Action 'Compliance gap: create the audit on the target (Copy-DbaInstanceAudit) and its folder.' -Fix 'Copy-DbaInstanceAudit' -Show { "$($args[0].Type) -> $($args[0].Path)" } -Diff { if ([int]$args[0].Enabled -ne [int]$args[1].Enabled) { , @("enabled=$($args[0].Enabled)", "enabled=$($args[1].Enabled)", 'Audit state differs.') } } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'audit_specs' $c 'Server audit specifications')) { Compare-NamedSet $T $c 'Server audit specification' $R.Sql.AuditSpecs $T.Sql.AuditSpecs { $args[0].Name } 'High' -Action 'Create the server audit specification (Copy-DbaInstanceAuditSpecification).' -Fix 'Copy-DbaInstanceAuditSpecification' }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'server_triggers' $c 'Server triggers')) { Compare-NamedSet $T $c 'Server trigger' $R.Sql.ServerTriggers $T.Sql.ServerTriggers { $args[0].Name } 'High' -Action 'Create the server/logon trigger (Copy-DbaInstanceTrigger). Test logon triggers carefully - a broken one locks everyone out (DAC can bypass).' -Fix 'Copy-DbaInstanceTrigger' }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'endpoints' $c 'Endpoints')) {
        $rm = @((ConvertTo-Arr $R.Sql.Endpoints) | Where-Object { $_.Type -eq 'DATABASE_MIRRORING' }); $tm = @((ConvertTo-Arr $T.Sql.Endpoints) | Where-Object { $_.Type -eq 'DATABASE_MIRRORING' })
        if ($isHa -and $tm.Count -eq 0) { Add-Finding $T $c 'HADR endpoint' "$(if ($rm.Count) { "$($rm[0].Name) port $($rm[0].Port) $($rm[0].Encryption)" } else { 'n/a' })" 'missing' 'Blocker' -Hard -Action 'Create the DATABASE_MIRRORING endpoint (same port, compatible encryption), start it and GRANT CONNECT to the partner engine accounts.' -Fix 'Copy-DbaEndpoint' }
        elseif ($rm.Count -and $tm.Count) {
            if ([string]$rm[0].Encryption -ne [string]$tm[0].Encryption) { Add-Finding $T $c 'HADR endpoint encryption' $rm[0].Encryption $tm[0].Encryption $(if ($isHa) { 'High' } else { 'Low' }) -Hard:$isHa -Action 'Endpoint encryption must be compatible between partners (e.g. both AES), otherwise replicas cannot connect.' } else { Add-Check $T $c }
            if ([string]$rm[0].Port -ne [string]$tm[0].Port) { Add-Finding $T $c 'HADR endpoint port' $rm[0].Port $tm[0].Port 'Low' -Action 'Different endpoint ports are allowed (endpoint_url) but firewall rules and docs must match.' } else { Add-Check $T $c }
            if ($tm[0].State -ne 'STARTED') { Add-Finding $T $c 'HADR endpoint state' $rm[0].State $tm[0].State $(if ($isHa) { 'High' } else { 'Low' }) -Action 'ALTER ENDPOINT ... STATE = STARTED.' } else { Add-Check $T $c }
        } elseif ($isMig -and $rm.Count -and -not $tm.Count) { Add-Finding $T $c 'HADR / mirroring endpoint' "$($rm[0].Name) port $($rm[0].Port)" 'missing' 'Low' -Action 'Needed only if the new server will join an AG / mirroring.' -Fix 'Copy-DbaEndpoint' }
        Compare-NamedSet $T $c 'Endpoint' @((ConvertTo-Arr $R.Sql.Endpoints) | Where-Object { $_.Type -ne 'DATABASE_MIRRORING' }) @((ConvertTo-Arr $T.Sql.Endpoints) | Where-Object { $_.Type -ne 'DATABASE_MIRRORING' }) { "$($args[0].Name) ($($args[0].Type))" } 'Medium' -Action 'Create the endpoint (Service Broker / TSQL) on the target.' -Fix 'Copy-DbaEndpoint'
    }

    # ================= Server objects
    $c = 'Server objects'
    if ($R.SqlOk -and $T.SqlOk) {
        if (Test-Sec $R $T 'Sql' 'startup_procs' $c 'Startup procedures') { Compare-NamedSet $T $c 'Startup procedure' @((ConvertTo-Arr $R.Sql.StartupProcs) | ForEach-Object { [pscustomobject]@{ Name = $_ } }) @((ConvertTo-Arr $T.Sql.StartupProcs) | ForEach-Object { [pscustomobject]@{ Name = $_ } }) { $args[0].Name } 'High' -Action 'Create the procedure in master and mark it with sp_procoption ... startup.' -Fix 'Copy-DbaStartupProcedure' }
        if (Test-Sec $R $T 'Sql' 'sysdb_objects' $c 'User objects in master/msdb') { Compare-NamedSet $T $c 'User object in system DB' $R.Sql.SysDbObjects $T.Sql.SysDbObjects { "$($args[0].Db).$($args[0].Name)" } 'Medium' -GroupAbove 8 -Action 'Forgotten DBA/helper objects in master/msdb (sp_WhoIsActive, maintenance procs...). Copy if still used.' -Fix 'Copy-DbaSysDbUserObject' -Show { $args[0].Type } }
        if (Test-Sec $R $T 'Sql' 'extended_procs' $c 'Extended stored procedures') { Compare-NamedSet $T $c 'Extended stored procedure' $R.Sql.XProcs $T.Sql.XProcs { $args[0].Name } 'High' -Action 'Copy the DLL (see Files & paths) and register with sp_addextendedproc. Consider replacing with CLR/PowerShell - extended procs are deprecated.' -Show { $args[0].Dll } }
        if (Test-Sec $R $T 'Sql' 'messages' $c 'Custom error messages') { Compare-NamedSet $T $c 'Custom message' $R.Sql.Messages $T.Sql.Messages { "$($args[0].Id) (lang $($args[0].Lang))" } 'Medium' -GroupAbove 8 -Action 'Add with sp_addmessage (prep script section 5) - RAISERROR(n, ...) calls fail without it.' -Fix 'Copy-DbaCustomError' -Show { "sev $($args[0].Severity)" } -Diff { if ($args[0].Text -ne $args[1].Text) { , @('text A', 'text differs', 'Message text differs.') } } -DiffSev 'Low' }
        if (Test-Sec $R $T 'Sql' 'xe_sessions' $c 'Extended Events sessions') { Compare-NamedSet $T $c 'XE session' $R.Sql.XeSessions $T.Sql.XeSessions { $args[0].Name } 'Low' -GroupAbove 5 -Action 'Create the XE session (Copy-DbaXESession) if it is still used for monitoring.' -Fix 'Copy-DbaXESession' -Show { if ($args[0].AutoStart) { 'starts with server' } else { 'manual start' } } }
        if (Test-Sec $R $T 'Sql' 'resource_governor' $c 'Resource Governor') {
            $rg = $R.Sql.ResourceGovernor; $tg = $T.Sql.ResourceGovernor
            if ([int]$rg.Enabled -ne [int]$tg.Enabled) { Add-Finding $T $c 'Resource Governor' "enabled=$($rg.Enabled) classifier=$($rg.Classifier)" "enabled=$($tg.Enabled) classifier=$($tg.Classifier)" 'Medium' -Action 'Copy pools, groups and the classifier function (Copy-DbaResourceGovernor), then ALTER RESOURCE GOVERNOR RECONFIGURE.' -Fix 'Copy-DbaResourceGovernor' } else { Add-Check $T $c }
            Compare-NamedSet $T $c 'Resource pool' $rg.Pools $tg.Pools { $args[0].Name } 'Medium' -Action 'Create the resource pool.' -Fix 'Copy-DbaResourceGovernor'
            Compare-NamedSet $T $c 'Workload group' $rg.Groups $tg.Groups { $args[0].Name } 'Medium' -Action 'Create the workload group.' -Fix 'Copy-DbaResourceGovernor'
        }
    }

    # ================= Databases
    $c = 'Databases'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'databases' $c 'Databases')) {
        $tDb = @{}; foreach ($d in $T.Sql.Databases) { $tDb[$d.Name] = $d }
        $minC = Get-MinCompat $tMaj
        foreach ($d in @((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Id -gt 4 })) {
            if ($isHa) {
                if ([int]$d.InAg -eq 1 -and -not $tDb.ContainsKey($d.Name)) { Add-Finding $T $c "AG database: $($d.Name)" 'in AG on primary' 'not on this replica' 'High' -Action 'Join/seed the database on this replica.' }
                elseif ([int]$d.InAg -eq 0 -and $d.Name -ne 'SSISDB' -and -not $tDb.ContainsKey($d.Name) -and $d.Distributor -eq 0) { Add-Finding $T $c "Database not protected: $($d.Name)" 'not in an AG' 'absent' 'Info' -Action 'This database does not fail over with the AG. Intended?' }
                else { Add-Check $T $c }
                continue
            }
            if ($tDb.ContainsKey($d.Name)) { Add-Finding $T $c "Database exists on target: $($d.Name)" "compat $($d.Compat), $(Format-MB $d.SizeMB)" "compat $($tDb[$d.Name].Compat), state $($tDb[$d.Name].State)" 'Low' -Action 'A database with this name already exists on the target (test restore?). The cut-over restore will need WITH REPLACE - confirm it is disposable.' } else { Add-Check $T $c }
            if ($minC -gt 0 -and [int]$d.Compat -lt $minC) { Add-Finding $T $c "Compatibility level: $($d.Name)" $d.Compat "minimum on SQL $($TI.VersionName) is $minC" 'High' -Action "Restore will raise the database to $minC. Test the application at that level first (new CE, deprecated syntax)." } else { Add-Check $T $c }
            if ($d.Owner -and $T.Sql.Sections['login_names'] -like 'ok*' -and $T.Sql.Logins -notcontains $d.Owner) { Add-Finding $T $c "Database owner: $($d.Name)" $d.Owner 'login missing' 'Medium' -Action 'Owner login does not exist on the target - restored DB will show an orphaned owner SID. Create the login (LoginSync) or ALTER AUTHORIZATION after restore.' }
            if ([int]$d.Trustworthy -eq 1 -and -not @((ConvertTo-Arr $R.Sql.Assemblies) | Where-Object { $_.Db -eq $d.Name }).Count) { Add-Finding $T $c "TRUSTWORTHY: $($d.Name)" 'ON' 'reset to OFF by restore' 'High' -Action 'Re-enable TRUSTWORTHY after restore (cross-database EXECUTE AS / impersonation depends on it), or better, sign the modules.' }
            if ([int]$d.DbChaining -eq 1) { Add-Finding $T $c "DB_CHAINING: $($d.Name)" 'ON' 'verify after restore' 'Medium' -Action 'Cross-database ownership chaining: confirm it is ON after restore and that the owners match across databases.' }
            if ([int]$d.Cdc -eq 1) { Add-Finding $T $c "Change Data Capture: $($d.Name)" 'enabled' 'needs KEEP_CDC + jobs' 'High' -Action 'Restore WITH KEEP_CDC and recreate the capture/cleanup jobs (sys.sp_cdc_add_job). Agent must run on the target.' }
            if ([int]$d.Published -gt 0 -or [int]$d.Subscribed -eq 1) { Add-Finding $T $c "Replication: $($d.Name)" $(if ($d.Published -gt 0) { 'published' } else { 'subscriber' }) 're-establish' 'High' -Action 'Replication does not move with a backup. Script publications/subscriptions, plan the distributor, use KEEP_REPLICATION only for log-shipping style moves.' }
            if ([int]$d.Broker -eq 1) { Add-Finding $T $c "Service Broker: $($d.Name)" 'enabled' 'verify after restore' 'Medium' -Action 'After restore check is_broker_enabled (ALTER DATABASE ... SET ENABLE_BROKER, or NEW_BROKER for a copy), routes and the broker endpoint.' }
            foreach ($x in @((ConvertTo-Arr $R.Sql.CrossDbRefs) | Where-Object { $_.Db -eq $d.Name })) {
                if (@((ConvertTo-Arr $R.Sql.Databases) | Where-Object { $_.Name -eq $x.RefDb }).Count) { Add-Finding $T $c "Cross-database dependency: $($d.Name) -> $($x.RefDb)" 'same instance' $(if ($tDb.ContainsKey($x.RefDb)) { 'present' } else { 'move together' }) 'Info' -Action 'Code in one database references the other by 3-part name; migrate them in the same wave (and keep owners/chaining consistent).' }
            }
            if ($d.Name -eq 'SSISDB') { Add-Finding $T $c 'SSISDB database' 'present' '' 'Medium' -Action 'Prefer redeploying projects to a new catalog. If restoring SSISDB: need the catalog master key password, CLR, the ##MS_SSISServerCleanupJobLogin## login and the maintenance job.' }
            if ($d.State -ne 'ONLINE') { Add-Finding $T $c "Database state: $($d.Name)" $d.State '' 'Info' -Action 'Not ONLINE on the reference; it was not scanned for features.' }
        }
        $mailDbs = @((ConvertTo-Arr $R.Sql.CodeUse) | Where-Object { $_.DbMail -gt 0 } | ForEach-Object { $_.Db })
        if ($mailDbs.Count -and -not @($T.Sql.Mail).Count) { Add-Finding $T $c 'Code sends Database Mail' ($mailDbs -join ', ') 'no mail profile on target' 'High' -Action 'Configure Database Mail on the target before cut-over (Copy-DbaDbMail).' -Fix 'Copy-DbaDbMail' }
    }

    # ================= Disk layout, tempdb, system databases
    $c = 'Hardware & storage'
    if ($R.OsOk -and $T.OsOk -and $R.Os.Sections['volumes'] -like 'ok*' -and $T.Os.Sections['volumes'] -like 'ok*') {
        $sqlVols = @{}
        foreach ($f in (ConvertTo-Arr $R.Sql.Files)) { $v = Get-VolumeFor (ConvertTo-Arr $R.Os.Volumes) ([string]$f.Path); if ($v) { $sqlVols[[string]$v.Name] = 1 } }
        foreach ($rv in (ConvertTo-Arr $R.Os.Volumes)) {
            $name = if ($isHa) { [string]$rv.Name } else { Convert-MappedPath ([string]$rv.Name) }
            $tv = ConvertTo-Arr $T.Os.Volumes | Where-Object { ([string]$_.Name).TrimEnd('\') -eq $name.TrimEnd('\') } | Select-Object -First 1
            $isSql = $sqlVols.ContainsKey([string]$rv.Name)
            if (-not $tv) {
                if (@($script:Findings | Where-Object { $_.Target -eq $T.Name -and $_.Item -like "Volume $name*" }).Count) { continue }
                Add-Finding $T $c "Disk layout: volume $name" ("{0} {1}, {2}" -f $rv.Label, (Format-MB $rv.SizeMB), $(if ($isSql) { 'holds SQL files' } else { 'no SQL files' })) 'missing' $(if ($isSql -and $isHa) { 'High' } elseif ($isSql) { 'Medium' } else { 'Low' }) -Action $(if ($isSql) { 'Volume holds SQL files (system DBs/tempdb/user DBs) on the reference. Create it with the same letter/mount point, or plan the moves (-PathMap).' } else { 'Volume exists only on the reference (backups, apps, staging?). Confirm nothing on the new server expects it.' })
                continue
            }
            Add-Check $T $c
            if ($isSql -and [int64]$tv.SizeMB -lt [int64]$rv.SizeMB * 0.9) { Add-Finding $T $c "Disk layout: size of $name" (Format-MB $rv.SizeMB) (Format-MB $tv.SizeMB) 'Low' -Action 'SQL volume is smaller than on the reference. Check growth headroom.' }
            if ($isSql -and $tv.BlockSize -and [int]$tv.BlockSize -ne 65536 -and -not @($script:Findings | Where-Object { $_.Target -eq $T.Name -and $_.Item -eq "Allocation unit on $([string]$tv.Name)" }).Count) { Add-Finding $T $c "Allocation unit on $([string]$tv.Name)" '64 KB recommended' ('{0} KB' -f ([int]$tv.BlockSize / 1024)) 'Low' -Action 'Format SQL data/log/tempdb volumes with 64 KB allocation units (reformat before data lands).' }
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'files' 'SQL instance' 'System database / tempdb files')) {
        $c = 'SQL instance'
        $rt = @(ConvertTo-Arr $R.Sql.Files | Where-Object { $_.DbId -eq 2 -and $_.Type -eq 'ROWS' }); $tt = @(ConvertTo-Arr $T.Sql.Files | Where-Object { $_.DbId -eq 2 -and $_.Type -eq 'ROWS' })
        $tCores = [int](Get-Prop $TI 'CpuCount' 0)
        if ($tt.Count -lt $rt.Count) { Add-Finding $T $c 'tempdb data files' "$($rt.Count) files, $(Format-MB (Get-Sum $rt SizeMB))" "$($tt.Count) files, $(Format-MB (Get-Sum $tt SizeMB))" 'Medium' -Action 'Add tempdb data files (equal size and growth) to match the reference - allocation contention otherwise.' }
        elseif ($tCores -gt 0 -and $tt.Count -lt [Math]::Min(8, $tCores)) { Add-Finding $T $c 'tempdb data files (target baseline)' "$($rt.Count) on reference" "$($tt.Count) for $tCores cores" 'Low' -Action 'Use one tempdb data file per core up to 8, same size and growth.' }
        else { Add-Check $T $c }
        $rSum = [int64](Get-Sum $rt SizeMB); $tSum = [int64](Get-Sum $tt SizeMB)
        if ($rSum -gt 0 -and $tSum -lt $rSum * 0.5) { Add-Finding $T $c 'tempdb pre-sized data' (Format-MB $rSum) (Format-MB $tSum) 'Low' -Action 'Pre-size tempdb like the reference so the workload does not wait on autogrowth after cut-over/failover.' } else { Add-Check $T $c }
        if (@($tt | Select-Object -ExpandProperty Growth -Unique).Count -gt 1 -or @($tt | Where-Object { [string]$_.Growth -like '*%' }).Count) { Add-Finding $T $c 'tempdb growth settings' (@($rt | Select-Object -ExpandProperty Growth -Unique) -join ', ') (@($tt | Select-Object -ExpandProperty Growth -Unique) -join ', ') 'Low' -Action 'Use the same fixed MB growth for every tempdb data file.' } else { Add-Check $T $c }
        foreach ($db in @(@{ Id = 1; N = 'master' }, @{ Id = 2; N = 'tempdb' }, @{ Id = 3; N = 'model' }, @{ Id = 4; N = 'msdb' })) {
            $rDirs = @(ConvertTo-Arr $R.Sql.Files | Where-Object { $_.DbId -eq $db.Id } | ForEach-Object { "$($_.Type): " + (Get-WinParent ([string]$_.Path)) } | Sort-Object -Unique)
            $tDirs = @(ConvertTo-Arr $T.Sql.Files | Where-Object { $_.DbId -eq $db.Id } | ForEach-Object { "$($_.Type): " + (Get-WinParent ([string]$_.Path)) } | Sort-Object -Unique)
            $onC = @(ConvertTo-Arr $T.Sql.Files | Where-Object { $_.DbId -eq $db.Id -and ([string]$_.Path) -like 'C:\*' }).Count -gt 0
            if (($rDirs -join ';') -ne ($tDirs -join ';')) { Add-Finding $T $c "File layout: $($db.N)" ($rDirs -join '; ') ($tDirs -join '; ') $(if ($onC -and $db.N -eq 'tempdb') { 'Medium' } elseif ($onC) { 'Low' } else { 'Info' }) -Action $(if ($onC) { "$($db.N) files sit on the C: drive on the target. Move them to the SQL data/log volumes like the reference (ALTER DATABASE ... MODIFY FILE; restart)." } else { 'Different location than on the reference - fine if intended (update runbooks/monitoring).' }) } else { Add-Check $T $c }
        }
        $rm = ConvertTo-Arr $R.Sql.Databases | Where-Object { $_.Name -eq 'model' } | Select-Object -First 1; $tmod = ConvertTo-Arr $T.Sql.Databases | Where-Object { $_.Name -eq 'model' } | Select-Object -First 1
        if ($rm -and $tmod -and $rm.Recovery -ne $tmod.Recovery) { Add-Finding $T $c 'model database recovery model' $rm.Recovery $tmod.Recovery 'Low' -Action 'New databases inherit model settings (recovery model, size, growth). Align model with the reference.' }
    }

    # ================= Agent: schedules, step definitions, categories, proxies detail, mail detail, Agent settings
    $c = 'Agent & automation'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'schedules' $c 'Job schedules')) {
        $tJobNames = @(ConvertTo-Arr $T.Sql.Jobs | ForEach-Object { $_.Name })
        foreach ($g in (ConvertTo-Arr $R.Sql.Schedules | Where-Object { $tJobNames -contains $_.Job } | Group-Object Job)) {
            $rOn = @($g.Group | Where-Object { [int]$_.Enabled -eq 1 } | ForEach-Object { $_.Text } | Sort-Object -Unique)
            $tAll = @(ConvertTo-Arr $T.Sql.Schedules | Where-Object { $_.Job -eq $g.Name })
            $tOn = @($tAll | Where-Object { [int]$_.Enabled -eq 1 } | ForEach-Object { $_.Text } | Sort-Object -Unique)
            if ($rOn.Count -eq 0) { Add-Check $T $c; continue }
            if ($tAll.Count -eq 0) { Add-Finding $T $c "Job schedule: $($g.Name)" ($rOn -join ' | ') 'no schedule' 'High' -Action 'Job exists but has no schedule on the target - it will never run. Attach the schedule (Copy-DbaAgentSchedule / Copy-DbaAgentJob -Force).' -Fix 'Copy-DbaAgentSchedule'; continue }
            if ($tOn.Count -eq 0) { Add-Finding $T $c "Job schedule: $($g.Name)" ($rOn -join ' | ') ('disabled: ' + ((@($tAll | ForEach-Object { $_.Text }) | Sort-Object -Unique) -join ' | ')) $(if ($isHa) { 'Low' } else { 'Medium' }) -Action $(if ($isHa) { 'Schedules disabled on this replica. Fine if failover runbooks enable them; most teams keep schedules on and test "is primary" in step 1.' } else { 'Enable the schedule at cut-over (and disable it on the old server).' }); continue }
            $miss = @($rOn | Where-Object { $tOn -notcontains $_ })
            if ($miss.Count) { Add-Finding $T $c "Job schedule: $($g.Name)" ($rOn -join ' | ') ($tOn -join ' | ') 'Medium' -Action 'Schedule differs from the reference (time/frequency). Align it, or confirm the change is intended.' -Fix 'Copy-DbaAgentSchedule' } else { Add-Check $T $c }
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'job_steps' $c 'Job step definitions')) {
        foreach ($j in $R.Sql.JobSignatures.Keys) {
            if (-not $T.Sql.JobSignatures.ContainsKey($j)) { continue }
            if ($R.Sql.JobSignatures[$j] -ne $T.Sql.JobSignatures[$j]) { Add-Finding $T $c "Job steps differ: $j" "signature $($R.Sql.JobSignatures[$j])" "signature $($T.Sql.JobSignatures[$j])" $(if ($T.Role -eq 'Synchronous') { 'High' } else { 'Medium' }) -Action 'Same job name, different step commands/databases/proxies/flow. The target runs different logic - re-copy the job (Copy-DbaAgentJob -Force) after reviewing which side is current.' -Fix 'Copy-DbaAgentJob -Force' } else { Add-Check $T $c }
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'job_categories' $c 'Job categories')) { Compare-NamedSet $T $c 'Job category' $R.Sql.JobCategories $T.Sql.JobCategories { "$($args[0].Name)" } 'Low' -Action 'Custom job/alert/operator category missing (job copy fails or lands in Uncategorized).' -Fix 'Copy-DbaAgentJobCategory' -GroupAbove 6 }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'proxy_subsystems' $c 'Proxy subsystems')) {
        $tProx = @(ConvertTo-Arr $T.Sql.Proxies | ForEach-Object { $_.Name })
        Compare-NamedSet $T $c 'Proxy subsystem' @(ConvertTo-Arr $R.Sql.ProxySubsystems | Where-Object { $tProx -contains $_.Proxy }) $T.Sql.ProxySubsystems { "$($args[0].Proxy) -> $($args[0].Subsystem)" } 'High' -Action 'The proxy exists but is not granted this subsystem - job steps of that type using it fail (sp_grant_proxy_to_subsystem).'
        if (Test-Sec $R $T 'Sql' 'proxy_logins' $c 'Proxy principals') { Compare-NamedSet $T $c 'Proxy principal' @(ConvertTo-Arr $R.Sql.ProxyLogins | Where-Object { $tProx -contains $_.Proxy }) $T.Sql.ProxyLogins { "$($args[0].Proxy) -> $($args[0].Principal)" } 'Medium' -Action 'Login/role allowed to use the proxy on the reference only (sp_grant_login_to_proxy). Non-sysadmin job owners fail without it.' }
    }
    $mailUsed = (@(ConvertTo-Arr $R.Sql.Operators | Where-Object { $_.Email }).Count -gt 0) -or (@(ConvertTo-Arr $R.Sql.CodeUse | Where-Object { $_.DbMail -gt 0 }).Count -gt 0)
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'dbmail_accounts' $c 'Database Mail accounts')) {
        Compare-NamedSet $T $c 'Database Mail account' $R.Sql.MailAccounts $T.Sql.MailAccounts { $args[0].Account } $(if ($mailUsed) { 'High' } else { 'Medium' }) -Action 'Create the mail account (SMTP password must be re-entered) and add it to the profile.' -Fix 'Copy-DbaDbMail' -Show { "$($args[0].Email) via $($args[0].SmtpServer):$($args[0].Port) ssl=$($args[0].Ssl)" } -Diff { $a = $args[0]; $b = $args[1]; $ka = "$($a.Email)|$($a.SmtpServer):$($a.Port)|ssl=$($a.Ssl)|$(if ($a.DefaultCreds) { 'windows' } elseif ($a.SmtpUser) { 'basic:' + $a.SmtpUser } else { 'anonymous' })"; $kb = "$($b.Email)|$($b.SmtpServer):$($b.Port)|ssl=$($b.Ssl)|$(if ($b.DefaultCreds) { 'windows' } elseif ($b.SmtpUser) { 'basic:' + $b.SmtpUser } else { 'anonymous' })"; if ($ka -ne $kb) { , @($ka, $kb, 'Account settings differ (address, relay, port, SSL or SMTP authentication). Confirm the target relays mail.') } } -DiffSev 'Medium'
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'dbmail_principals' $c 'Database Mail profile security')) {
        $rDef = @(ConvertTo-Arr $R.Sql.MailPrincipals | Where-Object { [int]$_.IsDefault -eq 1 -and $_.Principal -eq 'public' })
        $tDef = @(ConvertTo-Arr $T.Sql.MailPrincipals | Where-Object { [int]$_.IsDefault -eq 1 -and $_.Principal -eq 'public' })
        if ($rDef.Count -and -not $tDef.Count) { Add-Finding $T $c 'Default public mail profile' $rDef[0].Profile 'none' 'Medium' -Action 'sp_send_dbmail calls without @profile_name fail ("no global profile"). Mark the profile public + default (sysmail_add_principalprofile_sp).' } else { Add-Check $T $c }
        Compare-NamedSet $T $c 'Mail profile access' @(ConvertTo-Arr $R.Sql.MailPrincipals | Where-Object { $_.Principal -ne 'public' }) @(ConvertTo-Arr $T.Sql.MailPrincipals | Where-Object { $_.Principal -ne 'public' }) { "$($args[0].Profile) -> $($args[0].Principal)" } 'Low' -Action 'Private profile granted to this principal on the reference only.'
    }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'dbmail_config' $c 'Database Mail configuration')) {
        Compare-NamedSet $T $c 'Database Mail setting' $R.Sql.MailConfig $T.Sql.MailConfig { $args[0].Name } 'Low' -Action 'Set with sysmail_configure_sp.' -Diff { if ([string]$args[0].Value -ne [string]$args[1].Value) { , @([string]$args[0].Value, [string]$args[1].Value, 'Database Mail parameter differs (attachment size, prohibited extensions, retries, logging). Align with sysmail_configure_sp.') } } -DiffSev 'Low'
    }
    if ($R.OsOk -and $T.OsOk -and $R.Os.Sections['agent_reg'] -like 'ok*' -and $T.Os.Sections['agent_reg'] -like 'ok*') {
        $ra = $R.Os.AgentReg; $ta = $T.Os.AgentReg
        $rProf = [string]$ra['DatabaseMailProfile']; $tProf = [string]$ta['DatabaseMailProfile']
        if ($rProf -and ([string]$ta['UseDatabaseMail'] -ne '1' -or -not $tProf)) { Add-Finding $T $c 'Agent alert system: mail profile' "Database Mail, profile '$rProf'" $(if ($tProf) { "profile '$tProf', Database Mail off" } else { 'not configured' }) $(if (@(ConvertTo-Arr $R.Sql.Operators).Count) { 'High' } else { 'Medium' }) -Action 'SQL Agent > Properties > Alert System: enable the mail profile and restart Agent. Until then job/alert notifications are silently not sent.' }
        elseif ($rProf -and $tProf -ne $rProf) { Add-Finding $T $c 'Agent alert system: mail profile' $rProf $tProf 'Low' -Action 'Agent uses a different mail profile. Confirm it exists and works.' }
        else { Add-Check $T $c }
        if ([string]$ra['AlertFailSafeOperator'] -and [string]$ra['AlertFailSafeOperator'] -ne [string]$ta['AlertFailSafeOperator']) { Add-Finding $T $c 'Agent fail-safe operator' $ra['AlertFailSafeOperator'] $ta['AlertFailSafeOperator'] 'Low' -Action 'Set the fail-safe operator in Agent > Alert System.' } else { Add-Check $T $c }
    }
    if ($R.SqlOk -and $T.SqlOk -and $R.Sql.AgentProps.Count -and $T.Sql.AgentProps.Count) {
        foreach ($k in @('JobHistoryMaxRows', 'JobHistoryMaxRowsPerJob')) {
            if ([string]$R.Sql.AgentProps[$k] -ne [string]$T.Sql.AgentProps[$k]) { Add-Finding $T $c "Agent setting: $k" $R.Sql.AgentProps[$k] $T.Sql.AgentProps[$k] 'Low' -Action 'Job history retention differs (Agent > Properties > History).' } else { Add-Check $T $c }
        }
    }

    # ================= Linked server options + provider options
    $c = 'Linked servers & external'
    if ($R.SqlOk -and $T.SqlOk -and $R.Sql.Sections['linked_servers'] -like 'ok*' -and $T.Sql.Sections['linked_servers'] -like 'ok*') {
        foreach ($l in (ConvertTo-Arr $R.Sql.LinkedServers)) {
            $tl = ConvertTo-Arr $T.Sql.LinkedServers | Where-Object { $_.Name -eq $l.Name } | Select-Object -First 1
            if (-not $tl) { continue }
            $diff = @()
            foreach ($o in @('RpcOut', 'DataAccess', 'Rpc', 'CollationCompatible', 'RemoteCollation', 'LazySchema', 'DtcPromotion', 'ConnectTimeout', 'QueryTimeout')) { if ([string](Get-Prop $l $o) -ne [string](Get-Prop $tl $o)) { $diff += "$o $(Get-Prop $l $o)->$(Get-Prop $tl $o)" } }
            if ($diff.Count) { Add-Finding $T $c "Linked server options: $($l.Name)" (($diff | ForEach-Object { ($_ -split '->')[0] }) -join ', ') (($diff | ForEach-Object { ($_ -split '->')[1] }) -join ', ') $(if ($diff -match '^(RpcOut|DataAccess) ') { 'Medium' } else { 'Low' }) -Action 'Align with sp_serveroption (rpc out = EXEC ... AT / remote procs, data access = 4-part queries, remote proc transaction promotion = DTC).' } else { Add-Check $T $c }
        }
        if ($R.OsOk -and $T.OsOk -and $null -ne $R.Os.ProviderOptions -and $null -ne $T.Os.ProviderOptions) {
            foreach ($prov in @(ConvertTo-Arr $R.Sql.LinkedServers | ForEach-Object { [string]$_.Provider } | Where-Object { $_ } | Sort-Object -Unique)) {
                $ro = $R.Os.ProviderOptions[$prov]; $to = $T.Os.ProviderOptions[$prov]
                if (-not $ro) { continue }
                foreach ($k in $ro.Keys) {
                    $tv = if ($to -and $to.Contains($k)) { $to[$k] } else { 0 }
                    if ([int]$ro[$k] -ne [int]$tv) { Add-Finding $T $c "Provider option: $prov / $k" $ro[$k] $tv $(if ($k -eq 'AllowInProcess' -and [int]$ro[$k] -eq 1) { 'High' } else { 'Medium' }) -Action $(if ($k -eq 'AllowInProcess') { 'Linked servers using this provider fail ("Cannot create an instance of OLE DB provider") without Allow inprocess. Set it under Server Objects > Linked Servers > Providers (or sp_MSset_oledb_prop).' } else { 'Provider option differs - set it under Linked Servers > Providers.' }) } else { Add-Check $T $c }
                }
            }
        }
    }

    # ================= Certificates (SQL TLS + machine), firewall, service accounts
    $c = 'Security objects'
    if ($T.OsOk -and $T.Os.Sections['certificates'] -like 'ok*') {
        $now = Get-Date
        $tc = $T.Os.SqlCert
        if ($tc -and (Get-Prop $tc 'Missing')) { Add-Finding $T $c 'SQL TLS certificate' 'bound in Configuration Manager' "thumbprint $($tc.Thumbprint) not in LocalMachine\My" 'High' -Hard -Action 'SQL Server is configured for a certificate that is not in the machine store - the service may fail to start or fall back. Import the certificate (with private key) or fix the binding.' }
        elseif ($tc) {
            $days = [int]($tc.NotAfterDate - $now).TotalDays
            if ($days -lt 0) { Add-Finding $T $c 'SQL TLS certificate expiry' '' "expired $($tc.NotAfter)" 'High' -Action 'The certificate SQL Server uses for TLS has expired. Renew and rebind before cut-over (clients validating certificates will fail).' }
            elseif ($days -lt 45) { Add-Finding $T $c 'SQL TLS certificate expiry' '' "expires $($tc.NotAfter) ($days days)" 'Medium' -Action 'Renew the SQL TLS certificate before go-live.' } else { Add-Check $T $c }
            $names = @($tc.San) + @(([string]$tc.Subject -replace '^.*?CN=([^,]+).*$', '$1'))
            $fqdn = if ($T.Os.Info -and $T.Os.Info.Domain) { "$($T.Computer).$($T.Os.Info.Domain)" } else { $T.Computer }
            $want = @($T.Computer, $fqdn) + @(ConvertTo-Arr $T.Sql.Ag | Where-Object { $_.Listener } | ForEach-Object { $_.Listener } | Select-Object -Unique)
            $missNames = @($want | Where-Object { $_ } | Where-Object { $n = [string]$_; -not @($names | Where-Object { $sn = [string]$_; $sn -ieq $n -or $sn -like ($n + '.*') -or ($sn -like '`*.*' -and $n -like ('*' + $sn.Substring(1))) }).Count } | Select-Object -Unique)
            if ($missNames.Count) { Add-Finding $T $c 'SQL TLS certificate names' (($names | Select-Object -Unique) -join ', ') ("missing: " + ($missNames -join ', ')) 'Medium' -Action 'Clients connecting by these names get a certificate name mismatch (Encrypt=Strict/True with validation). Reissue with the host FQDN, short name and listener/CNAME in the SAN.' } else { Add-Check $T $c }
        }
        $rc = $R.Os.SqlCert
        if ($rc -and -not (Get-Prop $rc 'Missing') -and $rc.NotAfterDate -lt $now.AddDays(60)) { Add-Finding $T $c 'Reference SQL TLS certificate expiry' "expires $($rc.NotAfter)" '' 'Info' -Action 'Reference certificate expires soon - do not copy it to the new server; issue a new one.' }
        if ($R.OsOk -and $R.Os.Sections['certificates'] -like 'ok*') {
            $tSubj = @(ConvertTo-Arr $T.Os.Certificates | Where-Object { $_.NotAfterDate -gt $now } | ForEach-Object { $_.Subject })
            foreach ($mc in @(ConvertTo-Arr $R.Os.Certificates | Where-Object { $_.HasPrivateKey -and -not $_.SelfSigned -and $_.NotAfterDate -gt $now -and -not ($rc -and $_.Thumbprint -eq $rc.Thumbprint) })) {
                if ($tSubj -contains $mc.Subject) { Add-Check $T $c; continue }
                Add-Finding $T $c "Machine certificate: $($mc.Subject)" ("expires $($mc.NotAfter); $($mc.Eku)") 'no valid equivalent' 'Low' -Action 'Certificate with private key in LocalMachine\My on the reference only (SSRS/HTTPS, endpoint auth, app/ETL client auth?). Issue an equivalent for the new host if something uses it.' -Detail ("issuer: $($mc.Issuer)")
            }
        }
    }
    $c = 'Windows features & services'
    if ($T.OsOk -and $T.Os.Sections['firewall'] -like 'ok*' -and $T.Os.Firewall -and [int]$T.Os.Firewall.ProfilesOn -gt 0) {
        $fw = $T.Os.Firewall; $progOk = @($fw.SqlProgramRules).Count -gt 0
        $need = @()
        $tp = [string](Get-Prop $T.Sql.Net 'TcpPort' ''); if ($tp) { $need += @{ P = "TCP:$tp"; W = 'SQL Server TCP port'; S = $(if ($isHa) { 'High' } else { 'Medium' }); Prog = $true } }
        foreach ($ep in @(ConvertTo-Arr $T.Sql.Endpoints | Where-Object { $_.Type -eq 'DATABASE_MIRRORING' -and $_.Port })) { $need += @{ P = "TCP:$($ep.Port)"; W = 'HADR endpoint port'; S = $(if ($isHa) { 'High' } else { 'Low' }); Prog = $true } }
        foreach ($ag in @(ConvertTo-Arr $T.Sql.Ag | Where-Object { Get-Prop $_ 'ListenerPort' } | Select-Object -First 1)) { $need += @{ P = "TCP:$(Get-Prop $ag 'ListenerPort')"; W = 'AG listener port'; S = 'Medium'; Prog = $true } }
        if ($TI.InstanceName -and $R.OsOk -and $R.Os.Firewall -and @($R.Os.Firewall.Ports) -contains 'UDP:1434') { $need += @{ P = 'UDP:1434'; W = 'SQL Browser (named instances)'; S = 'Low'; Prog = $false } }
        foreach ($n in $need) {
            if (@($fw.Ports) -contains $n.P -or ($n.Prog -and $progOk)) { Add-Check $T $c }
            else { Add-Finding $T $c "Firewall: inbound $($n.P)" $n.W 'no enabled allow rule' $n.S -Action 'Windows Firewall is on and no enabled inbound allow rule covers this port (or sqlservr.exe). Add one (New-NetFirewallRule) - otherwise remote clients/replicas cannot connect.' }
        }
        if ($R.OsOk -and $R.Os.Firewall) {
            $refOwn = @("TCP:$(Get-Prop $R.Sql.Net 'TcpPort' '')") + @(ConvertTo-Arr $R.Sql.Endpoints | Where-Object { $_.Port } | ForEach-Object { "TCP:$($_.Port)" }) + @('UDP:1434')
            foreach ($p in @($R.Os.Firewall.Ports | Where-Object { @($fw.Ports) -notcontains $_ -and $_ -notin @($need | ForEach-Object { $_.P }) -and $_ -notin $refOwn })) { Add-Finding $T $c "Firewall: inbound $p" 'allowed on reference' 'no rule' 'Low' -Action 'Port opened on the reference only (SSRS 80/443, app ports, other instances?). Open it if the new server needs it.' }
        }
    }
    if ($R.OsOk -and $T.OsOk -and $R.Os.Sections['services'] -like 'ok*' -and $T.Os.Sections['services'] -like 'ok*') {
        foreach ($fam in @('Database Engine', 'SQL Server Agent', 'Integration Services', 'Reporting Services', 'Analysis Services')) {
            if ($isHa -and $fam -eq 'Database Engine') { continue }
            $rs = ConvertTo-Arr $R.Os.Services | Where-Object { (Get-SvcFamily $_.Name) -eq $fam } | Select-Object -First 1
            $ts = ConvertTo-Arr $T.Os.Services | Where-Object { (Get-SvcFamily $_.Name) -eq $fam } | Select-Object -First 1
            if (-not $rs -or -not $ts) { continue }
            if ([string]$rs.Account -ne [string]$ts.Account) { Add-Finding $T $c "Service account: $fam" $rs.Account $ts.Account $(if ($fam -eq 'Database Engine' -and $rs.Account -match '\\' -and $rs.Account -notmatch '^(NT |\.\\|LocalSystem)') { 'Medium' } else { 'Low' }) -Action 'Different service account: grant it what the old one had - backup/file shares, proxies/credentials, SPNs (Kerberos for linked servers/double hop), Lock pages / volume maintenance rights, cert private-key access.' } else { Add-Check $T $c }
        }
    }

    # ================= Backup devices, maintenance plans
    $c = 'Server objects'
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'backup_devices' $c 'Backup devices')) { Compare-NamedSet $T $c 'Backup device' $R.Sql.BackupDevices $T.Sql.BackupDevices { $args[0].Name } 'Medium' -Action 'Jobs/scripts that BACKUP TO <device> fail without it.' -Fix 'Copy-DbaBackupDevice' -Show { "$($args[0].Type) $($args[0].Path)" } }
    if ($R.SqlOk -and $T.SqlOk -and (Test-Sec $R $T 'Sql' 'maint_plans' 'Agent & automation' 'Maintenance plans')) { Compare-NamedSet $T 'Agent & automation' 'Maintenance plan' $R.Sql.MaintPlans $T.Sql.MaintPlans { $args[0].Name } $(if ($isHa) { 'Medium' } else { 'High' }) -Action 'Maintenance plans (SSIS packages in msdb + their subplan jobs) are not moved by restoring user databases. Recreate them, export/import the package, or replace with Ola Hallengren jobs.' -Show { "owner $($args[0].Owner)" } }

    # ================= Support lifecycle, patching, reboot, OS policy
    $c = 'Platform & OS'
    $today = Get-Date
    if ($T.SqlOk -and $script:SqlSupportEnd.ContainsKey($tMaj)) {
        $end = [datetime]$script:SqlSupportEnd[$tMaj]
        if ($end -lt $today) { Add-Finding $T 'SQL instance' 'SQL Server support lifecycle' '' "SQL $($TI.VersionName) out of extended support since $($script:SqlSupportEnd[$tMaj])" $(if ($isMig) { 'High' } else { 'Medium' }) -Hard:$isMig -Action $(if ($isMig) { 'The migration target runs an unsupported SQL Server version (no security fixes). Install a supported version before moving databases onto it.' } else { 'The whole AG runs an unsupported SQL Server version. Plan the upgrade (rolling upgrade or migration to a new AG).' }) }
        elseif ($end -lt $today.AddDays(365)) { Add-Finding $T 'SQL instance' 'SQL Server support lifecycle' '' "SQL $($TI.VersionName) extended support ends $($script:SqlSupportEnd[$tMaj])" 'Medium' -Action 'Support ends within a year - a new build should normally land on a newer version.' }
        else { Add-Check $T 'SQL instance' }
    }
    if ($R.SqlOk -and $isMig -and $script:SqlSupportEnd.ContainsKey($rMaj) -and ([datetime]$script:SqlSupportEnd[$rMaj]) -lt $today) { Add-Finding $T 'SQL instance' 'Reference support lifecycle' "SQL $($RI.VersionName) out of support since $($script:SqlSupportEnd[$rMaj])" '' 'Info' -Action 'Source is out of support - include version-upgrade testing (compat level, deprecated features) in the plan.' }
    if ($T.OsOk -and $T.Os.Info -and [string]$T.Os.Info.Caption -match '(?i)server') {
        $b = [int]("0" + $T.Os.Info.Build)
        if ($script:OsSupportEnd.ContainsKey($b)) {
            $e = [datetime]$script:OsSupportEnd[$b][1]; $nm = $script:OsSupportEnd[$b][0]
            if ($e -lt $today) { Add-Finding $T $c 'Windows support lifecycle' '' "$nm out of extended support since $($script:OsSupportEnd[$b][1])" $(if ($isMig) { 'High' } else { 'Medium' }) -Action 'Target OS no longer receives security updates. Rebuild on a supported Windows Server version.' }
            elseif ($e -lt $today.AddDays(365)) { Add-Finding $T $c 'Windows support lifecycle' '' "$nm extended support ends $($script:OsSupportEnd[$b][1])" $(if ($isMig) { 'Medium' } else { 'Low' }) -Action 'OS support ends within a year - a migration target should normally be on a newer Windows Server.' }
            else { Add-Check $T $c }
        }
    }
    if ($T.OsOk -and $T.Os.Sections['os_detail'] -like 'ok*') {
        $pr = @(ConvertTo-Arr $T.Os.PendingReboot)
        if ($pr.Count) { Add-Finding $T $c 'Pending reboot' '' ($pr -join ', ') $(if (@($pr | Where-Object { $_ -ne 'Pending file rename operations' }).Count) { 'Medium' } else { 'Low' }) -Action 'Reboot the target before installing/patching SQL Server or cutting over (SQL setup fails its "restart required" rule; half-applied updates).' } else { Add-Check $T $c }
        $rp = $R.Os.Patch; $tp = $T.Os.Patch
        if ($R.OsOk -and $rp -and $tp -and [string]$R.Os.Info.Build -eq [string]$T.Os.Info.Build -and $rp.Ubr -and $tp.Ubr -and [int64]$tp.Ubr -lt [int64]$rp.Ubr) { Add-Finding $T $c 'Windows patch level' "$($R.Os.Info.Build).$($rp.Ubr) (last $($rp.LastHotfix))" "$($T.Os.Info.Build).$($tp.Ubr) (last $($tp.LastHotfix))" $(if ($isHa) { 'Medium' } else { 'Low' }) -Action 'Target is behind the reference on Windows cumulative updates. Patch it (cluster nodes should match).' } else { Add-Check $T $c }
        if ($tp -and $tp.LastHotfix -and ([datetime]$tp.LastHotfix) -lt $today.AddDays(-90)) { Add-Finding $T $c 'Windows updates (target baseline)' '' "last update installed $($tp.LastHotfix)" 'Low' -Action 'No Windows update in 90+ days. Patch before go-live.' }
        if ($R.OsOk -and $rp -and $tp -and [string]$rp.InstallationType -ne [string]$tp.InstallationType) { Add-Finding $T $c 'Windows installation type' $rp.InstallationType $tp.InstallationType 'Info' -Action 'Server Core vs Desktop Experience: GUI tools (SSMS, Configuration Manager UI, vendor consoles) behave differently.' }
        if ($R.OsOk -and $R.Os.Sections['os_detail'] -like 'ok*' -and [string]$R.Os.Fips -ne [string]$T.Os.Fips) { Add-Finding $T $c 'FIPS mode (Lsa\FipsAlgorithmPolicy)' $(if ([int]("0$($R.Os.Fips)") -eq 1) { 'enabled' } else { 'disabled' }) $(if ([int]("0$($T.Os.Fips)") -eq 1) { 'enabled' } else { 'disabled' }) 'Medium' -Action 'FIPS policy differs: with FIPS on, non-compliant crypto fails (some SSIS components, older .NET apps, SSRS, MD5 hashing in code). Align with the security baseline and test.' } else { Add-Check $T $c }
        if ($R.OsOk -and $T.Os.SqlInstances -and $R.Os.SqlInstances) {
            $ownInst = if (Get-Prop $RI 'InstanceName') { [string](Get-Prop $RI 'InstanceName') } else { 'MSSQLSERVER' }
            $key = { param($x) ([string]$x -replace '\s*\(.*\)$', '').ToLowerInvariant() }
            $tKeys = @(ConvertTo-Arr $T.Os.SqlInstances | ForEach-Object { & $key $_ })
            foreach ($i in @(ConvertTo-Arr $R.Os.SqlInstances | Where-Object { (& $key $_) -ne ("engine: " + $ownInst.ToLowerInvariant()) -and $tKeys -notcontains (& $key $_) })) { Add-Finding $T $c "Other SQL component on reference host: $i" 'installed' 'not on target' $(if ($isMig) { 'Medium' } else { 'Info' }) -Action 'Another instance/component lives on the reference host. Decide where it moves - it is not covered by this comparison (run the tool for that instance too).' }
        }
        $c2 = 'Windows features & services'
        $rd = $R.Os.Msdtc; $td = $T.Os.Msdtc
        $dtcUsed = (@(ConvertTo-Arr $R.Sql.CodeUse | Where-Object { (Get-Prop $_ 'DistTx' 0) -gt 0 }).Count -gt 0) -or (@(ConvertTo-Arr $R.Sql.LinkedServers | Where-Object { [int](Get-Prop $_ 'DtcPromotion' 0) -eq 1 }).Count -gt 0)
        if ($R.OsOk -and $rd -and $td) {
            $diffs = @($rd.Keys | Where-Object { [int]("0" + $rd[$_]) -eq 1 -and [int]("0" + $td[$_]) -ne 1 })
            if ($diffs.Count) { Add-Finding $T $c2 'MSDTC network access' (($diffs | ForEach-Object { "$_=1" }) -join ', ') 'not enabled' $(if ($dtcUsed) { 'High' } else { 'Medium' }) -Action ('Enable the same MSDTC security settings (Component Services > Local DTC > Security) and open RPC/DTC in the firewall.' + $(if ($dtcUsed) { ' Code uses BEGIN DISTRIBUTED TRANSACTION or linked servers promote to DTC - those transactions fail otherwise.' } else { '' })) } else { Add-Check $T $c2 }
        }
    }
    if ($T.OsOk -and $T.Os.Sections['environment'] -like 'ok*' -and $R.OsOk -and $R.Os.Sections['environment'] -like 'ok*') {
        $c2 = 'Installed software'
        $skipVars = '^(?i)(COMPUTERNAME|NUMBER_OF_PROCESSORS|PROCESSOR_.*|OS|TEMP|TMP|windir|USERNAME|ComSpec|PATHEXT|DriverData|PSModulePath|Path)$'
        $jobText = (@(ConvertTo-Arr $R.Sql.JobSteps | ForEach-Object { "$($_.Paths -join ' ')" }) -join ' ')
        foreach ($k in @($R.Os.EnvVars.Keys | Where-Object { $_ -notmatch $skipVars })) {
            if ($T.Os.EnvVars.Contains($k)) { if ([string]$T.Os.EnvVars[$k] -ne [string](Convert-MappedPath $R.Os.EnvVars[$k])) { Add-Finding $T $c2 "Environment variable: $k" (Protect-Secret $R.Os.EnvVars[$k]) (Protect-Secret $T.Os.EnvVars[$k]) 'Low' -Action 'Machine environment variable has a different value - scripts/tools reading it behave differently.' } else { Add-Check $T $c2 } }
            else { Add-Finding $T $c2 "Environment variable: $k" (Protect-Secret $R.Os.EnvVars[$k]) 'missing' 'Medium' -Install -Action 'Machine-level environment variable exists only on the reference (JAVA_HOME, ORACLE_HOME, TNS_ADMIN, vendor tools?). Set it on the target after installing the product it belongs to.' }
        }
        $rPath = @(([string](Get-Prop $R.Os.EnvVars 'Path' '')) -split ';' | ForEach-Object { $_.Trim().TrimEnd('\') } | Where-Object { $_ -and $_ -notmatch '(?i)\\windows(\\|$)|%systemroot%|%windir%' })
        $tPath = @(([string](Get-Prop $T.Os.EnvVars 'Path' '')) -split ';' | ForEach-Object { $_.Trim().TrimEnd('\') } | Where-Object { $_ })
        $missP = @($rPath | Where-Object { $tPath -notcontains (Convert-MappedPath $_) -and $tPath -notcontains $_ } | Where-Object { $_ -notmatch '(?i)Microsoft SQL Server\\\d{2,3}\\' })
        if ($missP.Count) { Add-Finding $T $c2 'PATH entries missing' ("$($missP.Count) folder(s)") 'missing on target' $(if ($missP.Count -gt 0 -and @($missP | Where-Object { $jobText -like "*$_*" }).Count) { 'Medium' } else { 'Low' }) -Install -Action 'Folders on the reference machine PATH (tools called by bare name from jobs/scripts: sqlcmd, bcp, 7z, java, oracle client...). Install the tool / add the folder.' -Detail ($missP -join '; ') } else { Add-Check $T $c2 }
    }
    if ($T.OsOk -and $T.Os.Sections['scheduled_tasks'] -like 'ok*' -and $R.OsOk -and $R.Os.Sections['scheduled_tasks'] -like 'ok*') {
        Compare-NamedSet $T 'Agent & automation' 'Windows scheduled task' @(ConvertTo-Arr $R.Os.ScheduledTasks | Where-Object { $_.State -ne 'Disabled' }) $T.Os.ScheduledTasks { $args[0].Name } 'Medium' -Action 'Automation outside SQL Agent (exports, file moves, reports, maintenance). Recreate it on the target (Export-ScheduledTask / Register-ScheduledTask) with the same run-as account, or move it into Agent.' -Show { "run as $($args[0].RunAs): $($args[0].Actions)" } -Diff { if ([string]$args[0].RunAs -ne [string]$args[1].RunAs) { , @("run as $($args[0].RunAs)", "run as $($args[1].RunAs)", 'Task runs under a different account on the target.') } } -DiffSev 'Low' -GroupAbove 15
    }
    if ($T.OsOk -and $T.Os.Sections['defender'] -like 'ok*' -and $T.Os.Defender -and $T.Os.Defender.RealTime) {
        $c2 = 'Windows features & services'; $d = $T.Os.Defender
        if ($d.Hidden) { Add-Finding $T $c2 'Defender exclusions' '' 'not readable (hidden for non-admins)' 'Low' -Action 'Run as a local administrator to verify SQL Server antivirus exclusions.' }
        else {
            $sqlDirs = @(ConvertTo-Arr $T.Sql.Files | ForEach-Object { Get-WinParent ([string]$_.Path) }) + @(ConvertTo-Arr $R.Sql.Files | Where-Object { $_.DbId -gt 4 } | ForEach-Object { Convert-MappedPath (Get-WinParent ([string]$_.Path)) }) + @((Get-Prop $TI 'DefaultBackupPath'), (Get-Prop $TI 'DefaultDataPath'), (Get-Prop $TI 'DefaultLogPath'))
            $sqlDirs = @($sqlDirs | Where-Object { $_ } | ForEach-Object { ([string]$_).TrimEnd('\') } | Sort-Object -Unique)
            $covered = { param($dir) @($d.Paths | Where-Object { $ex = ([string]$_).TrimEnd('\'); $dir -like "$ex*" }).Count -gt 0 }
            $unc = @($sqlDirs | Where-Object { -not (& $covered $_) })
            $procOk = @($d.Processes | Where-Object { $_ -match '(?i)sqlservr\.exe' }).Count -gt 0
            $extOk = @('mdf', 'ldf', 'ndf') | Where-Object { @($d.Extensions | ForEach-Object { ([string]$_).TrimStart('.').ToLower() }) -contains $_ }
            $refMiss = @(if ($R.OsOk -and $R.Os.Defender) { @($R.Os.Defender.Paths | Where-Object { @($d.Paths) -notcontains (Convert-MappedPath ([string]$_)) -and @($d.Paths) -notcontains $_ }) })
            $aggr = ($unc.Count -and @($extOk).Count -lt 3)
            if ($aggr) { Add-Finding $T $c2 'Defender exclusions for SQL files' $(if ($R.OsOk -and $R.Os.Defender) { (@($R.Os.Defender.Paths) | Select-Object -First 6) -join '; ' } else { 'n/a' }) ("not excluded: " + ($unc -join '; ')) 'Medium' -Action 'Real-time scanning of data/log/tempdb/backup folders causes I/O stalls and file locks. Add the SQL folders (or .mdf/.ndf/.ldf/.bak/.trn) and sqlservr.exe as exclusions per Microsoft guidance / your security team.' -Detail $(if ($refMiss.Count) { 'excluded on reference only: ' + ($refMiss -join '; ') } else { '' }) } else { Add-Check $T $c2 }
            if (-not $procOk) { Add-Finding $T $c2 'Defender process exclusion: sqlservr.exe' $(if ($R.OsOk -and $R.Os.Defender -and @($R.Os.Defender.Processes | Where-Object { $_ -match 'sqlservr' }).Count) { 'excluded' } else { 'n/a' }) 'not excluded' 'Low' -Action 'Exclude sqlservr.exe (and SQLAGENT.EXE, MsDtsSrvr.exe/DTExec.exe where used) from real-time scanning.' } else { Add-Check $T $c2 }
            if (-not $aggr) { foreach ($x in $refMiss) { Add-Finding $T $c2 "Defender exclusion: $x" 'excluded on reference' 'missing' 'Low' -Action 'Path excluded on the reference only. Add it if the same data lives there on the target.' } }
        }
    }
    if ($R.OsOk -and $T.OsOk -and $R.Os.Sections['local_admins'] -like 'ok*' -and $T.Os.Sections['local_admins'] -like 'ok*') {
        $c2 = 'Security objects'
        $svcAcc = @(ConvertTo-Arr $R.Os.Services | Where-Object { Get-SvcFamily $_.Name } | ForEach-Object { [string]$_.Account } | Where-Object { $_ -match '\\' -and $_ -notmatch '^(?i)(NT |\.\\|LocalSystem)' } | ForEach-Object { $_.ToLower() })
        foreach ($m in @(ConvertTo-Arr $R.Os.LocalAdmins | Where-Object { @($T.Os.LocalAdmins) -notcontains $_ -and $_ -notmatch "^(?i)$([regex]::Escape(([string]$R.Computer -split '\.')[0]))\\" })) {
            $isSvc = $svcAcc -contains ($m.ToLower()) -or @($svcAcc | Where-Object { ($_ -split '\\')[1] -eq (($m -split '\\')[1]).ToLower() }).Count
            Add-Finding $T $c2 "Local Administrators: $m" 'member on reference' 'not a member' $(if ($isSvc) { 'Medium' } else { 'Low' }) -Action $(if ($isSvc) { 'A SQL service account is local admin on the reference - that masks missing rights (file shares, Lock Pages, volume maintenance, cert keys). Grant the specific rights on the target instead of admin, and test.' } else { 'Group/account administers the reference host only (DBA team, monitoring, backup service?). Add it if still needed.' })
        }
    }
    if ($T.OsOk -and $T.Os.Sections['disks'] -like 'ok*') {
        $c2 = 'Hardware & storage'
        foreach ($dk in @(ConvertTo-Arr $T.Os.Disks | Where-Object { $_.Offline -or $_.ReadOnly -or $_.Health -ne 'Healthy' })) { Add-Finding $T $c2 "Disk $($dk.Number): $($dk.Name)" '' ("{0}{1}{2}" -f $(if ($dk.Offline) { 'OFFLINE ' } else { '' }), $(if ($dk.ReadOnly) { 'READ-ONLY ' } else { '' }), $dk.Health) 'High' -Action 'Disk is offline, read-only or not healthy on the target. Fix before placing SQL files on it.' }
        if (-not @(ConvertTo-Arr $T.Os.PageFiles).Count -and $T.Os.Sections['disks'] -like 'ok*') { Add-Finding $T $c2 'Page file' $(if ($R.OsOk) { (@($R.Os.PageFiles) -join ', ') } else { '' }) 'none configured' 'Low' -Action 'No page file: no full crash dumps and less headroom. Configure a page file per your standard.' }
    }
    foreach ($k in @($T.PathInfo.Keys)) {
        $pi = $T.PathInfo[$k]
        if ($pi.IsDir -and ($pi.Compressed -or $pi.Encrypted)) {
            $isSqlDir = @(ConvertTo-Arr $R.Sql.Files | Where-Object { (Convert-MappedPath (Get-WinParent ([string]$_.Path))) -eq $k }).Count -gt 0 -or $k -in @((Get-Prop $TI 'DefaultDataPath' ''), (Get-Prop $TI 'DefaultLogPath' '')).ForEach({ ([string]$_).TrimEnd('\') })
            Add-Finding $T 'Files & paths' "NTFS attributes: $k" '' ("{0}{1}" -f $(if ($pi.Compressed) { 'COMPRESSED ' } else { '' }), $(if ($pi.Encrypted) { 'EFS-ENCRYPTED' } else { '' })) $(if ($isSqlDir) { 'High' } else { 'Low' }) -Action 'SQL Server does not support data/log files in NTFS-compressed or EFS-encrypted folders (only read-only filegroups on compressed). Remove the attribute before restoring.'
        }
    }
    $reqAll = @(if ($T.PathInfo.Count -and $R.SqlOk) { Get-RequiredPaths -R $R -T $T })
    foreach ($k in @($T.PathInfo.Keys)) {
        $tv = ConvertTo-VersionSafe $T.PathInfo[$k].Version
        $rk = @($reqAll | Where-Object { $_.Path -eq $k } | Select-Object -First 1)
        if (-not $tv -or -not $rk.Count -or -not $R.PathInfo.Contains([string]$rk[0].RefPath)) { continue }
        $rv = ConvertTo-VersionSafe $R.PathInfo[[string]$rk[0].RefPath].Version
        if ($rv -and $tv -lt $rv) { Add-Finding $T 'Files & paths' "Older binary: $k" $rv.ToString() $tv.ToString() 'Low' -Action 'The file exists on the target but is an older version than on the reference. Copy/install the same version.' -Detail $rk[0].Context }
    }

    # ================= Kerberos SPNs and DNS
    $c = 'Security objects'
    if ($T.Os.Sections['spn'] -like 'ok*') {
        $tSp = @(ConvertTo-Arr $T.Spns); $rSp = @(ConvertTo-Arr $R.Spns)
        $tHost = ([string]$T.Computer -split '\.')[0].ToLower()
        $engAcc = [string](Get-Prop $TI 'ServiceAccount' '')
        $expAcc = if ($engAcc -match '^(?i)(NT Service\\|NT AUTHORITY\\|LocalSystem|\.\\)') { ($tHost + '$') } elseif ($engAcc -match '\\') { ($engAcc -split '\\')[1] } else { $engAcc }
        $hostSpns = @($tSp | Where-Object { $_.Spn -match ('^(?i)MSSQLSvc/' + [regex]::Escape($tHost) + '(\.|:|$)') })
        if (-not $hostSpns.Count) { Add-Finding $T $c 'Kerberos SPN (MSSQLSvc)' $(if ($rSp.Count) { (($rSp | Select-Object -First 3 | ForEach-Object { "$($_.Spn) on $($_.Account)" }) -join '; ') } else { 'none' }) 'none registered' $(if ($rSp.Count) { 'Medium' } else { 'Low' }) -Action ("No MSSQLSvc SPN for the target host: connections fall back to NTLM - linked servers with 'be made using the login's current security context', SSRS/SSIS double hop and constrained delegation break. Register: setspn -S MSSQLSvc/<fqdn>:<port> <service account> (and MSSQLSvc/<fqdn>), or let a domain service account with SPN rights self-register.") }
        else {
            $wrong = @($hostSpns | Where-Object { $expAcc -and $_.Account -and $_.Account.ToLower() -ne $expAcc.ToLower() })
            if ($wrong.Count) { Add-Finding $T $c 'Kerberos SPN account' "expected on $expAcc (engine runs as $engAcc)" (($wrong | ForEach-Object { "$($_.Spn) on $($_.Account)" }) -join '; ') 'Medium' -Action 'SPN is registered on a different account than the one SQL Server runs as - Kerberos fails (KRB_AP_ERR_MODIFIED) and clients fall back or fail. Move the SPN to the service account.' } else { Add-Check $T $c }
            $tp2 = [string](Get-Prop $T.Sql.Net 'TcpPort' '')
            if ($tp2 -and -not @($hostSpns | Where-Object { $_.Spn -like "*:$tp2" }).Count) { Add-Finding $T $c 'Kerberos SPN port' "MSSQLSvc/<host>:$tp2" (($hostSpns | ForEach-Object { $_.Spn }) -join '; ') 'Low' -Action 'No SPN for the static TCP port - TCP Kerberos connections need MSSQLSvc/<fqdn>:<port>.' }
        }
        $dups = @($tSp | Group-Object Spn | Where-Object { @($_.Group | Select-Object -ExpandProperty Account -Unique).Count -gt 1 })
        foreach ($dg in $dups) { Add-Finding $T $c "Duplicate SPN: $($dg.Name)" '' ((@($dg.Group | Select-Object -ExpandProperty Account -Unique)) -join ', ') 'High' -Action 'The same SPN is registered on several accounts - Kerberos fails for everyone using it. Remove the stale one (setspn -D).' }
        foreach ($ag in @(ConvertTo-Arr $T.Sql.Ag | Where-Object { $_.Listener } | Select-Object -ExpandProperty Listener -Unique)) {
            if (-not @($tSp | Where-Object { $_.Spn -match ('^(?i)MSSQLSvc/' + [regex]::Escape([string]$ag) + '(\.|:|$)') }).Count) { Add-Finding $T $c "Kerberos SPN for listener $ag" '' 'none' 'Medium' -Action 'Register MSSQLSvc/<listener fqdn>:<port> on the (common) domain service account of all replicas, otherwise Kerberos through the listener fails.' }
        }
    }
    if ($T.Dns) {
        $c = 'Platform & OS'
        if ($T.Dns.Error) { Add-Finding $T $c 'DNS resolution' '' $T.Dns.Error 'High' -Action 'Target name does not resolve from the machine running this check. Fix DNS before cut-over (clients use the same resolution).' }
        else {
            $tHost = ([string]$T.Computer -split '\.')[0]
            $bad = @($T.Dns.Reverse | Where-Object { ([string]$_ -split '\.')[0] -ne $tHost })
            if ($bad.Count) { Add-Finding $T $c 'Reverse DNS (PTR)' ($T.Dns.Ips -join ', ') ($T.Dns.Reverse -join ', ') 'Low' -Action 'PTR record missing or pointing to another name - Kerberos/SPN troubleshooting, some monitoring and licensing tools depend on it.' } else { Add-Check $T $c }
        }
    }

    # ================= Source health, keys, legacy HA features (SQL)
    $c = 'Databases'
    if ($R.SqlOk -and $R.Sql.Sections['db_health'] -like 'ok*') {
        $now = Get-Date
        $tDbN = @(ConvertTo-Arr $T.Sql.Databases | ForEach-Object { $_.Name })
        foreach ($h in (ConvertTo-Arr $R.Sql.DbHealth)) {
            if ($h.Db -in @('SSISDB') -and -not $isMig) { continue }
            if ([int]$h.SuspectPages -gt 0) { Add-Finding $T $c "Suspect pages: $($h.Db)" "$($h.SuspectPages) in msdb.dbo.suspect_pages" '' 'High' -Hard -Action 'The source has recorded corruption events. Run DBCC CHECKDB and resolve before migrating/seeding - corruption travels with the backup.' }
            if ($isMig) {
                if (-not $h.LastFull) { Add-Finding $T $c "No full backup: $($h.Db)" 'never backed up (per msdb)' '' 'High' -Action 'No full backup recorded. You need a verified full backup for restore-based migration and as rollback.' }
                elseif (([datetime]$h.LastFull) -lt $now.AddDays(-7)) { Add-Finding $T $c "Old full backup: $($h.Db)" "last full $(([datetime]$h.LastFull).ToString('yyyy-MM-dd'))" '' 'Medium' -Action 'Last full backup is older than 7 days. Take (and test-restore) a fresh full before cut-over; check the backup jobs.' }
                if ($h.Recovery -eq 'FULL' -and (-not $h.LastLog -or ([datetime]$h.LastLog) -lt $now.AddDays(-1))) { Add-Finding $T $c "Log backups: $($h.Db)" "FULL recovery, last log $(if ($h.LastLog) { ([datetime]$h.LastLog).ToString('yyyy-MM-dd HH:mm') } else { 'never' })" '' 'Medium' -Action 'No log backup in 24 h: the log grows and point-in-time recovery / log-shipping style cut-over is not possible.' }
            }
            if ($R.Sql.Sections['db_health'] -like 'ok*' -and (Get-Prop $RI 'Major' 0) -ge 13) {
                if (-not $h.LastCheckDb -or ([datetime]$h.LastCheckDb).Year -lt 1901) { Add-Finding $T $c "DBCC CHECKDB never run: $($h.Db)" 'no last-known-good' '' $(if ($isMig) { 'Medium' } else { 'Low' }) -Action 'Run DBCC CHECKDB on the source before migrating, so you know you are not moving corruption.' }
                elseif (([datetime]$h.LastCheckDb) -lt $now.AddDays(-14)) { Add-Finding $T $c "DBCC CHECKDB old: $($h.Db)" "last good $(([datetime]$h.LastCheckDb).ToString('yyyy-MM-dd'))" '' 'Low' -Action 'Last clean CHECKDB is older than 14 days. Run one before cut-over.' }
            }
            if ([int]$h.DmkBySmk -eq 1) { Add-Finding $T 'Security objects' "Database master key: $($h.Db)" 'encrypted by the service master key' 'not opened automatically on the new instance' 'High' -Action 'After restore/join on another instance the DMK is not auto-opened (different SMK): symmetric-key decryption, signed modules and credentials using it fail. Run OPEN MASTER KEY DECRYPTION BY PASSWORD = ...; ALTER MASTER KEY ADD ENCRYPTION BY SERVICE MASTER KEY (needs the DMK password - find it now).' }
            if ([int]$h.Mirrored -gt 0) { Add-Finding $T $c "Database mirroring: $($h.Db)" 'mirrored' '' $(if ($isMig) { 'High' } else { 'Medium' }) -Action 'Mirroring does not move with a backup and is deprecated. Re-establish (or replace with an AG) after migration; remove the mirroring session on the source at cut-over.' }
        }
    }
    if ($R.SqlOk -and $R.Sql.Sections['log_shipping'] -like 'ok*') {
        foreach ($l in (ConvertTo-Arr $R.Sql.LogShipping)) { Add-Finding $T $c "Log shipping ($($l.Role)): $($l.Db)" 'configured on reference' '' $(if ($isMig) { 'High' } else { 'Info' }) -Action 'Log shipping jobs/monitor are not carried by the database backup. Re-create the configuration (and secondaries) after the move, and disable it on the old server at cut-over.' }
    }
    foreach ($d in @(ConvertTo-Arr $R.Sql.Databases | Where-Object { [int]$_.Distributor -eq 1 })) { Add-Finding $T $c "Replication distributor: $($d.Name)" 'reference is a distributor' '' $(if ($isMig) { 'High' } else { 'Medium' }) -Action 'The reference hosts a distribution database: publishers/subscribers point at this server. Plan the distributor move (script replication, re-create on target, re-point publishers).' }
    if ($R.SqlOk -and $T.SqlOk -and $R.Sql.Sections['backup_encryption'] -like 'ok*' -and $T.Sql.Sections['certificates'] -like 'ok*') {
        $tTh = @(ConvertTo-Arr $T.Sql.Certificates | ForEach-Object { $_.Thumbprint }) + @(ConvertTo-Arr $T.Sql.AsymKeys | ForEach-Object { $_.Thumbprint })
        foreach ($bc in (ConvertTo-Arr $R.Sql.BackupCerts)) {
            if ($tTh -contains $bc.Thumbprint) { Add-Check $T 'Security objects'; continue }
            $cn = ConvertTo-Arr $R.Sql.Certificates | Where-Object { $_.Thumbprint -eq $bc.Thumbprint } | Select-Object -First 1
            Add-Finding $T 'Security objects' "Backup encryption certificate: $(if ($cn) { $cn.Name } else { $bc.Thumbprint.Substring(0, 12) + '...' })" "$($bc.Backups) encrypted backups in 60 days" 'missing on target' $(if ($isMig) { 'High' } else { 'Medium' }) -Hard:$isMig -Action 'Recent backups are encrypted with this certificate - the target cannot restore them without it. Copy the certificate with its private key (Copy-DbaDbCertificate / BACKUP CERTIFICATE ... WITH PRIVATE KEY).' -Fix 'Copy-DbaDbCertificate'
        }
    }
    if ($R.SqlOk -and $T.SqlOk -and [int](Get-Prop $RI 'MasterDmk' 0) -eq 1 -and [int](Get-Prop $TI 'MasterDmk' 1) -eq 0 -and (@(ConvertTo-Arr $R.Sql.Certificates).Count -gt 0)) { Add-Finding $T 'Security objects' 'master database master key' 'present' 'missing' 'Medium' -Action 'Create a database master key in master on the target before importing certificates with private keys (TDE/backup encryption).' }
    foreach ($cu in @(ConvertTo-Arr $R.Sql.CodeUse | Where-Object { (Get-Prop $_ 'SelfRef' 0) -gt 0 })) { Add-Finding $T $c "Hard-coded server name in code: $($cu.Db)" "$($cu.SelfRef) module(s) contain '$((@(Get-Prop $RI 'OwnNames' @()) | Where-Object { $_ }) -join "'/'")'" '' $(if ($isMig) { 'Medium' } else { 'Low' }) -Action 'Procedures/views/functions reference the old server name (4-part names, OPENQUERY, sp_send_dbmail bodies, dynamic SQL). After migration they keep pointing at the old server. Search sys.sql_modules and replace with @@SERVERNAME, a synonym or the listener/CNAME.' }
    if ($R.SqlOk -and [int](Get-Prop $RI 'Policies' 0) -gt 0 -and [int](Get-Prop $TI 'Policies' 0) -eq 0) { Add-Finding $T 'Server objects' 'Policy-Based Management policies' "$(Get-Prop $RI 'Policies') user policies" 'none' 'Low' -Action 'Export the policies (SSMS > Management > Policy Management > Export) and import them on the target.' -Fix 'Copy-DbaPolicyManagement' }
    if ($R.SqlOk -and [int](Get-Prop $RI 'CmsServers' 0) -gt 0 -and [int](Get-Prop $TI 'CmsServers' 0) -eq 0) { Add-Finding $T 'Server objects' 'Central Management Server registrations' "$(Get-Prop $RI 'CmsServers') registered servers" 'none' 'Low' -Action 'This server is a CMS; move the registrations (Copy-DbaRegServer) and update tools pointing at it.' -Fix 'Copy-DbaRegServer' }

    # ================= Agent: hard-coded names, already-failing jobs
    $c = 'Agent & automation'
    foreach ($st in @(ConvertTo-Arr $R.Sql.JobSteps | Where-Object { @(@(Get-Prop $_ 'SelfRefs' @()) + @(Get-Prop $_ 'IpRefs' @()) | Where-Object { $_ }).Count })) {
        $what = @(@(Get-Prop $st 'SelfRefs' @()) + @(Get-Prop $st 'IpRefs' @()) | Where-Object { $_ }) -join ', '
        Add-Finding $T $c "Hard-coded server name/IP in job: $($st.Job) / $($st.StepId)" $what '' $(if ($isMig) { 'Medium' } else { 'Low' }) -Action $(if ($isMig) { 'The step names the old server (or an IP). After migration it still targets the old box - use @@SERVERNAME/$(ESCAPE_SQUOTE(SRVR)) tokens, the listener/CNAME, or update it at cut-over.' } else { 'Step names a specific replica/IP; after failover it may point to the wrong node. Prefer the listener or Agent tokens.' })
    }
    $failing = @(ConvertTo-Arr $R.Sql.Jobs | Where-Object { [int]$_.Enabled -eq 1 -and $null -ne (Get-Prop $_ 'LastOutcome') -and [int](Get-Prop $_ 'LastOutcome' 1) -eq 0 })
    if ($failing.Count) { Add-Finding $T $c "Jobs already failing on reference: $($failing.Count)" (($failing | Select-Object -First 8 | ForEach-Object { $_.Name }) -join '; ') '' 'Info' -Action 'These jobs failed on their last run on the reference. Fix or retire them first so post-migration failures are not blamed on the move.' -Detail (($failing | ForEach-Object { $_.Name }) -join '; ') }

    # ================= AG replica health (reference = primary)
    if ($isHa -and $R.SqlOk -and $R.Sql.Sections['ag_states'] -like 'ok*') {
        $c = 'SQL instance'
        $mine = @(ConvertTo-Arr $R.Sql.AgStates | Where-Object { ([string]$_.Replica -split '\\')[0] -eq ([string]$T.Name -split '[\\,]')[0] -or [string]$_.Replica -eq [string]$T.Name })
        foreach ($st in $mine) {
            $okState = if ($T.Role -eq 'Synchronous') { 'SYNCHRONIZED' } else { 'SYNCHRONIZING' }
            if ([int]$st.Suspended -eq 1 -or $st.Health -ne 'HEALTHY' -or ($st.SyncState -ne $okState -and $st.SyncState -ne 'SYNCHRONIZED')) { Add-Finding $T $c "AG replica state: $($st.Db)" "expected $okState / HEALTHY" "$($st.SyncState) / $($st.Health)$(if ([int]$st.Suspended) { ' / SUSPENDED' })" 'High' -Hard -Action 'Data movement to this replica is not healthy - a failover now would lose data or fail. Resume/repair before relying on this replica.' }
            elseif ([int64]("0" + $st.SendQueueKB) + [int64]("0" + $st.RedoQueueKB) -gt 1048576) { Add-Finding $T $c "AG replica lag: $($st.Db)" '' ("send {0} / redo {1}" -f (Format-MB ([int64]$st.SendQueueKB / 1024)), (Format-MB ([int64]$st.RedoQueueKB / 1024))) 'Medium' -Action 'Over 1 GB queued for this replica. Check network/IO; RPO/RTO is at risk.' }
            else { Add-Check $T $c }
        }
    }

    # ================= Config sanity on the target itself
    if ($T.SqlOk -and $T.Sql.Sections['config'] -like 'ok*') {
        $c = 'Configuration'
        foreach ($k in @($T.Sql.ConfigPending.Keys)) { Add-Finding $T $c "Pending configuration: $k" "configured $($T.Sql.ConfigPending[$k])" "running $($T.Sql.Config[$k])" 'Low' -Action 'Value changed with sp_configure but not active yet (needs RECONFIGURE or a service restart). Apply it or revert it so the running state is known.' }
        $md = [int]("0" + (Get-Prop $T.Sql.Config 'max degree of parallelism' '0')); $lc = [int](Get-Prop $TI 'CpuCount' 0); if ($T.OsOk -and $T.Os.Info) { $lc = [int]$T.Os.Info.Logical }
        if ($md -gt 0 -and $lc -gt 0 -and $md -gt $lc) { Add-Finding $T $c 'MAXDOP vs target CPUs' "MAXDOP $md" "$lc logical CPUs" 'Medium' -Action 'MAXDOP copied from a bigger server exceeds the target CPUs. Set it per the target NUMA layout (<= cores per NUMA node, max 8 typically).' }
        if ($R.SqlOk -and (Get-Prop $RI 'NumErrorLogs') -and [string](Get-Prop $RI 'NumErrorLogs') -ne [string](Get-Prop $TI 'NumErrorLogs')) { Add-Finding $T $c 'Error log retention (NumErrorLogs)' (Get-Prop $RI 'NumErrorLogs') $(if (Get-Prop $TI 'NumErrorLogs') { Get-Prop $TI 'NumErrorLogs' } else { 'default (6)' }) 'Low' -Action 'Number of retained error logs differs (Configure SQL Server Error Logs).' }
    }

    # ================= Files & paths
    $c = 'Files & paths'
    if ($Skip -notcontains 'Paths' -and $R.SqlOk -and $T.OsOk) {
        $req = Get-RequiredPaths -R $R -T $T
        $unk = 0
        foreach ($q in $req) {
            if ($isHa -and $q.Kind -eq 'Database folder') { $q.Path = $q.RefPath }
            if (-not $T.PathChecks.Contains($q.Path)) { $unk++; continue }
            $okT = $T.PathChecks[$q.Path]
            $okR = if ($q.RefPath -and $R.PathChecks.Contains($q.RefPath)) { $R.PathChecks[$q.RefPath] } else { $true }
            if ($null -eq $okT) { $unk++; continue }
            if ($okT) { Add-Check $T $c; continue }
            if ($okR -eq $false) { Add-Finding $T $c "$($q.Kind): $($q.Path)" 'missing on reference too' 'missing' 'Info' -Action 'Stale reference (already broken on the reference). Clean up the job/step.' -Detail $q.Context; continue }
            $sev = switch ($q.Kind) { 'Database folder' { if ($isHa) { 'High' } elseif ($PathMap.Count) { 'High' } else { 'Medium' } } 'Job step file' { 'High' } 'Job output folder' { 'Medium' } 'Audit folder' { 'High' } 'XE target folder' { 'Low' } 'Extended proc DLL' { 'High' } 'EKM provider DLL' { 'High' } 'Target default folder' { 'High' } default { 'Medium' } }
            $act = switch ($q.Kind) {
                'Database folder' { if ($isHa) { 'AG replicas need identical file paths: a file added on the primary in this folder suspends redo on this replica. Create the folder.' } else { 'Create the folder (with engine service account permissions) or RESTORE ... WITH MOVE to another path (-PathMap documents the mapping).' } }
                'Job step file' { 'Copy the executable/script/package (and its dependencies/config) to the target, or update the job step.' }
                'Job output folder' { 'Create the folder or the step fails writing its log.' }
                'Target default folder' { "The target's own default data/log/backup folder does not exist - fix the instance defaults." }
                'Extended proc DLL' { 'Copy the DLL into the Binn folder (or path) and register it.' }
                'EKM provider DLL' { 'Install the vendor EKM client.' }
                default { 'Create the folder/file on the target.' }
            }
            Add-Finding $T $c "$($q.Kind): $($q.Path)" $(if ($q.RefPath -and $q.RefPath -ne $q.Path) { "exists ($($q.RefPath))" } else { 'exists' }) 'missing' $sev -Install:($q.Kind -in @('Job step file', 'Extended proc DLL', 'EKM provider DLL')) -Action $act -Detail $q.Context
        }
        if ($unk -gt 0) { Add-Finding $T 'Collection gaps' 'Path checks' '' "$unk paths not verifiable" 'Low' -Action 'Some paths could not be checked through CIM (not local paths or WMI query refused).' }
    }
}

function Get-Readiness {
    param($Snapshots)
    $out = New-List
    foreach ($T in ($Snapshots | Select-Object -Skip 1)) {
        $f = @($script:Findings | Where-Object { $_.Target -eq $T.Name })
        $cnt = @{}; foreach ($s in @('Blocker', 'High', 'Medium', 'Low', 'Info')) { $cnt[$s] = @($f | Where-Object { $_.Severity -eq $s }).Count }
        $pen = 0; foreach ($s in $cnt.Keys) { $pen += [Math]::Min($cnt[$s] * $script:SevPenalty[$s], $(if ($s -eq 'Blocker') { 100 } elseif ($s -eq 'High') { 48 } elseif ($s -eq 'Medium') { 24 } else { 10 })) }
        $score = [Math]::Max(0, 100 - $pen)
        if ($cnt['Blocker'] -gt 0) { $score = [Math]::Min($score, 39) }
        $verdict = if ($cnt['Blocker'] -gt 0) { 'NOT READY' } elseif ($cnt['High'] -gt 0) { 'AT RISK' } elseif (($cnt['Medium'] + $cnt['Low']) -gt 0) { 'READY WITH NOTES' } else { 'READY' }
        if (-not $T.SqlOk -and $Skip -notcontains 'Sql') { $verdict = 'NOT VERIFIED'; $score = 0 }
        $checks = 0; foreach ($k in $script:Checks.Keys) { if ($k.StartsWith("$($T.Name)|")) { $checks += $script:Checks[$k] } }
        $cats = [ordered]@{}
        foreach ($cat in $script:Categories) {
            $cf = @($f | Where-Object { $_.Category -eq $cat })
            $worst = 'OK'; foreach ($x in $cf) { if ($script:SevOrder[$x.Severity] -gt $script:SevOrder[$worst]) { $worst = $x.Severity } }
            $n = 0; if ($script:Checks.ContainsKey("$($T.Name)|$cat")) { $n = $script:Checks["$($T.Name)|$cat"] }
            $cats[$cat] = [ordered]@{ Worst = $(if ($n -eq 0 -and $cf.Count -eq 0) { 'n/a' } else { $worst }); Findings = @($cf | Where-Object { $_.Severity -ne 'Info' }).Count; Checks = $n }
        }
        Add-To $out ([pscustomobject]@{ Target = $T.Name; Role = $T.Role; Computer = $T.Computer; Verdict = $verdict; Score = $score; Blocker = $cnt['Blocker']; High = $cnt['High']; Medium = $cnt['Medium']; Low = $cnt['Low']; Info = $cnt['Info']; Checks = $checks; Install = @($f | Where-Object { $_.Install }).Count; Categories = $cats })
    }
    return $out.ToArray()
}
#endregion

#region ---------------------------------------------------------------- demo data (no server needed)
function New-DemoSnapshot {
    param([string]$Name, [string]$Role, [string]$V)   # V = ref | new | sync | dr
    $s = New-Snapshot -Name $Name -Role $Role
    $s.Computer = $Name; $s.SqlOk = $true; $s.OsOk = $true
    $ok = @('connect', 'instance', 'sysinfo', 'services', 'config', 'registry', 'databases', 'files', 'dbscan', 'linked_servers', 'credentials', 'login_names', 'jobs', 'job_steps', 'proxies', 'operators', 'alerts', 'dbmail', 'ssisdb', 'ssisdb_env', 'msdb_packages', 'server_triggers', 'endpoints', 'startup_procs', 'sysdb_objects', 'extended_procs', 'audits', 'audit_specs', 'xe_sessions', 'resource_governor', 'crypto_providers', 'certificates', 'asym_keys', 'tde', 'messages', 'ag', 'trusted_assemblies')
    foreach ($k in $ok) { $s.Sql.Sections[$k] = 'ok' }
    foreach ($k in @('connect', 'os', 'volumes', 'power', 'services', 'features', 'software', 'dotnet', 'odbc', 'oledb', 'aliases', 'tls', 'cluster', 'ps_modules')) { $s.Os.Sections[$k] = 'ok' }
    $isNew = $V -eq 'new'
    $I = $s.Sql.Instance
    if ($isNew) {
        $I.ProductVersion = '16.0.4135.4'; $I.ProductLevel = 'RTM'; $I.ProductUpdateLevel = 'CU14'; $I.Major = 16; $I.Edition = 'Standard Edition (64-bit)'; $I.EngineEdition = 2
        $I.Collation = 'Latin1_General_CI_AS'; $I.FilestreamConfiguredLevel = 0; $I.IsFullTextInstalled = 0; $I.IsHadrEnabled = 0; $I.Ifi = 'N'; $I.MemoryModel = 'CONVENTIONAL'; $I.ResourceVersion = '16.00.4135'
    } else {
        $I.ProductVersion = $(if ($V -eq 'dr') { '13.0.6430.49' } else { '13.0.6435.1' }); $I.ProductLevel = 'SP3'; $I.ProductUpdateLevel = $(if ($V -eq 'dr') { 'GDR' } else { 'GDR' }); $I.Major = 13
        $I.Edition = 'Enterprise Edition: Core-based Licensing (64-bit)'; $I.EngineEdition = 3; $I.Collation = 'SQL_Latin1_General_CP1_CI_AS'
        $I.FilestreamConfiguredLevel = 2; $I.IsFullTextInstalled = 1; $I.IsHadrEnabled = 1; $I.Ifi = $(if ($V -eq 'dr') { 'N' } else { 'Y' }); $I.MemoryModel = 'LOCK_PAGES'
    }
    $I.Build = $I.ProductVersion; $I.VersionName = Get-SqlMajorName $I.Major; $I.ServerName = $Name; $I.PhysicalNode = $Name; $I.MachineName = $Name; $I.InstanceName = $null
    $I.IsClustered = 0; $I.IsPolyBaseInstalled = 0; $I.IsIntegratedSecurityOnly = 0
    $I.DefaultDataPath = 'D:\SQLData\'; $I.DefaultLogPath = 'L:\SQLLogs\'; $I.DefaultBackupPath = $(if ($isNew) { 'G:\Backup' } else { 'G:\Backup' })
    $I.BinnPath = $(if ($isNew) { 'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\Binn\' } else { 'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Binn\' })
    $I.CpuCount = $(if ($V -eq 'dr') { 8 } else { 16 }); $I.MemoryMB = $(if ($V -eq 'dr') { 65536 } elseif ($isNew) { 131072 } else { 131072 }); $I.ServiceAccount = 'CORP\svc_sqlprod'
    $s.Sql.Services = @([pscustomobject]@{ ServiceName = 'SQL Server (MSSQLSERVER)'; Account = 'CORP\svc_sqlprod'; StartMode = 'Automatic'; Status = 'Running'; Ifi = $I.Ifi; Binary = 'sqlservr.exe' })

    $cfg = [ordered]@{ 'clr enabled' = '1'; 'clr strict security' = '0'; 'xp_cmdshell' = '1'; 'Database Mail XPs' = '1'; 'max server memory (MB)' = '114688'; 'min server memory (MB)' = '0'; 'max degree of parallelism' = '8'; 'cost threshold for parallelism' = '50'; 'optimize for ad hoc workloads' = '1'; 'backup compression default' = '1'; 'remote admin connections' = '1'; 'Ad Hoc Distributed Queries' = '1'; 'Ole Automation Procedures' = '0'; 'contained database authentication' = '0'; 'filestream access level' = '2'; 'scan for startup procs' = '1'; 'show advanced options' = '1'; 'default language' = '0'; 'fill factor (%)' = '0'; 'cross db ownership chaining' = '0' }
    if ($isNew) { $cfg['clr enabled'] = '0'; $cfg['clr strict security'] = '1'; $cfg['xp_cmdshell'] = '0'; $cfg['max server memory (MB)'] = '2147483647'; $cfg['max degree of parallelism'] = '0'; $cfg['cost threshold for parallelism'] = '5'; $cfg['optimize for ad hoc workloads'] = '0'; $cfg['Ad Hoc Distributed Queries'] = '0'; $cfg['filestream access level'] = '0'; $cfg['scan for startup procs'] = '0'; $cfg['backup compression default'] = '0'; $cfg['remote admin connections'] = '0' }
    if ($V -eq 'dr') { $cfg['max server memory (MB)'] = '57344'; $cfg['cost threshold for parallelism'] = '5' }
    $s.Sql.Config = $cfg
    $s.Sql.TraceFlags = $(if ($isNew) { @(3226) } elseif ($V -eq 'dr') { @(1118, 3226) } else { @(1117, 1118, 3226, 7745) })
    $s.Sql.Net = [ordered]@{ TcpEnabled = '1'; TcpPort = $(if ($isNew) { '1433' } else { '14330' }); TcpDynamicPorts = ''; ForceEncryption = $(if ($isNew) { '0' } else { '1' }); Certificate = $(if ($isNew) { '' } else { 'A1B2C3D4E5' }) }

    $dbs = @(
        @{ Name = 'master'; Id = 1; Compat = 130 }, @{ Name = 'tempdb'; Id = 2; Compat = 130 }, @{ Name = 'model'; Id = 3; Compat = 130 }, @{ Name = 'msdb'; Id = 4; Compat = 130 },
        @{ Name = 'SalesDB'; Id = 5; Compat = 130; Size = 412000; Tw = 1; Enc = 1; Ag = 1; Owner = 'sa' },
        @{ Name = 'Finance'; Id = 6; Compat = 100; Size = 98000; Ag = 1; Owner = 'CORP\jdoe'; Cdc = 1 },
        @{ Name = 'DocStore'; Id = 7; Compat = 130; Size = 260000; Ag = 1; Owner = 'sa' },
        @{ Name = 'Staging'; Id = 8; Compat = 130; Size = 45000; Owner = 'sa'; Pub = 1 },
        @{ Name = 'ReportServer'; Id = 9; Compat = 130; Size = 4000; Owner = 'CORP\svc_ssrs' },
        @{ Name = 'SSISDB'; Id = 10; Compat = 130; Size = 9000; Owner = 'sa'; Ag = 1 }
    )
    $dbList = New-List
    foreach ($d in $dbs) {
        if ($isNew -and $d['Id'] -gt 4) { continue }
        if ($V -eq 'dr' -and $d['Name'] -eq 'DocStore') { continue }
        if ($V -in @('sync', 'dr') -and -not $d['Ag'] -and $d['Id'] -gt 4) { continue }
        Add-To $dbList ([pscustomobject]@{ Name = $d['Name']; Id = $d['Id']; Compat = $(if ($isNew -and $d['Id'] -le 4) { 160 } else { $d['Compat'] }); Collation = $I.Collation; Recovery = 'FULL'; State = 'ONLINE'; ReadOnly = 0; Trustworthy = [int]("0" + $d['Tw']); DbChaining = 0; Encrypted = [int]("0" + $d['Enc']); Broker = $(if ($d['Name'] -eq 'SalesDB') { 1 } else { 0 }); Cdc = [int]("0" + $d['Cdc']); Published = [int]("0" + $d['Pub']); Subscribed = 0; Distributor = 0; Containment = 'NONE'; Owner = $(if ($d['Owner']) { $d['Owner'] } else { 'sa' }); InAg = [int]("0" + $d['Ag']); SizeMB = [int64]("0" + $d['Size']) })
    }
    if ($isNew) { Add-To $dbList ([pscustomobject]@{ Name = 'SalesDB'; Id = 5; Compat = 130; Collation = 'SQL_Latin1_General_CP1_CI_AS'; Recovery = 'SIMPLE'; State = 'ONLINE'; ReadOnly = 0; Trustworthy = 0; DbChaining = 0; Encrypted = 0; Broker = 0; Cdc = 0; Published = 0; Subscribed = 0; Distributor = 0; Containment = 'NONE'; Owner = 'sa'; InAg = 0; SizeMB = 1024 }) }
    $s.Sql.Databases = $dbList.ToArray()
    $files = New-List
    foreach ($d in $s.Sql.Databases) {
        if ($d.Id -le 4) { continue }
        $dataDir = if ($d.Name -eq 'DocStore') { 'E:\SQLData' } else { 'D:\SQLData' }
        Add-To $files ([pscustomobject]@{ Db = $d.Name; DbId = $d.Id; Type = 'ROWS'; LogicalName = $d.Name; Path = "$dataDir\$($d.Name).mdf"; SizeMB = [int64]($d.SizeMB * 0.8) })
        Add-To $files ([pscustomobject]@{ Db = $d.Name; DbId = $d.Id; Type = 'LOG'; LogicalName = "$($d.Name)_log"; Path = "L:\SQLLogs\$($d.Name)_log.ldf"; SizeMB = [int64]($d.SizeMB * 0.2) })
        if ($d.Name -eq 'DocStore') { Add-To $files ([pscustomobject]@{ Db = $d.Name; DbId = $d.Id; Type = 'FILESTREAM'; LogicalName = 'DocStore_fs'; Path = 'E:\SQLData\DocStore_fs'; SizeMB = 0 }) }
    }
    $s.Sql.Files = $files.ToArray()
    if (-not $isNew) {
        $s.Sql.DbFeatures = @([pscustomobject]@{ Db = 'SalesDB'; FtCatalogs = 1; FsFilegroups = 0; XtpFilegroups = 0; BrokerQueues = 2; Assemblies = 2 }, [pscustomobject]@{ Db = 'DocStore'; FtCatalogs = 1; FsFilegroups = 1; XtpFilegroups = 0; BrokerQueues = 0; Assemblies = 0 }, [pscustomobject]@{ Db = 'Finance'; FtCatalogs = 0; FsFilegroups = 0; XtpFilegroups = 0; BrokerQueues = 0; Assemblies = 0 })
        $s.Sql.Assemblies = @([pscustomobject]@{ Db = 'SalesDB'; Name = 'RegexUtil'; ClrName = 'regexutil, version=1.0.0.0'; PermissionSet = 'SAFE_ACCESS'; Created = '2019-03-02'; Trustworthy = 1 }, [pscustomobject]@{ Db = 'SalesDB'; Name = 'FileExport'; ClrName = 'fileexport, version=2.1.0.0'; PermissionSet = 'EXTERNAL_ACCESS'; Created = '2020-07-14'; Trustworthy = 1 })
        $s.Sql.SkuFeatures = @([pscustomobject]@{ Db = 'SalesDB'; Feature = 'TransparentDataEncryption' }, [pscustomobject]@{ Db = 'SalesDB'; Feature = 'Partitioning' }, [pscustomobject]@{ Db = 'Finance'; Feature = 'ChangeCapture' }, [pscustomobject]@{ Db = 'DocStore'; Feature = 'Compression' })
        $s.Sql.LinkedRefs = @([pscustomobject]@{ Db = 'Finance'; LinkedServer = 'ORAFIN'; RefDb = $null }, [pscustomobject]@{ Db = 'SalesDB'; LinkedServer = 'LEGACYCRM'; RefDb = 'CRM' })
        $s.Sql.CrossDbRefs = @([pscustomobject]@{ Db = 'SalesDB'; RefDb = 'Finance' }, [pscustomobject]@{ Db = 'Staging'; RefDb = 'SalesDB' })
        $s.Sql.CodeUse = @([pscustomobject]@{ Db = 'SalesDB'; XpCmdshell = 2; OleAutomation = 0; AdHocQueries = 0; DbMail = 3; BulkInsert = 1 }, [pscustomobject]@{ Db = 'Staging'; XpCmdshell = 0; OleAutomation = 0; AdHocQueries = 4; DbMail = 0; BulkInsert = 6 })
        $s.Sql.LinkedServers = @([pscustomobject]@{ Name = 'ORAFIN'; Product = 'Oracle'; Provider = 'OraOLEDB.Oracle'; DataSource = 'FINPRD'; ProviderString = $null; Catalog = $null; RpcOut = 1; DataAccess = 1; StoredLogins = 1 }, [pscustomobject]@{ Name = 'LEGACYCRM'; Product = ''; Provider = 'SQLNCLI11'; DataSource = 'CRMSQL01'; ProviderString = $null; Catalog = 'CRM'; RpcOut = 1; DataAccess = 1; StoredLogins = 1 }, [pscustomobject]@{ Name = 'EXCELFEED'; Product = 'Excel'; Provider = 'Microsoft.ACE.OLEDB.12.0'; DataSource = 'G:\Feeds\rates.xlsx'; ProviderString = 'Excel 12.0;HDR=YES'; Catalog = $null; RpcOut = 0; DataAccess = 1; StoredLogins = 0 })
        if ($V -eq 'dr') { $s.Sql.LinkedServers = @($s.Sql.LinkedServers | Where-Object { $_.Name -ne 'EXCELFEED' }) }
        $s.Sql.Credentials = @([pscustomobject]@{ Name = 'AzureBackupCred'; CredIdentity = 'https://corpbackups.blob.core.windows.net/sql' }, [pscustomobject]@{ Name = 'FileExportCred'; CredIdentity = 'CORP\svc_fileexport' })
        $s.Sql.Proxies = @([pscustomobject]@{ Name = 'FileExportProxy'; Credential = 'FileExportCred'; Enabled = 1 })
        $s.Sql.Operators = @([pscustomobject]@{ Name = 'DBA Team'; Enabled = 1; Email = 'dba@corp.example' })
        $s.Sql.Alerts = @(@(17..25) | ForEach-Object { [pscustomobject]@{ Name = "Severity $_"; MessageId = 0; Severity = $_; Enabled = 1 } }) + @(823, 824, 825 | ForEach-Object { [pscustomobject]@{ Name = "Error $_"; MessageId = $_; Severity = 0; Enabled = 1 } })
        $s.Sql.Mail = @([pscustomobject]@{ Profile = 'DBA Mail'; Account = 'sqlalerts'; Email = 'sqlalerts@corp.example'; SmtpServer = 'smtp.corp.example'; Port = 25; Ssl = 0 })
        $s.Sql.SsisProjects = @([pscustomobject]@{ Folder = 'Finance'; Project = 'GL_Load'; Deployed = '2025-11-02' }, [pscustomobject]@{ Folder = 'Sales'; Project = 'NightlyExtract'; Deployed = '2026-02-19' })
        $s.Sql.SsisEnvironments = @([pscustomobject]@{ Folder = 'Finance'; Environment = 'PROD' })
        $s.Sql.SsisMsdb = @([pscustomobject]@{ Folder = 'Legacy'; Package = 'ImportRates' })
        $s.Sql.ServerTriggers = @([pscustomobject]@{ Name = 'trg_LogonAudit'; Disabled = 0 })
        $s.Sql.StartupProcs = @('usp_StartupWarmCache')
        $s.Sql.SysDbObjects = @([pscustomobject]@{ Db = 'master'; Name = 'dbo.sp_WhoIsActive'; Type = 'SQL_STORED_PROCEDURE' }, [pscustomobject]@{ Db = 'master'; Name = 'dbo.CommandLog'; Type = 'USER_TABLE' }, [pscustomobject]@{ Db = 'master'; Name = 'dbo.DatabaseBackup'; Type = 'SQL_STORED_PROCEDURE' })
        $s.Sql.XProcs = @([pscustomobject]@{ Name = 'xp_corp_hash'; Dll = 'xpcorphash.dll' })
        $s.Sql.Audits = @([pscustomobject]@{ Name = 'SOX_Audit'; Type = 'FILE'; OnFailure = 'CONTINUE'; Enabled = 1; Path = 'G:\Audit\' })
        $s.Sql.AuditSpecs = @([pscustomobject]@{ Name = 'SOX_ServerSpec'; Enabled = 1 })
        $s.Sql.XeSessions = @([pscustomobject]@{ Name = 'LongQueries'; AutoStart = 1; FilePath = 'G:\XE\LongQueries.xel' })
        $s.Sql.Certificates = @([pscustomobject]@{ Name = 'TDE_Sales'; Subject = 'TDE for SalesDB'; Thumbprint = '5F3C9A0B11D2E4F6A7B8C9D0E1F2A3B4C5D6E7F8'; Expiry = '2030-12-31'; PrivateKey = 'ENCRYPTED_BY_MASTER_KEY' }, [pscustomobject]@{ Name = 'BackupEncCert'; Subject = 'Backup encryption'; Thumbprint = '0AB1C2D3E4F5061728394A5B6C7D8E9F00112233'; Expiry = '2028-06-30'; PrivateKey = 'ENCRYPTED_BY_MASTER_KEY' })
        if ($V -eq 'dr') { $s.Sql.Certificates = @($s.Sql.Certificates | Where-Object { $_.Name -ne 'BackupEncCert' }) }
        $s.Sql.Tde = @([pscustomobject]@{ Db = 'SalesDB'; Thumbprint = '5F3C9A0B11D2E4F6A7B8C9D0E1F2A3B4C5D6E7F8'; State = 3 })
        $s.Sql.Messages = @([pscustomobject]@{ Id = 50001; Lang = 1033; Severity = 16; Logged = 0; Text = 'Order %d failed validation: %s' }, [pscustomobject]@{ Id = 50002; Lang = 1033; Severity = 10; Logged = 1; Text = 'Nightly load complete' })
        $s.Sql.Endpoints = @([pscustomobject]@{ Name = 'Hadr_endpoint'; Type = 'DATABASE_MIRRORING'; State = 'STARTED'; Port = 5022; Role = 'ALL'; Encryption = $(if ($V -eq 'dr') { 'RC4' } else { 'AES' }); Owner = 'sa' })
        $s.Sql.ResourceGovernor = [ordered]@{ Enabled = 1; Classifier = 'dbo.fn_rg_classifier'; Pools = @([pscustomobject]@{ Name = 'ReportingPool'; MinCpu = 0; MaxCpu = 40; MinMem = 0; MaxMem = 30 }); Groups = @([pscustomobject]@{ Name = 'Reporting'; Pool = 'ReportingPool' }) }
        $s.Sql.Ag = @([pscustomobject]@{ AgName = 'AG_PROD'; Replica = 'SQLPROD01'; Mode = 'SYNCHRONOUS_COMMIT'; Failover = 'AUTOMATIC'; Endpoint = 'TCP://SQLPROD01:5022'; Listener = 'SQLPRODLSN'; ListenerPort = 14330 })
    } else {
        $s.Sql.ResourceGovernor = [ordered]@{ Enabled = 0; Classifier = $null; Pools = @(); Groups = @() }
        $s.Sql.Operators = @([pscustomobject]@{ Name = 'DBA Team'; Enabled = 1; Email = 'dba@corp.example' })
        $s.Sql.Mail = @([pscustomobject]@{ Profile = 'DBA Mail'; Account = 'sqlalerts'; Email = 'sqlalerts@corp.example'; SmtpServer = 'smtp-relay.corp.example'; Port = 587; Ssl = 1 })
        $s.Sql.SysDbObjects = @([pscustomobject]@{ Db = 'master'; Name = 'dbo.sp_WhoIsActive'; Type = 'SQL_STORED_PROCEDURE' })
    }
    $jobs = @(
        @{ N = 'DBA - Full backups'; O = 'sa' }, @{ N = 'DBA - Integrity check'; O = 'sa' }, @{ N = 'DBA - Index maintenance'; O = 'sa' },
        @{ N = 'Sales - Nightly export'; O = 'CORP\svc_etl' }, @{ N = 'Finance - GL load (SSIS)'; O = 'CORP\svc_etl' }, @{ N = 'Ops - Archive feeds'; O = 'CORP\jdoe' }, @{ N = 'cdc.Finance_capture'; O = 'sa' }, @{ N = 'cdc.Finance_cleanup'; O = 'sa' }
    )
    $keep = switch ($V) { 'new' { @('DBA - Full backups', 'DBA - Integrity check') } 'dr' { @('DBA - Full backups', 'DBA - Integrity check', 'DBA - Index maintenance', 'Sales - Nightly export') } default { @($jobs | ForEach-Object { $_.N }) } }
    $s.Sql.Jobs = @($jobs | Where-Object { $keep -contains $_.N } | ForEach-Object { [pscustomobject]@{ Name = $_.N; Enabled = $(if ($V -in @('sync', 'dr') -and $_.N -notlike 'DBA*') { 0 } else { 1 }); Owner = $_.O; Category = '[Uncategorized (Local)]'; Steps = 2 } })
    if ($V -ne 'new') {
        $s.Sql.JobSteps = @(
            [pscustomobject]@{ Job = 'Sales - Nightly export'; StepId = 1; Step = 'Export'; Subsystem = 'CmdExec'; Proxy = 'FileExportProxy'; Paths = @('D:\Apps\Exporter\SalesExport.exe', 'D:\Apps\Exporter\SalesExport.config'); Unc = @('\\fileserver01\exports\sales'); Modules = @(); OutputDir = 'G:\JobLogs'; SsisProject = $null },
            [pscustomobject]@{ Job = 'Finance - GL load (SSIS)'; StepId = 1; Step = 'Run GL_Load'; Subsystem = 'SSIS'; Proxy = $null; Paths = @(); Unc = @(); Modules = @(); OutputDir = $null; SsisProject = 'Finance\GL_Load' },
            [pscustomobject]@{ Job = 'Ops - Archive feeds'; StepId = 1; Step = 'Zip + move'; Subsystem = 'PowerShell'; Proxy = $null; Paths = @('D:\Scripts\Archive-Feeds.ps1', 'C:\Program Files\7-Zip\7z.exe'); Unc = @(); Modules = @('dbatools', 'Posh-SSH'); OutputDir = 'G:\JobLogs'; SsisProject = $null }
        )
    }
    $s.Sql.Logins = @('sa', 'CORP\svc_etl', 'CORP\svc_sqlprod', 'CORP\DBA Team')
    if ($V -ne 'new') { $s.Sql.Logins += @('CORP\jdoe', 'CORP\svc_ssrs') }

    # ---- Windows side
    $s.Os.Info = [ordered]@{
        Caption = $(if ($isNew) { 'Microsoft Windows Server 2022 Standard' } else { 'Microsoft Windows Server 2016 Standard' }); Version = $(if ($isNew) { '10.0.20348' } else { '10.0.14393' }); Build = $(if ($isNew) { '20348' } else { '14393' }); Arch = '64-bit'
        Locale = '0409'; Language = 1033; Domain = 'corp.example'; PartOfDomain = $true; Manufacturer = 'VMware, Inc.'; Model = 'VMware7,1'; IsVirtual = $true
        Cores = $(if ($V -eq 'dr') { 8 } else { 16 }); Logical = $(if ($V -eq 'dr') { 8 } else { 16 }); Sockets = 2; CpuName = 'Intel(R) Xeon(R) Gold 6338 CPU @ 2.00GHz'
        MemoryMB = $(if ($V -eq 'dr') { 65536 } else { 131072 }); TimeZone = $(if ($V -eq 'dr') { '(UTC-06:00) Central Time (US & Canada)' } else { '(UTC-05:00) Eastern Time (US & Canada)' }); TimeZoneStd = ''; TzBias = -300; LastBoot = '2026-10-04 02:10'
    }
    $s.Os.PowerPlan = $(if ($isNew) { 'Balanced' } else { 'High performance' })
    $vol = { param($n, $l, $size, $free, $bs) [pscustomobject]@{ Name = $n; Letter = $l; Label = $l; SizeMB = $size; FreeMB = $free; BlockSize = $bs; FileSystem = 'NTFS' } }
    $s.Os.Volumes = switch ($V) {
        'new' { @((& $vol 'C:\' 'C:' 153600 90000 4096), (& $vol 'D:\' 'D:' 716800 700000 65536), (& $vol 'L:\' 'L:' 204800 200000 65536), (& $vol 'G:\' 'G:' 1048576 1000000 4096), (& $vol 'T:\' 'T:' 102400 100000 65536)) }
        'dr' { @((& $vol 'C:\' 'C:' 153600 60000 4096), (& $vol 'D:\' 'D:' 921600 300000 65536), (& $vol 'E:\' 'E:' 512000 240000 65536), (& $vol 'L:\' 'L:' 307200 150000 65536), (& $vol 'G:\' 'G:' 1048576 400000 65536)) }
        default { @((& $vol 'C:\' 'C:' 153600 60000 4096), (& $vol 'D:\' 'D:' 921600 300000 65536), (& $vol 'E:\' 'E:' 512000 220000 65536), (& $vol 'L:\' 'L:' 307200 150000 65536), (& $vol 'G:\' 'G:' 1048576 400000 65536)) }
    }
    $svc = { param($n, $d, $m, $a, $p) [pscustomobject]@{ Name = $n; DisplayName = $d; StartMode = $m; State = $(if ($m -eq 'Disabled') { 'Stopped' } else { 'Running' }); Account = $a; Path = $p } }
    $common = @((& $svc 'MSSQLSERVER' 'SQL Server (MSSQLSERVER)' 'Auto' 'CORP\svc_sqlprod' '"C:\Program Files\Microsoft SQL Server\MSSQL\Binn\sqlservr.exe" -sMSSQLSERVER'), (& $svc 'SQLSERVERAGENT' 'SQL Server Agent (MSSQLSERVER)' $(if ($isNew) { 'Manual' } else { 'Auto' }) 'CORP\svc_sqlagent' 'SQLAGENT.EXE'), (& $svc 'SQLWriter' 'SQL Server VSS Writer' 'Auto' 'LocalSystem' 'sqlwriter.exe'), (& $svc 'SQLTELEMETRY' 'SQL Server CEIP service' 'Auto' 'NT Service\SQLTELEMETRY' 'sqlceip.exe'), (& $svc 'MSDTC' 'Distributed Transaction Coordinator' 'Manual' 'NetworkService' 'C:\Windows\System32\msdtc.exe'))
    $extra = @()
    if (-not $isNew) { $extra += @((& $svc 'MsDtsServer130' 'SQL Server Integration Services 13.0' 'Auto' 'NT Service\MsDtsServer130' 'MsDtsSrvr.exe'), (& $svc 'MSSQLFDLauncher' 'SQL Full-text Filter Daemon Launcher (MSSQLSERVER)' 'Manual' 'NT Service\MSSQLFDLauncher' 'fdlauncher.exe'), (& $svc 'ClusSvc' 'Cluster Service' 'Auto' 'LocalSystem' 'C:\Windows\Cluster\clussvc.exe')) }
    if ($V -eq 'ref' -or $V -eq 'sync') { $extra += @((& $svc 'CVD_Instance001' 'Commvault Communications Service (CVD) (Instance001)' 'Auto' 'LocalSystem' 'C:\Program Files\Commvault\ContentStore\Base\cvd.exe'), (& $svc 'SQLMonitorAgent' 'Redgate SQL Monitor Base Monitor' 'Auto' 'CORP\svc_mon' 'D:\Redgate\SQL Monitor\Base Monitor\RedGate.Monitor.exe')) }
    if ($V -eq 'dr') { $extra += @((& $svc 'CVD_Instance001' 'Commvault Communications Service (CVD) (Instance001)' 'Auto' 'LocalSystem' 'C:\Program Files\Commvault\ContentStore\Base\cvd.exe')) }
    if ($isNew) { $extra += @((& $svc 'CSFalconService' 'CrowdStrike Falcon Sensor Service' 'Auto' 'LocalSystem' 'C:\Program Files\CrowdStrike\CSFalconService.exe')) }
    $s.Os.Services = @($common + $extra)
    $s.Os.Features = $(if ($isNew) { @('.NET Framework 4.8 Features', 'Windows PowerShell 5.1') } else { @('.NET Framework 4.8 Features', '.NET Framework 3.5 (includes .NET 2.0 and 3.0)', '.NET Framework 3.5 Features', 'Failover Clustering', 'Failover Cluster Module for Windows PowerShell', 'Windows PowerShell 5.1') })
    $sw = { param($n, $v, $p, $view) [pscustomobject]@{ Name = $n; Version = $v; Publisher = $p; View = $view } }
    $refSw = @(
        (& $sw 'Microsoft Visual C++ 2015-2022 Redistributable (x64) - 14.38.33130' '14.38.33130.0' 'Microsoft Corporation' 'x86'),
        (& $sw 'Microsoft Visual C++ 2015-2022 Redistributable (x86) - 14.38.33130' '14.38.33130.0' 'Microsoft Corporation' 'x86'),
        (& $sw 'Microsoft Visual C++ 2013 Redistributable (x64) - 12.0.40664' '12.0.40664.0' 'Microsoft Corporation' 'x86'),
        (& $sw 'Microsoft ODBC Driver 17 for SQL Server' '17.10.6.1' 'Microsoft Corporation' 'x64'),
        (& $sw 'Microsoft OLE DB Driver for SQL Server' '18.7.2.0' 'Microsoft Corporation' 'x64'),
        (& $sw 'Microsoft SQL Server 2012 Native Client' '11.4.7462.6' 'Microsoft Corporation' 'x64'),
        (& $sw 'Oracle Data Access Components 19.3 (64-bit)' '19.3.0' 'Oracle Corporation' 'x64'),
        (& $sw 'Microsoft Access database engine 2010 (English)' '14.0.7015.1000' 'Microsoft Corporation' 'x64'),
        (& $sw 'Java 8 Update 381 (64-bit)' '8.0.3810.9' 'Oracle Corporation' 'x64'),
        (& $sw 'Microsoft SQL Server 2016 Integration Services' '13.3.6300.1' 'Microsoft Corporation' 'x64'),
        (& $sw 'Microsoft System CLR Types for SQL Server 2016' '13.1.4001.0' 'Microsoft Corporation' 'x64'),
        (& $sw 'SQL Server Management Studio' '19.3.4.0' 'Microsoft Corporation' 'x86'),
        (& $sw 'Commvault ContentStore' '11.32.40' 'Commvault' 'x64'),
        (& $sw 'Redgate SQL Monitor Base Monitor' '13.0.72' 'Red Gate Software Ltd' 'x64'),
        (& $sw '7-Zip 23.01 (x64)' '23.01' 'Igor Pavlov' 'x64'),
        (& $sw 'WinSCP 6.1.2' '6.1.2' 'Martin Prikryl' 'x86'),
        (& $sw 'VMware Tools' '12.3.0.22234872' 'VMware, Inc.' 'x64'),
        (& $sw 'Security Update for Microsoft Visual C++ (KB2565063)' '1.0' 'Microsoft Corporation' 'x86')
    )
    $s.Os.Software = switch ($V) {
        'new' { @((& $sw 'Microsoft Visual C++ 2015-2022 Redistributable (x64) - 14.32.31332' '14.32.31332.0' 'Microsoft Corporation' 'x86'), (& $sw 'Microsoft ODBC Driver 18 for SQL Server' '18.3.2.1' 'Microsoft Corporation' 'x64'), (& $sw 'Microsoft ODBC Driver 17 for SQL Server' '17.10.6.1' 'Microsoft Corporation' 'x64'), (& $sw 'Microsoft OLE DB Driver for SQL Server' '18.7.2.0' 'Microsoft Corporation' 'x64'), (& $sw 'SQL Server Management Studio' '20.2.30.0' 'Microsoft Corporation' 'x86'), (& $sw 'CrowdStrike Windows Sensor' '7.10.17706' 'CrowdStrike, Inc.' 'x64'), (& $sw 'VMware Tools' '12.4.0' 'VMware, Inc.' 'x64'), (& $sw 'Microsoft SQL Server 2022 (64-bit)' '16.0.4135.4' 'Microsoft Corporation' 'x64')) }
        'dr' { @($refSw | Where-Object { $_.Name -notmatch 'Redgate|Access database|WinSCP' }) }
        default { $refSw }
    }
    $s.Os.DotNet = [ordered]@{ Release = $(if ($isNew) { 533320 } else { 528449 }); Version = $(if ($isNew) { '4.8.1' } else { '4.8' }); Net35 = (-not $isNew); StrongCrypto = 1 }
    $s.Os.OdbcDrivers = $(if ($isNew) { @([pscustomobject]@{ Name = 'ODBC Driver 17 for SQL Server'; Bitness = 'x64' }, [pscustomobject]@{ Name = 'ODBC Driver 18 for SQL Server'; Bitness = 'x64' }, [pscustomobject]@{ Name = 'SQL Server'; Bitness = 'x64' }) } else { @([pscustomobject]@{ Name = 'ODBC Driver 17 for SQL Server'; Bitness = 'x64' }, [pscustomobject]@{ Name = 'SQL Server Native Client 11.0'; Bitness = 'x64' }, [pscustomobject]@{ Name = 'Oracle in OraClient19Home1'; Bitness = 'x64' }, [pscustomobject]@{ Name = 'SQL Server'; Bitness = 'x64' }) })
    $s.Os.Dsns = $(if ($isNew) { @() } else { @([pscustomobject]@{ Name = 'FINPRD_ODBC'; Driver = 'Oracle in OraClient19Home1'; Server = 'FINPRD'; Bitness = 'x64' }) })
    $ole = [ordered]@{}
    foreach ($p in $script:KnownOleDb) { $ole[$p] = [ordered]@{ x64 = $false; x86 = $false } }
    foreach ($p in @('SQLOLEDB', 'MSDASQL', 'MSOLEDBSQL', 'ADsDSOObject', 'Search.CollatorDSO')) { $ole[$p] = [ordered]@{ x64 = $true; x86 = $true } }
    $ole['Microsoft.Jet.OLEDB.4.0'] = [ordered]@{ x64 = $false; x86 = $true }
    if (-not $isNew) { $ole['SQLNCLI11'] = [ordered]@{ x64 = $true; x86 = $true }; $ole['OraOLEDB.Oracle'] = [ordered]@{ x64 = $true; x86 = $false }; $ole['MSOLAP'] = [ordered]@{ x64 = $true; x86 = $false } }
    if ($V -ne 'new' -and $V -ne 'dr') { $ole['Microsoft.ACE.OLEDB.12.0'] = [ordered]@{ x64 = $true; x86 = $false } }
    $s.Os.OleDb = $ole
    $s.Os.Aliases = $(if ($isNew) { @() } else { @([pscustomobject]@{ Name = 'CRMSQL01'; Target = 'DBMSSOCN,crmsql01.corp.example,14331'; Bitness = 'x64' }) })
    $tls = [ordered]@{}; foreach ($p in @('SSL 3.0', 'TLS 1.0', 'TLS 1.1', 'TLS 1.2', 'TLS 1.3')) { foreach ($sd in @('Server', 'Client')) { $tls["$p $sd"] = 'OS default' } }
    $tls['SSL 3.0 Server'] = 'Disabled'; $tls['SSL 3.0 Client'] = 'Disabled'
    if ($isNew) { $tls['TLS 1.0 Server'] = 'Disabled'; $tls['TLS 1.0 Client'] = 'Disabled'; $tls['TLS 1.1 Server'] = 'Disabled'; $tls['TLS 1.1 Client'] = 'Disabled'; $tls['TLS 1.2 Server'] = 'Enabled'; $tls['TLS 1.2 Client'] = 'Enabled' }
    $s.Os.Tls = $tls
    $s.Os.ClusterName = $(if ($isNew) { $null } elseif ($V -eq 'dr') { 'CLPROD01' } else { 'CLPROD01' })
    $s.Os.PsModules = $(if ($isNew) { @('PackageManagement', 'PowerShellGet', 'SqlServer') } elseif ($V -eq 'dr') { @('PackageManagement', 'PowerShellGet', 'dbatools') } else { @('PackageManagement', 'PowerShellGet', 'SqlServer', 'dbatools', 'Posh-SSH') })

    # ---- extended facets (schedules, mail detail, layout, certificates, firewall ...)
    foreach ($k in @('schedules', 'proxy_subsystems', 'proxy_logins', 'dbmail_accounts', 'dbmail_principals', 'dbmail_config', 'backup_devices', 'maint_plans', 'job_categories')) { $s.Sql.Sections[$k] = 'ok' }
    foreach ($k in @('agent_reg', 'provider_options', 'certificates', 'firewall')) { $s.Os.Sections[$k] = 'ok' }
    $sysFiles = New-List
    $sysDir = if ($isNew) { 'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA' } else { 'D:\SQLSystem' }
    foreach ($x in @(@(1, 'master', 'master', 'ROWS', 'mdf'), @(1, 'master', 'mastlog', 'LOG', 'ldf'), @(3, 'model', 'modeldev', 'ROWS', 'mdf'), @(3, 'model', 'modellog', 'LOG', 'ldf'), @(4, 'msdb', 'MSDBData', 'ROWS', 'mdf'), @(4, 'msdb', 'MSDBLog', 'LOG', 'ldf'))) {
        Add-To $sysFiles ([pscustomobject]@{ Db = $x[1]; DbId = $x[0]; Type = $x[3]; LogicalName = $x[2]; Path = "$sysDir\$($x[2]).$($x[4])"; SizeMB = 64; Growth = '64 MB'; MaxSize = -1 })
    }
    $tdbCount = if ($isNew) { 1 } else { 8 }
    for ($n = 1; $n -le $tdbCount; $n++) { Add-To $sysFiles ([pscustomobject]@{ Db = 'tempdb'; DbId = 2; Type = 'ROWS'; LogicalName = $(if ($n -eq 1) { 'tempdev' } else { "temp$n" }); Path = $(if ($isNew) { "$sysDir\tempdb.mdf" } else { "T:\TempDB\temp$n.ndf" }); SizeMB = $(if ($isNew) { 8 } else { 8192 }); Growth = $(if ($isNew) { '64 MB' } else { '1024 MB' }); MaxSize = -1 }) }
    Add-To $sysFiles ([pscustomobject]@{ Db = 'tempdb'; DbId = 2; Type = 'LOG'; LogicalName = 'templog'; Path = $(if ($isNew) { "$sysDir\templog.ldf" } else { 'T:\TempDB\templog.ldf' }); SizeMB = $(if ($isNew) { 8 } else { 16384 }); Growth = '1024 MB'; MaxSize = -1 })
    $s.Sql.Files = @($sysFiles.ToArray()) + @($s.Sql.Files)
    if (-not $isNew) { $s.Os.Volumes = @($s.Os.Volumes) + @([pscustomobject]@{ Name = 'T:\'; Letter = 'T:'; Label = 'TempDB'; SizeMB = 204800; FreeMB = 60000; BlockSize = 65536; FileSystem = 'NTFS' }) }
    if ($V -eq 'dr') { $s.Os.Volumes = @($s.Os.Volumes | Where-Object { $_.Name -ne 'T:\' }) + @([pscustomobject]@{ Name = 'T:\'; Letter = 'T:'; Label = 'TempDB'; SizeMB = 102400; FreeMB = 30000; BlockSize = 4096; FileSystem = 'NTFS' }) }
    $sch = { param($j, $n, $e, $ft, $fi, $st, $sub, $si) [pscustomobject]@{ Job = $j; Schedule = $n; Enabled = $e; FreqType = $ft; FreqInterval = $fi; SubdayType = $sub; SubdayInterval = $si; RelInterval = 0; Recurrence = 1; StartTime = $st; EndTime = 235959; EndDate = 99991231 } }
    $schList = @(
        (& $sch 'DBA - Full backups' 'Nightly 01:00' 1 4 1 10000 1 0), (& $sch 'DBA - Integrity check' 'Sunday 03:00' 1 8 1 30000 1 0), (& $sch 'DBA - Index maintenance' 'Sat 22:00' 1 8 64 220000 1 0),
        (& $sch 'Sales - Nightly export' 'Nightly 02:30' 1 4 1 23000 1 0), (& $sch 'Finance - GL load (SSIS)' 'Every 15 min' 1 4 1 60000 4 15), (& $sch 'Ops - Archive feeds' 'Weekdays 19:00' 1 8 62 190000 1 0))
    if ($isNew) { $schList = @((& $sch 'DBA - Full backups' 'Nightly 01:00' 1 4 1 10000 1 0), (& $sch 'DBA - Integrity check' 'Sunday 03:00' 0 8 1 30000 1 0)) }
    if ($V -eq 'dr') { $schList = @($schList | Where-Object { $_.Job -ne 'Sales - Nightly export' }) + @((& $sch 'DBA - Index maintenance' 'Sat 22:00' 1 8 64 230000 1 0) | Where-Object { $false }) ; $schList = @($schList | ForEach-Object { if ($_.Job -eq 'DBA - Index maintenance') { $_.StartTime = 230000 }; $_ }) }
    foreach ($x in $schList) { $x | Add-Member -NotePropertyName Text -NotePropertyValue (Get-ScheduleText $x) }
    $s.Sql.Schedules = @($schList)
    $s.Sql.JobSignatures = @{}
    foreach ($j in $s.Sql.Jobs) { $s.Sql.JobSignatures[$j.Name] = $(if ($V -eq 'dr' -and $j.Name -eq 'Sales - Nightly export') { '9C41D07AA2E3' } else { ('{0:X12}' -f ([Math]::Abs($j.Name.GetHashCode()))).Substring(0, 12) }) }
    $s.Sql.JobCategories = $(if ($isNew) { @() } else { @([pscustomobject]@{ Name = 'ETL'; Class = 1 }, [pscustomobject]@{ Name = 'DBA Maintenance'; Class = 1 }) })
    if (-not $isNew) {
        $s.Sql.ProxySubsystems = @([pscustomobject]@{ Proxy = 'FileExportProxy'; Subsystem = 'CmdExec' }, [pscustomobject]@{ Proxy = 'FileExportProxy'; Subsystem = 'PowerShell' })
        if ($V -eq 'dr') { $s.Sql.ProxySubsystems = @($s.Sql.ProxySubsystems | Where-Object { $_.Subsystem -ne 'PowerShell' }) }
        $s.Sql.ProxyLogins = @([pscustomobject]@{ Proxy = 'FileExportProxy'; Principal = 'CORP\svc_etl' })
        $s.Sql.BackupDevices = @([pscustomobject]@{ Name = 'NightlyDump'; Type = 'DISK'; Path = 'G:\Backup\nightly.bak' })
        $s.Sql.MaintPlans = @([pscustomobject]@{ Name = 'Legacy Weekly Stats'; Owner = 'CORP\jdoe' })
    }
    $s.Sql.MailAccounts = @([pscustomobject]@{ Account = 'sqlalerts'; Email = 'sqlalerts@corp.example'; DisplayName = 'SQL Alerts'; ReplyTo = $null; SmtpServer = $(if ($isNew) { 'smtp-relay.corp.example' } else { 'smtp.corp.example' }); Port = $(if ($isNew) { 587 } else { 25 }); Ssl = $(if ($isNew) { 1 } else { 0 }); DefaultCreds = 0; SmtpUser = $null })
    $s.Sql.MailPrincipals = $(if ($isNew) { @() } else { @([pscustomobject]@{ Profile = 'DBA Mail'; Principal = 'public'; IsDefault = 1 }) })
    $s.Sql.MailConfig = @([pscustomobject]@{ Name = 'MaxFileSize'; Value = $(if ($isNew) { '1000000' } else { '10000000' }) }, [pscustomobject]@{ Name = 'AccountRetryAttempts'; Value = '1' })
    $s.Sql.AgentProps = [ordered]@{ JobHistoryMaxRows = $(if ($isNew) { '1000' } else { '50000' }); JobHistoryMaxRowsPerJob = $(if ($isNew) { '100' } else { '1000' }) }
    $s.Os.AgentReg = [ordered]@{ DatabaseMailProfile = $(if ($isNew) { $null } else { 'DBA Mail' }); UseDatabaseMail = $(if ($isNew) { 0 } else { 1 }); AlertFailSafeOperator = $(if ($isNew) { $null } else { 'DBA Team' }) }
    $s.Os.ProviderOptions = [ordered]@{ 'SQLNCLI11' = [ordered]@{}; 'MSOLEDBSQL' = [ordered]@{}; 'OraOLEDB.Oracle' = [ordered]@{ AllowInProcess = 1 }; 'Microsoft.ACE.OLEDB.12.0' = [ordered]@{ AllowInProcess = 1; DynamicParameters = 1 } }
    if ($V -eq 'dr') { $s.Os.ProviderOptions['OraOLEDB.Oracle'] = [ordered]@{} }
    if ($isNew) { $s.Os.ProviderOptions = [ordered]@{ 'MSOLEDBSQL' = [ordered]@{}; 'MSOLEDBSQL19' = [ordered]@{} } }
    $mk = { param($th, $subj, $iss, $days, $san, $key, $eku) [pscustomobject]@{ Thumbprint = $th; Subject = $subj; Issuer = $iss; NotAfter = (Get-Date).AddDays($days).ToString('yyyy-MM-dd'); NotAfterDate = (Get-Date).AddDays($days); SelfSigned = ($subj -eq $iss); San = $san; Eku = $eku; HasPrivateKey = $key; FriendlyName = $null } }
    $host2 = "$Name.corp.example".ToLower()
    $sqlc = & $mk 'A1B2C3D4E5' "CN=$host2" 'CN=Corp Issuing CA 1' $(if ($V -eq 'dr') { 20 } else { 300 }) @($host2, $Name.ToLower(), 'sqlprodlsn.corp.example') $true 'Server Authentication'
    $ssrs = & $mk 'F00DBEEF01' 'CN=reports.corp.example' 'CN=Corp Issuing CA 1' 200 @('reports.corp.example') $true 'Server Authentication'
    $s.Os.Certificates = $(if ($isNew) { @(& $mk '1122334455' "CN=$host2" "CN=$host2" 3650 @() $true 'Server Authentication') } else { @($sqlc, $ssrs) })
    $s.Os.SqlCert = $(if ($isNew) { $null } else { $sqlc })
    $s.Os.Firewall = [ordered]@{ ProfilesOn = 3; Ports = $(if ($isNew) { @('TCP:3389') } elseif ($V -eq 'dr') { @('TCP:14330', 'TCP:3389') } else { @('TCP:14330', 'TCP:5022', 'TCP:443', 'UDP:1434', 'TCP:3389') }); SqlProgramRules = @() }
    # ---- checklist-driven facets
    foreach ($k in @('db_health', 'log_shipping', 'backup_encryption', 'ag_states', 'server_misc', 'job_signatures')) { $s.Sql.Sections[$k] = 'ok' }
    foreach ($k in @('os_detail', 'environment', 'scheduled_tasks', 'defender', 'local_admins', 'disks', 'spn')) { $s.Os.Sections[$k] = 'ok' }
    $I.OwnNames = @($Name); $I.MasterDmk = $(if ($isNew) { 0 } else { 1 }); $I.Policies = $(if ($isNew) { 0 } else { 3 }); $I.CmsServers = 0; $I.NumaNodes = 2; $I.NumErrorLogs = $(if ($isNew) { $null } else { '30' })
    if (-not $isNew) {
        $s.Sql.DbHealth = @(
            [pscustomobject]@{ Db = 'SalesDB'; DmkBySmk = 1; Recovery = 'FULL'; LastFull = (Get-Date).AddDays(-1); LastLog = (Get-Date).AddMinutes(-15); LastCheckDb = (Get-Date).AddDays(-3); SuspectPages = 0; Mirrored = 0 },
            [pscustomobject]@{ Db = 'Finance'; DmkBySmk = 0; Recovery = 'FULL'; LastFull = (Get-Date).AddDays(-9); LastLog = (Get-Date).AddDays(-2); LastCheckDb = (Get-Date).AddDays(-40); SuspectPages = 0; Mirrored = 0 },
            [pscustomobject]@{ Db = 'DocStore'; DmkBySmk = 0; Recovery = 'FULL'; LastFull = (Get-Date).AddDays(-1); LastLog = (Get-Date).AddMinutes(-15); LastCheckDb = (Get-Date).AddDays(-3); SuspectPages = 2; Mirrored = 0 },
            [pscustomobject]@{ Db = 'Staging'; DmkBySmk = 0; Recovery = 'SIMPLE'; LastFull = (Get-Date).AddDays(-1); LastLog = $null; LastCheckDb = $null; SuspectPages = 0; Mirrored = 0 })
        $s.Sql.LogShipping = $(if ($V -eq 'ref') { @([pscustomobject]@{ Role = 'primary'; Db = 'Staging' }) } else { @() })
        $s.Sql.BackupCerts = @([pscustomobject]@{ Thumbprint = '0AB1C2D3E4F5061728394A5B6C7D8E9F00112233'; LastUsed = (Get-Date).AddDays(-1); Backups = 58 })
        $s.Sql.CodeUse = @($s.Sql.CodeUse | ForEach-Object { $_ | Add-Member -NotePropertyName DistTx -NotePropertyValue $(if ($_.Db -eq 'SalesDB') { 1 } else { 0 }) -Force; $_ | Add-Member -NotePropertyName SelfRef -NotePropertyValue $(if ($_.Db -eq 'SalesDB') { 3 } else { 0 }) -Force; $_ })
        foreach ($js in $s.Sql.JobSteps) { $js | Add-Member -NotePropertyName SelfRefs -NotePropertyValue $(if ($js.Job -eq 'Sales - Nightly export') { @('SQLPROD01') } else { @() }) -Force; $js | Add-Member -NotePropertyName IpRefs -NotePropertyValue $(if ($js.Job -eq 'Ops - Archive feeds') { @('10.20.30.41') } else { @() }) -Force }
        if ($V -eq 'ref') {
            $s.Sql.AgStates = @(
                [pscustomobject]@{ Replica = 'SQLPROD02'; Db = 'SalesDB'; SyncState = 'SYNCHRONIZED'; Health = 'HEALTHY'; SendQueueKB = 0; RedoQueueKB = 12; Suspended = 0 },
                [pscustomobject]@{ Replica = 'SQLDR01'; Db = 'SalesDB'; SyncState = 'SYNCHRONIZING'; Health = 'HEALTHY'; SendQueueKB = 2400000; RedoQueueKB = 300000; Suspended = 0 },
                [pscustomobject]@{ Replica = 'SQLDR01'; Db = 'Finance'; SyncState = 'NOT SYNCHRONIZING'; Health = 'NOT_HEALTHY'; SendQueueKB = 0; RedoQueueKB = 0; Suspended = 1 })
        }
    }
    $s.Sql.Jobs = @($s.Sql.Jobs | ForEach-Object { $_ | Add-Member -NotePropertyName LastOutcome -NotePropertyValue $(if ($_.Name -eq 'Ops - Archive feeds' -and $V -eq 'ref') { 0 } else { 1 }) -Force; $_ })
    if ($isNew) { $s.Sql.ConfigPending = [ordered]@{ 'max server memory (MB)' = '98304' } }
    $s.Os.Patch = [ordered]@{ InstallationType = 'Server'; DisplayVersion = $(if ($isNew) { '21H2' } else { '1607' }); Ubr = $(if ($isNew) { 2762 } elseif ($V -eq 'dr') { 7159 } else { 7428 }); LastHotfix = (Get-Date).AddDays($(if ($V -eq 'dr') { -120 } else { -20 })).ToString('yyyy-MM-dd') }
    $s.Os.PendingReboot = $(if ($isNew) { @('Windows Update') } else { @() })
    $s.Os.Fips = $(if ($isNew) { 1 } else { 0 })
    $s.Os.Msdtc = [ordered]@{ NetworkDtcAccess = $(if ($isNew) { 0 } else { 1 }); NetworkDtcAccessTransactions = $(if ($isNew) { 0 } else { 1 }); NetworkDtcAccessInbound = $(if ($isNew) { 0 } else { 1 }); NetworkDtcAccessOutbound = $(if ($isNew) { 0 } else { 1 }); XaTransactions = 0 }
    $s.Os.SqlInstances = $(if ($isNew) { @('Engine: MSSQLSERVER (MSSQL16.MSSQLSERVER)') } elseif ($V -eq 'ref') { @('Engine: MSSQLSERVER (MSSQL13.MSSQLSERVER)', 'Engine: REPORTS (MSSQL13.REPORTS)') } else { @('Engine: MSSQLSERVER (MSSQL13.MSSQLSERVER)') })
    $s.Os.EnvVars = [ordered]@{ Path = $(if ($isNew) { 'C:\Windows\system32;C:\Windows;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\' } else { 'C:\Windows\system32;C:\Windows;C:\Program Files\7-Zip;C:\oracle\product\19.0.0\client_1\bin;C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\' }); TEMP = 'C:\Windows\TEMP' }
    if (-not $isNew) { $s.Os.EnvVars['ORACLE_HOME'] = 'C:\oracle\product\19.0.0\client_1'; $s.Os.EnvVars['TNS_ADMIN'] = 'C:\oracle\network\admin' }
    $s.Os.ScheduledTasks = $(if ($isNew) { @() } else { @([pscustomobject]@{ Name = '\Corp\Upload feeds to SFTP'; State = 'Ready'; RunAs = 'CORP\svc_etl'; Actions = 'C:\Scripts\sftp-upload.cmd' }, [pscustomobject]@{ Name = '\Corp\Purge old backups'; State = 'Ready'; RunAs = 'SYSTEM'; Actions = 'powershell.exe -File D:\Scripts\Purge-Backups.ps1' }) })
    if ($V -eq 'dr') { $s.Os.ScheduledTasks = @($s.Os.ScheduledTasks | Where-Object { $_.Name -like '*Purge*' }) }
    $s.Os.Defender = [ordered]@{ RealTime = $true; Paths = $(if ($isNew) { @() } else { @('D:\SQLData', 'E:\SQLData', 'L:\SQLLogs', 'T:\TempDB', 'G:\Backup') }); Processes = $(if ($isNew) { @() } else { @('C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Binn\sqlservr.exe') }); Extensions = @(); Hidden = $false }
    $s.Os.LocalAdmins = $(if ($isNew) { @('SQLNEW01\Administrator', 'CORP\Domain Admins') } else { @("$Name\Administrator", 'CORP\Domain Admins', 'CORP\svc_sqlprod', 'CORP\SQL DBA Team') })
    $s.Os.Disks = @([pscustomobject]@{ Number = 0; Name = 'VMware Virtual disk'; SizeMB = 153600; Offline = $false; ReadOnly = $false; Health = 'Healthy'; Partition = 'GPT' }) + $(if ($isNew) { @([pscustomobject]@{ Number = 4; Name = 'VMware Virtual disk'; SizeMB = 1048576; Offline = $true; ReadOnly = $false; Health = 'Healthy'; Partition = 'GPT' }) } else { @() })
    $s.Os.PageFiles = @('C:\pagefile.sys (16384 MB)')
    $s.Spns = $(if ($isNew) { @() } else { @([pscustomobject]@{ Spn = "MSSQLSvc/$($Name.ToLower()).corp.example:14330"; Account = 'svc_sqlprod' }, [pscustomobject]@{ Spn = "MSSQLSvc/$($Name.ToLower()).corp.example"; Account = 'svc_sqlprod' }, [pscustomobject]@{ Spn = 'MSSQLSvc/sqlprodlsn.corp.example:14330'; Account = 'svc_sqlprod' }) })
    $s.Dns = [ordered]@{ Name = $Name; Ips = @('10.20.30.' + [string](40 + $Name.Length)); Reverse = @($(if ($V -eq 'dr') { 'sqldr-old.corp.example' } else { "$($Name.ToLower()).corp.example" })); Error = $null }
    return $s
}

function Get-DemoSnapshots {
    $R = New-DemoSnapshot 'SQLPROD01' 'Reference' 'ref'
    $N = New-DemoSnapshot 'SQLNEW01' 'Migration' 'new'
    $Y = New-DemoSnapshot 'SQLPROD02' 'Synchronous' 'sync'
    $D = New-DemoSnapshot 'SQLDR01' 'Asynchronous' 'dr'
    $all = @($R, $N, $Y, $D)
    # path checks as they would come back from CIM
    $R.PathChecks = [ordered]@{ 'D:\SQLData' = $true; 'E:\SQLData' = $true; 'E:\SQLData\DocStore_fs' = $true; 'L:\SQLLogs' = $true; 'D:\Apps\Exporter\SalesExport.exe' = $true; 'D:\Apps\Exporter\SalesExport.config' = $true; 'G:\JobLogs' = $true; 'D:\Scripts\Archive-Feeds.ps1' = $true; 'C:\Program Files\7-Zip\7z.exe' = $true; 'G:\Audit' = $true; 'G:\XE' = $true }
    $N.PathChecks = [ordered]@{ 'D:\SQLData' = $true; 'E:\SQLData' = $false; 'E:\SQLData\DocStore_fs' = $false; 'L:\SQLLogs' = $true; 'D:\Apps\Exporter\SalesExport.exe' = $false; 'D:\Apps\Exporter\SalesExport.config' = $false; 'G:\JobLogs' = $false; 'D:\Scripts\Archive-Feeds.ps1' = $false; 'C:\Program Files\7-Zip\7z.exe' = $false; 'G:\Audit' = $false; 'G:\XE' = $true; 'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\Binn\xpcorphash.dll' = $false; 'G:\Backup' = $true }
    $Y.PathChecks = [ordered]@{ 'D:\SQLData' = $true; 'E:\SQLData' = $true; 'E:\SQLData\DocStore_fs' = $true; 'L:\SQLLogs' = $true; 'D:\Apps\Exporter\SalesExport.exe' = $true; 'D:\Apps\Exporter\SalesExport.config' = $true; 'G:\JobLogs' = $true; 'D:\Scripts\Archive-Feeds.ps1' = $true; 'C:\Program Files\7-Zip\7z.exe' = $true; 'G:\Audit' = $true; 'G:\XE' = $true; 'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Binn\xpcorphash.dll' = $true; 'G:\Backup' = $true }
    $D.PathChecks = [ordered]@{ 'D:\SQLData' = $true; 'E:\SQLData' = $false; 'E:\SQLData\DocStore_fs' = $false; 'L:\SQLLogs' = $true; 'D:\Apps\Exporter\SalesExport.exe' = $true; 'D:\Apps\Exporter\SalesExport.config' = $false; 'G:\JobLogs' = $true; 'D:\Scripts\Archive-Feeds.ps1' = $false; 'C:\Program Files\7-Zip\7z.exe' = $true; 'G:\Audit' = $false; 'G:\XE' = $true; 'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Binn\xpcorphash.dll' = $false; 'G:\Backup' = $true }
    $N.PathInfo['D:\SQLData'] = [pscustomobject]@{ Compressed = $true; Encrypted = $false; Version = ''; IsDir = $true }
    $D.PathInfo['D:\Apps\Exporter\SalesExport.exe'] = [pscustomobject]@{ Compressed = $false; Encrypted = $false; Version = '2.3.0.0'; IsDir = $false }
    $R.PathInfo['D:\Apps\Exporter\SalesExport.exe'] = [pscustomobject]@{ Compressed = $false; Encrypted = $false; Version = '2.4.1.0'; IsDir = $false }
    $D.Sql.XProcs = @(); $D.Sql.StartupProcs = @('usp_StartupWarmCache'); $D.Sql.AuditSpecs = @()
    $Y.Sql.Messages = @($Y.Sql.Messages | Where-Object { $_.Id -ne 50002 })
    return $all
}
#endregion

#region ---------------------------------------------------------------- inventory (side-by-side view)
function Get-Inventory {
    # rows of { Facet; Item; Values = { server -> text } } - '' means absent on that server
    param($Snapshots)
    $rows = New-List
    $add = {
        param([string]$Facet, [string]$Item, [scriptblock]$Val)
        $vals = [ordered]@{}
        foreach ($s in $Snapshots) { $x = & $Val $s; $vals[$s.Name] = $(if ($null -eq $x) { '' } else { [string]$x }) }
        Add-To $rows ([pscustomobject]@{ Facet = $Facet; Item = $Item; Values = $vals })
    }
    $union = { param([scriptblock]$Get) $h = [ordered]@{}; foreach ($s in $Snapshots) { foreach ($k in (ConvertTo-Arr (& $Get $s))) { if ($k -and -not $h.Contains([string]$k)) { $h[[string]$k] = 1 } } }; return @($h.Keys | Sort-Object) }

    & $add 'Platform & OS' 'Windows' { param($s) if ($s.Os.Info) { "$($s.Os.Info.Caption) ($($s.Os.Info.Build))" } }
    & $add 'Platform & OS' 'Domain' { param($s) if ($s.Os.Info) { $s.Os.Info.Domain } }
    & $add 'Platform & OS' 'Time zone' { param($s) if ($s.Os.Info) { $s.Os.Info.TimeZone } }
    & $add 'Platform & OS' 'Locale' { param($s) if ($s.Os.Info) { $s.Os.Info.Locale } }
    & $add 'Platform & OS' 'Power plan' { param($s) $s.Os.PowerPlan }
    & $add 'Platform & OS' 'Cluster' { param($s) $s.Os.ClusterName }
    & $add 'Platform & OS' 'Hardware' { param($s) if ($s.Os.Info) { "$($s.Os.Info.Manufacturer) $($s.Os.Info.Model)" } }
    & $add 'Hardware & storage' 'Logical CPUs' { param($s) if ($s.Os.Info) { $s.Os.Info.Logical } else { Get-Prop $s.Sql.Instance 'CpuCount' } }
    & $add 'Hardware & storage' 'Memory' { param($s) if ($s.Os.Info) { Format-MB $s.Os.Info.MemoryMB } elseif (Get-Prop $s.Sql.Instance 'MemoryMB') { Format-MB (Get-Prop $s.Sql.Instance 'MemoryMB') } }
    foreach ($v in (& $union { param($s) (ConvertTo-Arr $s.Os.Volumes) | ForEach-Object { $_.Name } })) { & $add 'Hardware & storage' "Volume $v" { param($s) $x = (ConvertTo-Arr $s.Os.Volumes) | Where-Object { $_.Name -eq $v } | Select-Object -First 1; if ($x) { '{0} free of {1}, {2} KB AU' -f (Format-MB $x.FreeMB), (Format-MB $x.SizeMB), ([int]$x.BlockSize / 1024) } } }

    foreach ($k in @('ProductVersion', 'ProductUpdateLevel', 'Edition', 'Collation', 'IsHadrEnabled', 'FilestreamConfiguredLevel', 'IsFullTextInstalled', 'IsPolyBaseInstalled', 'IsIntegratedSecurityOnly', 'Ifi', 'MemoryModel', 'ServiceAccount', 'DefaultDataPath', 'DefaultLogPath', 'DefaultBackupPath')) {
        & $add 'SQL instance' $k { param($s) Get-Prop $s.Sql.Instance $k }
    }
    & $add 'SQL instance' 'TCP port' { param($s) Get-Prop $s.Sql.Net 'TcpPort' }
    & $add 'SQL instance' 'Force encryption' { param($s) Get-Prop $s.Sql.Net 'ForceEncryption' }
    & $add 'Configuration' 'Startup trace flags' { param($s) ((ConvertTo-Arr $s.Sql.TraceFlags) | ForEach-Object { "-T$_" }) -join ' ' }
    foreach ($k in (& $union { param($s) $s.Sql.Config.Keys })) { & $add 'Configuration' $k { param($s) Get-Prop $s.Sql.Config $k } }

    $swKey = { param($s) @((ConvertTo-Arr $s.Os.Software) | Where-Object { $_.Name -notmatch $script:SwNoise } | ForEach-Object { (Get-SwKey $_.Name $_.View).Clean + ' (' + (Get-SwKey $_.Name $_.View).Arch + ')' }) }
    foreach ($k in (& $union $swKey)) {
        & $add 'Installed software' $k { param($s) $x = (ConvertTo-Arr $s.Os.Software) | Where-Object { $_.Name -notmatch $script:SwNoise -and ((Get-SwKey $_.Name $_.View).Clean + ' (' + (Get-SwKey $_.Name $_.View).Arch + ')') -eq $k } | Select-Object -First 1; if ($x) { if ($x.Version) { $x.Version } else { 'installed' } } }
    }
    & $add 'Drivers & providers' '.NET Framework 4.x' { param($s) if ($s.Os.DotNet) { $s.Os.DotNet.Version } }
    & $add 'Drivers & providers' '.NET Framework 3.5' { param($s) if ($s.Os.DotNet) { if ($s.Os.DotNet.Net35) { 'installed' } else { '' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Os.OdbcDrivers) | ForEach-Object { "ODBC: $($_.Name) ($($_.Bitness))" } })) { & $add 'Drivers & providers' $k { param($s) if (@((ConvertTo-Arr $s.Os.OdbcDrivers) | Where-Object { "ODBC: $($_.Name) ($($_.Bitness))" -eq $k }).Count) { 'installed' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Os.Dsns) | ForEach-Object { "DSN: $($_.Name) ($($_.Bitness))" } })) { & $add 'Drivers & providers' $k { param($s) $x = (ConvertTo-Arr $s.Os.Dsns) | Where-Object { "DSN: $($_.Name) ($($_.Bitness))" -eq $k } | Select-Object -First 1; if ($x) { $x.Driver } } }
    foreach ($k in (& $union { param($s) $s.Os.OleDb.Keys })) { & $add 'Drivers & providers' "OLE DB: $k" { param($s) $x = $s.Os.OleDb[$k]; if ($x) { (@($(if ($x.x64) { 'x64' }), $(if ($x.x86) { 'x86' })) | Where-Object { $_ }) -join ' + ' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Os.Aliases) | ForEach-Object { "Alias: $($_.Name) ($($_.Bitness))" } })) { & $add 'Drivers & providers' $k { param($s) $x = (ConvertTo-Arr $s.Os.Aliases) | Where-Object { "Alias: $($_.Name) ($($_.Bitness))" -eq $k } | Select-Object -First 1; if ($x) { $x.Target } } }
    foreach ($k in (& $union { param($s) if ($s.Os.Tls) { $s.Os.Tls.Keys } })) { & $add 'Drivers & providers' "SCHANNEL $k" { param($s) if ($s.Os.Tls) { $s.Os.Tls[$k] } } }
    foreach ($k in (& $union { param($s) $s.Os.Features })) { & $add 'Windows features & services' "Feature: $k" { param($s) if ($s.Os.Features -contains $k) { 'installed' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Os.Services) | ForEach-Object { Get-SvcFamily $_.Name } })) { & $add 'Windows features & services' "Service: $k" { param($s) $x = (ConvertTo-Arr $s.Os.Services) | Where-Object { (Get-SvcFamily $_.Name) -eq $k } | Select-Object -First 1; if ($x) { "$($x.StartMode) / $($x.State) / $($x.Account)" } } }
    foreach ($k in (& $union { param($s) $s.Os.PsModules })) { & $add 'Agent & automation' "PS module: $k" { param($s) if ($s.Os.PsModules -contains $k) { 'installed' } } }

    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Assemblies) | ForEach-Object { "$($_.Db).$($_.Name)" } })) { & $add 'CLR & assemblies' $k { param($s) $x = (ConvertTo-Arr $s.Sql.Assemblies) | Where-Object { "$($_.Db).$($_.Name)" -eq $k } | Select-Object -First 1; if ($x) { $x.PermissionSet } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.LinkedServers) | ForEach-Object { $_.Name } })) { & $add 'Linked servers & external' "Linked server: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.LinkedServers) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { "$($x.Provider) -> $($x.DataSource)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Credentials) | ForEach-Object { $_.Name } })) { & $add 'Security objects' "Credential: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Credentials) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $x.CredIdentity } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Certificates) | ForEach-Object { $_.Name } })) { & $add 'Security objects' "Certificate: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Certificates) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { "expires $($x.Expiry)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Audits) | ForEach-Object { $_.Name } })) { & $add 'Security objects' "Audit: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Audits) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { "$($x.Type) $($x.Path)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.ServerTriggers) | ForEach-Object { $_.Name } })) { & $add 'Security objects' "Server trigger: $k" { param($s) if (@((ConvertTo-Arr $s.Sql.ServerTriggers) | Where-Object { $_.Name -eq $k }).Count) { 'present' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Endpoints) | ForEach-Object { $_.Name } })) { & $add 'Security objects' "Endpoint: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Endpoints) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { "$($x.Type) :$($x.Port) $($x.Encryption) $($x.State)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Jobs) | ForEach-Object { $_.Name } })) { & $add 'Agent & automation' "Job: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Jobs) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $(if ($x.Enabled) { 'enabled' } else { 'disabled' }) + " / $($x.Owner)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Proxies) | ForEach-Object { $_.Name } })) { & $add 'Agent & automation' "Proxy: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Proxies) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $x.Credential } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Operators) | ForEach-Object { $_.Name } })) { & $add 'Agent & automation' "Operator: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Operators) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $x.Email } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Mail) | ForEach-Object { $_.Profile } })) { & $add 'Agent & automation' "Mail profile: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Mail) | Where-Object { $_.Profile -eq $k } | Select-Object -First 1; if ($x) { "$($x.SmtpServer):$($x.Port)" } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.MailAccounts | ForEach-Object { $_.Account } })) { & $add 'Agent & automation' "Mail account: $k" { param($s) $x = ConvertTo-Arr $s.Sql.MailAccounts | Where-Object { $_.Account -eq $k } | Select-Object -First 1; if ($x) { "$($x.Email) via $($x.SmtpServer):$($x.Port) ssl=$($x.Ssl)" } } }
    & $add 'Agent & automation' 'Agent mail profile' { param($s) if ($s.Os.AgentReg) { "$($s.Os.AgentReg['DatabaseMailProfile']) (UseDatabaseMail=$($s.Os.AgentReg['UseDatabaseMail']))" } }
    & $add 'Agent & automation' 'Agent fail-safe operator' { param($s) if ($s.Os.AgentReg) { $s.Os.AgentReg['AlertFailSafeOperator'] } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.Schedules | ForEach-Object { "Schedule: $($_.Job) / $($_.Schedule)" } })) { & $add 'Agent & automation' $k { param($s) $x = ConvertTo-Arr $s.Sql.Schedules | Where-Object { "Schedule: $($_.Job) / $($_.Schedule)" -eq $k } | Select-Object -First 1; if ($x) { $(if ([int]$x.Enabled) { '' } else { '[disabled] ' }) + $x.Text } } }
    foreach ($k in (& $union { param($s) $s.Sql.JobSignatures.Keys })) { & $add 'Agent & automation' "Steps: $k" { param($s) if ($s.Sql.JobSignatures.ContainsKey($k)) { $s.Sql.JobSignatures[$k] } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.ProxySubsystems | ForEach-Object { "Proxy subsystem: $($_.Proxy) -> $($_.Subsystem)" } })) { & $add 'Agent & automation' $k { param($s) if (@(ConvertTo-Arr $s.Sql.ProxySubsystems | Where-Object { "Proxy subsystem: $($_.Proxy) -> $($_.Subsystem)" -eq $k }).Count) { 'granted' } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.MaintPlans | ForEach-Object { $_.Name } })) { & $add 'Agent & automation' "Maintenance plan: $k" { param($s) if (@(ConvertTo-Arr $s.Sql.MaintPlans | Where-Object { $_.Name -eq $k }).Count) { 'present' } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.Files | Where-Object { $_.DbId -le 4 } | ForEach-Object { "$($_.Db) / $($_.LogicalName)" } })) { & $add 'SQL instance' "System file: $k" { param($s) $x = ConvertTo-Arr $s.Sql.Files | Where-Object { "$($_.Db) / $($_.LogicalName)" -eq $k } | Select-Object -First 1; if ($x) { "$($x.Path)  $(Format-MB $x.SizeMB) grow $($x.Growth)" } } }
    & $add 'SQL instance' 'AG listener' { param($s) $x = ConvertTo-Arr $s.Sql.Ag | Where-Object { $_.Listener } | Select-Object -First 1; if ($x) { "$($x.Listener):$(Get-Prop $x 'ListenerPort')" } }
    & $add 'Security objects' 'SQL TLS certificate' { param($s) $c = $s.Os.SqlCert; if ($c) { if (Get-Prop $c 'Missing') { "MISSING $($c.Thumbprint)" } else { "$($c.Subject) exp $($c.NotAfter)" } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Os.Certificates | Where-Object { $_.HasPrivateKey } | ForEach-Object { $_.Subject } })) { & $add 'Security objects' "Machine cert: $k" { param($s) $x = ConvertTo-Arr $s.Os.Certificates | Where-Object { $_.Subject -eq $k } | Sort-Object NotAfterDate -Descending | Select-Object -First 1; if ($x) { "exp $($x.NotAfter)" } } }
    & $add 'Windows features & services' 'Firewall profiles on' { param($s) if ($s.Os.Firewall) { $s.Os.Firewall.ProfilesOn } }
    foreach ($k in (& $union { param($s) if ($s.Os.Firewall) { $s.Os.Firewall.Ports } })) { & $add 'Windows features & services' "Firewall allow $k" { param($s) if ($s.Os.Firewall -and @($s.Os.Firewall.Ports) -contains $k) { 'allowed' } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Os.Services | Where-Object { Get-SvcFamily $_.Name } | ForEach-Object { Get-SvcFamily $_.Name } })) { & $add 'Windows features & services' "Account: $k" { param($s) $x = ConvertTo-Arr $s.Os.Services | Where-Object { (Get-SvcFamily $_.Name) -eq $k } | Select-Object -First 1; if ($x) { $x.Account } } }
    foreach ($k in (& $union { param($s) if ($s.Os.ProviderOptions) { $s.Os.ProviderOptions.Keys } })) { & $add 'Linked servers & external' "Provider options: $k" { param($s) if ($s.Os.ProviderOptions -and $s.Os.ProviderOptions.Contains($k)) { $o = $s.Os.ProviderOptions[$k]; (@($o.Keys | ForEach-Object { "$_=$($o[$_])" }) -join ' ') } } }
    foreach ($k in (& $union { param($s) ConvertTo-Arr $s.Sql.BackupDevices | ForEach-Object { $_.Name } })) { & $add 'Server objects' "Backup device: $k" { param($s) $x = ConvertTo-Arr $s.Sql.BackupDevices | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $x.Path } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.SsisProjects) | ForEach-Object { "$($_.Folder)\$($_.Project)" } })) { & $add 'SSIS / SSRS / SSAS' "SSIS project: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.SsisProjects) | Where-Object { "$($_.Folder)\$($_.Project)" -eq $k } | Select-Object -First 1; if ($x) { "deployed $($x.Deployed)" } } }
    foreach ($k in (& $union { param($s) $s.Sql.StartupProcs })) { & $add 'Server objects' "Startup proc: $k" { param($s) if ($s.Sql.StartupProcs -contains $k) { 'present' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.SysDbObjects) | ForEach-Object { "$($_.Db).$($_.Name)" } })) { & $add 'Server objects' "Object: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.SysDbObjects) | Where-Object { "$($_.Db).$($_.Name)" -eq $k } | Select-Object -First 1; if ($x) { $x.Type } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.XProcs) | ForEach-Object { $_.Name } })) { & $add 'Server objects' "Extended proc: $k" { param($s) $x = (ConvertTo-Arr $s.Sql.XProcs) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { $x.Dll } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Messages) | ForEach-Object { "$($_.Id)/$($_.Lang)" } })) { & $add 'Server objects' "Message $k" { param($s) $x = (ConvertTo-Arr $s.Sql.Messages) | Where-Object { "$($_.Id)/$($_.Lang)" -eq $k } | Select-Object -First 1; if ($x) { "sev $($x.Severity)" } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.XeSessions) | ForEach-Object { $_.Name } })) { & $add 'Server objects' "XE session: $k" { param($s) if (@((ConvertTo-Arr $s.Sql.XeSessions) | Where-Object { $_.Name -eq $k }).Count) { 'present' } } }
    foreach ($k in (& $union { param($s) (ConvertTo-Arr $s.Sql.Databases) | Where-Object { $_.Id -gt 4 } | ForEach-Object { $_.Name } })) { & $add 'Databases' $k { param($s) $x = (ConvertTo-Arr $s.Sql.Databases) | Where-Object { $_.Name -eq $k } | Select-Object -First 1; if ($x) { "compat $($x.Compat), $(Format-MB $x.SizeMB), $($x.Recovery)$(if ($x.InAg) { ', AG' })$(if ($x.Encrypted) { ', TDE' })" } } }
    foreach ($k in (& $union { param($s) $s.PathChecks.Keys })) { & $add 'Files & paths' $k { param($s) if ($s.PathChecks.Contains($k)) { if ($null -eq $s.PathChecks[$k]) { '?' } elseif ($s.PathChecks[$k]) { 'exists' } else { 'MISSING' } } } }
    return $rows.ToArray()
}
#endregion

#region ---------------------------------------------------------------- simple outputs
function Get-SafeName { param([string]$n) return ($n -replace '[\\/:*?"<>|,]', '_') }
function Get-SnapshotExport {
    param($Snapshots)
    $o = New-List
    foreach ($s in $Snapshots) {
        $c = [ordered]@{}
        foreach ($k in $s.Keys) { if ($k -ne 'CimSession') { $c[$k] = $s[$k] } }
        Add-To $o $c
    }
    return $o.ToArray()
}

function Write-TxtReport {
    param($Path, $Meta, $Readiness, $Findings)
    $sb = New-Object System.Text.StringBuilder
    $w = { param($t) [void]$sb.AppendLine($t) }
    $wrap = {
        param([string]$Text, [int]$Width, [string]$Indent)
        $words = ([string]$Text) -split '\s+'; $line = ''
        foreach ($x in $words) { if (($line.Length + $x.Length + 1) -gt $Width) { [void]$sb.AppendLine($Indent + $line.TrimEnd()); $line = '' }; $line += $x + ' ' }
        if ($line.Trim()) { [void]$sb.AppendLine($Indent + $line.TrimEnd()) }
    }
    & $w ('=' * 100)
    & $w ("  SQL SERVER MIGRATION / HA READINESS REPORT   ($script:ToolName $script:ToolVersion, read-only)")
    & $w ('=' * 100)
    & $w ("  Generated : $($Meta.Generated)    Run by: $($Meta.RunBy) on $($Meta.RunHost)")
    & $w ("  Reference : $($Meta.Reference)" + $(if ($Meta.Demo) { '    ** DEMO DATA **' } else { '' }))
    if ($Meta.PathMap) { & $w ("  Path map  : $($Meta.PathMap)") }
    & $w ''
    & $w '  SUMMARY'
    & $w ('  ' + '-' * 96)
    & $w ('  {0,-22} {1,-13} {2,-17} {3,5}  {4,7} {5,5} {6,6} {7,4} {8,8}' -f 'Target', 'Role', 'Verdict', 'Score', 'Blocker', 'High', 'Medium', 'Low', 'Install')
    foreach ($r in $Readiness) { & $w ('  {0,-22} {1,-13} {2,-17} {3,5}  {4,7} {5,5} {6,6} {7,4} {8,8}' -f $r.Target, $r.Role, $r.Verdict, $r.Score, $r.Blocker, $r.High, $r.Medium, $r.Low, $r.Install) }
    foreach ($r in $Readiness) {
        $f = @($Findings | Where-Object { $_.Target -eq $r.Target })
        & $w ''; & $w ('#' * 100)
        & $w ("  TARGET $($r.Target)  [$($r.Role)]  ->  $($r.Verdict)  (score $($r.Score)/100, $($r.Checks) checks)")
        & $w ('#' * 100)
        & $w ''; & $w '  Facet status'
        foreach ($cat in $r.Categories.Keys) { $cs = $r.Categories[$cat]; & $w ('    {0,-30} {1,-8} {2,3} issue(s) / {3} checks' -f $cat, $cs.Worst, $cs.Findings, $cs.Checks) }
        $inst = @($f | Where-Object { $_.Install } | Sort-Object { - $script:SevOrder[$_.Severity] }, Category, Item)
        if ($inst.Count) {
            & $w ''; & $w '  NEEDS INSTALLATION ON THIS TARGET'
            foreach ($x in $inst) { & $w ('    [ ] {0,-8} {1}' -f $x.Severity, $x.Item); if ($x.Reference) { & $w ("             reference: $($x.Reference)") } }
        }
        foreach ($sev in @('Blocker', 'High', 'Medium', 'Low', 'Info')) {
            $list = @($f | Where-Object { $_.Severity -eq $sev } | Sort-Object Category, Item)
            if (-not $list.Count) { continue }
            & $w ''; & $w ("  $($sev.ToUpper()) ($($list.Count))"); & $w ('  ' + '-' * 96)
            foreach ($x in $list) {
                & $w ("  * [$($x.Category)] $($x.Item)")
                if ($x.Reference -or $x.TargetValue) { & $w ("      reference: $($x.Reference)  |  target: $($x.TargetValue)") }
                & $wrap ("action: " + $x.Action) 92 '      '
                if ($x.Fix) { & $w ("      tool: $($x.Fix)") }
                if ($x.Detail) { & $wrap ("detail: " + $x.Detail) 92 '      ' }
            }
        }
    }
    & $w ''; & $w ('#' * 100); & $w '  CHECKLIST COVERAGE (team checklist A-K)'; & $w ('#' * 100)
    foreach ($cr in (Get-CoverageRows)) { & $w ''; & $w ("  {0,-4} {1,-40} [{2}]" -f $cr.Ref, $cr.Area, $cr.Status); & $wrap ('checked: ' + $cr.Checked) 92 '       '; & $wrap ('manual : ' + $cr.Manual) 92 '       ' }
    & $w ''; & $w ('=' * 100)
    & $w '  Read-only collection: catalog SELECTs + CIM/registry reads. Nothing was changed on any server.'
    & $w '  Logins/SIDs/permissions are covered by LoginSync.ps1.'
    [System.IO.File]::WriteAllText($Path, $sb.ToString().Replace("`n", "`r`n").Replace("`r`r`n", "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

function ConvertTo-SqlLit { param($v) if ($null -eq $v) { return 'NULL' } return "N'" + ([string]$v).Replace("'", "''") + "'" }

function Write-PrepSql {
    param($Path, $R, $T, $Findings)
    $sb = New-Object System.Text.StringBuilder
    $w = { param($t) [void]$sb.AppendLine($t) }
    $srv = [string](Get-Prop $T.Sql.Instance 'ServerName' $T.Name)
    & $w "/* ==========================================================================================="
    & $w "   PREP SCRIPT for target $($T.Name) [$($T.Role)]  -  reference $($R.Name)"
    & $w "   Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $script:ToolName $script:ToolVersion"
    & $w ""
    & $w "   REVIEW BEFORE USE. The script refuses to run unless:"
    & $w "     1. @@SERVERNAME = '$srv'  and"
    & $w "     2. you set @Reviewed = 1 below after reading every section."
    & $w "   Everything is IF NOT EXISTS / idempotent; anything that needs a secret or downtime is commented."
    & $w "   Installs (drivers, providers, runtimes, Windows features) are NOT T-SQL - see the report's"
    & $w "   'Needs installation' list."
    & $w "   =========================================================================================== */"
    & $w "SET NOCOUNT ON;"
    & $w "DECLARE @Reviewed bit = 0;   -- <<< set to 1 after review"
    & $w "IF @@SERVERNAME <> $(ConvertTo-SqlLit $srv) BEGIN RAISERROR('Wrong server: this script is for $srv.', 16, 1); SET NOEXEC ON; END"
    & $w "IF @Reviewed = 0 BEGIN RAISERROR('Set @Reviewed = 1 after reviewing the script.', 16, 1); SET NOEXEC ON; END"
    & $w "GO"
    & $w ""
    # 1 sp_configure
    & $w "/* ---- 1. sp_configure differences (reference value) ---- */"
    $cfg = @($Findings | Where-Object { $_.Item -like 'sp_configure:*' -and $_.Severity -ne 'Info' })
    if ($cfg.Count) {
        & $w "EXEC sys.sp_configure N'show advanced options', 1; RECONFIGURE;"
        foreach ($c in $cfg) {
            $name = $c.Item.Substring(14)
            $line = "EXEC sys.sp_configure $(ConvertTo-SqlLit $name), $($c.Reference);   -- target now: $($c.TargetValue)"
            if ($name -match 'max server memory|min server memory|affinity|max worker threads|priority boost|lightweight pooling') { $line = "-- $line   (host-specific: size for THIS server)" }
            elseif ($name -match '^(clr strict security|xp_cmdshell|Ole Automation Procedures|Ad Hoc Distributed Queries|cross db ownership chaining|remote access|c2 audit mode|common criteria compliance enabled)$') { $line = "-- $line   (security-sensitive: enable only if the report shows it is really used$(if ($name -eq 'clr strict security') { '; prefer signing assemblies / sp_add_trusted_assembly' }))" }
            & $w $line
        }
        & $w "RECONFIGURE;"
    } else { & $w "-- none" }
    & $w "GO"; & $w ""
    # 2 trace flags
    & $w "/* ---- 2. Startup trace flags (set in SQL Server Configuration Manager > Startup Parameters; restart) ---- */"
    $tf = @($Findings | Where-Object { $_.Item -like 'Startup trace flag*' -and $_.Reference -eq 'on' })
    if ($tf.Count) { foreach ($x in $tf) { & $w ("--   {0}   [{1}] {2}" -f ($x.Item -replace 'Startup trace flag ', ''), $x.Severity, $x.Action) }; & $w "-- For the current session only (lost on restart):"; foreach ($x in $tf) { if ($x.Severity -ne 'Info') { & $w ("-- DBCC TRACEON({0}, -1);" -f ($x.Item -replace '.*-T', '')) } } } else { & $w "-- none" }
    & $w ""
    # 3 operators / alerts / mail
    & $w "/* ---- 3. Agent operators and alerts ---- */"
    foreach ($o in $R.Sql.Operators) {
        if (@((ConvertTo-Arr $T.Sql.Operators) | Where-Object { $_.Name -eq $o.Name }).Count) { continue }
        & $w "IF NOT EXISTS (SELECT 1 FROM msdb.dbo.sysoperators WHERE name = $(ConvertTo-SqlLit $o.Name))"
        & $w "    EXEC msdb.dbo.sp_add_operator @name = $(ConvertTo-SqlLit $o.Name), @enabled = $([int]$o.Enabled), @email_address = $(ConvertTo-SqlLit $o.Email);"
    }
    foreach ($a in $R.Sql.Alerts) {
        if (@((ConvertTo-Arr $T.Sql.Alerts) | Where-Object { $_.Name -eq $a.Name }).Count) { continue }
        & $w "IF NOT EXISTS (SELECT 1 FROM msdb.dbo.sysalerts WHERE name = $(ConvertTo-SqlLit $a.Name))"
        if ([int]$a.MessageId -gt 0) { & $w "    EXEC msdb.dbo.sp_add_alert @name = $(ConvertTo-SqlLit $a.Name), @message_id = $($a.MessageId), @enabled = $([int]$a.Enabled), @include_event_description_in = 1;" }
        else { & $w "    EXEC msdb.dbo.sp_add_alert @name = $(ConvertTo-SqlLit $a.Name), @severity = $($a.Severity), @enabled = $([int]$a.Enabled), @include_event_description_in = 1;" }
        & $w "-- EXEC msdb.dbo.sp_add_notification @alert_name = $(ConvertTo-SqlLit $a.Name), @operator_name = N'<operator>', @notification_method = 1;"
    }
    if (@($R.Sql.Mail).Count -and -not @($T.Sql.Mail).Count) { & $w "-- Database Mail: run Copy-DbaDbMail -Source $($R.Name) -Destination $($T.Name), or create profile/account manually (SMTP password needed)." }
    & $w "GO"; & $w ""
    # 4 custom messages
    & $w "/* ---- 4. Custom error messages ---- */"
    foreach ($m in $R.Sql.Messages) {
        if (@((ConvertTo-Arr $T.Sql.Messages) | Where-Object { $_.Id -eq $m.Id -and $_.Lang -eq $m.Lang }).Count) { continue }
        $lang = if ([int]$m.Lang -eq 1033) { 'us_english' } else { "<language for lcid $($m.Lang)>" }
        & $w "IF NOT EXISTS (SELECT 1 FROM sys.messages WHERE message_id = $($m.Id) AND language_id = $($m.Lang))"
        & $w "    EXEC sys.sp_addmessage @msgnum = $($m.Id), @severity = $($m.Severity), @msgtext = $(ConvertTo-SqlLit $m.Text), @lang = N'$lang', @with_log = $(if ([int]$m.Logged) { "'TRUE'" } else { "'FALSE'" });"
    }
    & $w "GO"; & $w ""
    # 5 linked servers + credentials
    & $w "/* ---- 5. Linked servers (skeleton - remote passwords are not readable; prefer Copy-DbaLinkedServer) ---- */"
    foreach ($l in $R.Sql.LinkedServers) {
        if (@((ConvertTo-Arr $T.Sql.LinkedServers) | Where-Object { $_.Name -eq $l.Name }).Count) { continue }
        $prov = [string]$l.Provider; if ($prov -like 'SQLNCLI*' -and [int](Get-Prop $T.Sql.Instance 'Major' 0) -ge 16) { $prov = 'MSOLEDBSQL' }
        & $w "IF NOT EXISTS (SELECT 1 FROM sys.servers WHERE name = $(ConvertTo-SqlLit $l.Name))"
        & $w "    EXEC master.dbo.sp_addlinkedserver @server = $(ConvertTo-SqlLit $l.Name), @srvproduct = $(ConvertTo-SqlLit $(if ($l.Product) { $l.Product } else { '' })), @provider = $(ConvertTo-SqlLit $prov), @datasrc = $(ConvertTo-SqlLit $l.DataSource)$(if ($l.Catalog) { ', @catalog = ' + (ConvertTo-SqlLit $l.Catalog) } else { '' })$(if ($l.ProviderString) { ', @provstr = ' + (ConvertTo-SqlLit $l.ProviderString) } else { '' });"
        if ([int]$l.RpcOut) { & $w "EXEC master.dbo.sp_serveroption @server = $(ConvertTo-SqlLit $l.Name), @optname = N'rpc out', @optvalue = N'true';" }
        & $w "-- EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname = $(ConvertTo-SqlLit $l.Name), @useself = N'False', @locallogin = NULL, @rmtuser = N'<remote user>', @rmtpassword = N'<secret>';"
    }
    & $w ""
    & $w "/* Credentials (secrets must be re-entered) */"
    foreach ($c in $R.Sql.Credentials) {
        if (@((ConvertTo-Arr $T.Sql.Credentials) | Where-Object { $_.Name -eq $c.Name }).Count) { continue }
        & $w "-- IF NOT EXISTS (SELECT 1 FROM sys.credentials WHERE name = $(ConvertTo-SqlLit $c.Name)) CREATE CREDENTIAL [$($c.Name.Replace(']', ']]'))] WITH IDENTITY = $(ConvertTo-SqlLit $c.CredIdentity), SECRET = N'<secret>';"
    }
    & $w "GO"; & $w ""
    # 6 TDE / certificates
    & $w "/* ---- 6. Certificates (TDE certificates must exist BEFORE restore/seeding) ---- */"
    foreach ($f in @($Findings | Where-Object { $_.Item -like 'TDE certificate:*' -or $_.Item -like 'Certificate in master:*' })) {
        $cn = ($f.Item -split ':\s*', 2)[1]
        & $w "-- On $($R.Name):  BACKUP CERTIFICATE [$cn] TO FILE = N'<share>\$cn.cer' WITH PRIVATE KEY (FILE = N'<share>\$cn.pvk', ENCRYPTION BY PASSWORD = N'<strong password>');"
        & $w "-- On $($T.Name):  IF NOT EXISTS (SELECT 1 FROM sys.symmetric_keys WHERE name = N'##MS_DatabaseMasterKey##') CREATE MASTER KEY ENCRYPTION BY PASSWORD = N'<strong password>';"
        & $w "--                 CREATE CERTIFICATE [$cn] FROM FILE = N'<share>\$cn.cer' WITH PRIVATE KEY (FILE = N'<share>\$cn.pvk', DECRYPTION BY PASSWORD = N'<strong password>');"
    }
    & $w ""
    # 7 post restore
    & $w "/* ---- 7. After the databases are restored on this server (uncomment as needed) ---- */"
    foreach ($f in @($Findings | Where-Object { $_.Item -like 'TRUSTWORTHY:*' })) { $db = ($f.Item -split ':\s*', 2)[1]; & $w "-- ALTER DATABASE [$($db.Replace(']', ']]'))] SET TRUSTWORTHY ON;   -- prefer signing modules" }
    foreach ($f in @($Findings | Where-Object { $_.Item -like 'Service Broker:*' })) { $db = ($f.Item -split ':\s*', 2)[1]; & $w "-- ALTER DATABASE [$($db.Replace(']', ']]'))] SET ENABLE_BROKER WITH ROLLBACK IMMEDIATE;" }
    foreach ($f in @($Findings | Where-Object { $_.Item -like 'Change Data Capture:*' })) { $db = ($f.Item -split ':\s*', 2)[1]; & $w "-- USE [$($db.Replace(']', ']]'))]; EXEC sys.sp_cdc_add_job @job_type = N'capture'; EXEC sys.sp_cdc_add_job @job_type = N'cleanup';" }
    foreach ($f in @($Findings | Where-Object { $_.Item -like 'Database owner:*' })) { $db = ($f.Item -split ':\s*', 2)[1]; & $w "-- ALTER AUTHORIZATION ON DATABASE::[$($db.Replace(']', ']]'))] TO [$($f.Reference)];   -- create the login first" }
    & $w "-- Then: DBCC CHECKDB, update statistics, re-run $script:ToolName and LoginSync as proof."
    & $w "GO"
    & $w "SET NOEXEC OFF;"
    [System.IO.File]::WriteAllText($Path, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
}
#endregion

#region ---------------------------------------------------------------- Word (.docx) writer - plain OpenXML, no Office needed
$script:SevFill = @{ 'Blocker' = 'F4C7C3'; 'High' = 'FBD9BE'; 'Medium' = 'FCEFB4'; 'Low' = 'D6E6F5'; 'Info' = 'E7E7E7'; 'OK' = 'CFEBD3'; 'n/a' = 'F2F2F2' }
$script:VerdictFill = @{ 'NOT READY' = 'F4C7C3'; 'AT RISK' = 'FBD9BE'; 'READY WITH NOTES' = 'FCEFB4'; 'READY' = 'CFEBD3'; 'NOT VERIFIED' = 'E7E7E7' }

function ConvertTo-XmlText { param([string]$t) if ($null -eq $t) { return '' } return [System.Security.SecurityElement]::Escape(($t -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')) }

function New-DocxPara {
    param([string]$Text, [string]$Style = '', [switch]$Bold, [string]$Color = '', [int]$Size = 0, [switch]$PageBreakBefore)
    $ppr = ''
    if ($Style -or $PageBreakBefore) { $ppr = '<w:pPr>' + $(if ($Style) { "<w:pStyle w:val=`"$Style`"/>" }) + $(if ($PageBreakBefore) { '<w:pageBreakBefore/>' }) + '</w:pPr>' }
    $rpr = ''
    if ($Bold -or $Color -or $Size) { $rpr = '<w:rPr>' + $(if ($Bold) { '<w:b/>' }) + $(if ($Color) { "<w:color w:val=`"$Color`"/>" }) + $(if ($Size) { "<w:sz w:val=`"$Size`"/>" }) + '</w:rPr>' }
    return "<w:p>$ppr<w:r>$rpr<w:t xml:space=`"preserve`">$(ConvertTo-XmlText $Text)</w:t></w:r></w:p>"
}

function New-DocxTable {
    # $Rows: array of string arrays; $Fill: scriptblock (rowIndex, colIndex, text) -> hex fill or ''
    param([string[]]$Headers, $Rows, [int[]]$Widths, [scriptblock]$Fill = $null)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<w:tbl><w:tblPr><w:tblStyle w:val="ReportTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="70" w:type="dxa"/><w:right w:w="70" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>')
    foreach ($w in $Widths) { [void]$sb.Append("<w:gridCol w:w=`"$w`"/>") }
    [void]$sb.Append('</w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr>')
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$($Widths[$i])`" w:type=`"dxa`"/><w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"1F3A5F`"/></w:tcPr><w:p><w:pPr><w:spacing w:before=`"20`" w:after=`"20`"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val=`"FFFFFF`"/><w:sz w:val=`"17`"/></w:rPr><w:t xml:space=`"preserve`">$(ConvertTo-XmlText $Headers[$i])</w:t></w:r></w:p></w:tc>")
    }
    [void]$sb.Append('</w:tr>')
    $ri = 0
    foreach ($row in $Rows) {
        [void]$sb.Append('<w:tr><w:trPr><w:cantSplit/></w:trPr>')
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $txt = [string]$row[$i]
            $cellFill = if ($Fill) { [string](& $Fill $ri $i $txt) } else { '' }
            if (-not $cellFill -and ($ri % 2) -eq 1) { $cellFill = 'F7F9FC' }
            $shd = if ($cellFill) { "<w:shd w:val=`"clear`" w:color=`"auto`" w:fill=`"$cellFill`"/>" } else { '' }
            [void]$sb.Append("<w:tc><w:tcPr><w:tcW w:w=`"$($Widths[$i])`" w:type=`"dxa`"/>$shd</w:tcPr><w:p><w:pPr><w:spacing w:before=`"10`" w:after=`"10`"/></w:pPr><w:r><w:rPr><w:sz w:val=`"16`"/></w:rPr><w:t xml:space=`"preserve`">$(ConvertTo-XmlText $txt)</w:t></w:r></w:p></w:tc>")
        }
        [void]$sb.Append('</w:tr>')
        $ri++
    }
    [void]$sb.Append('</w:tbl>')
    [void]$sb.Append('<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>')
    return $sb.ToString()
}

function Write-DocxReport {
    param($Path, $Meta, $Readiness, $Findings)
    $b = New-Object System.Text.StringBuilder
    $sevFillFn = { param($r, $c, $t) if ($c -eq 0 -and $script:SevFill.ContainsKey($t)) { return $script:SevFill[$t] } return '' }
    [void]$b.Append((New-DocxPara 'SQL Server Migration / HA Readiness' 'Title'))
    [void]$b.Append((New-DocxPara ("Reference: {0}    Generated: {1}    Run by: {2} on {3}{4}" -f $Meta.Reference, $Meta.Generated, $Meta.RunBy, $Meta.RunHost, $(if ($Meta.Demo) { '    DEMO DATA' } else { '' })) 'Subtle'))
    if ($Meta.PathMap) { [void]$b.Append((New-DocxPara "Path map: $($Meta.PathMap)" 'Subtle')) }
    [void]$b.Append((New-DocxPara 'Summary' 'Heading1'))
    [void]$b.Append((New-DocxPara 'Each target is compared with the reference server across platform, hardware, software, drivers, services, SQL Server configuration and server-level objects. Severity: Blocker = the move/failover cannot succeed; High = something will break after cut-over; Medium = should be fixed or consciously accepted; Low/Info = review.'))
    $rows = @($Readiness | ForEach-Object { , @($_.Target, $_.Role, $_.Verdict, "$($_.Score)", "$($_.Blocker)", "$($_.High)", "$($_.Medium)", "$($_.Low)", "$($_.Install)") })
    [void]$b.Append((New-DocxTable @('Target', 'Role', 'Verdict', 'Score', 'Blocker', 'High', 'Medium', 'Low', 'To install') $rows @(2600, 1600, 2300, 900, 1000, 900, 1000, 900, 1200) { param($r, $c, $t) if ($c -eq 2 -and $script:VerdictFill.ContainsKey($t)) { return $script:VerdictFill[$t] } if ($c -eq 4 -and $t -ne '0') { return $script:SevFill['Blocker'] } if ($c -eq 5 -and $t -ne '0') { return $script:SevFill['High'] } return '' }))

    [void]$b.Append((New-DocxPara 'Facet heat map (worst severity per facet)' 'Heading2'))
    $tc = @($Readiness).Count
    $colW = [int][Math]::Floor((14400 - 3400) / [Math]::Max(1, $tc)); if ($colW -gt 2600) { $colW = 2600 }
    $hm = @($script:Categories | ForEach-Object { $cat = $_; , (@($cat) + @($Readiness | ForEach-Object { $x = $_.Categories[$cat]; if ($x.Worst -eq 'n/a') { 'n/a' } else { "$($x.Worst)  ($($x.Findings)/$($x.Checks))" } })) })
    [void]$b.Append((New-DocxTable (@('Facet') + @($Readiness | ForEach-Object { "$($_.Target) ($($_.Role))" })) $hm (@(3400) + @(1..$tc | ForEach-Object { $colW })) { param($r, $c, $t) if ($c -gt 0) { $w = ($t -split ' ')[0]; if ($script:SevFill.ContainsKey($w)) { return $script:SevFill[$w] } } return '' }))
    [void]$b.Append((New-DocxPara 'Cell text: worst severity (issues / checks). Details per target follow.' 'Subtle'))
    foreach ($r in $Readiness) {
        $f = @($Findings | Where-Object { $_.Target -eq $r.Target })
        [void]$b.Append((New-DocxPara ("{0}  -  {1}  ({2}/100)" -f $r.Target, $r.Verdict, $r.Score) 'Heading1' -PageBreakBefore))
        [void]$b.Append((New-DocxPara ("Role: {0}    Windows host: {1}    Checks run: {2}    Findings: {3} blocker, {4} high, {5} medium, {6} low, {7} info" -f $r.Role, $r.Computer, $r.Checks, $r.Blocker, $r.High, $r.Medium, $r.Low, $r.Info)))
        [void]$b.Append((New-DocxPara 'Facet status' 'Heading2'))
        $cr = @($r.Categories.Keys | ForEach-Object { , @($_, $r.Categories[$_].Worst, "$($r.Categories[$_].Findings)", "$($r.Categories[$_].Checks)") })
        [void]$b.Append((New-DocxTable @('Facet', 'Worst', 'Issues', 'Checks') $cr @(4200, 1600, 1200, 1200) { param($ri, $c, $t) if ($c -eq 1 -and $script:SevFill.ContainsKey($t)) { return $script:SevFill[$t] } return '' }))
        $inst = @($f | Where-Object { $_.Install } | Sort-Object { - $script:SevOrder[$_.Severity] }, Category, Item)
        if ($inst.Count) {
            [void]$b.Append((New-DocxPara 'Needs installation' 'Heading2'))
            [void]$b.Append((New-DocxTable @('Sev', 'Category', 'What', 'On reference', 'Action') @($inst | ForEach-Object { , @($_.Severity, $_.Category, $_.Item, $_.Reference, $_.Action) }) @(900, 1900, 3700, 2600, 5300) $sevFillFn))
        }
        $top = @($f | Where-Object { $_.Severity -in @('Blocker', 'High') -and -not $_.Install } | Sort-Object { - $script:SevOrder[$_.Severity] }, Category, Item)
        [void]$b.Append((New-DocxPara 'Blockers and high-risk items' 'Heading2'))
        if ($top.Count) { [void]$b.Append((New-DocxTable @('Sev', 'Category', 'Item', 'Reference', 'Target', 'Action') @($top | ForEach-Object { , @($_.Severity, $_.Category, $_.Item, $_.Reference, $_.TargetValue, ($_.Action + $(if ($_.Fix) { "  [$($_.Fix)]" } else { '' }))) }) @(900, 1700, 3000, 2000, 1800, 5000) $sevFillFn)) }
        else { [void]$b.Append((New-DocxPara 'None.')) }
        $mid = @($f | Where-Object { $_.Severity -in @('Medium', 'Low') -and -not $_.Install } | Sort-Object { - $script:SevOrder[$_.Severity] }, Category, Item)
        if ($mid.Count) {
            [void]$b.Append((New-DocxPara 'Medium and low items' 'Heading2'))
            [void]$b.Append((New-DocxTable @('Sev', 'Category', 'Item', 'Reference / Target', 'Action') @($mid | ForEach-Object { , @($_.Severity, $_.Category, $_.Item, ("$($_.Reference) / $($_.TargetValue)"), $_.Action) }) @(900, 1700, 3200, 2600, 6000) $sevFillFn))
        }
        $info = @($f | Where-Object { $_.Severity -eq 'Info' })
        if ($info.Count) {
            [void]$b.Append((New-DocxPara 'Informational' 'Heading2'))
            [void]$b.Append((New-DocxTable @('Category', 'Item', 'Note') @($info | Sort-Object Category, Item | ForEach-Object { , @($_.Category, $_.Item, $_.Action) }) @(2200, 4500, 7700)))
        }
    }
    [void]$b.Append((New-DocxPara 'Checklist coverage' 'Heading1' -PageBreakBefore))
    [void]$b.Append((New-DocxPara 'How each section of the migration checklist (A-K) is covered: Automated = verified by this run; Partial = critical items verified, the rest manual; Manual = owner/runbook task.'))
    [void]$b.Append((New-DocxTable @('Ref', 'Area', 'Coverage', 'Checked by this tool', 'Still manual') @(Get-CoverageRows | ForEach-Object { , @($_.Ref, $_.Area, $_.Status, $_.Checked, $_.Manual) }) @(700, 2300, 1200, 6200, 4000) { param($r, $c, $t) if ($c -eq 2) { switch ($t) { 'Automated' { return $script:SevFill['OK'] } 'Partial' { return $script:SevFill['Medium'] } default { return $script:SevFill['Info'] } } } return '' }))
    [void]$b.Append((New-DocxPara 'About this report' 'Heading1' -PageBreakBefore))
    [void]$b.Append((New-DocxPara "Produced by $script:ToolName $script:ToolVersion. Collection is read-only: static SELECTs against catalog views/DMVs/msdb and WMI/CIM + registry reads. Nothing was installed, changed, started or stopped. Logins, SIDs and server permissions are covered by LoginSync.ps1. Secrets (linked server passwords, credential secrets, SMTP passwords) cannot be read and must be re-entered on the target."))

    $sect = '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>'
    $doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' + $b.ToString() + $sect + '</w:body></w:document>'
    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="19"/><w:szCs w:val="19"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtle"><w:name w:val="Subtle"/><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="666666"/><w:sz w:val="17"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3A5F"/><w:sz w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2E5A88"/><w:sz w:val="23"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="ReportTable"><w:name w:val="Report Table"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFC7D1"/></w:tblBorders></w:tblPr></w:style>
</w:styles>
'@
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>'
    $drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>SQL Server Migration / HA Readiness - ' + (ConvertTo-XmlText $Meta.Reference) + '</dc:title><dc:creator>' + $script:ToolName + '</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created></cp:coreProperties>'

    try { Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue; Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue } catch { }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $enc = New-Object System.Text.UTF8Encoding($false)
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($e in @(@('[Content_Types].xml', $ct), @('_rels/.rels', $rels), @('word/document.xml', $doc), @('word/_rels/document.xml.rels', $drels), @('word/styles.xml', $styles.Trim()), @('docProps/core.xml', $core))) {
            $entry = $zip.CreateEntry($e[0])
            $st = $entry.Open(); $bytes = $enc.GetBytes($e[1]); $st.Write($bytes, 0, $bytes.Length); $st.Dispose()
        }
    } finally { $zip.Dispose() }
}
#endregion

#region ---------------------------------------------------------------- HTML report (single self-contained file)
$script:HtmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Readiness</title>
<style>
:root{--bg:#f5f6f8;--panel:#fff;--ink:#1b2430;--muted:#5d6b7c;--line:#dfe3e9;--accent:#1f4e79;--chip:#eef1f5;
--blocker:#b42318;--blocker-bg:#fde4e1;--high:#c2570c;--high-bg:#feead9;--medium:#a16207;--medium-bg:#fdf3c7;--low:#2563a8;--low-bg:#e2edf9;--info:#5d6b7c;--info-bg:#eceff3;--ok:#1d7a3e;--ok-bg:#dcf2e3;--na:#9aa5b1;--na-bg:#f3f4f6;
--mono:ui-monospace,SFMono-Regular,Consolas,"Cascadia Mono",Menlo,monospace;--sans:"Segoe UI",system-ui,-apple-system,Roboto,Helvetica,Arial,sans-serif}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#11161d;--panel:#18202a;--ink:#e6ebf1;--muted:#97a3b3;--line:#2a3542;--accent:#7fb2e5;--chip:#222c38;
--blocker:#ff8a80;--blocker-bg:#4a1d1a;--high:#ffb27a;--high-bg:#4a2c14;--medium:#f2cf66;--medium-bg:#433812;--low:#8cc1f5;--low-bg:#1b3550;--info:#a8b3c1;--info-bg:#28323e;--ok:#7fd69b;--ok-bg:#163a24;--na:#6c7785;--na-bg:#1f2731}}
:root[data-theme="dark"]{--bg:#11161d;--panel:#18202a;--ink:#e6ebf1;--muted:#97a3b3;--line:#2a3542;--accent:#7fb2e5;--chip:#222c38;
--blocker:#ff8a80;--blocker-bg:#4a1d1a;--high:#ffb27a;--high-bg:#4a2c14;--medium:#f2cf66;--medium-bg:#433812;--low:#8cc1f5;--low-bg:#1b3550;--info:#a8b3c1;--info-bg:#28323e;--ok:#7fd69b;--ok-bg:#163a24;--na:#6c7785;--na-bg:#1f2731}
*{box-sizing:border-box}html,body{margin:0}body{background:var(--bg);color:var(--ink);font:14px/1.45 var(--sans)}
header{padding:22px 28px 14px;border-bottom:1px solid var(--line);background:var(--panel)}
h1{margin:0;font-size:21px;letter-spacing:-.01em}h2{font-size:15px;margin:26px 0 10px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.meta{color:var(--muted);font-size:12.5px;margin-top:4px;display:flex;flex-wrap:wrap;gap:4px 18px}.meta b{color:var(--ink);font-weight:600}
.demo{background:var(--medium-bg);color:var(--medium);font-weight:700;padding:1px 8px;border-radius:4px;font-size:11px;letter-spacing:.05em}
main{padding:6px 28px 40px;max-width:1600px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:16px;display:grid;grid-template-columns:86px 1fr;gap:14px;align-items:center;cursor:pointer;border-left:5px solid var(--na)}
.card:hover{border-color:var(--accent)}.card .name{font-weight:700;font-size:15.5px;word-break:break-all}.card .role{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.verdict{display:inline-block;margin:6px 0 8px;font-weight:700;font-size:12px;padding:2px 9px;border-radius:999px;letter-spacing:.04em}
.counts{display:flex;flex-wrap:wrap;gap:5px}.cnt{font:600 11.5px var(--mono);padding:1px 7px;border-radius:4px}
.v-NOT-READY{border-left-color:var(--blocker)}.v-NOT-READY .verdict{background:var(--blocker-bg);color:var(--blocker)}
.v-AT-RISK{border-left-color:var(--high)}.v-AT-RISK .verdict{background:var(--high-bg);color:var(--high)}
.v-READY-WITH-NOTES{border-left-color:var(--medium)}.v-READY-WITH-NOTES .verdict{background:var(--medium-bg);color:var(--medium)}
.v-READY{border-left-color:var(--ok)}.v-READY .verdict{background:var(--ok-bg);color:var(--ok)}
.v-NOT-VERIFIED .verdict{background:var(--info-bg);color:var(--info)}
.s-Blocker{background:var(--blocker-bg);color:var(--blocker)}.s-High{background:var(--high-bg);color:var(--high)}.s-Medium{background:var(--medium-bg);color:var(--medium)}.s-Low{background:var(--low-bg);color:var(--low)}.s-Info{background:var(--info-bg);color:var(--info)}.s-OK{background:var(--ok-bg);color:var(--ok)}.s-na{background:var(--na-bg);color:var(--na)}
.heat{border-collapse:separate;border-spacing:3px;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:8px;width:100%;overflow-x:auto;display:block}
.heat th{font-size:12px;color:var(--muted);font-weight:600;text-align:left;padding:4px 8px;white-space:nowrap}.heat td.cat{font-weight:600;font-size:13px;padding:4px 10px 4px 4px;white-space:nowrap}
.heat td.cell{min-width:120px;text-align:center;border-radius:5px;font:600 12px var(--mono);padding:7px 6px;cursor:pointer}.heat td.cell:hover{outline:2px solid var(--accent)}
.heat td.cell small{display:block;font:400 10.5px var(--sans);opacity:.8}
.tabs{display:flex;gap:2px;border-bottom:1px solid var(--line);margin-top:28px;flex-wrap:wrap}.tab{padding:9px 16px;cursor:pointer;border:1px solid transparent;border-bottom:none;border-radius:8px 8px 0 0;color:var(--muted);font-weight:600;background:none;font:inherit;font-weight:600}
.tab[aria-selected="true"]{background:var(--panel);border-color:var(--line);color:var(--ink);margin-bottom:-1px}
.pane{background:var(--panel);border:1px solid var(--line);border-top:none;border-radius:0 0 10px 10px;padding:16px}.pane[hidden]{display:none}
.filters{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:12px}
select,input[type=search]{font:inherit;padding:6px 9px;border:1px solid var(--line);border-radius:6px;background:var(--bg);color:var(--ink)}input[type=search]{min-width:240px}
.chip{border:1px solid var(--line);background:var(--chip);color:var(--ink);border-radius:999px;padding:3px 11px;cursor:pointer;font:600 12px var(--sans)}.chip[aria-pressed="false"]{opacity:.42}
table.grid{width:100%;border-collapse:collapse;font-size:13px}table.grid th{position:sticky;top:0;background:var(--panel);text-align:left;font-size:11.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);border-bottom:2px solid var(--line);padding:7px 8px}
table.grid td{border-bottom:1px solid var(--line);padding:7px 8px;vertical-align:top}table.grid tr.f:hover td{background:var(--chip)}
.sev{display:inline-block;font:700 11px var(--sans);padding:2px 7px;border-radius:4px;min-width:62px;text-align:center}
.mono{font-family:var(--mono);font-size:12px;word-break:break-word}.act{color:var(--ink)}.fix{display:inline-block;margin-top:4px;font:600 11px var(--mono);background:var(--chip);padding:1px 6px;border-radius:4px;color:var(--accent)}
.detail{color:var(--muted);font-size:12px;margin-top:3px}.tgt{font-weight:600;white-space:nowrap}.cat-col{white-space:nowrap;color:var(--muted)}
.empty{padding:28px;text-align:center;color:var(--muted)}
.inst h3{margin:18px 0 8px;font-size:15px}.inst ul{list-style:none;margin:0;padding:0}.inst li{display:grid;grid-template-columns:26px 74px 1fr;gap:8px;padding:8px 4px;border-bottom:1px solid var(--line);align-items:start}
.inst li.done{opacity:.45;text-decoration:line-through}.inst input{margin-top:3px;width:16px;height:16px}
.sbs td.diff{background:var(--medium-bg)}.sbs td.miss{background:var(--blocker-bg);color:var(--blocker)}.sbs td.refcol{background:var(--chip)}.sbs th.refcol{color:var(--accent)}
.legend{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}.note{color:var(--muted);font-size:12.5px}
.count-badge{font:600 11px var(--mono);color:var(--muted);margin-left:6px}
.themebtn{float:right;border:1px solid var(--line);background:var(--chip);color:var(--ink);border-radius:6px;padding:4px 10px;cursor:pointer;font:inherit;font-size:12px}
.wrap{overflow-x:auto}
@media (max-width:700px){header,main{padding-left:16px;padding-right:16px}.card{grid-template-columns:70px 1fr}input[type=search]{min-width:0;width:100%}}
@media print{.tabs,.filters,.themebtn{display:none}.pane[hidden]{display:block}.pane{border:none;padding:0}body{background:#fff}.card{break-inside:avoid}}
</style>
</head>
<body>
<header>
<button class="themebtn" id="themeBtn" type="button">Theme</button>
<h1>SQL Server Migration / HA Readiness</h1>
<div class="meta" id="meta"></div>
<div class="legend" id="legend"></div>
</header>
<main>
<h2>Targets</h2>
<div class="cards" id="cards"></div>
<h2>Facet heat map <span class="note">(worst severity per facet - click a cell to filter)</span></h2>
<div class="wrap"><table class="heat" id="heat"></table></div>
<div class="tabs" role="tablist" id="tabs">
<button class="tab" role="tab" data-pane="pFind" aria-selected="true">Findings<span class="count-badge" id="cFind"></span></button>
<button class="tab" role="tab" data-pane="pInst" aria-selected="false">Needs installation<span class="count-badge" id="cInst"></span></button>
<button class="tab" role="tab" data-pane="pSbs" aria-selected="false">Side-by-side inventory</button>
<button class="tab" role="tab" data-pane="pAbout" aria-selected="false">Coverage &amp; notes</button>
</div>
<section class="pane" id="pFind">
<div class="filters">
<select id="fTarget"></select><select id="fCat"></select>
<span id="fSev"></span>
<input type="search" id="fText" placeholder="Search item, value, action...">
</div>
<div class="wrap"><table class="grid" id="tFind"><thead><tr><th>Severity</th><th>Target</th><th>Facet</th><th>Item</th><th>Reference</th><th>Target value</th><th>Action</th></tr></thead><tbody></tbody></table></div>
</section>
<section class="pane inst" id="pInst" hidden><p class="note">Everything the reference has that the target is missing and that needs an installer, a file copy or a Windows feature. Tick items as you go (stored only in this browser).</p><div id="instBody"></div></section>
<section class="pane" id="pSbs" hidden>
<div class="filters"><select id="sFacet"></select><label class="note"><input type="checkbox" id="sDiff" checked> differences only</label><input type="search" id="sText" placeholder="Filter items..."></div>
<div class="wrap"><table class="grid sbs" id="tSbs"></table></div>
</section>
<section class="pane" id="pAbout" hidden><div id="about"></div></section>
</main>
<script>
const D = __DATA__;
const SEV = ["Blocker","High","Medium","Low","Info"], ORDER = {Blocker:5,High:4,Medium:3,Low:2,Info:1,OK:0,"n/a":-1};
const $ = s => document.querySelector(s), esc = s => String(s==null?"":s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const store = {get(k){try{return localStorage.getItem(k)}catch(e){return null}}, set(k,v){try{localStorage.setItem(k,v)}catch(e){}}};
(function theme(){const t=store.get("mr-theme"); if(t) document.documentElement.dataset.theme=t;
 $("#themeBtn").onclick=()=>{const cur=document.documentElement.dataset.theme||(matchMedia("(prefers-color-scheme: dark)").matches?"dark":"light"); const n=cur==="dark"?"light":"dark"; document.documentElement.dataset.theme=n; store.set("mr-theme",n);};})();
const M = D.meta;
$("#meta").innerHTML = `<span>Reference <b>${esc(M.Reference)}</b> (${esc(M.ReferenceVersion)})</span><span>Generated <b>${esc(M.Generated)}</b></span><span>Run by <b>${esc(M.RunBy)}</b> on ${esc(M.RunHost)}</span>${M.PathMap?`<span>Path map <b class="mono">${esc(M.PathMap)}</b></span>`:""}${M.Demo?'<span class="demo">DEMO DATA</span>':""}<span>${esc(M.Tool)} ${esc(M.Version)} &middot; read-only</span>`;
$("#legend").innerHTML = SEV.map(s=>`<span class="sev s-${s}">${s}</span>`).join("") + `<span class="note" style="align-self:center">Blocker = move/failover cannot succeed &middot; High = breaks after cut-over &middot; Medium = fix or accept &middot; Low/Info = review</span>`;
// cards
function ring(score){const r=30,c=2*Math.PI*r,p=Math.max(0,Math.min(100,score))/100;const col=score<40?"var(--blocker)":score<70?"var(--high)":score<90?"var(--medium)":"var(--ok)";
 return `<svg viewBox="0 0 80 80" width="80" height="80" role="img" aria-label="score ${score}"><circle cx="40" cy="40" r="${r}" fill="none" stroke="var(--line)" stroke-width="8"/><circle cx="40" cy="40" r="${r}" fill="none" stroke="${col}" stroke-width="8" stroke-linecap="round" stroke-dasharray="${c*p} ${c}" transform="rotate(-90 40 40)"/><text x="40" y="45" text-anchor="middle" font-size="18" font-weight="700" fill="var(--ink)">${score}</text></svg>`;}
$("#cards").innerHTML = D.readiness.map(r=>`<div class="card v-${r.Verdict.replace(/ /g,"-")}" data-t="${esc(r.Target)}" tabindex="0">${ring(r.Score)}<div><div class="name">${esc(r.Target)}</div><div class="role">${esc(r.Role)}${r.Computer&&r.Computer!==r.Target?" &middot; host "+esc(r.Computer):""}</div><div class="verdict">${esc(r.Verdict)}</div><div class="counts">${SEV.slice(0,4).map(s=>r[s]?`<span class="cnt s-${s}">${r[s]} ${s}</span>`:"").join("")}${r.Install?`<span class="cnt s-Info">${r.Install} to install</span>`:""}<span class="cnt s-na">${r.Checks} checks</span></div></div></div>`).join("");
document.querySelectorAll(".card").forEach(c=>c.onclick=()=>{$("#fTarget").value=c.dataset.t;$("#fCat").value="";showTab("pFind");renderFind();$("#tabs").scrollIntoView({behavior:"smooth"});});
// heat
const cats = D.categories;
$("#heat").innerHTML = `<tr><th></th>${D.readiness.map(r=>`<th>${esc(r.Target)}<br><span class="note">${esc(r.Role)}</span></th>`).join("")}</tr>` + cats.map(c=>`<tr><td class="cat">${esc(c)}</td>${D.readiness.map(r=>{const x=r.Categories[c]||{Worst:"n/a",Findings:0,Checks:0};const cls=x.Worst==="n/a"?"na":x.Worst;return `<td class="cell s-${cls}" data-t="${esc(r.Target)}" data-c="${esc(c)}">${x.Worst==="n/a"?"&ndash;":esc(x.Worst)}<small>${x.Findings} issue${x.Findings==1?"":"s"} / ${x.Checks} checks</small></td>`}).join("")}</tr>`).join("");
document.querySelectorAll("#heat td.cell").forEach(td=>td.onclick=()=>{$("#fTarget").value=td.dataset.t;$("#fCat").value=td.dataset.c;sevOn.Info=true;drawSev();showTab("pFind");renderFind();$("#tabs").scrollIntoView({behavior:"smooth"});});
// tabs
function showTab(id){document.querySelectorAll(".tab").forEach(t=>t.setAttribute("aria-selected",t.dataset.pane===id));document.querySelectorAll(".pane").forEach(p=>p.hidden=p.id!==id);}
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>showTab(t.dataset.pane));
// findings
$("#fTarget").innerHTML = `<option value="">All targets</option>` + D.readiness.map(r=>`<option>${esc(r.Target)}</option>`).join("");
$("#fCat").innerHTML = `<option value="">All facets</option>` + cats.map(c=>`<option>${esc(c)}</option>`).join("");
const sevOn = {Blocker:true,High:true,Medium:true,Low:true,Info:false};
function drawSev(){$("#fSev").innerHTML = SEV.map(s=>`<button class="chip s-${s}" type="button" aria-pressed="${sevOn[s]}" data-s="${s}">${s} ${D.findings.filter(f=>f.Severity===s).length}</button>`).join(" ");document.querySelectorAll("#fSev .chip").forEach(b=>b.onclick=()=>{sevOn[b.dataset.s]=!sevOn[b.dataset.s];drawSev();renderFind();});}
drawSev();
function renderFind(){const t=$("#fTarget").value,c=$("#fCat").value,q=$("#fText").value.toLowerCase();
 const rows=D.findings.filter(f=>sevOn[f.Severity]&&(!t||f.Target===t)&&(!c||f.Category===c)&&(!q||[f.Item,f.Reference,f.TargetValue,f.Action,f.Detail,f.Fix].join(" ").toLowerCase().includes(q))).sort((a,b)=>ORDER[b.Severity]-ORDER[a.Severity]||a.Target.localeCompare(b.Target)||a.Category.localeCompare(b.Category)||a.Item.localeCompare(b.Item));
 $("#cFind").textContent=rows.length;
 $("#tFind tbody").innerHTML = rows.length? rows.map(f=>`<tr class="f"><td><span class="sev s-${f.Severity}">${f.Severity}</span>${f.Install?'<div class="detail">install</div>':""}</td><td class="tgt">${esc(f.Target)}</td><td class="cat-col">${esc(f.Category)}</td><td>${esc(f.Item)}${f.Detail?`<div class="detail">${esc(f.Detail)}</div>`:""}</td><td class="mono">${esc(f.Reference)}</td><td class="mono">${esc(f.TargetValue)}</td><td class="act">${esc(f.Action)}${f.Fix?`<div><span class="fix">${esc(f.Fix)}</span></div>`:""}</td></tr>`).join("") : `<tr><td colspan="7" class="empty">Nothing matches the filters.</td></tr>`;}
["#fTarget","#fCat"].forEach(s=>$(s).onchange=renderFind); $("#fText").oninput=renderFind; renderFind();
// install
(function(){const inst=D.findings.filter(f=>f.Install); $("#cInst").textContent=inst.length;
 if(!inst.length){$("#instBody").innerHTML='<div class="empty">Nothing to install - every component found on the reference is present on the targets.</div>';return;}
 $("#instBody").innerHTML = D.readiness.map(r=>{const l=inst.filter(f=>f.Target===r.Target).sort((a,b)=>ORDER[b.Severity]-ORDER[a.Severity]||a.Category.localeCompare(b.Category));
  return `<h3>${esc(r.Target)} <span class="note">${l.length} item(s)</span></h3>`+(l.length?`<ul>${l.map((f,i)=>{const k="mr-i|"+M.Generated+"|"+f.Target+"|"+f.Item;const d=store.get(k)==="1";return `<li class="${d?"done":""}"><input type="checkbox" data-k="${esc(k)}" ${d?"checked":""} aria-label="done"><span class="sev s-${f.Severity}">${f.Severity}</span><div><b>${esc(f.Item)}</b><div class="detail">${esc(f.Category)}${f.Reference?" &middot; reference: "+esc(f.Reference):""}</div><div>${esc(f.Action)}</div></div></li>`}).join("")}</ul>`:'<p class="note">Nothing to install.</p>');}).join("");
 document.querySelectorAll("#instBody input").forEach(cb=>cb.onchange=()=>{store.set(cb.dataset.k,cb.checked?"1":"0");cb.closest("li").classList.toggle("done",cb.checked);});})();
// side-by-side
const servers = D.servers, facets=[...new Set(D.inventory.map(r=>r.Facet))];
$("#sFacet").innerHTML = facets.map(f=>`<option>${esc(f)}</option>`).join("");
function renderSbs(){const fa=$("#sFacet").value,only=$("#sDiff").checked,q=$("#sText").value.toLowerCase();const ref=servers[0];
 let rows=D.inventory.filter(r=>r.Facet===fa&&(!q||r.Item.toLowerCase().includes(q)));
 if(only) rows=rows.filter(r=>servers.some(s=>(r.Values[s]||"")!==(r.Values[ref]||"")));
 $("#tSbs").innerHTML = `<thead><tr><th>Item</th>${servers.map((s,i)=>`<th class="${i?"":"refcol"}">${esc(s)}${i?"":" (reference)"}</th>`).join("")}</tr></thead><tbody>`+(rows.length?rows.map(r=>`<tr><td>${esc(r.Item)}</td>${servers.map((s,i)=>{const v=r.Values[s]||"",rv=r.Values[ref]||"";const cls=i===0?"refcol":(v===rv?"":(!v&&rv?"miss":"diff"));return `<td class="mono ${cls}">${v?esc(v):(rv&&i?"&mdash; missing":"&ndash;")}</td>`}).join("")}</tr>`).join(""):`<tr><td colspan="${servers.length+1}" class="empty">${only?"No differences in this facet.":"No items."}</td></tr>`)+"</tbody>";}
$("#sFacet").onchange=renderSbs;$("#sDiff").onchange=renderSbs;$("#sText").oninput=renderSbs;renderSbs();
// about
$("#about").innerHTML = `<h3>Coverage per server</h3><div class="wrap"><table class="grid"><thead><tr><th>Server</th><th>Role</th><th>Windows host</th><th>SQL</th><th>Windows</th><th>Sections not collected</th></tr></thead><tbody>${D.coverage.map(c=>`<tr><td class="tgt">${esc(c.Server)}</td><td>${esc(c.Role)}</td><td>${esc(c.Computer)}</td><td>${c.SqlOk?'<span class="sev s-OK">ok</span>':'<span class="sev s-Blocker">no</span>'}</td><td>${c.OsOk?'<span class="sev s-OK">ok</span>':'<span class="sev s-Blocker">no</span>'}</td><td class="mono">${c.Gaps.length?c.Gaps.map(esc).join("<br>"):"none"}</td></tr>`).join("")}</tbody></table></div>
<h3>Checklist coverage (team checklist A-K)</h3><p class="note">What this tool verifies automatically for each section of the migration checklist, and what stays a manual / owner task.</p><div class="wrap"><table class="grid"><thead><tr><th>Ref</th><th>Area</th><th>Coverage</th><th>Checked by this tool</th><th>Still manual</th></tr></thead><tbody>${(D.checklist||[]).map(c=>`<tr><td class="tgt">${esc(c.Ref)}</td><td>${esc(c.Area)}</td><td><span class="sev ${c.Status==='Automated'?'s-OK':c.Status==='Partial'?'s-Medium':'s-Info'}">${esc(c.Status)}</span></td><td>${esc(c.Checked)}</td><td class="note">${esc(c.Manual)}</td></tr>`).join("")}</tbody></table></div>
<h3>How to read this</h3><p class="note">Each target is compared with the reference. Severity is weighted by role: drift on an <b>asynchronous</b> DR replica is rated one level lower than on a synchronous partner or migration target, except for hard blockers (version, TDE certificate, FILESTREAM, space, HADR endpoint). Logins/SIDs are out of scope here - run LoginSync.ps1. Secrets (linked-server passwords, credentials, SMTP passwords) are never read.</p>
<p class="note">Collection was read-only: static SELECTs on catalog views/DMVs/msdb, CIM/WMI queries and StdRegProv read methods. Nothing was installed, changed, started or stopped.</p>`;
</script>
</body>
</html>
'@

function Write-HtmlReport {
    param($Path, $Payload)
    $json = $Payload | ConvertTo-Json -Depth 12 -Compress
    $json = $json.Replace('</', '<\/').Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
    $html = $script:HtmlTemplate.Replace('__DATA__', $json)
    [System.IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($false)))
}
#endregion

#region ---------------------------------------------------------------- checklist coverage (maps the team's "Exhaustive checklist" A-K)
$script:ChecklistCoverage = @(
    @('A1', 'OS identity & compatibility', 'Automated', 'Name/domain, edition/build/UBR, Core vs Desktop, architecture, OS support lifecycle, last update date, pending reboot, locale, time zone, FIPS', 'Secure Boot/TPM/VBS, clock drift vs domain, legacy compatibility modes'),
    @('A2', 'Windows roles & features', 'Automated', 'Every installed Windows feature compared (IIS, .NET 3.5, Failover Clustering, MPIO, MSMQ, WAS...); severity raised only where SQL needs it', 'Feature payload availability, features needed by apps not on the reference'),
    @('A3', 'OS configuration & policy', 'Partial', 'Windows Firewall profiles + SQL/HADR/listener/Browser rules, SCHANNEL TLS, Defender real-time + SQL exclusions, machine environment variables + PATH, scheduled tasks (non-Microsoft), local Administrators, power plan, page file, certificates in LocalMachine\My', 'GPO/secpol user rights (log on as a service, LPIM, volume maintenance are checked via SQL), UAC, SMB signing, Credential Guard, WER, event-log sizes, service recovery actions'),
    @('A4', 'Host compute & virtualization', 'Partial', 'Physical vs virtual, model, sockets/cores/logical CPUs, RAM, NUMA nodes (SQL), max memory vs RAM, MAXDOP vs CPUs', 'Hypervisor reservations/limits, CPU ready/steal, snapshots - check in vCenter/Hyper-V'),
    @('A5', 'Disks, volumes & file systems', 'Automated', 'Every volume (letter/mount, label, size, free, 64 KB AU), disk online/read-only/health, free space per volume for the databases (+25% headroom), data/log/tempdb/system file placement, NTFS compressed/EFS-encrypted SQL folders', 'IOPS/latency (see H2), BitLocker, dedup, NTFS ACLs on SQL folders, cloud temp-disk persistence'),
    @('A6', 'Windows services & identities', 'Partial', 'SQL family services + start modes, third-party services, service accounts per SQL service, SQL service account in local Administrators, Kerberos SPNs (LDAP read)', 'Service dependencies/recovery actions, gMSA/password management, group memberships, share/cert-key ACLs'),
    @('B1', 'SQL Server inventory', 'Automated', 'Instance name/ID, version/CU, edition, support lifecycle, components installed (services + Full-Text/PolyBase/SSIS/SSRS/SSAS), other instances on the host, default paths, startup parameters, aliases, edition features in use', 'Setup bootstrap logs, failed patches'),
    @('B2', 'Engine configuration', 'Automated', 'Every sp_configure option (severity by actual use), pending (configured <> running) values, MAXDOP vs CPUs, max memory vs RAM, startup trace flags (obsolete ones recognised), Resource Governor, error-log retention', 'Database-scoped configurations (travel with the database), session trace flags'),
    @('B3', 'SQL networking & connectivity', 'Automated', 'TCP/NP enabled, static/dynamic port, Browser, force encryption, TLS certificate (present, expiry, SAN vs host/FQDN/listener), endpoints, AG listener port, SPNs, firewall rules, DNS forward/reverse, client aliases', 'Network ACLs/security groups between subnets, app-side connection strings, MultiSubnetFailover'),
    @('B4', 'SQL security configuration', 'Partial', 'Auth mode, credentials, proxies, TDE/backup/other certificates in master, master DMK, per-DB DMK encrypted by SMK, EKM, audits + specifications, server/logon triggers, TRUSTWORTHY, DB chaining, contained DBs', 'Logins, SIDs, server roles/permissions, orphaned users -> LoginSync.ps1; Always Encrypted client keys'),
    @('B5', 'SQL Agent & automation', 'Automated', 'Jobs, step definitions (signature), schedules, owners, proxies + subsystems + principals, categories, operators, alerts, fail-safe operator, Agent mail profile, maintenance plans, Database Mail profiles/accounts/settings, files/exes/scripts/modules/UNC shares used by steps, hard-coded server names/IPs, jobs already failing', 'Token usage review, overlap/maintenance-window planning, third-party schedulers'),
    @('B6', 'Instance-level objects', 'Automated', 'Linked servers (+options, logins, provider + provider options), endpoints, server triggers, credentials, replication distributor/publications, log shipping, mirroring, AGs, Resource Governor, XE sessions, audits, PBM policies, CMS registrations, custom messages, startup procs, extended procs, backup devices', 'Inbound linked-server connectivity tests, FCI internals'),
    @('C1', 'SQL Server registry', 'Partial', 'Instance registry root discovered dynamically, startup params, network/TCP, cert binding, Agent settings, provider options, aliases (x64/x86), other instances', 'Registry ACLs, orphaned keys from old installs'),
    @('C2', 'Windows registry & app config', 'Partial', 'Uninstall keys (x64/x86), ODBC drivers + system DSNs, OLE DB provider CLSIDs (x64/x86), MSDTC security, FIPS, TLS, .NET release keys', 'App-specific vendor keys, COM+/COM registrations, user-scope DSNs'),
    @('C3', 'SQL & application binaries', 'Partial', 'Existence + file version of every exe/script/package/DLL referenced by job steps, extended procs and EKM providers; PowerShell modules (all-users)', 'Hashes/signatures, native DLL dependencies, AppLocker/WDAC'),
    @('C4', 'Drivers, runtimes & frameworks', 'Automated', '.NET 4.x/3.5 + strong crypto, VC++ runtimes, Java/Python/.NET Core, ODBC/OLE DB drivers + bitness, Oracle/DB2/ACE/SAP providers, SQLNCLI removal on 2022+', 'JDBC drivers inside applications'),
    @('C5', 'Installed products & patches', 'Automated', 'Uninstall inventory compared (missing / older version), classified as drivers, runtimes, SQL components, agents, tooling; Windows update level', 'Licensing/activation tied to hostname, installation media availability'),
    @('C6', 'SQL CLR & external code', 'Automated', 'User assemblies + permission sets, clr enabled/strict security, trusted assemblies, TRUSTWORTHY, extended procs + DLLs, OLE automation and xp_cmdshell usage in code', 'Native dependencies of CLR assemblies'),
    @('D1', 'Database inventory & state', 'Partial', 'State, size, files/paths, compat level vs target minimum, collation, containment, recovery model, owner, encryption, CDC, broker, replication, AG membership, snapshots excluded', 'Page verify, auto-close/shrink, Query Store settings (travel with the backup)'),
    @('D2', 'Objects & schema', 'Partial', 'Full-text catalogs, FILESTREAM, In-Memory filegroups, Service Broker queues, assemblies, edition features in use, cross-database + linked-server references, hard-coded server names', 'Object-level schema review -> DMA / schema compare'),
    @('D3', 'SQL behaviour & compatibility', 'Partial', 'Compat level jump, collation differences, sp_configure user options/two digit year/default language, deprecated SQLNCLI', 'Deprecated T-SQL, CE regressions -> DMA + Query Store A/B testing'),
    @('D4', 'Data integrity & size', 'Automated', 'Last full/log backup, last good CHECKDB, suspect pages, sizes vs target free space, backup-encryption certificates needed on target', 'Row-count reconciliation, transfer/restore time estimates'),
    @('E', 'Security, identity & encryption', 'Partial', 'Certificates/keys (TDE, backup, DMK/SMK), credentials, proxies, audits, SPNs, service accounts, local admins, TLS; secrets are never read', 'Logins/SIDs/permissions -> LoginSync.ps1; delegation settings, share ACLs, Credential Manager, vault access'),
    @('F1', 'Network & name resolution', 'Partial', 'DNS forward + reverse from the runner, ports + firewall, listener, aliases, UNC shares in jobs listed', 'Routing/VLAN/ACL/load balancer, app-host resolution, DNS TTL'),
    @('F2', 'Application & client compatibility', 'Manual', 'Client-side items surfaced indirectly: TLS/driver requirements, SQLNCLI removal, instance name/port changes', 'Connection strings, app drivers, retry/failover settings - application owners'),
    @('F3', 'External integrations', 'Partial', 'Linked servers + providers, SSIS projects/packages, SSRS key step, Database Mail SMTP, UNC/file paths, scheduled tasks', 'APIs, SFTP endpoints, message queues, Power BI gateways'),
    @('G1', 'HA / DR', 'Automated', 'AG enabled, replicas, modes, endpoints (port/encryption/state), listener + SPN, replica health/lag (from the primary), cluster membership, identical file paths, async severity weighting', 'Quorum/witness, cluster networks, RPO/RTO measurements, failover tests'),
    @('G2', 'Backup & restore readiness', 'Partial', 'Backup recency, log backups for FULL recovery, backup-encryption certs, backup devices, backup jobs/agents/UNC paths', 'Restore tests, retention/immutability, throughput'),
    @('G3', 'Monitoring & operations', 'Partial', 'Operators, alerts, Database Mail, Agent alert system, XE sessions, monitoring agents/services, maintenance jobs/plans, error-log retention', 'Thresholds, central logging, runbooks, ownership'),
    @('H', 'Performance, capacity & sizing', 'Manual', 'CPU/RAM/disk capacity compared statically (configuration-based)', 'Workload baselines (CPU, IO latency, waits, Query Store) over a business cycle - use your monitoring tool / sql-dba MCP'),
    @('I', 'Feature & product dependencies', 'Automated', 'For each feature: installed on reference, actually used (code scan, objects, jobs), present on target - severity follows "used"; tested-on-target stays manual', 'Functional testing'),
    @('J', 'Migration method, cutover & rollback', 'Manual', 'Prep T-SQL per target, install list, readiness verdict as the go/no-go input', 'Method choice, downtime window, cutover runbook, rollback plan'),
    @('K', 'Easy-to-overlook dependencies', 'Partial', 'Job scripts/modules/exes/UNC paths, scheduled tasks, env vars/PATH, DSNs/providers (32/64-bit), aliases, certificates, hard-coded names/IPs in jobs and code, local admins, AV exclusions, third-party agents', 'IIS apps on the host, mapped-drive scripts, GPO-deployed software, vendor IP allowlists')
)
function Get-CoverageRows { return @($script:ChecklistCoverage | ForEach-Object { [pscustomobject]@{ Ref = $_[0]; Area = $_[1]; Status = $_[2]; Checked = $_[3]; Manual = $_[4] } }) }
#endregion

#region ---------------------------------------------------------------- run pipeline (shared by headless and GUI)
function Invoke-Readiness {
    # $Specs: array of @{ Name; Role }  - first one is the reference
    param($Specs, [switch]$UseDemo)
    $script:Findings = New-List
    $script:Checks = @{}
    if ($UseDemo) { $snaps = Get-DemoSnapshots }
    else {
        $snaps = @($Specs | ForEach-Object { New-Snapshot -Name $_.Name -Role $_.Role })
        Invoke-Collection -Snapshots $snaps
    }
    $R = $snaps[0]
    if (-not $UseDemo -and -not $R.SqlOk -and -not $R.OsOk) { throw "Reference $($R.Name) could not be reached (SQL and Windows). Nothing to compare against." }
    foreach ($T in ($snaps | Select-Object -Skip 1)) {
        Write-Log "Comparing $($T.Name) [$($T.Role)] with $($R.Name)..."
        try { Compare-Target -R $R -T $T }
        catch {
            Write-Log ("  comparison error on $($T.Name): " + $_.Exception.Message) 'WARN'
            Add-Finding $T 'Collection gaps' 'Comparison engine error' '' ($_.Exception.Message) 'High' -Hard -Action ('Part of the comparison did not complete (' + ($_.InvocationInfo.PositionMessage -replace '\s+', ' ') + '). Results for this target are incomplete - please report it.')
        }
    }
    $ready = @(Get-Readiness -Snapshots $snaps)
    return @{ Snapshots = @($snaps); Readiness = $ready; Findings = @($script:Findings.ToArray()); Demo = [bool]$UseDemo }
}

function Export-Results {
    param($Result, [string]$OutDir, [string[]]$Formats)
    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
    $snaps = $Result.Snapshots; $R = $snaps[0]
    $meta = [ordered]@{
        Tool = $script:ToolName; Version = $script:ToolVersion; Generated = (Get-Date -Format 'yyyy-MM-dd HH:mm'); RunBy = $(if ($env:USERDOMAIN) { "$env:USERDOMAIN\" } else { '' }) + [Environment]::UserName; RunHost = [Environment]::MachineName
        Reference = $R.Name; ReferenceVersion = ("SQL {0} {1} {2}" -f (Get-Prop $R.Sql.Instance 'VersionName' '?'), (Get-Prop $R.Sql.Instance 'ProductLevel' ''), (Get-Prop $R.Sql.Instance 'ProductUpdateLevel' '')).Trim()
        PathMap = ($PathMap -join '; '); Demo = $Result.Demo; Targets = @($snaps | Select-Object -Skip 1 | ForEach-Object { "$($_.Name) [$($_.Role)]" })
    }
    $coverage = @($snaps | ForEach-Object {
            $s = $_
            $g = @(); foreach ($a in @('Sql', 'Os')) { foreach ($k in $s[$a].Sections.Keys) { if ([string]$s[$a].Sections[$k] -notlike 'ok*') { $g += "$a/$k : $($s[$a].Sections[$k])" } } }
            [ordered]@{ Server = $s.Name; Role = $s.Role; Computer = $s.Computer; SqlOk = $s.SqlOk; OsOk = $s.OsOk; Gaps = $g }
        })
    $inv = @(Get-Inventory -Snapshots $snaps)
    $files = New-List
    $base = Join-Path $OutDir $script:ToolName
    if ($Formats -contains 'Html') {
        $payload = [ordered]@{ meta = $meta; readiness = $Result.Readiness; findings = $Result.Findings; categories = $script:Categories; servers = @($snaps | ForEach-Object { $_.Name }); inventory = $inv; coverage = $coverage; checklist = @(Get-CoverageRows) }
        Write-HtmlReport -Path "$base.html" -Payload $payload; Add-To $files "$base.html"
    }
    if ($Formats -contains 'Docx') { Write-DocxReport -Path "$base.docx" -Meta $meta -Readiness $Result.Readiness -Findings $Result.Findings; Add-To $files "$base.docx" }
    if ($Formats -contains 'Txt') { Write-TxtReport -Path "$base.txt" -Meta $meta -Readiness $Result.Readiness -Findings $Result.Findings; Add-To $files "$base.txt" }
    if ($Formats -contains 'Csv') {
        $Result.Findings | Select-Object Target, Role, Severity, BaseSeverity, Category, Item, Reference, TargetValue, Action, Fix, Install, Detail | Export-Csv -Path (Join-Path $OutDir 'findings.csv') -NoTypeInformation -Encoding UTF8
        $Result.Readiness | Select-Object Target, Role, Computer, Verdict, Score, Blocker, High, Medium, Low, Info, Install, Checks | Export-Csv -Path (Join-Path $OutDir 'readiness.csv') -NoTypeInformation -Encoding UTF8
        @($Result.Findings | Where-Object { $_.Install }) | Select-Object Target, Severity, Category, Item, Reference, Action | Export-Csv -Path (Join-Path $OutDir 'install_list.csv') -NoTypeInformation -Encoding UTF8
        $inv | ForEach-Object { $o = [ordered]@{ Facet = $_.Facet; Item = $_.Item }; foreach ($k in $_.Values.Keys) { $o[$k] = $_.Values[$k] }; [pscustomobject]$o } | Export-Csv -Path (Join-Path $OutDir 'inventory.csv') -NoTypeInformation -Encoding UTF8
        foreach ($n in @('findings.csv', 'readiness.csv', 'install_list.csv', 'inventory.csv')) { Add-To $files (Join-Path $OutDir $n) }
    }
    if ($Formats -contains 'Json') {
        $j = [ordered]@{ meta = $meta; readiness = $Result.Readiness; findings = $Result.Findings; coverage = $coverage; snapshots = (Get-SnapshotExport $snaps) }
        [System.IO.File]::WriteAllText((Join-Path $OutDir 'result.json'), ($j | ConvertTo-Json -Depth 12), (New-Object System.Text.UTF8Encoding($false)))
        Add-To $files (Join-Path $OutDir 'result.json')
    }
    if ($Formats -contains 'Sql') {
        foreach ($T in ($snaps | Select-Object -Skip 1)) {
            if (-not $T.SqlOk) { continue }
            $p = Join-Path $OutDir ('prep_' + (Get-SafeName $T.Name) + '.sql')
            Write-PrepSql -Path $p -R $R -T $T -Findings @($Result.Findings | Where-Object { $_.Target -eq $T.Name })
            Add-To $files $p
        }
    }
    return $files.ToArray()
}

function Get-ExitCode {
    param($Result)
    if ($FailOnSeverity -eq 'Never') { return 0 }
    $th = $script:SevOrder[$FailOnSeverity]
    if (@($Result.Findings | Where-Object { $script:SevOrder[$_.Severity] -ge $th }).Count -gt 0) { return 2 }
    return 0
}
#endregion

#region ---------------------------------------------------------------- GUI (WinForms; loaded only in GUI mode)
function Get-AutomationText {
    param([string]$Ref, $Targets, [string]$Out)
    $script = $PSCommandPath; if (-not $script) { $script = '.\MigrationReadiness.ps1' }
    $mig = @($Targets | Where-Object { $_.Role -eq 'Migration' } | ForEach-Object { $_.Name })
    $syn = @($Targets | Where-Object { $_.Role -eq 'Synchronous' } | ForEach-Object { $_.Name })
    $asy = @($Targets | Where-Object { $_.Role -eq 'Asynchronous' } | ForEach-Object { $_.Name })
    $args1 = "-ReferenceServer $Ref"
    if ($mig.Count) { $args1 += " -TargetServer $($mig -join ',') -TargetMode Migration" }
    elseif ($syn.Count) { $args1 += " -TargetServer $($syn -join ',') -TargetMode Synchronous" }
    if ($asy.Count) { $args1 += " -AsyncTarget $($asy -join ',')" }
    if ($PathMap.Count) { $args1 += " -PathMap " + (($PathMap | ForEach-Object { "'$_'" }) -join ',') }
    $outRoot = 'D:\DBA\Reports\MigrationReadiness'
    return @"
# 1) One-off, interactive (same engine as this window):
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$script" $args1 -OutputPath "$Out" -Open

# 2) Scheduled re-check before cut-over (Task Scheduler, run as a gMSA with VIEW SERVER STATE / VIEW ANY DEFINITION
#    on the SQL instances and WMI read rights on the hosts):
`$action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "$script" $args1 -OutputPath "$outRoot\run" -FailOnSeverity Blocker'
`$trigger = New-ScheduledTaskTrigger -Daily -At 06:00
`$princ   = New-ScheduledTaskPrincipal -UserId 'CORP\gmsa_dbatools$' -LogonType Password
Register-ScheduledTask -TaskName 'MigrationReadiness $Ref' -Action `$action -Trigger `$trigger -Principal `$princ

# 3) SQL Agent job step (type: Operating system (CmdExec)) on a DBA utility server - exit code 2 fails the step,
#    so the job's failure notification alerts you:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$script" $args1 -OutputPath "$outRoot\agent" -Format Html,Csv,Json -FailOnSeverity High

# 4) After fixing items, re-run and diff findings.csv between runs; keep LoginSync.ps1 for logins.
# Exit codes: 0 = nothing at/above -FailOnSeverity, 2 = findings at/above it, 1 = run failed.
"@
}

function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $font = New-Object System.Drawing.Font('Segoe UI', 9)
    $mono = New-Object System.Drawing.Font('Consolas', 9)
    $sevColor = @{ 'Blocker' = [System.Drawing.Color]::FromArgb(244, 199, 195); 'High' = [System.Drawing.Color]::FromArgb(251, 217, 190); 'Medium' = [System.Drawing.Color]::FromArgb(252, 239, 180); 'Low' = [System.Drawing.Color]::FromArgb(214, 230, 245); 'Info' = [System.Drawing.Color]::FromArgb(231, 231, 231); 'OK' = [System.Drawing.Color]::FromArgb(207, 235, 211); 'n/a' = [System.Drawing.Color]::FromArgb(245, 245, 245) }
    $verdictColor = @{ 'NOT READY' = $sevColor['Blocker']; 'AT RISK' = $sevColor['High']; 'READY WITH NOTES' = $sevColor['Medium']; 'READY' = $sevColor['OK']; 'NOT VERIFIED' = $sevColor['Info'] }
    $state = @{ Result = $null; Files = @(); OutDir = $null }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$script:ToolName $script:ToolVersion - migration / HA readiness (read-only)"
    $form.Size = New-Object System.Drawing.Size(1280, 860); $form.StartPosition = 'CenterScreen'; $form.Font = $font
    $form.MinimumSize = New-Object System.Drawing.Size(980, 640)

    # ---------- top: servers + options
    $top = New-Object System.Windows.Forms.TableLayoutPanel
    $top.Dock = 'Top'; $top.Height = 230; $top.ColumnCount = 3; $top.Padding = New-Object System.Windows.Forms.Padding(8)
    [void]$top.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 42)))
    [void]$top.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 38)))
    [void]$top.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 20)))

    $gServers = New-Object System.Windows.Forms.GroupBox; $gServers.Text = 'Servers  (first row = reference: current server / AG primary)'; $gServers.Dock = 'Fill'
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'; $grid.AllowUserToAddRows = $true; $grid.RowHeadersVisible = $false; $grid.AutoSizeColumnsMode = 'Fill'; $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
    $cName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $cName.HeaderText = 'SQL Server (name, name\inst, name,port, listener)'; $cName.FillWeight = 65
    $cRole = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $cRole.HeaderText = 'Role'; $cRole.FillWeight = 35
    [void]$cRole.Items.AddRange(@('Reference', 'Migration', 'Synchronous', 'Asynchronous'))
    [void]$grid.Columns.Add($cName); [void]$grid.Columns.Add($cRole)
    $grid.Add_DefaultValuesNeeded({ param($s, $e) $e.Row.Cells[1].Value = $(if ($grid.Rows.Count -le 1) { 'Reference' } else { 'Migration' }) })
    $gServers.Controls.Add($grid)
    $top.Controls.Add($gServers, 0, 0)

    $gOpt = New-Object System.Windows.Forms.GroupBox; $gOpt.Text = 'Options'; $gOpt.Dock = 'Fill'
    $opt = New-Object System.Windows.Forms.FlowLayoutPanel; $opt.Dock = 'Fill'; $opt.FlowDirection = 'TopDown'; $opt.WrapContents = $false; $opt.AutoScroll = $true
    $lblMap = New-Object System.Windows.Forms.Label; $lblMap.Text = 'Path map old=new (comma separated), e.g. D:\=E:\, L:\=F:\'; $lblMap.AutoSize = $true
    $txtMap = New-Object System.Windows.Forms.TextBox; $txtMap.Width = 420
    $lblCm = New-Object System.Windows.Forms.Label; $lblCm.Text = 'Windows host overrides SQLNAME=HOST (comma separated, optional)'; $lblCm.AutoSize = $true
    $txtCm = New-Object System.Windows.Forms.TextBox; $txtCm.Width = 420
    $chkSw = New-Object System.Windows.Forms.CheckBox; $chkSw.Text = 'Installed software (registry; slower over WAN)'; $chkSw.Checked = $true; $chkSw.AutoSize = $true
    $chkPaths = New-Object System.Windows.Forms.CheckBox; $chkPaths.Text = 'Check referenced files/folders on targets'; $chkPaths.Checked = $true; $chkPaths.AutoSize = $true
    $chkCode = New-Object System.Windows.Forms.CheckBox; $chkCode.Text = 'Scan module code (xp_cmdshell, OPENROWSET, sp_OA*, dbmail)'; $chkCode.Checked = $true; $chkCode.AutoSize = $true
    $chkEnc = New-Object System.Windows.Forms.CheckBox; $chkEnc.Text = 'Encrypt SQL connections + trust server certificate'; $chkEnc.AutoSize = $true
    $chkDcom = New-Object System.Windows.Forms.CheckBox; $chkDcom.Text = 'Use DCOM instead of WinRM for Windows facets'; $chkDcom.AutoSize = $true
    $btnCred = New-Object System.Windows.Forms.Button; $btnCred.Text = 'SQL login... (default: Windows auth)'; $btnCred.AutoSize = $true
    $opt.Controls.AddRange(@($lblMap, $txtMap, $lblCm, $txtCm, $chkSw, $chkPaths, $chkCode, $chkEnc, $chkDcom, $btnCred))
    $gOpt.Controls.Add($opt)
    $top.Controls.Add($gOpt, 1, 0)

    $gAct = New-Object System.Windows.Forms.GroupBox; $gAct.Text = 'Run'; $gAct.Dock = 'Fill'
    $act = New-Object System.Windows.Forms.FlowLayoutPanel; $act.Dock = 'Fill'; $act.FlowDirection = 'TopDown'; $act.WrapContents = $false
    $mk = { param($t) $b = New-Object System.Windows.Forms.Button; $b.Text = $t; $b.Width = 200; $b.Height = 28; $b }
    $btnRun = & $mk 'Run readiness check'; $btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $btnAg = & $mk 'Discover AG from server...'
    $btnDemo = & $mk 'Load demo data'
    $btnHtml = & $mk 'Open HTML report'; $btnDocx = & $mk 'Open Word report'; $btnTxt = & $mk 'Open text report (Notepad)'; $btnDir = & $mk 'Open output folder'
    foreach ($b in @($btnHtml, $btnDocx, $btnTxt, $btnDir)) { $b.Enabled = $false }
    $act.Controls.AddRange(@($btnRun, $btnAg, $btnDemo, $btnHtml, $btnDocx, $btnTxt, $btnDir))
    $gAct.Controls.Add($act)
    $top.Controls.Add($gAct, 2, 0)

    # ---------- bottom: tabs
    $tabs = New-Object System.Windows.Forms.TabControl; $tabs.Dock = 'Fill'
    $newTab = { param($t) $p = New-Object System.Windows.Forms.TabPage; $p.Text = $t; [void]$tabs.TabPages.Add($p); $p }
    $newGrid = { $g = New-Object System.Windows.Forms.DataGridView; $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.RowHeadersVisible = $false; $g.SelectionMode = 'FullRowSelect'; $g.AutoSizeColumnsMode = 'Fill'; $g.BackgroundColor = [System.Drawing.SystemColors]::Window; $g.DefaultCellStyle.WrapMode = 'True'; $g.AutoSizeRowsMode = 'AllCells'; $g }
    $tSum = & $newTab 'Summary'; $gSum = & $newGrid; $tSum.Controls.Add($gSum)
    $tMat = & $newTab 'Facet matrix'; $gMat = & $newGrid; $tMat.Controls.Add($gMat)
    $tFind = & $newTab 'Findings'
    $fp = New-Object System.Windows.Forms.FlowLayoutPanel; $fp.Dock = 'Top'; $fp.Height = 34
    $cbT = New-Object System.Windows.Forms.ComboBox; $cbT.DropDownStyle = 'DropDownList'; $cbT.Width = 200
    $cbC = New-Object System.Windows.Forms.ComboBox; $cbC.DropDownStyle = 'DropDownList'; $cbC.Width = 220
    $cbS = New-Object System.Windows.Forms.ComboBox; $cbS.DropDownStyle = 'DropDownList'; $cbS.Width = 150; [void]$cbS.Items.AddRange(@('Low and above', 'Medium and above', 'High and above', 'Blocker only', 'Everything incl. Info')); $cbS.SelectedIndex = 0
    $txtF = New-Object System.Windows.Forms.TextBox; $txtF.Width = 260
    $lblF = New-Object System.Windows.Forms.Label; $lblF.Text = 'search:'; $lblF.AutoSize = $true; $lblF.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $fp.Controls.AddRange(@($cbT, $cbC, $cbS, $lblF, $txtF))
    $gFind = & $newGrid; $tFind.Controls.Add($gFind); $tFind.Controls.Add($fp)
    $tInst = & $newTab 'Needs installation'; $gInst = & $newGrid; $tInst.Controls.Add($gInst)
    $tSql = & $newTab 'Prep T-SQL'
    $sp = New-Object System.Windows.Forms.FlowLayoutPanel; $sp.Dock = 'Top'; $sp.Height = 34
    $cbSql = New-Object System.Windows.Forms.ComboBox; $cbSql.DropDownStyle = 'DropDownList'; $cbSql.Width = 260
    $btnCopy = New-Object System.Windows.Forms.Button; $btnCopy.Text = 'Copy to clipboard'; $btnCopy.AutoSize = $true
    $sp.Controls.AddRange(@($cbSql, $btnCopy))
    $txtSql = New-Object System.Windows.Forms.TextBox; $txtSql.Multiline = $true; $txtSql.ScrollBars = 'Both'; $txtSql.WordWrap = $false; $txtSql.Dock = 'Fill'; $txtSql.Font = $mono; $txtSql.ReadOnly = $true
    $tSql.Controls.Add($txtSql); $tSql.Controls.Add($sp)
    $tAuto = & $newTab 'Automate'
    $txtAuto = New-Object System.Windows.Forms.TextBox; $txtAuto.Multiline = $true; $txtAuto.ScrollBars = 'Both'; $txtAuto.WordWrap = $false; $txtAuto.Dock = 'Fill'; $txtAuto.Font = $mono; $txtAuto.ReadOnly = $true
    $tAuto.Controls.Add($txtAuto)
    $tLog = & $newTab 'Log'
    $txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Multiline = $true; $txtLog.ScrollBars = 'Vertical'; $txtLog.Dock = 'Fill'; $txtLog.Font = $mono; $txtLog.ReadOnly = $true
    $tLog.Controls.Add($txtLog)

    $status = New-Object System.Windows.Forms.StatusStrip
    $stLbl = New-Object System.Windows.Forms.ToolStripStatusLabel; $stLbl.Text = 'Read-only. Enter servers (first row = reference) or load demo data.'; $stLbl.Spring = $true; $stLbl.TextAlign = 'MiddleLeft'
    [void]$status.Items.Add($stLbl)

    $form.Controls.Add($tabs); $form.Controls.Add($top); $form.Controls.Add($status)

    $script:LogSink = { param($line) $txtLog.AppendText($line + "`r`n"); $stLbl.Text = $line; [System.Windows.Forms.Application]::DoEvents() }

    # ---------- rendering
    $fillFind = {
        $res = $state.Result; if (-not $res) { return }
        $min = switch ($cbS.SelectedIndex) { 0 { 2 } 1 { 3 } 2 { 4 } 3 { 5 } default { 1 } }
        $t = [string]$cbT.SelectedItem; $c = [string]$cbC.SelectedItem; $q = $txtF.Text.ToLower()
        $dt = New-Object System.Data.DataTable
        foreach ($col in @('Severity', 'Target', 'Facet', 'Item', 'Reference', 'Target value', 'Action', 'Tool')) { [void]$dt.Columns.Add($col) }
        foreach ($f in ($res.Findings | Sort-Object @{ e = { $script:SevOrder[$_.Severity] }; Descending = $true }, Target, Category, Item)) {
            if ($script:SevOrder[$f.Severity] -lt $min) { continue }
            if ($t -and $t -ne '(all targets)' -and $f.Target -ne $t) { continue }
            if ($c -and $c -ne '(all facets)' -and $f.Category -ne $c) { continue }
            if ($q -and ("$($f.Item) $($f.Reference) $($f.TargetValue) $($f.Action) $($f.Detail)").ToLower().IndexOf($q) -lt 0) { continue }
            [void]$dt.Rows.Add($f.Severity, $f.Target, $f.Category, $f.Item, $f.Reference, $f.TargetValue, $f.Action, $f.Fix)
        }
        $gFind.DataSource = $dt
        if ($gFind.Columns.Count) { $gFind.Columns[0].FillWeight = 40; $gFind.Columns[1].FillWeight = 55; $gFind.Columns[2].FillWeight = 60; $gFind.Columns[3].FillWeight = 110; $gFind.Columns[4].FillWeight = 70; $gFind.Columns[5].FillWeight = 60; $gFind.Columns[6].FillWeight = 190; $gFind.Columns[7].FillWeight = 55 }
    }
    $gFind.Add_CellFormatting({ param($s, $e) if ($e.ColumnIndex -eq 0 -and $sevColor.ContainsKey([string]$e.Value)) { $e.CellStyle.BackColor = $sevColor[[string]$e.Value] } })
    $gSum.Add_CellFormatting({ param($s, $e) if ($e.ColumnIndex -eq 2 -and $verdictColor.ContainsKey([string]$e.Value)) { $e.CellStyle.BackColor = $verdictColor[[string]$e.Value] } })
    $gMat.Add_CellFormatting({ param($s, $e) if ($e.ColumnIndex -gt 0) { $w = ([string]$e.Value -split ' ')[0]; if ($sevColor.ContainsKey($w)) { $e.CellStyle.BackColor = $sevColor[$w] } } })
    $gInst.Add_CellFormatting({ param($s, $e) if ($e.ColumnIndex -eq 1 -and $sevColor.ContainsKey([string]$e.Value)) { $e.CellStyle.BackColor = $sevColor[[string]$e.Value] } })
    $gMat.Add_CellDoubleClick({ param($s, $e)
            if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 1) { return }
            $cbT.SelectedItem = $gMat.Columns[$e.ColumnIndex].HeaderText; $cbC.SelectedItem = [string]$gMat.Rows[$e.RowIndex].Cells[0].Value; $cbS.SelectedIndex = 4
            $tabs.SelectedTab = $tFind; & $fillFind })
    $cbT.Add_SelectedIndexChanged($fillFind); $cbC.Add_SelectedIndexChanged($fillFind); $cbS.Add_SelectedIndexChanged($fillFind); $txtF.Add_TextChanged($fillFind)

    $render = {
        $res = $state.Result
        $dt = New-Object System.Data.DataTable
        foreach ($col in @('Target', 'Role', 'Verdict', 'Score', 'Blocker', 'High', 'Medium', 'Low', 'Info', 'To install', 'Checks', 'Windows host')) { [void]$dt.Columns.Add($col) }
        foreach ($r in $res.Readiness) { [void]$dt.Rows.Add($r.Target, $r.Role, $r.Verdict, $r.Score, $r.Blocker, $r.High, $r.Medium, $r.Low, $r.Info, $r.Install, $r.Checks, $r.Computer) }
        $gSum.DataSource = $dt
        $dm = New-Object System.Data.DataTable; [void]$dm.Columns.Add('Facet')
        foreach ($r in $res.Readiness) { [void]$dm.Columns.Add($r.Target) }
        foreach ($cat in $script:Categories) { $row = @($cat); foreach ($r in $res.Readiness) { $x = $r.Categories[$cat]; $row += ("{0}  ({1} issues / {2} checks)" -f $x.Worst, $x.Findings, $x.Checks) }; [void]$dm.Rows.Add($row) }
        $gMat.DataSource = $dm
        $di = New-Object System.Data.DataTable
        foreach ($col in @('Target', 'Severity', 'Facet', 'What', 'On reference', 'Action')) { [void]$di.Columns.Add($col) }
        foreach ($f in ($res.Findings | Where-Object { $_.Install } | Sort-Object Target, @{ e = { $script:SevOrder[$_.Severity] }; Descending = $true }, Category)) { [void]$di.Rows.Add($f.Target, $f.Severity, $f.Category, $f.Item, $f.Reference, $f.Action) }
        $gInst.DataSource = $di
        $cbT.Items.Clear(); [void]$cbT.Items.Add('(all targets)'); foreach ($r in $res.Readiness) { [void]$cbT.Items.Add($r.Target) }; $cbT.SelectedIndex = 0
        $cbC.Items.Clear(); [void]$cbC.Items.Add('(all facets)'); foreach ($c in $script:Categories) { [void]$cbC.Items.Add($c) }; $cbC.SelectedIndex = 0
        $cbSql.Items.Clear(); foreach ($f in ($state.Files | Where-Object { $_ -like '*.sql' })) { [void]$cbSql.Items.Add((Split-Path -Leaf $f)) }; if ($cbSql.Items.Count) { $cbSql.SelectedIndex = 0 } else { $txtSql.Text = '' }
        $R = $res.Snapshots[0]
        $txtAuto.Text = (Get-AutomationText -Ref $R.Name -Targets @($res.Snapshots | Select-Object -Skip 1) -Out $state.OutDir).Replace("`n", "`r`n")
        & $fillFind
        foreach ($b in @($btnHtml, $btnDocx, $btnTxt, $btnDir)) { $b.Enabled = $true }
        $tabs.SelectedTab = $tSum
    }
    $cbSql.Add_SelectedIndexChanged({ $p = Join-Path $state.OutDir ([string]$cbSql.SelectedItem); if (Test-Path $p) { $txtSql.Text = [System.IO.File]::ReadAllText($p).Replace("`r`n", "`n").Replace("`n", "`r`n") } })
    $btnCopy.Add_Click({ if ($txtSql.Text) { [System.Windows.Forms.Clipboard]::SetText($txtSql.Text); $stLbl.Text = 'T-SQL copied. Review it before running anything.' } })

    $runIt = {
        param([switch]$UseDemo)
        $specs = New-List
        if (-not $UseDemo) {
            foreach ($row in $grid.Rows) {
                if ($row.IsNewRow) { continue }
                $n = ([string]$row.Cells[0].Value).Trim(); if (-not $n) { continue }
                $role = [string]$row.Cells[1].Value; if (-not $role) { $role = 'Migration' }
                Add-To $specs @{ Name = $n; Role = $role }
            }
            if ($specs.Count -lt 2) { [void][System.Windows.Forms.MessageBox]::Show('Enter at least a reference and one target.', $script:ToolName); return }
            $specs[0].Role = 'Reference'
            foreach ($x in ($specs | Select-Object -Skip 1)) { if ($x.Role -eq 'Reference') { $x.Role = 'Migration' } }
        }
        $script:PathMap = @($txtMap.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '=' })
        $script:ComputerMap = @($txtCm.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '=' })
        $sk = @(); if (-not $chkSw.Checked) { $sk += 'Software' }; if (-not $chkPaths.Checked) { $sk += 'Paths' }; if (-not $chkCode.Checked) { $sk += 'CodeScan' }
        $script:Skip = $sk
        $script:Encrypt = [switch]$chkEnc.Checked; $script:TrustServerCertificate = [switch]$chkEnc.Checked
        $script:CimProtocol = $(if ($chkDcom.Checked) { 'Dcom' } else { 'Wsman' })
        $state.OutDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ('MigrationReadiness_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $btnRun.Enabled = $false; $btnDemo.Enabled = $false; $tabs.SelectedTab = $tLog; $txtLog.Clear()
        try {
            $res = Invoke-Readiness -Specs $specs.ToArray() -UseDemo:$UseDemo
            $state.Result = $res
            Write-Log "Writing reports to $($state.OutDir)..."
            $state.Files = Export-Results -Result $res -OutDir $state.OutDir -Formats @('Html', 'Docx', 'Txt', 'Csv', 'Json', 'Sql')
            & $render
            $nr = @($res.Readiness | Where-Object { $_.Verdict -eq 'NOT READY' }).Count
            $stLbl.Text = "Done. $(@($res.Readiness).Count) target(s), $nr not ready. Reports in $($state.OutDir)"
        } catch {
            Write-Log ("FAILED: " + $_.Exception.Message) 'ERROR'
            [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "$script:ToolName - run failed", 'OK', 'Error')
        } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $btnRun.Enabled = $true; $btnDemo.Enabled = $true }
    }
    $btnRun.Add_Click({ & $runIt })
    $btnDemo.Add_Click({ & $runIt -UseDemo })
    $btnCred.Add_Click({
            $c = $null; try { $c = Get-Credential -Message 'SQL Server login used for every server (cancel = Windows authentication)' } catch { }
            $script:SqlCredential = $c; $btnCred.Text = $(if ($c) { "SQL login: $($c.UserName)" } else { 'SQL login... (default: Windows auth)' })
        })
    $btnAg.Add_Click({
            $srv = [Microsoft.VisualBasic.Interaction]::InputBox('Any replica or listener name:', 'Discover availability group', '')
            if (-not $srv) { return }
            try {
                $topo = Get-AgTopology -Server $srv
                $grid.Rows.Clear()
                [void]$grid.Rows.Add($topo.Primary, 'Reference')
                foreach ($r in $topo.Replicas) { [void]$grid.Rows.Add($r.Name, $r.Mode) }
                $stLbl.Text = "AG $($topo.AgName): primary $($topo.Primary), $(@($topo.Replicas).Count) secondary replica(s)."
            } catch { [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'AG discovery failed', 'OK', 'Error') }
        })
    $open = { param($ext) $f = $state.Files | Where-Object { $_ -like "*.$ext" } | Select-Object -First 1; if ($f) { if ($ext -eq 'txt') { Start-Process notepad.exe -ArgumentList "`"$f`"" } else { Invoke-Item $f } } }
    $btnHtml.Add_Click({ & $open 'html' }); $btnDocx.Add_Click({ & $open 'docx' }); $btnTxt.Add_Click({ & $open 'txt' }); $btnDir.Add_Click({ if ($state.OutDir) { Invoke-Item $state.OutDir } })
    Add-Type -AssemblyName Microsoft.VisualBasic
    [void]$grid.Rows.Add('', 'Reference')
    [void]$form.ShowDialog()
    $script:LogSink = $null
}
#endregion

#region ---------------------------------------------------------------- entry point
if ($PSCmdlet.ParameterSetName -eq 'Gui') {
    if ([Environment]::OSVersion.Platform -ne 'Win32NT') { Write-Host 'The GUI needs Windows. Use -ReferenceServer/-TargetServer or -Demo for headless mode.'; exit 1 }
    Show-Gui
    exit 0
}

try {
    $specs = New-List
    switch ($PSCmdlet.ParameterSetName) {
        'Manual' {
            Add-To $specs @{ Name = $ReferenceServer; Role = 'Reference' }
            foreach ($t in $TargetServer) { if ($t) { Add-To $specs @{ Name = $t; Role = $TargetMode } } }
            foreach ($t in $AsyncTarget) { if ($t) { Add-To $specs @{ Name = $t; Role = 'Asynchronous' } } }
            if ($specs.Count -lt 2) { throw 'Give at least one -TargetServer or -AsyncTarget.' }
        }
        'Ag' {
            $topo = Get-AgTopology -Server $DiscoverAgFrom -Ag $AvailabilityGroup
            Write-Log "AG $($topo.AgName): primary $($topo.Primary) is the reference."
            Add-To $specs @{ Name = $topo.Primary; Role = 'Reference' }
            foreach ($r in $topo.Replicas) { if ($SyncReplicasOnly -and $r.Mode -ne 'Synchronous') { continue }; Add-To $specs @{ Name = $r.Name; Role = $r.Mode } }
            if ($specs.Count -lt 2) { throw 'No secondary replicas to compare.' }
        }
    }
    $res = Invoke-Readiness -Specs $specs.ToArray() -UseDemo:$Demo
    Write-Log "Writing reports to $OutputPath ..."
    $files = Export-Results -Result $res -OutDir $OutputPath -Formats $Format
    Write-Host ''
    foreach ($r in $res.Readiness) {
        $col = switch ($r.Verdict) { 'NOT READY' { 'Red' } 'AT RISK' { 'Yellow' } 'READY WITH NOTES' { 'Cyan' } 'READY' { 'Green' } default { 'Gray' } }
        Write-Host ('  {0,-24} {1,-13} {2,-17} score {3,3}  B{4} H{5} M{6} L{7}  install {8}' -f $r.Target, $r.Role, $r.Verdict, $r.Score, $r.Blocker, $r.High, $r.Medium, $r.Low, $r.Install) -ForegroundColor $col
    }
    Write-Host ''
    foreach ($f in $files) { Write-Host "  $f" }
    if ($Open) { $h = $files | Where-Object { $_ -like '*.html' } | Select-Object -First 1; if ($h) { Invoke-Item $h } }
    exit (Get-ExitCode -Result $res)
} catch {
    Write-Log ("Run failed: " + $_.Exception.Message) 'ERROR'
    if ($_.InvocationInfo) { Write-Log ($_.InvocationInfo.PositionMessage -replace '\s+', ' ') 'ERROR' }
    exit 1
}
#endregion

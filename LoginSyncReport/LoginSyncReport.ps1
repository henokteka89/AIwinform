<#
.SYNOPSIS
    LoginSync - read-only SQL Server login synchronization check for Always On AG replicas or any set of
    servers (e.g. migration: current vs newly provisioned). One file, two modes:
      * no parameters (or -Gui)  -> WinForms window: enter/discover servers, visualise drift, review fixes
      * with parameters          -> headless: interactive HTML report + CSV/JSON + remediation .sql
    Remediation T-SQL is generated for review and is NEVER executed by this tool.

.DESCRIPTION
    Compares server-level security between a REFERENCE server (AG primary / migration source) and one or
    more TARGET servers (AG secondaries / migration target):

      * Login existence (SQL, Windows user, Windows group) and custom server roles
      * SID match  -> the thing that orphans database users after an AG failover or a restore
      * Password hash (SQL logins, compared by SHA-256 digest; requires CONTROL SERVER to read)
      * Disabled state, default database / language, CHECK_POLICY / CHECK_EXPIRATION
      * Server role membership and server-level permissions (GRANT / DENY)
      * Orphan risk: every database user on the reference, checked against each server's logins by SID
      * Logins that exist only on a target (reported, never dropped)

    Severity is weighted by replica commit mode: drift on a SYNCHRONOUS (automatic-failover) partner or a
    migration target is rated one level higher than the same drift on an ASYNCHRONOUS (DR) replica.

    PRODUCTION GUARDRAILS (by design, not by option):
      * Only static SELECT statements against catalog views. Every query passes a guard that refuses
        anything other than a single SELECT/WITH statement (no DML/DDL/EXEC/GRANT/DBCC/xp_/sp_ ...).
      * Session is set to LOCK_TIMEOUT (default 3000 ms), DEADLOCK_PRIORITY LOW and READ UNCOMMITTED:
        if a catalog is locked the query gives up quickly instead of queuing behind (or in front of) work.
      * One connection at a time, connection pooling off, closed immediately. Application Name is
        'LoginSyncCheck (read-only)' so the session is easy to identify in sys.dm_exec_sessions.
      * No user-table access, no DMVs that scan the plan cache, no temp tables, no writes anywhere.
      * Database-user inventory runs only on the reference, only on ONLINE databases you can access,
        capped by -MaxDatabases; a locked or unreadable database is skipped, not waited on.
      * Remediation is written to .sql files only. Each file refuses to run on the wrong server
        (@@SERVERNAME check + SET NOEXEC ON) and destructive steps are commented out.

    Permissions: VIEW ANY DEFINITION is enough for everything except password hashes, which need
    CONTROL SERVER (sysadmin) on the reference/targets. VIEW SERVER STATE is used for AG discovery.

.PARAMETER ReferenceServer
    The source of truth (AG primary, or the current server in a migration).

.PARAMETER TargetServer
    One or more servers to compare against the reference. Mode for these comes from -TargetMode.

.PARAMETER AsyncTarget
    Targets that are asynchronous-commit replicas (lower severity).

.PARAMETER TargetMode
    Mode label for -TargetServer entries: Synchronous, Asynchronous or Standalone (migration / peer).

.PARAMETER DiscoverAgFrom
    Any replica (or listener) of the AG. The primary becomes the reference, all other replicas targets,
    with their real commit mode.

.PARAMETER AvailabilityGroup
    AG name when the instance hosts more than one AG.

.PARAMETER SyncReplicasOnly
    With -DiscoverAgFrom: compare only synchronous-commit replicas.

.PARAMETER Demo
    Generate a report from built-in sample data (4-replica AG). No server is contacted.

.PARAMETER Gui
    Open the WinForms window (also the default on Windows when no server parameters are given).

.PARAMETER FailOnSeverity
    For automation: exit code 2 when any finding is at or above this severity.

.EXAMPLE
    .\LoginSync.ps1
    Opens the GUI (same as -Gui).

.EXAMPLE
    .\LoginSync.ps1 -DiscoverAgFrom SQLAGLISTENER01
    Discover the AG, use the primary as reference and compare every replica.

.EXAMPLE
    .\LoginSync.ps1 -ReferenceServer SQLOLD01 -TargetServer SQLNEW01
    Migration: compare the newly provisioned server against the current one.

.EXAMPLE
    .\LoginSync.ps1 -ReferenceServer SQL01 -TargetServer SQL02 -TargetMode Synchronous -AsyncTarget SQL03,SQL04

.EXAMPLE
    .\LoginSync.ps1 -DiscoverAgFrom SQLAG01 -NoOpen -FailOnSeverity High -OutputPath \\share\dba\LoginSync
    Scheduled / SQL Agent use: exit code 2 when High or Critical drift is found.

.EXAMPLE
    .\LoginSync.ps1 -Demo

.NOTES
    Version 1.1.0 (single file: engine + GUI). Windows PowerShell 5.1 or PowerShell 7+. No module dependencies (System.Data.SqlClient).
    The .sql remediation files contain password hashes when CONTROL SERVER was available: treat them as
    secrets and delete after use. The HTML report and JSON never contain hashes (digests only).
#>
[CmdletBinding(DefaultParameterSetName = 'Manual')]
param(
    [Parameter(ParameterSetName = 'Manual')][string]$ReferenceServer,
    [Parameter(ParameterSetName = 'Manual')][string[]]$TargetServer,
    [Parameter(ParameterSetName = 'Manual')][string[]]$AsyncTarget,
    [Parameter(ParameterSetName = 'Manual')][ValidateSet('Synchronous', 'Asynchronous', 'Standalone')][string]$TargetMode = 'Standalone',

    [Parameter(ParameterSetName = 'AG')][string]$DiscoverAgFrom,
    [Parameter(ParameterSetName = 'AG')][string]$AvailabilityGroup,
    [Parameter(ParameterSetName = 'AG')][switch]$SyncReplicasOnly,

    [Parameter(ParameterSetName = 'Demo')][switch]$Demo,

    [pscredential]$Credential,
    [switch]$Encrypt,
    [switch]$TrustServerCertificate,
    [ValidateSet('Auto', 'AG', 'AllUser', 'None')][string]$DatabaseScope = 'Auto',
    [string[]]$ExcludeLogin = @('##*', 'NT SERVICE\*', 'NT AUTHORITY\*', 'BUILTIN\*'),
    [switch]$SkipPasswordCompare,
    [ValidateRange(500, 30000)][int]$LockTimeoutMs = 3000,
    [ValidateRange(5, 300)][int]$QueryTimeoutSec = 30,
    [ValidateRange(0, 5000)][int]$MaxDatabases = 300,
    [string]$OutputPath,
    [switch]$NoOpen,
    [switch]$PassThru,
    [ValidateSet('None', 'Low', 'Medium', 'High', 'Critical')][string]$FailOnSeverity = 'None',

    [switch]$Gui
)

#region ---------- constants ----------
$script:LSVersion = '1.1.0'
$script:LSAppName = 'LoginSyncCheck (read-only)'
$script:LSSevRank = @{ Info = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }
$script:LSDefaultExclude = @('##*', 'NT SERVICE\*', 'NT AUTHORITY\*', 'BUILTIN\*')

# Anything matching this (outside string literals and [quoted] identifiers) is refused before it reaches SQL Server.
$script:LSBlockedPattern = '(?i)\b(INSERT|UPDATE|DELETE|MERGE|DROP|ALTER|CREATE|EXEC|EXECUTE|GRANT|REVOKE|DENY|TRUNCATE|BACKUP|RESTORE|DBCC|SHUTDOWN|KILL|RECONFIGURE|WAITFOR|OPENROWSET|OPENQUERY|OPENDATASOURCE|OPENXML|BULK|INTO|USE|xp_\w*|sp_\w*)\b'

$script:LSQ_Info = @"
SELECT @@SERVERNAME AS ServerName,
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128)) AS Version,
       CAST(SERVERPROPERTY('Edition') AS nvarchar(128)) AS Edition,
       CAST(SERVERPROPERTY('MachineName') AS nvarchar(128)) AS MachineName,
       CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS int) AS IsHadrEnabled,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW ANY DEFINITION') AS CanViewDefinition,
       HAS_PERMS_BY_NAME(NULL, NULL, 'CONTROL SERVER') AS CanControlServer,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER STATE') AS CanViewServerState
"@

$script:LSQ_Principals = @"
SELECT sp.name, sp.sid, sp.type, sp.type_desc, sp.is_disabled, sp.default_database_name,
       sp.default_language_name, sp.is_fixed_role, sl.is_policy_checked, sl.is_expiration_checked,
       sl.password_hash, sp.create_date, sp.modify_date
FROM sys.server_principals AS sp
LEFT JOIN sys.sql_logins AS sl ON sl.principal_id = sp.principal_id
WHERE sp.type IN ('S','U','G','R','C','K')
"@

$script:LSQ_RoleMembers = @"
SELECT r.name AS RoleName, m.name AS MemberName
FROM sys.server_role_members AS rm
JOIN sys.server_principals AS r ON r.principal_id = rm.role_principal_id
JOIN sys.server_principals AS m ON m.principal_id = rm.member_principal_id
"@

$script:LSQ_Permissions = @"
SELECT pr.name AS Grantee, pe.class, pe.class_desc, pe.permission_name, pe.state_desc, pe.major_id,
       CASE pe.class
            WHEN 101 THEN (SELECT p2.name FROM sys.server_principals AS p2 WHERE p2.principal_id = pe.major_id)
            WHEN 105 THEN (SELECT e.name FROM sys.endpoints AS e WHERE e.endpoint_id = pe.major_id)
            ELSE NULL END AS TargetName
FROM sys.server_permissions AS pe
JOIN sys.server_principals AS pr ON pr.principal_id = pe.grantee_principal_id
WHERE pr.type IN ('S','U','G','R','C','K')
  AND NOT (pe.permission_name = 'CONNECT SQL' AND pe.state = 'G')
"@

$script:LSQ_Databases = @"
SELECT d.name, CASE WHEN d.replica_id IS NULL THEN 0 ELSE 1 END AS IsAgDb, ag.name AS AgName,
       HAS_DBACCESS(d.name) AS HasAccess
FROM sys.databases AS d
LEFT JOIN sys.availability_replicas AS ar ON ar.replica_id = d.replica_id
LEFT JOIN sys.availability_groups AS ag ON ag.group_id = ar.group_id
WHERE d.database_id > 4 AND d.state = 0
"@

# {0} = database name literal, {1} = [quoted] database name
$script:LSQ_DbUsers = @"
SELECT {0} AS DatabaseName, dp.name AS UserName, dp.type AS UserType, dp.type_desc AS TypeDesc, dp.sid
FROM {1}.sys.database_principals AS dp
WHERE dp.type IN ('S','U','G') AND dp.sid IS NOT NULL AND dp.authentication_type IN (1,3)
  AND (dp.principal_id > 4 OR dp.principal_id = 1)
"@

$script:LSQ_AgTopology = @"
SELECT ag.name AS AgName, ar.replica_server_name, ar.endpoint_url, ar.availability_mode_desc,
       ar.failover_mode_desc, ar.secondary_role_allow_connections_desc, ags.primary_replica,
       ars.role_desc, ars.connected_state_desc
FROM sys.availability_groups AS ag
JOIN sys.availability_replicas AS ar ON ar.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_group_states AS ags ON ags.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states AS ars ON ars.replica_id = ar.replica_id
"@
#endregion

#region ---------- small helpers ----------
function Write-LSLog {
    param([string]$Message, [ValidateSet('Info', 'Warn', 'Error')][string]$Level = 'Info')
    $line = '{0:HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level.ToUpper(), $Message
    $queue = $null
    try { $queue = Get-Variable -Name LSLogQueue -Scope Global -ValueOnly -ErrorAction Stop } catch { $queue = $null }
    if ($null -ne $queue) { $queue.Enqueue($line); return }
    switch ($Level) {
        'Error' { Write-Host $line -ForegroundColor Red }
        'Warn' { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line -ForegroundColor DarkGray }
    }
}

function ConvertFrom-LSDbNull { param($Value) if ($Value -is [System.DBNull]) { return $null } return $Value }

function ConvertTo-LSHex {
    param($Bytes)
    if ($null -eq $Bytes -or $Bytes -is [System.DBNull]) { return $null }
    $b = [byte[]]$Bytes
    if ($b.Length -eq 0) { return $null }
    return '0x' + ([System.BitConverter]::ToString($b) -replace '-', '')
}

function Get-LSDigest {
    param($Bytes)
    if ($null -eq $Bytes -or $Bytes -is [System.DBNull]) { return $null }
    $b = [byte[]]$Bytes
    if ($b.Length -eq 0) { return $null }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([System.BitConverter]::ToString($sha.ComputeHash($b)) -replace '-', '').Substring(0, 16) }
    finally { $sha.Dispose() }
}

function Format-LSName { param([string]$Name) return '[' + $Name.Replace(']', ']]') + ']' }
function Format-LSLiteral { param([string]$Value) if ($null -eq $Value) { return 'NULL' } return "N'" + $Value.Replace("'", "''") + "'" }

function Get-LSSeverity {
    param([string]$Base, [string]$Mode)
    if ($Mode -eq 'Asynchronous') {
        switch ($Base) { 'Critical' { return 'High' } 'High' { return 'Medium' } 'Medium' { return 'Low' } }
    }
    return $Base
}

function Test-LSExcluded {
    param([string]$Name, [string[]]$Patterns)
    foreach ($p in $Patterns) { if ($p -and $Name -like $p) { return $true } }
    return $false
}
#endregion

#region ---------- guarded, read-only data access ----------
function Assert-LSReadOnlyQuery {
    param([string]$Query)
    # Neutralise string literals and [bracketed] identifiers so a database called [Drop Zone] does not trip the guard.
    $scan = [regex]::Replace($Query, "N?'(?:[^']|'')*'", "''")
    $scan = [regex]::Replace($scan, '\[(?:[^\]]|\]\])*\]', '[x]')
    if ($scan -notmatch '^\s*(SELECT|WITH)\s') { throw 'Guardrail: only a single SELECT/WITH statement is permitted.' }
    if ($scan.Contains(';')) { throw 'Guardrail: multiple statements are not permitted.' }
    if ($scan.Contains('--') -or $scan.Contains('/*')) { throw 'Guardrail: comments are not permitted in engine queries.' }
    $m = [regex]::Match($scan, $script:LSBlockedPattern)
    if ($m.Success) { throw "Guardrail: keyword '$($m.Value)' is not permitted. Query refused." }
}

function Open-LSConnection {
    param([string]$Server, [pscredential]$Credential, [switch]$Encrypt, [switch]$TrustServerCertificate,
        [int]$ConnectTimeoutSec = 10, [int]$LockTimeoutMs = 3000)
    $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $b['Data Source'] = $Server
    $b['Initial Catalog'] = 'master'
    $b['Application Name'] = $script:LSAppName
    $b['Connect Timeout'] = $ConnectTimeoutSec
    $b['Pooling'] = $false
    $b['Encrypt'] = [bool]$Encrypt
    $b['TrustServerCertificate'] = [bool]$TrustServerCertificate
    if ($Credential) {
        $b['Integrated Security'] = $false
        $pw = $Credential.Password.Copy(); $pw.MakeReadOnly()
        $sqlCred = New-Object System.Data.SqlClient.SqlCredential($Credential.UserName, $pw)
        $conn = New-Object System.Data.SqlClient.SqlConnection($b.ConnectionString, $sqlCred)
    }
    else {
        $b['Integrated Security'] = $true
        $conn = New-Object System.Data.SqlClient.SqlConnection($b.ConnectionString)
    }
    try {
        $conn.Open()
        # Fixed session settings (constant text, not user input): give up fast on locks, always be the deadlock victim.
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SET NOCOUNT ON; SET LOCK_TIMEOUT $([int]$LockTimeoutMs); SET DEADLOCK_PRIORITY LOW; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;"
        $cmd.CommandTimeout = 10
        [void]$cmd.ExecuteNonQuery()
        $cmd.Dispose()
    }
    catch { $conn.Dispose(); throw }
    return $conn
}

function Invoke-LSQuery {
    param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Query, [int]$TimeoutSec = 30)
    Assert-LSReadOnlyQuery -Query $Query
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = $TimeoutSec
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    try { [void]$da.Fill($dt) } finally { $da.Dispose(); $cmd.Dispose() }
    return , $dt
}

function New-LSSnapshot {
    param([string]$Name, [string]$Role, [string]$Mode)
    return [pscustomobject]@{
        Name = $Name; Role = $Role; Mode = $Mode; Connected = $false; Error = $null
        ServerName = $null; MachineName = $null; Version = $null; Edition = $null; IsHadrEnabled = $false
        CanViewDefinition = $false; CanControlServer = $false; CanViewServerState = $false
        Principals = @(); RoleMembers = @(); Permissions = @(); Databases = @(); DbUsers = @()
        Warnings = (New-Object System.Collections.Generic.List[string]); ElapsedMs = 0; QueryCount = 0
    }
}

function Get-LSServerSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Server, [string]$Role = 'Target', [string]$Mode = 'Standalone',
        [pscredential]$Credential, [switch]$Encrypt, [switch]$TrustServerCertificate,
        [int]$LockTimeoutMs = 3000, [int]$QueryTimeoutSec = 30,
        [switch]$CollectDbUsers, [string]$DatabaseScope = 'Auto', [int]$MaxDatabases = 300)

    $snap = New-LSSnapshot -Name $Server -Role $Role -Mode $Mode
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $conn = $null
    try {
        Write-LSLog "Connecting to $Server ($Role / $Mode)"
        $conn = Open-LSConnection -Server $Server -Credential $Credential -Encrypt:$Encrypt -TrustServerCertificate:$TrustServerCertificate -LockTimeoutMs $LockTimeoutMs

        $info = (Invoke-LSQuery $conn $script:LSQ_Info $QueryTimeoutSec).Rows[0]; $snap.QueryCount++
        $snap.ServerName = [string]$info['ServerName']
        $snap.MachineName = [string]$info['MachineName']
        $snap.Version = [string]$info['Version']
        $snap.Edition = [string]$info['Edition']
        $snap.IsHadrEnabled = ([int]$info['IsHadrEnabled'] -eq 1)
        $snap.CanViewDefinition = ([int](ConvertFrom-LSDbNull $info['CanViewDefinition']) -eq 1)
        $snap.CanControlServer = ([int](ConvertFrom-LSDbNull $info['CanControlServer']) -eq 1)
        $snap.CanViewServerState = ([int](ConvertFrom-LSDbNull $info['CanViewServerState']) -eq 1)
        if (-not $snap.CanViewDefinition) { $snap.Warnings.Add('No VIEW ANY DEFINITION: only your own login and fixed roles are visible; results will be incomplete.') }
        if (-not $snap.CanControlServer) { $snap.Warnings.Add('No CONTROL SERVER: password hashes not readable; password comparison and hash-based scripting skipped.') }

        $plist = New-Object System.Collections.Generic.List[object]
        foreach ($r in (Invoke-LSQuery $conn $script:LSQ_Principals $QueryTimeoutSec).Rows) {
            $hash = ConvertFrom-LSDbNull $r['password_hash']
            $plist.Add([pscustomobject]@{
                    Name = [string]$r['name']; Type = ([string]$r['type']).Trim(); TypeDesc = [string]$r['type_desc']
                    Sid = ConvertTo-LSHex (ConvertFrom-LSDbNull $r['sid'])
                    IsDisabled = [bool](ConvertFrom-LSDbNull $r['is_disabled'])
                    DefaultDb = ConvertFrom-LSDbNull $r['default_database_name']
                    DefaultLang = ConvertFrom-LSDbNull $r['default_language_name']
                    IsFixedRole = [bool](ConvertFrom-LSDbNull $r['is_fixed_role'])
                    PolicyChecked = ConvertFrom-LSDbNull $r['is_policy_checked']
                    ExpirationChecked = ConvertFrom-LSDbNull $r['is_expiration_checked']
                    PwdHash = ConvertTo-LSHex $hash
                    PwdDigest = Get-LSDigest $hash
                    CreateDate = ConvertFrom-LSDbNull $r['create_date']
                    ModifyDate = ConvertFrom-LSDbNull $r['modify_date']
                })
        }
        $snap.Principals = $plist.ToArray(); $snap.QueryCount++

        $snap.RoleMembers = @(foreach ($r in (Invoke-LSQuery $conn $script:LSQ_RoleMembers $QueryTimeoutSec).Rows) {
                [pscustomobject]@{ RoleName = [string]$r['RoleName']; MemberName = [string]$r['MemberName'] } })
        $snap.QueryCount++

        $snap.Permissions = @(foreach ($r in (Invoke-LSQuery $conn $script:LSQ_Permissions $QueryTimeoutSec).Rows) {
                $p = [pscustomobject]@{
                    Grantee = [string]$r['Grantee']; Class = [int]$r['class']; ClassDesc = [string]$r['class_desc']
                    Permission = [string]$r['permission_name']; State = [string]$r['state_desc']
                    MajorId = [int]$r['major_id']; TargetName = ConvertFrom-LSDbNull $r['TargetName']; Key = $null
                }
                $tgt = if ($p.TargetName) { $p.TargetName } elseif ($p.Class -eq 100) { '' } else { "#$($p.MajorId)" }
                $p.Key = '{0}|{1}|{2}|{3}' -f $p.State, $p.Permission, $p.ClassDesc, $tgt
                $p
            })
        $snap.QueryCount++

        if ($CollectDbUsers -and $DatabaseScope -ne 'None') {
            $dbs = @(foreach ($r in (Invoke-LSQuery $conn $script:LSQ_Databases $QueryTimeoutSec).Rows) {
                    [pscustomobject]@{ Name = [string]$r['name']; IsAgDb = ([int]$r['IsAgDb'] -eq 1); AgName = ConvertFrom-LSDbNull $r['AgName']; HasAccess = ([int](ConvertFrom-LSDbNull $r['HasAccess']) -eq 1) } })
            $snap.QueryCount++
            $scope = $DatabaseScope
            if ($scope -eq 'Auto') { $scope = if (@($dbs | Where-Object IsAgDb).Count -gt 0) { 'AG' } else { 'AllUser' } }
            $selected = @($dbs | Where-Object { $_.HasAccess -and ($scope -eq 'AllUser' -or $_.IsAgDb) })
            $noAccess = @($dbs | Where-Object { -not $_.HasAccess }).Count
            if ($noAccess -gt 0) { $snap.Warnings.Add("$noAccess database(s) skipped: no access from the connecting login.") }
            if ($MaxDatabases -gt 0 -and $selected.Count -gt $MaxDatabases) {
                $snap.Warnings.Add("Database user scan capped at $MaxDatabases of $($selected.Count) databases (-MaxDatabases).")
                $selected = $selected[0..($MaxDatabases - 1)]
            }
            $snap.Databases = $selected
            Write-LSLog "  scanning users in $($selected.Count) database(s) (scope: $scope)"
            $ulist = New-Object System.Collections.Generic.List[object]
            foreach ($db in $selected) {
                $q = $script:LSQ_DbUsers -f (Format-LSLiteral $db.Name), (Format-LSName $db.Name)
                try {
                    foreach ($r in (Invoke-LSQuery $conn $q $QueryTimeoutSec).Rows) {
                        $ulist.Add([pscustomobject]@{
                                Database = $db.Name; IsAgDb = $db.IsAgDb; AgName = $db.AgName
                                User = [string]$r['UserName']; Type = ([string]$r['UserType']).Trim(); TypeDesc = [string]$r['TypeDesc']
                                Sid = ConvertTo-LSHex (ConvertFrom-LSDbNull $r['sid'])
                            })
                    }
                    $snap.QueryCount++
                }
                catch { $snap.Warnings.Add("Database [$($db.Name)] skipped: $($_.Exception.GetBaseException().Message)") }
            }
            $snap.DbUsers = $ulist.ToArray()
        }
        $snap.Connected = $true
        Write-LSLog ("  {0}: {1} principals, {2} role memberships, {3} permissions{4}" -f $snap.ServerName, $snap.Principals.Count, $snap.RoleMembers.Count, $snap.Permissions.Count, $(if ($CollectDbUsers) { ", $($snap.DbUsers.Count) db users" } else { '' }))
    }
    catch {
        $snap.Error = $_.Exception.GetBaseException().Message
        Write-LSLog "  $Server failed: $($snap.Error)" 'Error'
    }
    finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
        $sw.Stop(); $snap.ElapsedMs = [int]$sw.ElapsedMilliseconds
    }
    return $snap
}

function Get-LSAgTopology {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Server, [pscredential]$Credential, [switch]$Encrypt, [switch]$TrustServerCertificate,
        [int]$LockTimeoutMs = 3000, [int]$QueryTimeoutSec = 30)
    $conn = $null
    try {
        $conn = Open-LSConnection -Server $Server -Credential $Credential -Encrypt:$Encrypt -TrustServerCertificate:$TrustServerCertificate -LockTimeoutMs $LockTimeoutMs
        $rows = (Invoke-LSQuery $conn $script:LSQ_AgTopology $QueryTimeoutSec).Rows
        $out = foreach ($r in $rows) {
            $replica = [string]$r['replica_server_name']
            $primary = [string](ConvertFrom-LSDbNull $r['primary_replica'])
            $modeRaw = [string]$r['availability_mode_desc']
            $mode = switch -Wildcard ($modeRaw) { 'SYNCHRONOUS*' { 'Synchronous' } 'ASYNCHRONOUS*' { 'Asynchronous' } 'CONFIGURATION*' { 'ConfigurationOnly' } default { $modeRaw } }
            [pscustomobject]@{
                AgName = [string]$r['AgName']; Replica = $replica; Mode = $mode
                FailoverMode = [string]$r['failover_mode_desc']; Readable = [string]$r['secondary_role_allow_connections_desc']
                EndpointUrl = [string]$r['endpoint_url']; PrimaryReplica = $primary
                IsPrimary = ($primary -and $replica -eq $primary)
                Connected = [string](ConvertFrom-LSDbNull $r['connected_state_desc'])
            }
        }
        return @($out)
    }
    finally { if ($conn) { $conn.Close(); $conn.Dispose() } }
}
#endregion

#region ---------- comparison ----------
function New-LSIndex {
    param($Snapshot, [string[]]$ExcludeLogin)
    $idx = @{ AllByName = @{}; AllBySid = @{}; ByName = @{}; Scoped = (New-Object System.Collections.Generic.List[object]); Roles = @{}; Perms = @{} }
    foreach ($p in $Snapshot.Principals) {
        $idx.AllByName[$p.Name] = $p
        if ($p.Sid -and $p.Type -ne 'R' -and -not $idx.AllBySid.ContainsKey($p.Sid)) { $idx.AllBySid[$p.Sid] = $p }
        if ($p.Type -eq 'R' -and ($p.IsFixedRole -or $p.Name -eq 'public')) { continue }
        if (Test-LSExcluded $p.Name $ExcludeLogin) { continue }
        $idx.ByName[$p.Name] = $p
        $idx.Scoped.Add($p)
    }
    foreach ($rm in $Snapshot.RoleMembers) {
        if (-not $idx.Roles.ContainsKey($rm.MemberName)) { $idx.Roles[$rm.MemberName] = @{} }
        $idx.Roles[$rm.MemberName][$rm.RoleName] = $true
    }
    foreach ($pe in $Snapshot.Permissions) {
        if (-not $idx.Perms.ContainsKey($pe.Grantee)) { $idx.Perms[$pe.Grantee] = @{} }
        $idx.Perms[$pe.Grantee][$pe.Key] = $pe
    }
    return $idx
}

function New-LSCreatePrincipalSql {
    param($P)
    $n = Format-LSName $P.Name
    $lit = Format-LSLiteral $P.Name
    $sb = New-Object System.Text.StringBuilder
    $lang = if ($P.DefaultLang) { ", DEFAULT_LANGUAGE = $(Format-LSName $P.DefaultLang)" } else { '' }
    switch ($P.Type) {
        'R' {
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = $lit) CREATE SERVER ROLE $n;")
            return $sb.ToString()
        }
        'S' {
            $pol = if ($P.PolicyChecked) { 'ON' } else { 'OFF' }
            $exp = if ($P.ExpirationChecked) { 'ON' } else { 'OFF' }
            if ($P.PwdHash) {
                [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = $lit)")
                [void]$sb.AppendLine("    CREATE LOGIN $n WITH PASSWORD = $($P.PwdHash) HASHED, SID = $($P.Sid), DEFAULT_DATABASE = [master]$lang, CHECK_POLICY = $pol, CHECK_EXPIRATION = $exp;")
            }
            else {
                [void]$sb.AppendLine("-- [MANUAL] Password hash for $n was not readable (needs CONTROL SERVER on the reference). Supply the password:")
                [void]$sb.AppendLine("-- CREATE LOGIN $n WITH PASSWORD = N'<password>', SID = $($P.Sid), DEFAULT_DATABASE = [master]$lang, CHECK_POLICY = $pol, CHECK_EXPIRATION = $exp;")
            }
        }
        { $_ -in 'U', 'G' } {
            [void]$sb.AppendLine("IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = $lit)")
            [void]$sb.AppendLine("    CREATE LOGIN $n FROM WINDOWS WITH DEFAULT_DATABASE = [master]$lang;")
        }
        default {
            [void]$sb.AppendLine("-- [MANUAL] $n is a $($P.TypeDesc); create the certificate/asymmetric key in master first, then CREATE LOGIN ... FROM CERTIFICATE/ASYMMETRIC KEY.")
            return $sb.ToString()
        }
    }
    if ($P.DefaultDb -and $P.DefaultDb -ne 'master') {
        [void]$sb.AppendLine("IF DB_ID($(Format-LSLiteral $P.DefaultDb)) IS NOT NULL AND EXISTS (SELECT 1 FROM sys.server_principals WHERE name = $lit) ALTER LOGIN $n WITH DEFAULT_DATABASE = $(Format-LSName $P.DefaultDb);")
    }
    if ($P.IsDisabled) { [void]$sb.AppendLine("IF EXISTS (SELECT 1 FROM sys.server_principals WHERE name = $lit) ALTER LOGIN $n DISABLE;") }
    return $sb.ToString()
}

function New-LSPermSql {
    param($Perm, [string]$Grantee, [hashtable]$RefByName, [switch]$Revoke)
    $on = $null
    switch ($Perm.Class) {
        100 { $on = '' }
        101 {
            if ($Perm.TargetName) {
                $tp = $RefByName[$Perm.TargetName]
                $on = if ($tp -and $tp.Type -eq 'R') { " ON SERVER ROLE::$(Format-LSName $Perm.TargetName)" } else { " ON LOGIN::$(Format-LSName $Perm.TargetName)" }
            }
        }
        105 { if ($Perm.TargetName) { $on = " ON ENDPOINT::$(Format-LSName $Perm.TargetName)" } }
    }
    if ($null -eq $on) { return "-- [MANUAL] $($Perm.State) $($Perm.Permission) on $($Perm.ClassDesc) id $($Perm.MajorId) for $(Format-LSName $Grantee) - not scriptable automatically." }
    $g = Format-LSName $Grantee
    if ($Revoke) { return "-- REVOKE $($Perm.Permission)$on FROM $g;   -- exists only on target ($($Perm.State))" }
    $verb = if ($Perm.State -eq 'DENY') { 'DENY' } else { 'GRANT' }
    $wgo = if ($Perm.State -eq 'GRANT_WITH_GRANT_OPTION') { ' WITH GRANT OPTION' } else { '' }
    return "$verb $($Perm.Permission)$on TO $g$wgo;"
}

function Compare-LSSnapshots {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Reference, [object[]]$Targets, [string[]]$ExcludeLogin, [switch]$SkipPasswordCompare)

    $issues = New-Object System.Collections.Generic.List[object]
    $matrix = [ordered]@{}
    $fixes = @{}
    $summaries = New-Object System.Collections.Generic.List[object]

    $all = @($Reference) + @($Targets)
    $machines = @{}
    foreach ($s in $all) {
        if ($s.MachineName) { $machines[$s.MachineName] = $true }
        $hostPart = ($s.Name -replace '^tcp:', '' -split '[\\,]')[0]
        if ($hostPart) { $machines[($hostPart -split '\.')[0]] = $true }
    }
    $isHostLocal = {
        param($p)
        if ($p.Type -notin 'U', 'G' -or -not $p.Name.Contains('\')) { return $false }
        return $machines.ContainsKey($p.Name.Split('\')[0])
    }

    $refIdx = New-LSIndex $Reference $ExcludeLogin
    $indexes = @{ $Reference.Name = $refIdx }

    $newIssue = {
        param($Severity, $Target, $Mode, $Object, $ObjectType, $Category, $Detail, $Suggestion)
        $issues.Add([pscustomobject]@{
                Severity = $Severity; Target = $Target; Mode = $Mode; Object = $Object; ObjectType = $ObjectType
                Category = $Category; Detail = $Detail; Suggestion = $Suggestion
            })
    }

    foreach ($p in $refIdx.Scoped) {
        $matrix[$p.Name] = [pscustomobject]@{
            Name = $p.Name; Type = $p.Type; TypeDesc = $p.TypeDesc; HostLocal = [bool](& $isHostLocal $p)
            Disabled = $p.IsDisabled; Cells = [ordered]@{}
        }
        $refNotes = @("SID $($p.Sid)")
        if ($p.IsDisabled) { $refNotes += 'Disabled' }
        if ($p.DefaultDb) { $refNotes += "Default DB: $($p.DefaultDb)" }
        if ($p.PwdDigest) { $refNotes += "Password digest: $($p.PwdDigest)" }
        $roles = if ($refIdx.Roles.ContainsKey($p.Name)) { @($refIdx.Roles[$p.Name].Keys | Sort-Object) } else { @() }
        if ($roles.Count) { $refNotes += "Roles: $($roles -join ', ')" }
        $matrix[$p.Name].Cells[$Reference.Name] = [pscustomobject]@{ S = 'REF'; N = $refNotes }
    }

    foreach ($t in $Targets) {
        $mode = $t.Mode
        $label = $t.Name
        $fx = @{
            Roles = New-Object System.Collections.Generic.List[string]; Missing = New-Object System.Collections.Generic.List[string]
            Renamed = New-Object System.Collections.Generic.List[string]; Sid = New-Object System.Collections.Generic.List[string]
            Password = New-Object System.Collections.Generic.List[string]; Props = New-Object System.Collections.Generic.List[string]
            RoleAdd = New-Object System.Collections.Generic.List[string]; RoleDrop = New-Object System.Collections.Generic.List[string]
            PermAdd = New-Object System.Collections.Generic.List[string]; PermRevoke = New-Object System.Collections.Generic.List[string]
            Extras = New-Object System.Collections.Generic.List[string]; Remap = New-Object System.Collections.Generic.List[string]
        }
        $fixes[$t.Name] = $fx

        if (-not $t.Connected) {
            foreach ($row in $matrix.Values) { $row.Cells[$t.Name] = [pscustomobject]@{ S = 'ERR'; N = @("Not reachable: $($t.Error)") } }
            & $newIssue 'High' $label $mode $t.Name 'Server' 'Unreachable' "Could not collect data: $($t.Error)" 'Check name/port, firewall, permissions (VIEW ANY DEFINITION) and that the instance is online.'
            continue
        }

        $tIdx = New-LSIndex $t $ExcludeLogin
        $indexes[$t.Name] = $tIdx

        foreach ($p in $refIdx.Scoped) {
            $row = $matrix[$p.Name]
            $notes = New-Object System.Collections.Generic.List[string]
            $status = 'OK'
            $n = Format-LSName $p.Name
            $lit = Format-LSLiteral $p.Name
            $otype = $p.TypeDesc

            if ($row.HostLocal) {
                $row.Cells[$t.Name] = [pscustomobject]@{ S = 'INFO'; N = @('Host-local Windows account: SID is machine-specific, excluded from scoring. Verify an equivalent account exists on this node.') }
                continue
            }

            $tp = $tIdx.AllByName[$p.Name]
            if (-not $tp) {
                $alt = $null
                if ($p.Type -ne 'R' -and $p.Sid) { $alt = $tIdx.AllBySid[$p.Sid] }
                if ($alt) {
                    $row.Cells[$t.Name] = [pscustomobject]@{ S = 'RENAMED'; N = @("Same SID exists here as [$($alt.Name)]") }
                    & $newIssue (Get-LSSeverity 'High' $mode) $label $mode $p.Name $otype 'Renamed' "SID $($p.Sid) exists on target under the name [$($alt.Name)]. Database users map fine, but applications connecting as [$($p.Name)] will fail." "Rename the target login to match the reference: ALTER LOGIN [$($alt.Name)] WITH NAME = $n."
                    $fx.Renamed.Add("-- [MANUAL] Login with reference SID is named [$($alt.Name)] here. Confirm nothing connects with that name, then:`r`n-- ALTER LOGIN $(Format-LSName $alt.Name) WITH NAME = $n;")
                }
                else {
                    $row.Cells[$t.Name] = [pscustomobject]@{ S = 'MISSING'; N = @('Not present on this server') }
                    $sev = if ($p.Type -eq 'R') { 'High' } else { 'Critical' }
                    $what = if ($p.Type -eq 'R') { 'Custom server role is missing.' } else { 'Login is missing: after failover/cutover this principal cannot connect and its database users are orphaned.' }
                    & $newIssue (Get-LSSeverity $sev $mode) $label $mode $p.Name $otype 'Missing' $what 'Create it with the reference SID (see remediation script section 2), then add role membership and permissions.'
                    if ($p.Type -eq 'R') { $fx.Roles.Add((New-LSCreatePrincipalSql $p)) } else { $fx.Missing.Add((New-LSCreatePrincipalSql $p)) }
                    if ($refIdx.Roles.ContainsKey($p.Name)) {
                        foreach ($r in ($refIdx.Roles[$p.Name].Keys | Sort-Object)) { $fx.RoleAdd.Add("ALTER SERVER ROLE $(Format-LSName $r) ADD MEMBER $n;") }
                    }
                    if ($refIdx.Perms.ContainsKey($p.Name)) {
                        foreach ($pe in $refIdx.Perms[$p.Name].Values) { $fx.PermAdd.Add((New-LSPermSql $pe $p.Name $refIdx.AllByName)) }
                    }
                }
                continue
            }

            $sevRank = 0
            if ($p.Type -ne $tp.Type) {
                $status = 'SID'; $notes.Add("Type differs: reference $($p.TypeDesc), target $($tp.TypeDesc)")
                & $newIssue (Get-LSSeverity 'High' $mode) $label $mode $p.Name $otype 'Type mismatch' "Reference is $($p.TypeDesc), target is $($tp.TypeDesc)." 'Drop and recreate on the target to match the reference (manual).'
            }
            elseif ($p.Type -ne 'R' -and $p.Sid -ne $tp.Sid) {
                $status = 'SID'; $notes.Add("SID differs: ref $($p.Sid) vs $($tp.Sid)")
                $isSql = ($p.Type -eq 'S')
                $detail = if ($isSql) { "SID mismatch: database users replicated from the reference map to $($p.Sid); on this server they would be orphaned after failover/restore." } else { "Windows SID mismatch: the AD account/group was probably recreated. Users mapped to the old SID are orphaned." }
                & $newIssue (Get-LSSeverity 'Critical' $mode) $label $mode $p.Name $otype 'SID mismatch' $detail 'Recreate the target login with the reference SID in a quiet window (section 3, commented out on purpose).'
                $blk = New-Object System.Text.StringBuilder
                [void]$blk.AppendLine("-- [MANUAL] $n  target SID $($tp.Sid)  ->  reference SID $($p.Sid)")
                [void]$blk.AppendLine("--   Impact: drops the target login (its server roles/permissions are re-added below). Database users on THIS server")
                [void]$blk.AppendLine("--   that are not from the reference (target-only databases) and map to the old SID would become orphaned.")
                [void]$blk.AppendLine("-- IF NOT EXISTS (SELECT 1 FROM sys.dm_exec_sessions WHERE login_name = $lit)")
                [void]$blk.AppendLine("-- BEGIN")
                [void]$blk.AppendLine("--     DROP LOGIN $n;")
                foreach ($ln in ((New-LSCreatePrincipalSql $p) -split "`r?`n")) { if ($ln.Trim()) { [void]$blk.AppendLine("--     " + ($ln -replace '^-- ', '')) } }
                if ($refIdx.Roles.ContainsKey($p.Name)) { foreach ($r in ($refIdx.Roles[$p.Name].Keys | Sort-Object)) { [void]$blk.AppendLine("--     ALTER SERVER ROLE $(Format-LSName $r) ADD MEMBER $n;") } }
                if ($refIdx.Perms.ContainsKey($p.Name)) { foreach ($pe in $refIdx.Perms[$p.Name].Values) { [void]$blk.AppendLine("--     " + ((New-LSPermSql $pe $p.Name $refIdx.AllByName) -replace '^-- ', '')) } }
                [void]$blk.AppendLine("-- END")
                [void]$blk.AppendLine("-- ELSE RAISERROR(N'Login $($p.Name.Replace("'", "''")) has active sessions; retry later.', 10, 1) WITH NOWAIT;")
                $fx.Sid.Add($blk.ToString())
                if ($t.Mode -eq 'Standalone') {
                    foreach ($u in @($Reference.DbUsers | Where-Object { $_.Sid -eq $p.Sid -and $_.User -ne 'dbo' })) {
                        $fx.Remap.Add("-- USE $(Format-LSName $u.Database); ALTER USER $(Format-LSName $u.User) WITH LOGIN = $n;   -- only if the database is restored here and you prefer remapping over fixing the SID")
                    }
                }
            }

            if ($p.Type -ne 'R' -and $p.Type -eq $tp.Type) {
                if ($p.IsDisabled -ne $tp.IsDisabled) {
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    if ($tp.IsDisabled) {
                        $notes.Add('Disabled here, enabled on reference')
                        & $newIssue (Get-LSSeverity 'High' $mode) $label $mode $p.Name $otype 'Disabled state' 'Login is DISABLED on the target but enabled on the reference: connections fail after failover.' 'ALTER LOGIN ... ENABLE on the target.'
                        $fx.Props.Add("ALTER LOGIN $n ENABLE;")
                    }
                    else {
                        $notes.Add('Enabled here, disabled on reference')
                        & $newIssue (Get-LSSeverity 'Medium' $mode) $label $mode $p.Name $otype 'Disabled state' 'Login is enabled on the target but disabled on the reference (larger attack surface on the target).' 'ALTER LOGIN ... DISABLE on the target, unless intentionally different.'
                        $fx.Props.Add("ALTER LOGIN $n DISABLE;")
                    }
                }
                if ("$($p.DefaultDb)" -ne "$($tp.DefaultDb)") {
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Default DB: $($tp.DefaultDb) (ref $($p.DefaultDb))")
                    & $newIssue (Get-LSSeverity 'Low' $mode) $label $mode $p.Name $otype 'Default database' "Default database is [$($tp.DefaultDb)] on target, [$($p.DefaultDb)] on reference." 'Align so the login lands in the same database after failover.'
                    if ($p.DefaultDb) { $fx.Props.Add("IF DB_ID($(Format-LSLiteral $p.DefaultDb)) IS NOT NULL ALTER LOGIN $n WITH DEFAULT_DATABASE = $(Format-LSName $p.DefaultDb);") }
                }
                if ("$($p.DefaultLang)" -ne "$($tp.DefaultLang)" -and $p.DefaultLang) {
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Default language: $($tp.DefaultLang) (ref $($p.DefaultLang))")
                    & $newIssue (Get-LSSeverity 'Low' $mode) $label $mode $p.Name $otype 'Default language' "Default language differs ($($tp.DefaultLang) vs $($p.DefaultLang)); date parsing can change after failover." 'Align DEFAULT_LANGUAGE.'
                    $fx.Props.Add("ALTER LOGIN $n WITH DEFAULT_LANGUAGE = $(Format-LSName $p.DefaultLang);")
                }
                if ($p.Type -eq 'S') {
                    if ([bool]$p.PolicyChecked -ne [bool]$tp.PolicyChecked -or [bool]$p.ExpirationChecked -ne [bool]$tp.ExpirationChecked) {
                        if ($status -eq 'OK') { $status = 'DIFF' }
                        $notes.Add('Password policy/expiration flags differ')
                        & $newIssue (Get-LSSeverity 'Low' $mode) $label $mode $p.Name $otype 'Password policy' "CHECK_POLICY/CHECK_EXPIRATION: target $([bool]$tp.PolicyChecked)/$([bool]$tp.ExpirationChecked), reference $([bool]$p.PolicyChecked)/$([bool]$p.ExpirationChecked)." 'Align policy flags (expiration on only one side can lock an app out on a different day).'
                        $pol = if ($p.PolicyChecked) { 'ON' } else { 'OFF' }; $exp = if ($p.ExpirationChecked) { 'ON' } else { 'OFF' }
                        if ($exp -eq 'OFF') { $fx.Props.Add("ALTER LOGIN $n WITH CHECK_EXPIRATION = OFF;"); $fx.Props.Add("ALTER LOGIN $n WITH CHECK_POLICY = $pol;") }
                        else { $fx.Props.Add("ALTER LOGIN $n WITH CHECK_POLICY = $pol;"); $fx.Props.Add("ALTER LOGIN $n WITH CHECK_EXPIRATION = $exp;") }
                    }
                    if (-not $SkipPasswordCompare -and $p.PwdDigest -and $tp.PwdDigest -and $p.PwdDigest -ne $tp.PwdDigest -and $status -ne 'SID') {
                        if ($status -eq 'OK') { $status = 'DIFF' }
                        $notes.Add("Password hash differs ($($tp.PwdDigest) vs ref $($p.PwdDigest))")
                        & $newIssue (Get-LSSeverity 'High' $mode) $label $mode $p.Name $otype 'Password' 'Password hash differs. Hashes are salted, so the same password typed twice also differs - but if this was synced by hash it means the password really changed on one side.' 'Re-apply the reference hash on the target (section 4) so the app password works after failover.'
                        if ($p.PwdHash) { $fx.Password.Add("ALTER LOGIN $n WITH PASSWORD = $($p.PwdHash) HASHED;") }
                    }
                }
            }

            # role membership
            $rRoles = if ($refIdx.Roles.ContainsKey($p.Name)) { $refIdx.Roles[$p.Name] } else { @{} }
            $tRoles = if ($tIdx.Roles.ContainsKey($p.Name)) { $tIdx.Roles[$p.Name] } else { @{} }
            foreach ($r in ($rRoles.Keys | Sort-Object)) {
                if (-not $tRoles.ContainsKey($r)) {
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Missing role: $r")
                    $base = if ($r -eq 'sysadmin') { 'Critical' } else { 'High' }
                    & $newIssue (Get-LSSeverity $base $mode) $label $mode $p.Name $otype 'Role membership' "Member of [$r] on reference but not on target." "ALTER SERVER ROLE [$r] ADD MEMBER $n on the target."
                    $fx.RoleAdd.Add("ALTER SERVER ROLE $(Format-LSName $r) ADD MEMBER $n;")
                }
            }
            foreach ($r in ($tRoles.Keys | Sort-Object)) {
                if (-not $rRoles.ContainsKey($r)) {
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Extra role on target: $r")
                    $base = if ($r -eq 'sysadmin') { 'High' } else { 'Medium' }
                    & $newIssue (Get-LSSeverity $base $mode) $label $mode $p.Name $otype 'Role membership' "Member of [$r] on target only (more privilege than on the reference)." 'Review; remove if not intended (commented in section 5).'
                    $fx.RoleDrop.Add("-- ALTER SERVER ROLE $(Format-LSName $r) DROP MEMBER $n;   -- exists only on target")
                }
            }

            # server permissions
            $rPerm = if ($refIdx.Perms.ContainsKey($p.Name)) { $refIdx.Perms[$p.Name] } else { @{} }
            $tPerm = if ($tIdx.Perms.ContainsKey($p.Name)) { $tIdx.Perms[$p.Name] } else { @{} }
            foreach ($k in ($rPerm.Keys | Sort-Object)) {
                if (-not $tPerm.ContainsKey($k)) {
                    $pe = $rPerm[$k]
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Missing permission: $($pe.State) $($pe.Permission)$(if ($pe.TargetName) { ' ON ' + $pe.TargetName })")
                    $base = if ($pe.Permission -eq 'CONTROL SERVER' -or $pe.State -eq 'DENY') { 'High' } else { 'Medium' }
                    & $newIssue (Get-LSSeverity $base $mode) $label $mode $p.Name $otype 'Server permission' "$($pe.State) $($pe.Permission) ($($pe.ClassDesc)$(if ($pe.TargetName) { ': ' + $pe.TargetName })) exists on reference only." 'Grant on the target (section 6).'
                    $fx.PermAdd.Add((New-LSPermSql $pe $p.Name $refIdx.AllByName))
                }
            }
            foreach ($k in ($tPerm.Keys | Sort-Object)) {
                if (-not $rPerm.ContainsKey($k)) {
                    $pe = $tPerm[$k]
                    if ($status -eq 'OK') { $status = 'DIFF' }
                    $notes.Add("Extra permission on target: $($pe.State) $($pe.Permission)")
                    & $newIssue (Get-LSSeverity 'Low' $mode) $label $mode $p.Name $otype 'Server permission' "$($pe.State) $($pe.Permission) ($($pe.ClassDesc)) exists on target only." 'Review; revoke if not intended (commented in section 6).'
                    $fx.PermRevoke.Add((New-LSPermSql $pe $p.Name $tIdx.AllByName -Revoke))
                }
            }

            $row.Cells[$t.Name] = [pscustomobject]@{ S = $status; N = @($notes) }
        }

        # principals that exist only on the target
        foreach ($tp in $tIdx.Scoped) {
            if ($refIdx.AllByName.ContainsKey($tp.Name)) { continue }
            if (-not $matrix.Contains($tp.Name)) {
                $matrix[$tp.Name] = [pscustomobject]@{ Name = $tp.Name; Type = $tp.Type; TypeDesc = $tp.TypeDesc; HostLocal = [bool](& $isHostLocal $tp); Disabled = $tp.IsDisabled; Cells = [ordered]@{} }
                $matrix[$tp.Name].Cells[$Reference.Name] = [pscustomobject]@{ S = 'ABSENT'; N = @('Not on reference') }
            }
            if ($matrix[$tp.Name].HostLocal) { $matrix[$tp.Name].Cells[$t.Name] = [pscustomobject]@{ S = 'INFO'; N = @('Host-local Windows account of this node (machine-specific SID).') }; continue }
            $matrix[$tp.Name].Cells[$t.Name] = [pscustomobject]@{ S = 'EXTRA'; N = @('Exists only on target', "SID $($tp.Sid)") }
            if (-not $matrix[$tp.Name].HostLocal) {
                & $newIssue 'Info' $label $mode $tp.Name $tp.TypeDesc 'Target only' 'Exists on the target but not on the reference.' 'Review: leftover from an old failover/migration or target-specific. Never dropped automatically.'
                $fx.Extras.Add("--   $(Format-LSName $tp.Name)  ($($tp.TypeDesc), created $($tp.CreateDate))")
            }
        }
    }

    # ---------- orphan risk: reference database users vs every server's logins (by SID) ----------
    $orphans = New-Object System.Collections.Generic.List[object]
    $orphanAgg = @{}
    foreach ($u in $Reference.DbUsers) {
        $refLogin = $refIdx.AllBySid[$u.Sid]
        $lookupName = if ($refLogin) { $refLogin.Name } elseif ($u.User -ne 'dbo') { $u.User } else { $null }
        $orow = [pscustomobject]@{ Database = $u.Database; IsAgDb = $u.IsAgDb; User = $u.User; Type = $u.TypeDesc; Sid = $u.Sid; Login = $(if ($refLogin) { $refLogin.Name } else { $null }); Cells = [ordered]@{} }
        $anyIssue = $false
        foreach ($s in $all) {
            if (-not $s.Connected) { $orow.Cells[$s.Name] = [pscustomobject]@{ S = 'ERR'; N = @('Server not reachable') }; continue }
            $idx = $indexes[$s.Name]
            $hit = $idx.AllBySid[$u.Sid]
            if ($hit) { $orow.Cells[$s.Name] = [pscustomobject]@{ S = 'OK'; N = @("Maps to login [$($hit.Name)]") }; continue }
            $byName = if ($lookupName) { $idx.AllByName[$lookupName] } else { $null }
            if ($byName) { $st = 'SID'; $msg = "Login [$($byName.Name)] exists with a different SID ($($byName.Sid))" }
            elseif ($u.User -eq 'dbo') { $st = 'OWNER'; $msg = 'Database owner SID has no login here' }
            elseif ($u.Type -eq 'U') { $st = 'GROUP'; $msg = 'No individual login: fine only if access is granted through a Windows group login' }
            else { $st = 'MISSING'; $msg = 'No login with this SID' }
            $orow.Cells[$s.Name] = [pscustomobject]@{ S = $st; N = @($msg) }
            $anyIssue = $true
            $key = '{0}|{1}|{2}' -f $s.Name, $u.Sid, $st
            if (-not $orphanAgg.ContainsKey($key)) {
                $orphanAgg[$key] = [pscustomobject]@{ Server = $s; Status = $st; Login = $lookupName; Sid = $u.Sid; Type = $u.TypeDesc; Users = (New-Object System.Collections.Generic.List[string]) }
            }
            $orphanAgg[$key].Users.Add("$($u.Database).$($u.User)")
        }
        $orow | Add-Member -NotePropertyName AtRisk -NotePropertyValue $anyIssue
        $orphans.Add($orow)
    }
    $refBroken = @{}
    foreach ($a in $orphanAgg.Values) { if ($a.Server.Name -eq $Reference.Name) { $refBroken[$a.Sid] = $true } }
    foreach ($a in @($orphanAgg.Values | Sort-Object { $_.Server.Name })) {
        $s = $a.Server; $isRef = ($s.Name -eq $Reference.Name)
        # Already broken on the reference: report once (on the reference), not again on every target.
        if (-not $isRef -and $refBroken.ContainsKey($a.Sid)) { continue }
        $count = $a.Users.Count
        $list = (@($a.Users) | Select-Object -First 8) -join ', '
        if ($count -gt 8) { $list += ", +$($count - 8) more" }
        $obj = if ($a.Login) { $a.Login } else { "SID $($a.Sid)" }
        switch ($a.Status) {
            'GROUP' { $base = 'Low'; $sug = 'Confirm the user is a member of a Windows group that has a login on this server.' }
            'OWNER' { $base = 'Medium'; $sug = 'Set a valid owner after failover/restore: ALTER AUTHORIZATION ON DATABASE::[db] TO [login] (on the primary).' }
            default { $base = 'Critical'; $sug = 'Create the login with the reference SID (fixing the login fixes every database at once; no ALTER USER needed on an AG).' }
        }
        if ($isRef) {
            $sev = if ($base -eq 'Critical') { 'High' } else { $base }
            $detail = "$count database user(s) are ALREADY orphaned on the reference: $list"
            $sug = 'Recreate/fix the login on the reference first, or remove the stale users.'
        }
        else {
            $sev = Get-LSSeverity $base $s.Mode
            $detail = "$count database user(s) would be orphaned here after failover/restore: $list"
            # Fold into the login-level finding when one exists (Missing / SID mismatch / Renamed) instead of a second row.
            $loginIssue = $null
            if ($a.Login) { $loginIssue = $issues | Where-Object { $_.Target -eq $s.Name -and $_.Object -eq $a.Login -and $_.Category -in 'Missing', 'SID mismatch', 'Renamed', 'Type mismatch' } | Select-Object -First 1 }
            if ($loginIssue) { $loginIssue.Detail += " Orphans $count database user(s): $list."; continue }
        }
        & $newIssue $sev $s.Name $s.Mode $obj $a.Type 'Orphan risk' $detail $sug
    }

    # ---------- per-server summary ----------
    $inScope = @($refIdx.Scoped | Where-Object { -not $matrix[$_.Name].HostLocal }).Count
    foreach ($s in $all) {
        $isRef = ($s.Name -eq $Reference.Name)
        $mine = @($issues | Where-Object { $_.Target -eq $s.Name })
        $ok = 0
        if (-not $isRef -and $s.Connected) {
            foreach ($p in $refIdx.Scoped) { $c = $matrix[$p.Name].Cells[$s.Name]; if ($c -and $c.S -eq 'OK') { $ok++ } }
        }
        $score = if ($isRef) { $null } elseif (-not $s.Connected) { $null } elseif ($inScope -eq 0) { 100 } else { [int][math]::Floor(100 * $ok / $inScope) }
        $summaries.Add([pscustomobject]@{
                Name = $s.Name; ServerName = $s.ServerName; Role = $s.Role; Mode = $s.Mode; Connected = $s.Connected; Error = $s.Error
                Version = $s.Version; Edition = $s.Edition; CanViewDefinition = $s.CanViewDefinition; CanControlServer = $s.CanControlServer
                Score = $score; InScope = $inScope; InSync = $ok
                Critical = @($mine | Where-Object Severity -eq 'Critical').Count
                High = @($mine | Where-Object Severity -eq 'High').Count
                Medium = @($mine | Where-Object Severity -eq 'Medium').Count
                Low = @($mine | Where-Object Severity -eq 'Low').Count
                Info = @($mine | Where-Object Severity -eq 'Info').Count
                OrphanRisk = @($orphans | Where-Object { $_.Cells[$s.Name] -and $_.Cells[$s.Name].S -notin 'OK', 'ERR' }).Count
                ElapsedMs = $s.ElapsedMs; QueryCount = $s.QueryCount; Warnings = @($s.Warnings)
            })
    }

    $sorted = @($issues | Sort-Object @{ Expression = { $script:LSSevRank[$_.Severity] }; Descending = $true }, Target, Object)

    $scripts = [ordered]@{}
    foreach ($t in $Targets) { if ($t.Connected) { $scripts[$t.Name] = New-LSRemediationScript -Reference $Reference -Target $t -Fix $fixes[$t.Name] } }

    return [pscustomobject]@{
        Meta = $null
        Servers = $summaries.ToArray()
        Matrix = @($matrix.Values)
        Issues = $sorted
        Orphans = $orphans.ToArray()
        Scripts = $scripts
        Recommendations = @()
    }
}

function New-LSRemediationScript {
    param($Reference, $Target, $Fix)
    $sb = New-Object System.Text.StringBuilder
    $tname = if ($Target.ServerName) { $Target.ServerName } else { $Target.Name }
    $L = { param($s) [void]$sb.AppendLine($s) }
    & $L '/*'
    & $L "  LoginSync remediation  |  TARGET: $tname ($($Target.Mode))  |  REFERENCE: $($Reference.ServerName)"
    & $L "  Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME  (tool v$script:LSVersion)"
    & $L ''
    & $L '  NOTHING IN THIS FILE HAS BEEN EXECUTED. Review every statement before running it on the target.'
    & $L '  * The batch below refuses to run anywhere except the target (@@SERVERNAME check + SET NOEXEC ON).'
    & $L '  * Every CREATE is guarded by IF NOT EXISTS, so re-running is safe.'
    & $L '  * Lines starting with "-- [MANUAL]" or "-- ALTER/REVOKE/DROP" are deliberately commented: destructive or judgement calls.'
    & $L '  * May contain SQL password hashes: treat as a secret, store securely and delete after use.'
    & $L '  * Logins live in master and are NOT replicated by Always On; run this on each secondary that shows drift.'
    & $L '*/'
    & $L 'SET NOCOUNT ON;'
    & $L "IF @@SERVERNAME <> $(Format-LSLiteral $tname)"
    & $L 'BEGIN'
    & $L "    RAISERROR(N'Wrong server: this script is for $($tname.Replace("'", "''")). Nothing was executed.', 16, 1) WITH NOWAIT;"
    & $L '    SET NOEXEC ON;'
    & $L 'END'
    & $L 'GO'
    $sections = @(
        @('1. Custom server roles', $Fix.Roles),
        @('2. Missing logins (created with the reference SID, password hash, policy, default DB, disabled state)', $Fix.Missing),
        @('3. SID mismatches  [MANUAL - drop & recreate during a quiet window]', $Fix.Sid),
        @('3b. Logins present under another name  [MANUAL]', $Fix.Renamed),
        @('4. Password hash re-sync', $Fix.Password),
        @('5. Login properties (enabled/disabled, default database/language, policy)', $Fix.Props),
        @('5b. Server role membership', $Fix.RoleAdd),
        @('5c. Role membership present only on target  [review]', $Fix.RoleDrop),
        @('6. Server-level permissions', $Fix.PermAdd),
        @('6b. Permissions present only on target  [review]', $Fix.PermRevoke),
        @('7. Migration only: remap database users instead of fixing SIDs  [MANUAL]', $Fix.Remap)
    )
    $any = $false
    foreach ($sec in $sections) {
        $items = $sec[1]
        if ($items.Count -eq 0) { continue }
        $any = $true
        & $L ''
        & $L ('-- ' + ('=' * 100))
        & $L "-- $($sec[0])"
        & $L ('-- ' + ('=' * 100))
        foreach ($i in $items) { [void]$sb.Append($i); if (-not $i.EndsWith("`n")) { [void]$sb.AppendLine() } }
        & $L 'GO'
    }
    if ($Fix.Extras.Count) {
        & $L ''
        & $L '-- Principals that exist ONLY on this target (informational - not dropped):'
        foreach ($i in $Fix.Extras) { & $L $i }
    }
    if (-not $any) { & $L ''; & $L '-- No differences requiring changes were found for this target.' }
    & $L ''
    & $L 'SET NOEXEC OFF;'
    & $L 'GO'
    return $sb.ToString()
}

function Get-LSRecommendations {
    param($Result)
    $rec = New-Object System.Collections.Generic.List[object]
    $add = { param($lvl, $txt) $rec.Add([pscustomobject]@{ Level = $lvl; Text = $txt }) }
    $targets = @($Result.Servers | Where-Object { $_.Role -ne 'Reference' -and $_.Role -ne 'Primary' })
    foreach ($s in $targets) {
        if (-not $s.Connected) { & $add 'High' "$($s.Name): not reachable - nothing was compared. Fix connectivity before trusting this report."; continue }
        $mine = @($Result.Issues | Where-Object Target -eq $s.Name)
        $miss = @($mine | Where-Object Category -eq 'Missing').Count
        $sid = @($mine | Where-Object Category -eq 'SID mismatch').Count
        $pwd = @($mine | Where-Object Category -eq 'Password').Count
        $tag = switch ($s.Mode) { 'Synchronous' { 'synchronous / failover partner' } 'Asynchronous' { 'asynchronous / DR' } default { $s.Mode.ToLower() } }
        if ($s.Critical + $s.High -eq 0 -and $s.Medium -eq 0) { & $add 'Good' "$($s.Name) ($tag): in sync ($($s.Score)% of $($s.InScope) principals identical)."; continue }
        $lvl = if ($s.Critical -gt 0) { 'Critical' } elseif ($s.High -gt 0) { 'High' } else { 'Medium' }
        $parts = @()
        if ($miss) { $parts += "$miss missing" }
        if ($sid) { $parts += "$sid SID mismatch" }
        if ($pwd) { $parts += "$pwd password drift" }
        $why = if ($s.Mode -eq 'Synchronous') { 'An automatic failover today would break these logins.' } elseif ($s.Mode -eq 'Asynchronous') { 'Fix before the next DR test or forced failover.' } else { 'Fix before cutover.' }
        & $add $lvl ("$($s.Name) ($tag): $($s.Score)% in sync" + $(if ($parts) { ' - ' + ($parts -join ', ') } else { '' }) + ". $why Review remediation_$($s.Name -replace '[\\/:*?""<>|,]', '_').sql.")
    }
    if (@($Result.Issues | Where-Object Category -eq 'SID mismatch').Count) {
        & $add 'Info' 'SID mismatches cannot be fixed with ALTER LOGIN. Drop and recreate the target login with the reference SID during a quiet window (the script block is commented out on purpose and checks for active sessions first). On an AG, never ALTER USER on the database - fix the login and every replica database lines up.'
    }
    if (@($Result.Issues | Where-Object { $_.Category -eq 'Orphan risk' -and $_.Target -eq $Result.Meta.Reference }).Count) {
        & $add 'High' 'Some database users are already orphaned on the reference itself - these will be orphaned everywhere. Fix or clean up on the primary first.'
    }
    if (@($Result.Servers | Where-Object { $_.Connected -and -not $_.CanControlServer }).Count) {
        & $add 'Info' 'Password hashes were not readable on at least one server (needs CONTROL SERVER). Run with a sysadmin login to compare passwords and to script missing SQL logins with their real hash.'
    }
    & $add 'Automation' 'Schedule this script (Task Scheduler or a SQL Agent PowerShell/CmdExec step) with -DiscoverAgFrom <listener> -NoOpen -FailOnSeverity High -OutputPath \\share\LoginSync. Exit code 2 = drift: alert on it and attach the HTML.'
    & $add 'Automation' 'For hands-off sync after review, dbatools Copy-DbaLogin (SID + hash preserved) or Sync-DbaAvailabilityGroup can push logins primary -> secondaries; keep this report as the read-only before/after check.'
    & $add 'Automation' 'Make login creation part of the change process: create on the primary with an explicit SID, then on every replica with the same SID and hash. Re-run this check after every security change and after every failover.'
    $order = @{ Critical = 0; High = 1; Medium = 2; Low = 3; Good = 4; Info = 5; Automation = 6 }
    return @($rec | Sort-Object { $order[$_.Level] })
}

function Invoke-LSLoginSync {
    [CmdletBinding()]
    param([string]$ReferenceServer, [string]$ReferenceRole = 'Reference', [object[]]$Targets, [string]$AgName,
        [pscredential]$Credential, [switch]$Encrypt, [switch]$TrustServerCertificate,
        [string]$DatabaseScope = 'Auto', [string[]]$ExcludeLogin = $script:LSDefaultExclude, [switch]$SkipPasswordCompare,
        [int]$LockTimeoutMs = 3000, [int]$QueryTimeoutSec = 30, [int]$MaxDatabases = 300, [switch]$Demo)

    $started = Get-Date
    if ($Demo) {
        Write-LSLog 'Demo mode: using built-in sample data (no server contacted).'
        $snaps = New-LSDemoSnapshots
        $ref = $snaps[0]; $tSnaps = @($snaps[1..($snaps.Count - 1)]); $AgName = 'AG_Orders'
    }
    else {
        if (-not $ReferenceServer) { throw 'A reference server is required.' }
        if (-not $Targets -or @($Targets).Count -eq 0) { throw 'At least one target server is required.' }
        $common = @{ Credential = $Credential; Encrypt = $Encrypt; TrustServerCertificate = $TrustServerCertificate; LockTimeoutMs = $LockTimeoutMs; QueryTimeoutSec = $QueryTimeoutSec }
        $ref = Get-LSServerSnapshot -Server $ReferenceServer -Role $ReferenceRole -Mode $ReferenceRole -CollectDbUsers -DatabaseScope $DatabaseScope -MaxDatabases $MaxDatabases @common
        if (-not $ref.Connected) { throw "Reference server $ReferenceServer is not reachable: $($ref.Error)" }
        $tSnaps = @(foreach ($t in $Targets) { Get-LSServerSnapshot -Server $t.Name -Role 'Target' -Mode $t.Mode @common })
    }
    Write-LSLog 'Comparing...'
    $result = Compare-LSSnapshots -Reference $ref -Targets $tSnaps -ExcludeLogin $ExcludeLogin -SkipPasswordCompare:$SkipPasswordCompare
    $result.Meta = [pscustomobject]@{
        Tool = 'LoginSync'; Version = $script:LSVersion; GeneratedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); DurationSec = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
        RunBy = "$env:USERDOMAIN\$env:USERNAME"; RunFrom = $env:COMPUTERNAME; Reference = $ref.Name; ReferenceServerName = $ref.ServerName
        AgName = $AgName; Demo = [bool]$Demo; DatabaseScope = $DatabaseScope; ExcludeLogin = @($ExcludeLogin); LockTimeoutMs = $LockTimeoutMs
        QueryTimeoutSec = $QueryTimeoutSec; PasswordCompare = (-not $SkipPasswordCompare); DatabasesScanned = @($ref.Databases).Count
        DbUsersScanned = @($ref.DbUsers).Count; Warnings = @(foreach ($s in @($ref) + $tSnaps) { foreach ($w in $s.Warnings) { "$($s.Name): $w" } })
    }
    $result.Recommendations = Get-LSRecommendations $result
    $worst = 'Info'
    foreach ($i in $result.Issues) { if ($script:LSSevRank[$i.Severity] -gt $script:LSSevRank[$worst]) { $worst = $i.Severity } }
    $result.Meta | Add-Member -NotePropertyName WorstSeverity -NotePropertyValue $(if ($result.Issues.Count) { $worst } else { 'None' })
    Write-LSLog ("Done: {0} finding(s), worst = {1}" -f $result.Issues.Count, $result.Meta.WorstSeverity)
    return $result
}
#endregion

#region ---------- report output ----------
function Get-LSRedactedScript { param([string]$Sql) return [regex]::Replace($Sql, 'PASSWORD = 0x[0-9A-Fa-f]+ HASHED', 'PASSWORD = 0x<hash in .sql file> HASHED') }

function ConvertTo-LSHtml {
    param($Result)
    $scripts = [ordered]@{}
    foreach ($k in $Result.Scripts.Keys) { $scripts[$k] = Get-LSRedactedScript $Result.Scripts[$k] }
    $payload = [pscustomobject]@{
        Meta = $Result.Meta; Servers = $Result.Servers; Matrix = $Result.Matrix; Issues = $Result.Issues
        Orphans = $Result.Orphans; Scripts = $scripts; Recommendations = $Result.Recommendations
    }
    $json = ConvertTo-Json -InputObject $payload -Depth 8 -Compress
    $json = $json.Replace('</', '<\/')
    return $script:LSHtmlTemplate.Replace('/*__LSDATA__*/', $json)
}

function Export-LSReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][string]$OutputPath)
    $safe = { param($s) ($s -replace '[\\/:*?"<>|,]', '_') }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $folder = Join-Path $OutputPath ("LoginSync_{0}_{1}" -f (& $safe $Result.Meta.Reference), $stamp)
    [void](New-Item -ItemType Directory -Path $folder -Force)
    $utf8 = New-Object System.Text.UTF8Encoding($true)

    $html = Join-Path $folder 'LoginSyncReport.html'
    [System.IO.File]::WriteAllText($html, (ConvertTo-LSHtml $Result), $utf8)

    $sqlFiles = @()
    foreach ($k in $Result.Scripts.Keys) {
        $f = Join-Path $folder ("remediation_{0}.sql" -f (& $safe $k))
        [System.IO.File]::WriteAllText($f, $Result.Scripts[$k], $utf8)
        $sqlFiles += $f
    }

    $csv = Join-Path $folder 'findings.csv'
    $Result.Issues | Select-Object Severity, Target, Mode, Object, ObjectType, Category, Detail, Suggestion | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

    $mcsv = Join-Path $folder 'login_matrix.csv'
    $names = @($Result.Servers | ForEach-Object Name)
    $Result.Matrix | ForEach-Object {
        $o = [ordered]@{ Principal = $_.Name; Type = $_.TypeDesc }
        foreach ($n in $names) { $c = $_.Cells[$n]; $o[$n] = if ($c) { $c.S } else { 'ABSENT' } }
        [pscustomobject]$o
    } | Export-Csv -Path $mcsv -NoTypeInformation -Encoding UTF8

    $jsonPath = Join-Path $folder 'result.json'
    $red = [ordered]@{}; foreach ($k in $Result.Scripts.Keys) { $red[$k] = Get-LSRedactedScript $Result.Scripts[$k] }
    $j = [pscustomobject]@{ Meta = $Result.Meta; Servers = $Result.Servers; Issues = $Result.Issues; Recommendations = $Result.Recommendations; Scripts = $red }
    [System.IO.File]::WriteAllText($jsonPath, (ConvertTo-Json -InputObject $j -Depth 6), $utf8)

    return [pscustomobject]@{ Folder = $folder; Html = $html; Sql = $sqlFiles; Csv = $csv; MatrixCsv = $mcsv; Json = $jsonPath }
}
#endregion

#region ---------- demo data ----------
function New-LSDemoSnapshots {
    $hashOf = { param($seed) $sha = [System.Security.Cryptography.SHA512]::Create(); $b = [byte[]](@(2, 0) + [byte[]]$sha.ComputeHash([Text.Encoding]::UTF8.GetBytes("salt$seed")) + [byte[]](1, 2, 3, 4)); $sha.Dispose(); , $b }
    $mk = {
        param($Name, $Type, $Sid, [bool]$Disabled = $false, $Db = 'master', $Pwd = $null, $Policy = $true)
        $td = @{ S = 'SQL_LOGIN'; U = 'WINDOWS_LOGIN'; G = 'WINDOWS_GROUP'; R = 'SERVER_ROLE' }[$Type]
        $h = if ($Type -eq 'S' -and $Pwd) { & $hashOf $Pwd } else { $null }
        [pscustomobject]@{ Name = $Name; Type = $Type; TypeDesc = $td; Sid = $Sid; IsDisabled = $Disabled; DefaultDb = $(if ($Type -eq 'R') { $null } else { $Db })
            DefaultLang = $(if ($Type -eq 'R') { $null } else { 'us_english' }); IsFixedRole = $false
            PolicyChecked = $(if ($Type -eq 'S') { $Policy } else { $null }); ExpirationChecked = $(if ($Type -eq 'S') { $false } else { $null })
            PwdHash = (ConvertTo-LSHex $h); PwdDigest = (Get-LSDigest $h); CreateDate = (Get-Date).AddDays(-400); ModifyDate = (Get-Date).AddDays(-30) }
    }
    $fixed = { param($n) [pscustomobject]@{ Name = $n; Type = 'R'; TypeDesc = 'SERVER_ROLE'; Sid = $null; IsDisabled = $false; DefaultDb = $null; DefaultLang = $null; IsFixedRole = $true; PolicyChecked = $null; ExpirationChecked = $null; PwdHash = $null; PwdDigest = $null; CreateDate = $null; ModifyDate = $null } }
    $perm = { param($g, $perm, $state = 'GRANT') [pscustomobject]@{ Grantee = $g; Class = 100; ClassDesc = 'SERVER'; Permission = $perm; State = $state; MajorId = 0; TargetName = $null; Key = "$state|$perm|SERVER|" } }
    $rm = { param($r, $m) [pscustomobject]@{ RoleName = $r; MemberName = $m } }

    $base = @(
        (& $mk 'sa' 'S' '0x01' $true 'master' 'sa-pwd'),
        (& $mk 'app_orders' 'S' '0x5A1C0E6F2B9D4F11A0C3D2E4F5061728' $false 'OrdersDB' 'orders-pwd'),
        (& $mk 'app_reporting' 'S' '0x7E2F1A3B4C5D6E7F8091A2B3C4D5E6F7' $false 'ReportingDB' 'rpt-pwd'),
        (& $mk 'etl_service' 'S' '0x11223344556677889900AABBCCDDEEFF' $false 'master' 'etl-pwd'),
        (& $mk 'monitor_ro' 'S' '0x9988776655443322110FFEEDDCCBBAA0' $false 'master' 'mon-pwd'),
        (& $mk 'legacy_batch' 'S' '0xABCDEF0123456789ABCDEF0123456789' $true 'master' 'legacy-pwd' $false),
        (& $mk 'CORP\svc_orders_app' 'U' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2A1040000'),
        (& $mk 'CORP\SQL_DBA_Team' 'G' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2B2050000'),
        (& $mk 'CORP\Report_Readers' 'G' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2C3060000'),
        (& $mk 'AppOperators' 'R' '0x0B00000000000000000000000000000001'),
        (& $mk 'NT SERVICE\SQLSERVERAGENT' 'U' '0x0106000000000005500000001111'),
        (& $mk '##MS_PolicyEventProcessingLogin##' 'S' '0xDAD97D283865624AB48CD292DCB2BE5F' $true)
    )
    $fixedRoles = @('sysadmin', 'securityadmin', 'serveradmin', 'dbcreator', 'processadmin', 'public') | ForEach-Object { & $fixed $_ }
    $baseRoles = @((& $rm 'sysadmin' 'CORP\SQL_DBA_Team'), (& $rm 'sysadmin' 'sa'), (& $rm 'AppOperators' 'etl_service'), (& $rm 'processadmin' 'AppOperators'), (& $rm 'dbcreator' 'etl_service'))
    $basePerms = @((& $perm 'monitor_ro' 'VIEW SERVER STATE'), (& $perm 'monitor_ro' 'VIEW ANY DEFINITION'), (& $perm 'app_reporting' 'VIEW SERVER STATE'), (& $perm 'legacy_batch' 'CONNECT SQL' 'DENY'), (& $perm 'AppOperators' 'ALTER TRACE'))
    $clone = { param($list) @($list | ForEach-Object { $_.PSObject.Copy() }) }

    $mkSnap = {
        param($Name, $Role, $Mode, $Machine, $Principals, $Roles, $Perms)
        $s = New-LSSnapshot -Name $Name -Role $Role -Mode $Mode
        $s.Connected = $true; $s.ServerName = $Name; $s.MachineName = $Machine; $s.Version = '16.0.4185.3'; $s.Edition = 'Enterprise Edition: Core-based Licensing (64-bit)'
        $s.IsHadrEnabled = $true; $s.CanViewDefinition = $true; $s.CanControlServer = $true; $s.CanViewServerState = $true
        $s.Principals = @($Principals) + $fixedRoles; $s.RoleMembers = @($Roles); $s.Permissions = @($Perms)
        $s.ElapsedMs = Get-Random -Minimum 180 -Maximum 650; $s.QueryCount = 5
        $s
    }

    # Reference / primary
    $p1 = & $clone $base
    $p1 += (& $mk 'SQLPROD01\local_ops' 'U' '0x0105000000000005150000009999888877776666555544440A030000')
    $ref = & $mkSnap 'SQLPROD01' 'Primary' 'Primary' 'SQLPROD01' $p1 $baseRoles $basePerms
    $ref.Databases = @([pscustomobject]@{ Name = 'OrdersDB'; IsAgDb = $true; AgName = 'AG_Orders'; HasAccess = $true }, [pscustomobject]@{ Name = 'ReportingDB'; IsAgDb = $true; AgName = 'AG_Orders'; HasAccess = $true })
    $u = { param($db, $user, $type, $sid) [pscustomobject]@{ Database = $db; IsAgDb = $true; AgName = 'AG_Orders'; User = $user; Type = $type; TypeDesc = @{ S = 'SQL_USER'; U = 'WINDOWS_USER'; G = 'WINDOWS_GROUP' }[$type]; Sid = $sid } }
    $ref.DbUsers = @(
        (& $u 'OrdersDB' 'dbo' 'S' '0x01'),
        (& $u 'OrdersDB' 'app_orders' 'S' '0x5A1C0E6F2B9D4F11A0C3D2E4F5061728'),
        (& $u 'OrdersDB' 'etl_service' 'S' '0x11223344556677889900AABBCCDDEEFF'),
        (& $u 'OrdersDB' 'svc_orders_app' 'U' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2A1040000'),
        (& $u 'OrdersDB' 'CORP\Report_Readers' 'G' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2C3060000'),
        (& $u 'ReportingDB' 'dbo' 'U' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F2FF990000'),
        (& $u 'ReportingDB' 'app_reporting' 'S' '0x7E2F1A3B4C5D6E7F8091A2B3C4D5E6F7'),
        (& $u 'ReportingDB' 'etl_service' 'S' '0x11223344556677889900AABBCCDDEEFF'),
        (& $u 'ReportingDB' 'CORP\jdoe' 'U' '0x010500000000000515000000A1B2C3D4E5F6A7B8C9D0E1F277880000'),
        (& $u 'ReportingDB' 'old_vendor' 'S' '0x0F0E0D0C0B0A09080706050403020100')
    )
    $ref.Warnings.Add('Demo data - not a real server.')

    # Synchronous secondary: SID mismatch, missing ETL login, missing sysadmin, password drift, extra login
    $p2 = & $clone ($base | Where-Object Name -ne 'etl_service')
    ($p2 | Where-Object Name -eq 'app_reporting').Sid = '0x3C4D5E6F708192A3B4C5D6E7F8091A2B'
    $mon = $p2 | Where-Object Name -eq 'monitor_ro'; $h = & $hashOf 'mon-pwd-OLD'; $mon.PwdHash = ConvertTo-LSHex $h; $mon.PwdDigest = Get-LSDigest $h
    $p2 += (& $mk 'temp_vendor' 'S' '0x77777777777777777777777777777777' $false 'master' 'vendor')
    $p2 += (& $mk 'SQLPROD02\local_ops' 'U' '0x0105000000000005150000001111222233334444555566660A030000')
    $r2 = @($baseRoles | Where-Object { -not ($_.RoleName -eq 'sysadmin' -and $_.MemberName -eq 'CORP\SQL_DBA_Team') -and $_.MemberName -ne 'etl_service' })
    $sec = & $mkSnap 'SQLPROD02' 'Target' 'Synchronous' 'SQLPROD02' $p2 $r2 ($basePerms | Where-Object Grantee -ne 'etl_service')

    # Async DR 1: role missing, legacy login enabled
    $p3 = & $clone ($base | Where-Object Name -ne 'AppOperators')
    ($p3 | Where-Object Name -eq 'legacy_batch').IsDisabled = $false
    $r3 = @($baseRoles | Where-Object { $_.RoleName -ne 'AppOperators' -and $_.MemberName -ne 'AppOperators' })
    $dr1 = & $mkSnap 'SQLDR01' 'Target' 'Asynchronous' 'SQLDR01' $p3 $r3 ($basePerms | Where-Object Grantee -ne 'AppOperators')

    # Async DR 2: app login missing, default DB drift, missing permission, extra sysadmin
    $p4 = & $clone ($base | Where-Object Name -ne 'app_orders')
    ($p4 | Where-Object Name -eq 'monitor_ro').DefaultDb = 'tempdb'
    $r4 = @($baseRoles) + @(& $rm 'sysadmin' 'monitor_ro')
    $pe4 = @($basePerms | Where-Object { -not ($_.Grantee -eq 'monitor_ro' -and $_.Permission -eq 'VIEW SERVER STATE') })
    $dr2 = & $mkSnap 'SQLDR02' 'Target' 'Asynchronous' 'SQLDR02' $p4 $r4 $pe4

    return @($ref, $sec, $dr1, $dr2)
}
#endregion

#region ---------- HTML template (self-contained: no CDN, works on locked-down jump hosts) ----------
$script:LSHtmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login Sync Report</title>
<style>
:root{
  --bg:#f5f6f8; --panel:#ffffff; --ink:#1d2330; --muted:#647086; --line:#e2e6ec; --soft:#eef1f5;
  --accent:#2563eb; --ok:#15803d; --ok-bg:#e3f4e8; --warn:#b45309; --warn-bg:#fdf0dc; --bad:#c02626; --bad-bg:#fbe4e4;
  --ren:#c2410c; --ren-bg:#fde9dc; --info:#1d4ed8; --info-bg:#e1eafd; --gray:#6b7280; --gray-bg:#eceef1;
  --crit:#b91c1c; --high:#ea580c; --med:#ca8a04; --low:#2563eb; --inf:#6b7280;
  --sync:#0f766e; --async:#7c3aed; --mono: ui-monospace, "Cascadia Code", Consolas, monospace;
}
@media (prefers-color-scheme: dark){ :root{
  --bg:#0f1218; --panel:#171b23; --ink:#e6e9ef; --muted:#97a1b3; --line:#2a303c; --soft:#1e232d;
  --accent:#60a5fa; --ok:#4ade80; --ok-bg:#12301d; --warn:#fbbf24; --warn-bg:#352812; --bad:#f87171; --bad-bg:#3a1717;
  --ren:#fb923c; --ren-bg:#3a2112; --info:#93c5fd; --info-bg:#172a4a; --gray:#9ca3af; --gray-bg:#252a34;
  --crit:#f87171; --high:#fb923c; --med:#facc15; --low:#60a5fa; --inf:#9ca3af; --sync:#2dd4bf; --async:#c4b5fd; } }
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 "Segoe UI",system-ui,-apple-system,sans-serif}
header{padding:22px 28px 14px;border-bottom:1px solid var(--line);background:var(--panel);display:flex;gap:18px;align-items:flex-end;flex-wrap:wrap}
header h1{margin:0;font-size:22px;letter-spacing:-.01em}
header .sub{color:var(--muted);font-size:13px;margin-top:4px}
.pill{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:999px;font-weight:600;font-size:12px}
.grow{flex:1}
main{padding:20px 28px 60px;max-width:1500px;margin:0 auto}
.topo{display:flex;gap:14px;align-items:stretch;flex-wrap:wrap;margin-bottom:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px 16px;min-width:230px;flex:1 1 230px;position:relative}
.card.ref{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent) inset}
.card h3{margin:0 0 2px;font-size:15px;word-break:break-all}
.card .meta{color:var(--muted);font-size:12px}
.chip{display:inline-block;font-size:11px;font-weight:700;letter-spacing:.04em;padding:2px 7px;border-radius:5px;text-transform:uppercase}
.chip.Primary,.chip.Reference{background:var(--info-bg);color:var(--info)}
.chip.Synchronous{background:color-mix(in srgb,var(--sync) 16%,transparent);color:var(--sync)}
.chip.Asynchronous{background:color-mix(in srgb,var(--async) 16%,transparent);color:var(--async)}
.chip.Standalone{background:var(--gray-bg);color:var(--gray)}
.ringrow{display:flex;gap:14px;align-items:center;margin-top:10px}
.ring{width:64px;height:64px;flex:none}
.counts{display:grid;grid-template-columns:auto auto;gap:2px 10px;font-size:12px}
.counts b{font-variant-numeric:tabular-nums}
.arrow{align-self:center;color:var(--muted);font-size:20px}
.recs{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:6px 16px;margin-bottom:18px}
.rec{display:flex;gap:10px;padding:8px 0;border-bottom:1px solid var(--line);align-items:flex-start}
.rec:last-child{border-bottom:0}
.rec .lvl{flex:none;width:86px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--line);margin-bottom:14px;flex-wrap:wrap}
.tab{padding:9px 14px;border:0;background:none;color:var(--muted);font:inherit;font-weight:600;cursor:pointer;border-bottom:2px solid transparent}
.tab.on{color:var(--ink);border-bottom-color:var(--accent)}
.tab .n{font-weight:500;color:var(--muted);margin-left:4px}
.panel{display:none}.panel.on{display:block}
.toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:10px}
input[type=search],select{background:var(--panel);color:var(--ink);border:1px solid var(--line);border-radius:7px;padding:6px 10px;font:inherit}
input[type=search]{min-width:240px}
label.cb{display:flex;gap:6px;align-items:center;color:var(--muted)}
.tablewrap{overflow:auto;background:var(--panel);border:1px solid var(--line);border-radius:10px;max-height:70vh}
table{border-collapse:separate;border-spacing:0;width:100%}
th,td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{position:sticky;top:0;background:var(--soft);z-index:1;font-size:12px;color:var(--muted);font-weight:600;white-space:nowrap}
th .chip{margin-top:3px}
tbody tr:hover td{background:color-mix(in srgb,var(--accent) 5%,transparent)}
td.name{font-weight:600;white-space:nowrap}
#itab td.name{white-space:normal;min-width:140px;max-width:220px;word-break:break-word}
#itab td:nth-child(5){min-width:340px}
#itab td:nth-child(6){min-width:260px}
td.c{text-align:center;white-space:nowrap}
.st{display:inline-block;min-width:54px;padding:3px 8px;border-radius:6px;font-weight:700;font-size:12px;cursor:pointer;font-family:var(--mono)}
.st.OK{background:var(--ok-bg);color:var(--ok)}
.st.DIFF{background:var(--warn-bg);color:var(--warn)}
.st.SID,.st.MISSING,.st.OWNER{background:var(--bad-bg);color:var(--bad)}
.st.RENAMED{background:var(--ren-bg);color:var(--ren)}
.st.INFO,.st.REF,.st.GROUP{background:var(--info-bg);color:var(--info)}
.st.EXTRA{background:var(--gray-bg);color:var(--ink)}
.st.ABSENT,.st.ERR{background:transparent;color:var(--gray)}
.sev{font-weight:700;font-size:12px}
.sev.Critical{color:var(--crit)}.sev.High{color:var(--high)}.sev.Medium{color:var(--med)}.sev.Low{color:var(--low)}.sev.Info{color:var(--inf)}
.sev.Good{color:var(--ok)}.sev.Automation{color:var(--accent)}
.sevchip{border:1px solid var(--line);background:var(--panel);color:var(--ink);border-radius:999px;padding:4px 11px;cursor:pointer;font:inherit;font-size:12px}
.sevchip.on{background:var(--ink);color:var(--panel)}
.legend{display:flex;gap:12px;flex-wrap:wrap;color:var(--muted);font-size:12px;margin:6px 0 10px}
.detail{margin-top:12px;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:12px 16px;display:none}
.detail.on{display:block}
.detail ul{margin:6px 0 0 18px;padding:0}
pre{background:var(--soft);border:1px solid var(--line);border-radius:10px;padding:14px;overflow:auto;max-height:65vh;font:12.5px/1.5 var(--mono);white-space:pre}
.btn{background:var(--accent);color:#fff;border:0;border-radius:7px;padding:7px 13px;font:inherit;font-weight:600;cursor:pointer}
.btn.ghost{background:var(--panel);color:var(--ink);border:1px solid var(--line)}
.note{color:var(--muted);font-size:12.5px}
.guard{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:12px}
.guard .card{min-width:0}
.guard li{margin:4px 0}
.muted{color:var(--muted)}
.mono{font-family:var(--mono);font-size:12px}
.empty{padding:30px;text-align:center;color:var(--muted)}
footer{color:var(--muted);font-size:12px;text-align:center;padding:20px}
</style>
</head>
<body>
<header>
  <div class="grow">
    <h1>Login Sync Report</h1>
    <div class="sub" id="sub"></div>
  </div>
  <div id="overall"></div>
</header>
<main>
  <section class="topo" id="topo" aria-label="Servers"></section>
  <section class="recs" id="recs" aria-label="Recommendations"></section>

  <nav class="tabs" id="tabs">
    <button class="tab on" data-p="matrix">Login matrix<span class="n" id="n-matrix"></span></button>
    <button class="tab" data-p="issues">Findings<span class="n" id="n-issues"></span></button>
    <button class="tab" data-p="orphans">Orphan risk<span class="n" id="n-orphans"></span></button>
    <button class="tab" data-p="fix">Remediation T-SQL</button>
    <button class="tab" data-p="guard">Run &amp; guardrails</button>
  </nav>

  <div class="panel on" id="p-matrix">
    <div class="toolbar">
      <input type="search" id="mq" placeholder="Filter principals..." aria-label="Filter principals">
      <select id="mtype" aria-label="Type"><option value="">All types</option><option value="S">SQL logins</option><option value="U">Windows logins</option><option value="G">Windows groups</option><option value="R">Server roles</option></select>
      <label class="cb"><input type="checkbox" id="mdrift" checked> Only rows with drift</label>
      <span class="grow"></span><span class="note">Click any cell for details</span>
    </div>
    <div class="legend" id="legend"></div>
    <div class="tablewrap"><table id="mtab"></table></div>
    <div class="detail" id="mdetail"></div>
  </div>

  <div class="panel" id="p-issues">
    <div class="toolbar" id="sevbar"></div>
    <div class="toolbar">
      <select id="itarget" aria-label="Server"><option value="">All servers</option></select>
      <select id="icat" aria-label="Category"><option value="">All categories</option></select>
      <input type="search" id="iq" placeholder="Search findings..." aria-label="Search findings">
    </div>
    <div class="tablewrap"><table id="itab"></table></div>
  </div>

  <div class="panel" id="p-orphans">
    <div class="toolbar">
      <input type="search" id="oq" placeholder="Filter database or user..." aria-label="Filter">
      <label class="cb"><input type="checkbox" id="orisk" checked> Only users at risk</label>
      <span class="grow"></span><span class="note" id="onote"></span>
    </div>
    <div class="tablewrap"><table id="otab"></table></div>
  </div>

  <div class="panel" id="p-fix">
    <div class="toolbar" id="fixbar"></div>
    <p class="note">Generated, never executed. Password hashes are redacted here; the full statements are in the <span class="mono">remediation_&lt;server&gt;.sql</span> file next to this report. Each script aborts unless <span class="mono">@@SERVERNAME</span> matches its target.</p>
    <pre id="fixpre"></pre>
  </div>

  <div class="panel" id="p-guard"><div class="guard" id="guard"></div></div>
</main>
<footer id="foot"></footer>

<script type="application/json" id="lsdata">/*__LSDATA__*/</script>
<script>
(function(){
const D = JSON.parse(document.getElementById('lsdata').textContent);
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const arr = v => v == null ? [] : (Array.isArray(v) ? v : [v]);
const servers = arr(D.Servers), matrix = arr(D.Matrix), issues = arr(D.Issues), orphans = arr(D.Orphans), recs = arr(D.Recommendations);
const M = D.Meta || {};
const RANK = {Critical:4, High:3, Medium:2, Low:1, Info:0};
const LABEL = {OK:'\u2713', DIFF:'\u0394', SID:'SID', MISSING:'\u2717', RENAMED:'\u21aa', INFO:'i', EXTRA:'+', ERR:'?', REF:'\u25cf', ABSENT:'\u2013', GROUP:'grp', OWNER:'own'};
const DESC = {REF:'present on reference', OK:'identical', DIFF:'exists, properties/roles/permissions differ', SID:'SID (or type) mismatch', MISSING:'missing', RENAMED:'same SID, different name', EXTRA:'only on target', INFO:'host-local / informational', ERR:'not reachable', ABSENT:'not present'};
const roleOf = s => (s.Role === 'Primary' || s.Role === 'Reference') ? s.Role : s.Mode;

const worst = M.WorstSeverity || 'None';
const ov = $('#overall');
const sevColor = {Critical:'var(--crit)',High:'var(--high)',Medium:'var(--med)',Low:'var(--low)',Info:'var(--inf)',None:'var(--ok)'};
ov.innerHTML = '<span class="pill" style="background:color-mix(in srgb,'+sevColor[worst]+' 15%,transparent);color:'+sevColor[worst]+'">' +
  (worst === 'None' || worst === 'Info' ? 'All in sync' : 'Worst finding: ' + esc(worst)) + '</span>';
$('#sub').innerHTML = 'Reference <b>' + esc(M.ReferenceServerName || M.Reference) + '</b>' + (M.AgName ? ' &middot; AG <b>' + esc(M.AgName) + '</b>' : '') +
  ' &middot; ' + esc(M.GeneratedAt) + ' &middot; by ' + esc(M.RunBy) + (M.Demo ? ' &middot; <b>DEMO DATA</b>' : '');

function ring(score){
  if (score == null) return '<svg class="ring" viewBox="0 0 64 64"><circle cx="32" cy="32" r="26" fill="none" stroke="var(--line)" stroke-width="7"/><text x="32" y="37" text-anchor="middle" font-size="14" fill="var(--muted)" font-weight="700">\u2014</text></svg>';
  const c = 2*Math.PI*26, col = score >= 100 ? 'var(--ok)' : score >= 90 ? 'var(--med)' : 'var(--bad)';
  return '<svg class="ring" viewBox="0 0 64 64" role="img" aria-label="'+score+'% in sync"><circle cx="32" cy="32" r="26" fill="none" stroke="var(--line)" stroke-width="7"/>' +
    '<circle cx="32" cy="32" r="26" fill="none" stroke="'+col+'" stroke-width="7" stroke-linecap="round" stroke-dasharray="'+(c*score/100)+' '+c+'" transform="rotate(-90 32 32)"/>' +
    '<text x="32" y="37" text-anchor="middle" font-size="14" fill="var(--ink)" font-weight="700">'+score+'%</text></svg>';
}
const topo = $('#topo');
servers.forEach((s, i) => {
  const isRef = (s.Role === 'Primary' || s.Role === 'Reference');
  if (i === 1) topo.insertAdjacentHTML('beforeend', '<div class="arrow" aria-hidden="true">\u279c</div>');
  let body;
  if (isRef) {
    body = '<div class="meta" style="margin-top:10px">Source of truth &middot; ' + esc(s.InScope) + ' principals in scope</div>' +
      '<div class="meta">' + esc(M.DatabasesScanned) + ' databases / ' + esc(M.DbUsersScanned) + ' users scanned for orphan risk</div>';
  } else if (!s.Connected) {
    body = '<div class="meta" style="margin-top:10px;color:var(--bad)">Not reachable: ' + esc(s.Error) + '</div>';
  } else {
    body = '<div class="ringrow">' + ring(s.Score) + '<div class="counts">' +
      '<span class="sev Critical">Critical</span><b>'+s.Critical+'</b><span class="sev High">High</span><b>'+s.High+'</b>' +
      '<span class="sev Medium">Medium</span><b>'+s.Medium+'</b><span class="muted">Orphan risk</span><b>'+s.OrphanRisk+'</b></div></div>' +
      '<div class="meta">' + s.InSync + ' of ' + s.InScope + ' principals identical</div>';
  }
  topo.insertAdjacentHTML('beforeend', '<div class="card'+(isRef?' ref':'')+'"><span class="chip '+esc(roleOf(s))+'">'+esc(isRef ? s.Role : s.Mode)+'</span>' +
    '<h3>'+esc(s.Name)+'</h3><div class="meta">'+esc(s.Version || '')+'</div>' + body + '</div>');
});

$('#recs').innerHTML = recs.map(r => '<div class="rec"><span class="lvl sev '+esc(r.Level)+'">'+esc(r.Level)+'</span><span>'+esc(r.Text)+'</span></div>').join('');

// tabs
document.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('on', x === b));
  document.querySelectorAll('.panel').forEach(p => p.classList.toggle('on', p.id === 'p-' + b.dataset.p));
}));

// matrix
const names = servers.map(s => s.Name);
$('#legend').innerHTML = ['REF','OK','DIFF','SID','MISSING','RENAMED','EXTRA','INFO','ABSENT','ERR'].map(k => '<span><span class="st '+k+'" style="min-width:34px;cursor:default">'+LABEL[k]+'</span> '+DESC[k]+'</span>').join('');
const cellOf = (row, n) => (row.Cells && row.Cells[n]) || {S:'ABSENT', N:['Not present']};
const isDrift = row => !row.HostLocal && names.some(n => { const s = cellOf(row, n).S; return !(s === 'OK' || s === 'REF' || s === 'INFO' || (s === 'ERR')); });
function renderMatrix(){
  const q = $('#mq').value.toLowerCase(), t = $('#mtype').value, drift = $('#mdrift').checked;
  const rows = matrix.filter(r => (!q || r.Name.toLowerCase().includes(q)) && (!t || r.Type === t) && (!drift || isDrift(r)));
  $('#n-matrix').textContent = matrix.filter(isDrift).length + ' drift / ' + matrix.length;
  let h = '<thead><tr><th>Principal</th><th>Type</th>' + servers.map(s => '<th>'+esc(s.Name)+'<br><span class="chip '+esc(roleOf(s))+'">'+esc((s.Role==='Primary'||s.Role==='Reference')?s.Role:s.Mode)+'</span></th>').join('') + '</tr></thead><tbody>';
  if (!rows.length) h += '<tr><td colspan="'+(names.length+2)+'" class="empty">'+(drift ? 'No drift - every principal matches the reference.' : 'No principals match the filter.')+'</td></tr>';
  rows.forEach((r, i) => {
    h += '<tr><td class="name">'+esc(r.Name)+(r.Disabled ? ' <span class="muted" title="disabled on reference">(disabled)</span>' : '')+'</td><td class="muted">'+esc((r.TypeDesc||'').replace('_',' ').toLowerCase())+'</td>';
    names.forEach(n => { const c = cellOf(r, n); h += '<td class="c"><span class="st '+esc(c.S)+'" data-r="'+esc(r.Name)+'" data-s="'+esc(n)+'" title="'+esc(arr(c.N).join('\n'))+'">'+LABEL[c.S]+'</span></td>'; });
    h += '</tr>';
  });
  $('#mtab').innerHTML = h + '</tbody>';
}
$('#mtab').addEventListener('click', e => {
  const el = e.target.closest('.st'); if (!el) return;
  const row = matrix.find(r => r.Name === el.dataset.r); const c = cellOf(row, el.dataset.s);
  const rel = issues.filter(i => i.Target === el.dataset.s && i.Object === row.Name);
  $('#mdetail').className = 'detail on';
  $('#mdetail').innerHTML = '<b>'+esc(row.Name)+'</b> on <b>'+esc(el.dataset.s)+'</b> &mdash; <span class="st '+esc(c.S)+'">'+LABEL[c.S]+'</span> '+esc(DESC[c.S]||'') +
    '<ul>' + arr(c.N).map(n => '<li>'+esc(n)+'</li>').join('') + '</ul>' +
    (rel.length ? '<div style="margin-top:8px"><b>Suggested action</b><ul>' + rel.map(i => '<li><span class="sev '+esc(i.Severity)+'">'+esc(i.Severity)+'</span> '+esc(i.Suggestion)+'</li>').join('') + '</ul></div>' : '');
});
['#mq','#mtype','#mdrift'].forEach(s => $(s).addEventListener('input', renderMatrix));
renderMatrix();

// findings
const sevOn = new Set(['Critical','High','Medium','Low']);
$('#sevbar').innerHTML = ['Critical','High','Medium','Low','Info'].map(s => '<button class="sevchip'+(sevOn.has(s)?' on':'')+'" data-s="'+s+'"><span class="sev '+s+'">\u25cf</span> '+s+' ('+issues.filter(i=>i.Severity===s).length+')</button>').join('');
$('#sevbar').addEventListener('click', e => { const b = e.target.closest('.sevchip'); if (!b) return; const s = b.dataset.s; sevOn.has(s) ? sevOn.delete(s) : sevOn.add(s); b.classList.toggle('on'); renderIssues(); });
[...new Set(issues.map(i => i.Target))].forEach(t => $('#itarget').insertAdjacentHTML('beforeend', '<option>'+esc(t)+'</option>'));
[...new Set(issues.map(i => i.Category))].sort().forEach(t => $('#icat').insertAdjacentHTML('beforeend', '<option>'+esc(t)+'</option>'));
function renderIssues(){
  const t = $('#itarget').value, c = $('#icat').value, q = $('#iq').value.toLowerCase();
  const rows = issues.filter(i => sevOn.has(i.Severity) && (!t || i.Target === t) && (!c || i.Category === c) &&
     (!q || (i.Object + ' ' + i.Detail + ' ' + i.Category).toLowerCase().includes(q)));
  $('#n-issues').textContent = issues.filter(i => i.Severity !== 'Info').length;
  let h = '<thead><tr><th>Severity</th><th>Server</th><th>Principal</th><th>Category</th><th>Detail</th><th>Suggestion</th></tr></thead><tbody>';
  if (!rows.length) h += '<tr><td colspan="6" class="empty">No findings for this filter.</td></tr>';
  rows.forEach(i => { h += '<tr><td><span class="sev '+esc(i.Severity)+'">'+esc(i.Severity)+'</span></td><td style="white-space:nowrap">'+esc(i.Target)+'<br><span class="chip '+esc(i.Mode)+'">'+esc(i.Mode)+'</span></td><td class="name">'+esc(i.Object)+'<div class="muted" style="font-weight:400">'+esc(i.ObjectType)+'</div></td><td>'+esc(i.Category)+'</td><td>'+esc(i.Detail)+'</td><td class="muted">'+esc(i.Suggestion)+'</td></tr>'; });
  $('#itab').innerHTML = h + '</tbody>';
}
['#itarget','#icat','#iq'].forEach(s => $(s).addEventListener('input', renderIssues));
renderIssues();

// orphans
function renderOrphans(){
  const q = $('#oq').value.toLowerCase(), risk = $('#orisk').checked;
  const rows = orphans.filter(o => (!q || (o.Database + ' ' + o.User + ' ' + (o.Login||'')).toLowerCase().includes(q)) && (!risk || o.AtRisk));
  $('#n-orphans').textContent = orphans.filter(o => o.AtRisk).length;
  $('#onote').textContent = (M.DbUsersScanned || 0) + ' users in ' + (M.DatabasesScanned || 0) + ' reference databases, matched by SID against every server';
  let h = '<thead><tr><th>Database</th><th>User</th><th>Reference login</th>' + servers.map(s => '<th>'+esc(s.Name)+'</th>').join('') + '</tr></thead><tbody>';
  if (!rows.length) h += '<tr><td colspan="'+(names.length+3)+'" class="empty">'+(orphans.length ? 'No database user is at risk.' : 'No database users were scanned (scope None, or no accessible databases).')+'</td></tr>';
  rows.forEach(o => {
    h += '<tr><td>'+esc(o.Database)+'</td><td class="name">'+esc(o.User)+'<div class="muted" style="font-weight:400">'+esc(o.Type)+'</div></td><td>'+(o.Login ? esc(o.Login) : '<span style="color:var(--bad)">none</span>')+'</td>';
    names.forEach(n => { const c = (o.Cells && o.Cells[n]) || {S:'ERR',N:[]}; h += '<td class="c"><span class="st '+esc(c.S)+'" style="cursor:default" title="'+esc(arr(c.N).join('\n'))+'">'+(LABEL[c.S]||esc(c.S))+'</span></td>'; });
    h += '</tr>';
  });
  $('#otab').innerHTML = h + '</tbody>';
}
['#oq','#orisk'].forEach(s => $(s).addEventListener('input', renderOrphans));
renderOrphans();

// remediation
const scripts = D.Scripts || {};
const keys = Object.keys(scripts);
let cur = keys[0];
function renderFix(){
  $('#fixbar').innerHTML = keys.map(k => '<button class="btn'+(k===cur?'':' ghost')+'" data-k="'+esc(k)+'">'+esc(k)+'</button>').join('') +
    '<span class="grow"></span>' + (keys.length ? '<button class="btn ghost" id="copy">Copy</button>' : '');
  $('#fixpre').textContent = keys.length ? scripts[cur] : 'No reachable targets - nothing to script.';
}
$('#fixbar').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  if (b.id === 'copy') {
    const txt = $('#fixpre').textContent;
    const done = () => { b.textContent = 'Copied'; setTimeout(() => b.textContent = 'Copy', 1400); };
    if (navigator.clipboard) navigator.clipboard.writeText(txt).then(done, () => {}); return;
  }
  cur = b.dataset.k; renderFix();
});
renderFix();

// guardrails
const g = $('#guard');
g.innerHTML = '<div class="card"><h3>What this run did</h3><ul>' +
  '<li>Read-only <b>SELECT</b> on catalog views only (sys.server_principals, sys.sql_logins, sys.server_role_members, sys.server_permissions, sys.databases, &lt;db&gt;.sys.database_principals, AG catalog views).</li>' +
  '<li>Every statement passed a guard refusing anything but a single SELECT (no DML/DDL/EXEC/GRANT/DBCC/xp_/sp_).</li>' +
  '<li>Session: <span class="mono">LOCK_TIMEOUT '+esc(M.LockTimeoutMs)+'</span>, <span class="mono">DEADLOCK_PRIORITY LOW</span>, <span class="mono">READ UNCOMMITTED</span>; query timeout '+esc(M.QueryTimeoutSec)+'s.</li>' +
  '<li>One connection at a time, pooling off, closed immediately. Application name <span class="mono">LoginSyncCheck (read-only)</span>.</li>' +
  '<li>Remediation scripts written to files only; nothing executed. Hashes never stored in this HTML.</li></ul></div>' +
  '<div class="card"><h3>Scope</h3><ul><li>Database scope: <b>'+esc(M.DatabaseScope)+'</b> &mdash; '+esc(M.DatabasesScanned)+' databases scanned on the reference.</li>' +
  '<li>Excluded patterns: <span class="mono">'+esc(arr(M.ExcludeLogin).join('  '))+'</span></li>' +
  '<li>Password comparison: '+(M.PasswordCompare ? 'on (SHA-256 digest of the hash)' : 'off')+'</li>' +
  '<li>Total run time: '+esc(M.DurationSec)+' s from '+esc(M.RunFrom)+'</li></ul></div>' +
  '<div class="card"><h3>Per server</h3><table><thead><tr><th>Server</th><th>Queries</th><th>ms</th><th>VIEW ANY DEF</th><th>CONTROL SERVER</th></tr></thead><tbody>' +
  servers.map(s => '<tr><td>'+esc(s.Name)+'</td><td>'+esc(s.QueryCount)+'</td><td>'+esc(s.ElapsedMs)+'</td><td>'+(s.CanViewDefinition?'\u2713':'\u2717')+'</td><td>'+(s.CanControlServer?'\u2713':'\u2717')+'</td></tr>').join('') +
  '</tbody></table></div>' +
  '<div class="card"><h3>Warnings</h3>' + (arr(M.Warnings).length ? '<ul>' + arr(M.Warnings).map(w => '<li>'+esc(w)+'</li>').join('') + '</ul>' : '<p class="muted">None.</p>') + '</div>';
$('#foot').textContent = 'LoginSync v' + (M.Version || '') + ' \u00b7 read-only report \u00b7 ' + (M.GeneratedAt || '');
})();
</script>
</body>
</html>
'@
#endregion

#region ---------- background worker support ----------
function New-LSWorkerSessionState {
    # Copies this session's engine functions and constants into a new runspace, so the GUI's background
    # worker never depends on a file path (works even when the script was pasted or run as an ISE selection).
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    foreach ($f in @(Get-ChildItem Function: | Where-Object { $_.Name -match '^\w+-LS[A-Z]' })) {
        $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($f.Name, $f.Definition)))
    }
    foreach ($v in @(Get-Variable -Scope Script | Where-Object { $_.Name -like 'LS*' })) {
        $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($v.Name, $v.Value, '')))
    }
    return $iss
}
#endregion

# ---------------------------------------------------------------------------------------------------
# Dot-sourced (e.g. by the WinForms front end): expose the functions and stop here.
if ($MyInvocation.InvocationName -eq '.' -or $MyInvocation.Line -match '^\s*\.\s') { return }

# =============================================================================================
#  Mode: GUI (no server parameters, or -Gui) or headless (anything else - scheduled tasks, SQL Agent)
# =============================================================================================
$script:SelfPath = $PSCommandPath
if (-not $script:SelfPath -and $MyInvocation.MyCommand.Path) { $script:SelfPath = $MyInvocation.MyCommand.Path }
$onWindows = ($PSVersionTable.PSEdition -ne 'Core') -or $IsWindows
$noServerArgs = ($PSCmdlet.ParameterSetName -eq 'Manual' -and -not $ReferenceServer -and -not $TargetServer -and -not $AsyncTarget)
$launchGui = $Gui -or ($noServerArgs -and $onWindows -and [Environment]::UserInteractive)

if (-not $launchGui) {
# ======================================= HEADLESS =======================================
try {
    if (-not $OutputPath) {
        # Documents is empty for some service accounts (SQL Agent, Task Scheduler as SYSTEM): fall back to TEMP.
        $docs = [Environment]::GetFolderPath('MyDocuments')
        $root = if ($docs) { $docs } elseif ($env:TEMP) { $env:TEMP } else { (Get-Location).Path }
        $OutputPath = Join-Path $root 'LoginSyncReports'
    }
    $targets = @()
    $ref = $ReferenceServer
    $refRole = 'Reference'
    $agName = $null
    $common = @{ Credential = $Credential; Encrypt = $Encrypt; TrustServerCertificate = $TrustServerCertificate; LockTimeoutMs = $LockTimeoutMs; QueryTimeoutSec = $QueryTimeoutSec }

    if ($PSCmdlet.ParameterSetName -eq 'AG') {
        Write-LSLog "Discovering availability group topology from $DiscoverAgFrom"
        $topo = @(Get-LSAgTopology -Server $DiscoverAgFrom @common)
        if (-not $topo.Count) { throw "$DiscoverAgFrom hosts no availability groups. Use -ReferenceServer/-TargetServer instead." }
        $ags = @($topo | Select-Object -ExpandProperty AgName -Unique)
        if ($AvailabilityGroup) { $agName = $AvailabilityGroup }
        elseif ($ags.Count -eq 1) { $agName = $ags[0] }
        else { throw "Multiple availability groups found ($($ags -join ', ')). Pick one with -AvailabilityGroup." }
        $replicas = @($topo | Where-Object AgName -eq $agName)
        if (-not $replicas.Count) { throw "Availability group '$agName' not found on $DiscoverAgFrom." }
        $primary = $replicas | Where-Object IsPrimary | Select-Object -First 1
        if (-not $primary) { throw "Could not determine the primary replica of '$agName' (needs VIEW SERVER STATE)." }
        $ref = $primary.Replica; $refRole = 'Primary'
        foreach ($r in $replicas) {
            if ($r.IsPrimary) { continue }
            if ($r.Mode -eq 'ConfigurationOnly') { Write-LSLog "  skipping configuration-only replica $($r.Replica)" 'Warn'; continue }
            if ($SyncReplicasOnly -and $r.Mode -ne 'Synchronous') { continue }
            $targets += [pscustomobject]@{ Name = $r.Replica; Mode = $r.Mode }
        }
        Write-LSLog ("  AG {0}: primary {1}; targets {2}" -f $agName, $ref, (($targets | ForEach-Object { "$($_.Name) [$($_.Mode)]" }) -join ', '))
        if (-not $targets.Count) { throw 'No secondary replicas to compare.' }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Manual') {
        if (-not $ReferenceServer -or (-not $TargetServer -and -not $AsyncTarget)) {
            throw 'Specify -ReferenceServer with -TargetServer/-AsyncTarget, or -DiscoverAgFrom, or -Demo (or run without parameters / with -Gui for the window).'
        }
        foreach ($t in @($TargetServer)) { if ($t) { $targets += [pscustomobject]@{ Name = $t; Mode = $TargetMode } } }
        foreach ($t in @($AsyncTarget)) { if ($t) { $targets += [pscustomobject]@{ Name = $t; Mode = 'Asynchronous' } } }
    }

    $result = Invoke-LSLoginSync -ReferenceServer $ref -ReferenceRole $refRole -Targets $targets -AgName $agName -DatabaseScope $DatabaseScope `
        -ExcludeLogin $ExcludeLogin -SkipPasswordCompare:$SkipPasswordCompare -MaxDatabases $MaxDatabases -Demo:$Demo @common
    $files = Export-LSReport -Result $result -OutputPath $OutputPath

    Write-Host ''
    Write-Host ($result.Servers | Format-Table Name, Role, Mode, Connected, Score, Critical, High, Medium, Low, OrphanRisk -AutoSize | Out-String -Width 220)
    Write-Host "Report : $($files.Html)" -ForegroundColor Cyan
    foreach ($f in $files.Sql) { Write-Host "Script : $f" -ForegroundColor Cyan }
    if (@($files.Sql).Count) { Write-Host 'The .sql files may contain password hashes - store securely, delete after use.' -ForegroundColor Yellow }

    if (-not $NoOpen) { try { Invoke-Item $files.Html } catch { Write-LSLog "Could not open browser: $($_.Exception.Message)" 'Warn' } }
    if ($PassThru) { $result | Add-Member -NotePropertyName Files -NotePropertyValue $files -PassThru }

    if ($FailOnSeverity -ne 'None') {
        $hit = @($result.Issues | Where-Object { $script:LSSevRank[$_.Severity] -ge $script:LSSevRank[$FailOnSeverity] })
        if ($hit.Count) { Write-LSLog "$($hit.Count) finding(s) at or above $FailOnSeverity - exit code 2" 'Warn'; exit 2 }
    }
}
catch {
    Write-LSLog $_.Exception.Message 'Error'
    exit 1
}
}
else {
# ========================================= GUI ==========================================
if (-not $onWindows) { Write-Error 'The GUI needs Windows (WinForms). Pass server parameters to run headless.'; return }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if (-not $script:SelfPath -and (Get-Variable -Name psISE -ErrorAction SilentlyContinue) -and $psISE.CurrentFile -and -not $psISE.CurrentFile.IsUntitled) { $script:SelfPath = $psISE.CurrentFile.FullPath }
# WinForms (clipboard, dialogs) needs a single-threaded apartment (ISE and powershell.exe already are).
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA' -and $script:SelfPath) {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$script:SelfPath`"", '-Gui')
    return
}
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ---------- look & feel ----------
function C([string]$Hex) { return [System.Drawing.ColorTranslator]::FromHtml($Hex) }
$script:Font = New-Object System.Drawing.Font('Segoe UI', 9)
$script:FontB = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:FontH = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$script:Mono = New-Object System.Drawing.Font('Consolas', 9.5)
$script:Accent = C '#2563EB'

$script:StatusStyle = @{
    OK = @('#E3F4E8', '#15803D'); DIFF = @('#FDF0DC', '#B45309'); SID = @('#FBE4E4', '#C02626'); MISSING = @('#FBE4E4', '#C02626')
    OWNER = @('#FBE4E4', '#C02626'); RENAMED = @('#FDE9DC', '#C2410C'); INFO = @('#E1EAFD', '#1D4ED8'); REF = @('#E1EAFD', '#1D4ED8')
    GROUP = @('#E1EAFD', '#1D4ED8'); EXTRA = @('#ECEEF1', '#1D2330'); ABSENT = @('#FFFFFF', '#9CA3AF'); ERR = @('#FFFFFF', '#9CA3AF')
}
$script:StatusText = @{
    OK = "$([char]0x2713) OK"; DIFF = "$([char]0x0394) diff"; SID = 'SID'; MISSING = "$([char]0x2717) missing"; RENAMED = "$([char]0x21AA) renamed"
    INFO = 'i  info'; REF = "$([char]0x25CF) reference"; EXTRA = '+ only here'; ABSENT = [string][char]0x2013; ERR = '? n/a'; GROUP = 'via group?'; OWNER = 'owner'
}
$script:SevColor = @{ Critical = '#B91C1C'; High = '#EA580C'; Medium = '#CA8A04'; Low = '#2563EB'; Info = '#6B7280'; Good = '#15803D'; Automation = '#2563EB'; None = '#15803D' }
$script:SevRank = @{ Info = 0; Low = 1; Medium = 2; High = 3; Critical = 4 }

function New-Ctl {
    param([string]$Type, [hashtable]$Props = @{})
    $c = New-Object "System.Windows.Forms.$Type"
    # Deterministic order: ranges/modes first, dependent values (Value, Height, SelectedIndex) last.
    $first = @('Minimum', 'Maximum', 'Multiline', 'DropDownStyle', 'Style', 'AutoSize', 'Dock')
    $last = @('Value', 'Height', 'Width', 'Size', 'Text', 'SelectedIndex', 'Checked')
    $keys = @($first | Where-Object { $Props.ContainsKey($_) }) + @($Props.Keys | Where-Object { $_ -notin $first -and $_ -notin $last }) + @($last | Where-Object { $Props.ContainsKey($_) })
    foreach ($k in $keys) { $c.$k = $Props[$k] }
    return $c
}

function New-Grid {
    $g = New-Ctl DataGridView @{
        Dock = 'Fill'; ReadOnly = $true; AllowUserToAddRows = $false; AllowUserToDeleteRows = $false; AllowUserToResizeRows = $false
        RowHeadersVisible = $false; SelectionMode = 'CellSelect'; BackgroundColor = [System.Drawing.Color]::White; BorderStyle = 'None'
        AutoSizeColumnsMode = 'Fill'; EnableHeadersVisualStyles = $false; ColumnHeadersHeightSizeMode = 'AutoSize'; GridColor = (C '#E2E6EC')
        CellBorderStyle = 'SingleHorizontal'
    }
    $g.ColumnHeadersDefaultCellStyle.BackColor = C '#EEF1F5'
    $g.ColumnHeadersDefaultCellStyle.Font = $script:FontB
    $g.ColumnHeadersDefaultCellStyle.WrapMode = 'True'
    $g.DefaultCellStyle.SelectionBackColor = C '#CFE0FF'
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
    $pi = $g.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
    $pi.SetValue($g, $true, $null)
    $g.Add_DataError({ $_.ThrowException = $false })
    return $g
}

function New-Flow {
    param([int]$Height = 36)
    return New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = $Height; WrapContents = $false; Padding = (New-Object System.Windows.Forms.Padding(4, 4, 4, 0)) }
}

function New-Label { param([string]$Text, [int]$Top = 7) return New-Ctl Label @{ Text = $Text; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(6, $Top, 3, 0)) } }
function New-Button { param([string]$Text) return New-Ctl Button @{ Text = $Text; AutoSize = $true; Height = 28; Padding = (New-Object System.Windows.Forms.Padding(6, 0, 6, 0)); FlatStyle = 'System' } }
#endregion

#region ---------- form layout ----------
$ui = @{}
$form = New-Ctl Form @{
    Text = 'Login Sync Check  -  read-only'; Size = (New-Object System.Drawing.Size(1340, 900)); MinimumSize = (New-Object System.Drawing.Size(1100, 720))
    StartPosition = 'CenterScreen'; Font = $script:Font; BackColor = (C '#F5F6F8'); AutoScaleMode = 'Dpi'
}

# ----- tabs (fill) -----
$tabs = New-Ctl TabControl @{ Dock = 'Fill'; Padding = (New-Object System.Drawing.Point(14, 5)) }
$ui.Tabs = $tabs
function New-Tab { param([string]$Text) $t = New-Ctl TabPage @{ Text = $Text; BackColor = [System.Drawing.Color]::White; Padding = (New-Object System.Windows.Forms.Padding(6)) }; [void]$tabs.TabPages.Add($t); return $t }
$tabOverview = New-Tab 'Overview'
$tabMatrix = New-Tab 'Login matrix'
$tabIssues = New-Tab 'Findings'
$tabOrph = New-Tab 'Orphan risk'
$tabSql = New-Tab 'Remediation T-SQL'
$tabAuto = New-Tab 'Automate'
$tabLog = New-Tab 'Log'

# Overview
$split = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal'; SplitterWidth = 6 }
$ui.GridSum = New-Grid
$ui.GridSum.SelectionMode = 'FullRowSelect'
$ui.RtbRec = New-Ctl RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; BorderStyle = 'None'; BackColor = [System.Drawing.Color]::White; Font = $script:Font }
$recHead = New-Ctl Label @{ Text = 'Recommendations'; Dock = 'Top'; Height = 26; Font = $script:FontH }
$split.Panel1.Controls.Add($ui.GridSum)
$split.Panel2.Controls.Add($ui.RtbRec)
$split.Panel2.Controls.Add($recHead)
$ui.Welcome = New-Ctl Label @{ Dock = 'Fill'; TextAlign = 'MiddleCenter'; ForeColor = (C '#647086'); Font = $script:FontH
    Text = "Enter a reference server and targets (or Discover AG replicas), then Run comparison.`n`nEverything is read-only: catalog SELECTs only, LOCK_TIMEOUT + DEADLOCK_PRIORITY LOW, one connection at a time.`nTry 'Load demo data' to see the report without touching a server." }
$tabOverview.Controls.Add($ui.Welcome)

# Matrix
$ui.GridMatrix = New-Grid
$ui.GridMatrix.AutoSizeColumnsMode = 'DisplayedCells'
$ui.TxtMDetail = New-Ctl TextBox @{ Dock = 'Bottom'; Height = 110; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; BackColor = (C '#F8F9FB'); Font = $script:Font; Text = 'Click a cell to see what differs and the suggested action.' }
$fm = New-Flow
$ui.TxtMFilter = New-Ctl TextBox @{ Width = 220 }
$ui.ChkMDrift = New-Ctl CheckBox @{ Text = 'Only rows with drift'; Checked = $true; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(12, 6, 3, 0)) }
$ui.CmbMType = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 150 }
[void]$ui.CmbMType.Items.AddRange(@('All types', 'SQL logins', 'Windows logins', 'Windows groups', 'Server roles')); $ui.CmbMType.SelectedIndex = 0
$fm.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'Filter:'), $ui.TxtMFilter, $ui.CmbMType, $ui.ChkMDrift, (New-Label "   $([char]0x2713) identical   $([char]0x0394) differs   SID mismatch   $([char]0x2717) missing   $([char]0x21AA) renamed   + only on target   i host-local")))
$tabMatrix.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.GridMatrix, $ui.TxtMDetail, $fm))

# Findings
$ui.GridIssues = New-Grid
$ui.GridIssues.AutoSizeRowsMode = 'AllCells'
$fi = New-Flow
$ui.CmbSev = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 110 }
[void]$ui.CmbSev.Items.AddRange(@('Info', 'Low', 'Medium', 'High', 'Critical')); $ui.CmbSev.SelectedIndex = 1
$ui.CmbISrv = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 200 }
$ui.TxtISearch = New-Ctl TextBox @{ Width = 220 }
$fi.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'Minimum severity:'), $ui.CmbSev, (New-Label 'Server:'), $ui.CmbISrv, (New-Label 'Search:'), $ui.TxtISearch))
$tabIssues.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.GridIssues, $fi))

# Orphans
$ui.GridOrph = New-Grid
$fo = New-Flow
$ui.ChkORisk = New-Ctl CheckBox @{ Text = 'Only users at risk'; Checked = $true; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(6, 6, 3, 0)) }
$ui.LblONote = New-Label ''
$ui.LblONote.ForeColor = C '#647086'
$fo.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.ChkORisk, $ui.LblONote))
$tabOrph.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.GridOrph, $fo))

# Remediation
$ui.RtbSql = New-Ctl RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; WordWrap = $false; Font = $script:Mono; BackColor = (C '#F8F9FB'); BorderStyle = 'None' }
$fs = New-Flow
$ui.CmbSql = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 220 }
$ui.ChkReveal = New-Ctl CheckBox @{ Text = 'Show password hashes'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(12, 6, 3, 0)) }
$ui.BtnCopySql = New-Button 'Copy'
$ui.BtnSaveSql = New-Button 'Save .sql as...'
$sqlNote = New-Label '   Generated, never executed. Aborts unless @@SERVERNAME = target. Destructive steps are commented out.'
$sqlNote.ForeColor = C '#647086'
$fs.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'Target:'), $ui.CmbSql, $ui.ChkReveal, $ui.BtnCopySql, $ui.BtnSaveSql, $sqlNote))
$tabSql.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.RtbSql, $fs))

# Automate
$ui.RtbAuto = New-Ctl RichTextBox @{ Dock = 'Fill'; ReadOnly = $true; WordWrap = $false; Font = $script:Mono; BackColor = (C '#F8F9FB'); BorderStyle = 'None' }
$fa = New-Flow
$ui.BtnAutoRefresh = New-Button 'Build from current inputs'
$ui.BtnAutoCopy = New-Button 'Copy'
$fa.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.BtnAutoRefresh, $ui.BtnAutoCopy))
$tabAuto.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.RtbAuto, $fa))

# Log
$ui.TxtLog = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono; BackColor = (C '#F8F9FB') }
$tabLog.Controls.Add($ui.TxtLog)

# ----- action bar -----
$bar = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 46; WrapContents = $false; Padding = (New-Object System.Windows.Forms.Padding(8, 8, 8, 0)) }
$ui.BtnRun = New-Ctl Button @{ Text = 'Run comparison'; AutoSize = $true; Height = 30; FlatStyle = 'Flat'; BackColor = $script:Accent; ForeColor = [System.Drawing.Color]::White; Font = $script:FontB; Padding = (New-Object System.Windows.Forms.Padding(10, 0, 10, 0)) }
$ui.BtnRun.FlatAppearance.BorderSize = 0
$ui.BtnCancel = New-Button 'Cancel'; $ui.BtnCancel.Enabled = $false
$ui.BtnDemo = New-Button 'Load demo data'
$ui.BtnHtml = New-Button 'Open HTML report'; $ui.BtnHtml.Enabled = $false
$ui.BtnFolder = New-Button 'Open output folder'; $ui.BtnFolder.Enabled = $false
$ui.Progress = New-Ctl ProgressBar @{ Style = 'Marquee'; Width = 140; Height = 16; Visible = $false; Margin = (New-Object System.Windows.Forms.Padding(12, 9, 3, 0)) }
$ui.LblStatus = New-Label 'Ready.' 9
$ui.LblStatus.ForeColor = C '#647086'
$bar.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.BtnRun, $ui.BtnCancel, $ui.BtnDemo, $ui.BtnHtml, $ui.BtnFolder, $ui.Progress, $ui.LblStatus))

# ----- top: servers + options -----
$top = New-Ctl Panel @{ Dock = 'Top'; Height = 250; Padding = (New-Object System.Windows.Forms.Padding(8, 8, 8, 0)) }

$grpSrv = New-Ctl GroupBox @{ Text = 'Servers to compare'; Dock = 'Fill'; Padding = (New-Object System.Windows.Forms.Padding(8)) }
$fr = New-Flow 34
$ui.TxtRef = New-Ctl TextBox @{ Width = 260 }
$ui.BtnDiscover = New-Button 'Discover AG replicas'
$ui.CmbAg = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 170; Visible = $false }
$ui.LblRefInfo = New-Label ''
$ui.LblRefInfo.ForeColor = C '#647086'
$fr.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'Reference (AG primary / migration source):'), $ui.TxtRef, $ui.BtnDiscover, $ui.CmbAg, $ui.LblRefInfo))

$g = New-Ctl DataGridView @{
    Dock = 'Fill'; AllowUserToAddRows = $true; RowHeadersVisible = $false; AutoSizeColumnsMode = 'Fill'; BackgroundColor = [System.Drawing.Color]::White
    BorderStyle = 'FixedSingle'; EditMode = 'EditOnEnter'; EnableHeadersVisualStyles = $false; SelectionMode = 'CellSelect'; GridColor = (C '#E2E6EC')
}
$g.ColumnHeadersDefaultCellStyle.BackColor = C '#EEF1F5'
$g.ColumnHeadersDefaultCellStyle.Font = $script:FontB
$cUse = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $cUse.Name = 'Use'; $cUse.HeaderText = 'Compare'; $cUse.FillWeight = 10
$cSrv = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $cSrv.Name = 'Server'; $cSrv.HeaderText = 'Target server  (name, name\instance or name,port)'; $cSrv.FillWeight = 42
$cMode = New-Object System.Windows.Forms.DataGridViewComboBoxColumn; $cMode.Name = 'Mode'; $cMode.HeaderText = 'Commit mode / role'; $cMode.FillWeight = 18; $cMode.FlatStyle = 'Flat'
[void]$cMode.Items.AddRange(@('Synchronous', 'Asynchronous', 'Standalone'))
$cInfo = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $cInfo.Name = 'Info'; $cInfo.HeaderText = 'Details'; $cInfo.FillWeight = 40; $cInfo.ReadOnly = $true
$cInfo.DefaultCellStyle.ForeColor = C '#647086'
foreach ($col in @($cUse, $cSrv, $cMode, $cInfo)) { [void]$g.Columns.Add($col) }
$g.Add_DefaultValuesNeeded({ $_.Row.Cells['Use'].Value = $true; $_.Row.Cells['Mode'].Value = $(if ($script:AgName) { 'Synchronous' } else { 'Standalone' }) })
$g.Add_DataError({ $_.ThrowException = $false })
$g.Add_CurrentCellDirtyStateChanged({ if ($ui.GridSrv.IsCurrentCellDirty -and $ui.GridSrv.CurrentCell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) { [void]$ui.GridSrv.CommitEdit('CurrentCellChange') } })
$ui.GridSrv = $g

$hint = New-Ctl Label @{ Dock = 'Bottom'; Height = 22; ForeColor = (C '#647086'); TextAlign = 'MiddleLeft'
    Text = 'Synchronous = automatic-failover partner (drift rated highest). Asynchronous = DR replica (one level lower). Standalone = migration target / peer.' }
$grpSrv.Controls.AddRange([System.Windows.Forms.Control[]]@($g, $hint, $fr))

$grpOpt = New-Ctl GroupBox @{ Text = 'Connection & options'; Dock = 'Right'; Width = 430; Padding = (New-Object System.Windows.Forms.Padding(8)) }
$fl = New-Ctl FlowLayoutPanel @{ Dock = 'Fill'; FlowDirection = 'TopDown'; WrapContents = $false; AutoScroll = $true }
function New-Row { $r = New-Ctl FlowLayoutPanel @{ AutoSize = $true; WrapContents = $false; Margin = (New-Object System.Windows.Forms.Padding(0)) }; return $r }
$ui.RbWin = New-Ctl RadioButton @{ Text = 'Windows authentication'; Checked = $true; AutoSize = $true }
$ui.RbSql = New-Ctl RadioButton @{ Text = 'SQL login'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(12, 3, 3, 3)) }
$ui.BtnCred = New-Button 'Set credential...'; $ui.BtnCred.Enabled = $false
$ui.LblCred = New-Label '' 8; $ui.LblCred.ForeColor = C '#647086'
$r1 = New-Row; $r1.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.RbWin, $ui.RbSql, $ui.BtnCred, $ui.LblCred))
$ui.ChkEncrypt = New-Ctl CheckBox @{ Text = 'Encrypt connection'; AutoSize = $true }
$ui.ChkTrust = New-Ctl CheckBox @{ Text = 'Trust server certificate'; AutoSize = $true }
$r2 = New-Row; $r2.Controls.AddRange([System.Windows.Forms.Control[]]@($ui.ChkEncrypt, $ui.ChkTrust))
$ui.CmbScope = New-Ctl ComboBox @{ DropDownStyle = 'DropDownList'; Width = 90 }
[void]$ui.CmbScope.Items.AddRange(@('Auto', 'AG', 'AllUser', 'None')); $ui.CmbScope.SelectedIndex = 0
$ui.NumLock = New-Ctl NumericUpDown @{ Minimum = 500; Maximum = 30000; Increment = 500; Value = 3000; Width = 70 }
$r3 = New-Row; $r3.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'DB users scope:' 6), $ui.CmbScope, (New-Label 'Lock timeout ms:' 6), $ui.NumLock))
$ui.ChkExclude = New-Ctl CheckBox @{ Text = 'Skip system principals (##, NT SERVICE, NT AUTHORITY, BUILTIN)'; Checked = $true; AutoSize = $true }
$ui.ChkPwd = New-Ctl CheckBox @{ Text = 'Compare password hashes (needs CONTROL SERVER)'; Checked = $true; AutoSize = $true }
$ui.TxtOut = New-Ctl TextBox @{ Width = 270; Text = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'LoginSyncReports') }
$ui.BtnOut = New-Button '...'
$r4 = New-Row; $r4.Controls.AddRange([System.Windows.Forms.Control[]]@((New-Label 'Output:' 6), $ui.TxtOut, $ui.BtnOut))
$fl.Controls.AddRange([System.Windows.Forms.Control[]]@($r1, $r2, $r3, $ui.ChkExclude, $ui.ChkPwd, $r4))
$grpOpt.Controls.Add($fl)

$top.Controls.AddRange([System.Windows.Forms.Control[]]@($grpSrv, $grpOpt))
$form.Controls.AddRange([System.Windows.Forms.Control[]]@($tabs, $bar, $top))   # docking: last added is outermost
#endregion

#region ---------- state & helpers ----------
$script:R = $null; $script:Files = $null; $script:Job = $null; $script:Cred = $null
$script:AgName = $null; $script:DiscoveredRef = $null; $script:Topo = @()
$script:SettingsPath = Join-Path $env:APPDATA 'LoginSync\gui-settings.json'

function Add-Log { param([string]$Line) $ui.TxtLog.AppendText($Line + [Environment]::NewLine) }
function Set-Status { param([string]$Text, [string]$Color = '#647086') $ui.LblStatus.Text = $Text; $ui.LblStatus.ForeColor = C $Color }

function Set-Busy {
    param([bool]$Busy, [string]$Text = 'Working...')
    foreach ($c in @($ui.BtnRun, $ui.BtnDemo, $ui.BtnDiscover, $ui.GridSrv, $ui.TxtRef)) { $c.Enabled = -not $Busy }
    $ui.BtnCancel.Enabled = $Busy
    $ui.Progress.Visible = $Busy
    if ($Busy) { Set-Status $Text '#1D4ED8'; $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting } else { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Show-Message { param([string]$Text, [string]$Icon = 'Information') [void][System.Windows.Forms.MessageBox]::Show($form, $Text, 'Login Sync', 'OK', $Icon) }

function Get-GridTargets {
    $ui.GridSrv.EndEdit() | Out-Null
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in $ui.GridSrv.Rows) {
        if ($row.IsNewRow) { continue }
        $name = [string]$row.Cells['Server'].Value
        if (-not $name -or -not $name.Trim()) { continue }
        if (-not [bool]$row.Cells['Use'].Value) { continue }
        $mode = [string]$row.Cells['Mode'].Value
        if (-not $mode) { $mode = 'Standalone' }
        $list.Add([pscustomobject]@{ Name = $name.Trim(); Mode = $mode })
    }
    return , $list.ToArray()
}

function Get-Options {
    $ref = $ui.TxtRef.Text.Trim()
    $isAg = ($script:AgName -and $script:DiscoveredRef -and $ref -eq $script:DiscoveredRef)
    return @{
        Reference = $ref; ReferenceRole = $(if ($isAg) { 'Primary' } else { 'Reference' }); AgName = $(if ($isAg) { $script:AgName } else { $null })
        Targets = (Get-GridTargets); Credential = $(if ($ui.RbSql.Checked) { $script:Cred } else { $null })
        Encrypt = $ui.ChkEncrypt.Checked; Trust = $ui.ChkTrust.Checked; Scope = [string]$ui.CmbScope.SelectedItem
        Exclude = $(if ($ui.ChkExclude.Checked) { @('##*', 'NT SERVICE\*', 'NT AUTHORITY\*', 'BUILTIN\*') } else { @() })
        ComparePwd = $ui.ChkPwd.Checked; LockMs = [int]$ui.NumLock.Value; OutputPath = $ui.TxtOut.Text.Trim(); Demo = $false
    }
}

function Save-Settings {
    try {
        $dir = Split-Path -Parent $script:SettingsPath
        if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        $o = [pscustomobject]@{
            Reference = $ui.TxtRef.Text; Targets = @(foreach ($row in $ui.GridSrv.Rows) { if (-not $row.IsNewRow -and $row.Cells['Server'].Value) { [pscustomobject]@{ Name = [string]$row.Cells['Server'].Value; Mode = [string]$row.Cells['Mode'].Value; Use = [bool]$row.Cells['Use'].Value; Info = [string]$row.Cells['Info'].Value } } })
            Encrypt = $ui.ChkEncrypt.Checked; Trust = $ui.ChkTrust.Checked; Scope = [string]$ui.CmbScope.SelectedItem; LockMs = [int]$ui.NumLock.Value
            OutputPath = $ui.TxtOut.Text; SqlAuth = $ui.RbSql.Checked; SqlUser = $(if ($script:Cred) { $script:Cred.UserName } else { $null })   # never the password
            AgName = $script:AgName; DiscoveredRef = $script:DiscoveredRef
        }
        [System.IO.File]::WriteAllText($script:SettingsPath, (ConvertTo-Json -InputObject $o -Depth 4))
    }
    catch { Add-Log "Could not save settings: $($_.Exception.Message)" }
}

function Import-Settings {
    if (-not (Test-Path $script:SettingsPath)) { return }
    try {
        $o = Get-Content -Raw -LiteralPath $script:SettingsPath | ConvertFrom-Json
        if ($o.Reference) { $ui.TxtRef.Text = $o.Reference }
        foreach ($t in @($o.Targets)) { if ($t.Name) { [void]$ui.GridSrv.Rows.Add([bool]$t.Use, $t.Name, $(if ($t.Mode) { $t.Mode } else { 'Standalone' }), $t.Info) } }
        $ui.ChkEncrypt.Checked = [bool]$o.Encrypt; $ui.ChkTrust.Checked = [bool]$o.Trust
        if ($o.Scope) { $ui.CmbScope.SelectedItem = $o.Scope }
        if ($o.LockMs) { $ui.NumLock.Value = [decimal]$o.LockMs }
        if ($o.OutputPath) { $ui.TxtOut.Text = $o.OutputPath }
        if ($o.SqlAuth) { $ui.RbSql.Checked = $true; if ($o.SqlUser) { $ui.LblCred.Text = "$($o.SqlUser) (password needed)" } }
        if ($o.AgName -and $o.DiscoveredRef) { $script:AgName = $o.AgName; $script:DiscoveredRef = $o.DiscoveredRef; $ui.LblRefInfo.Text = "AG $($o.AgName) (primary at last discovery)" }
        Add-Log "Loaded last-used settings from $script:SettingsPath"
    }
    catch { Add-Log "Could not read settings: $($_.Exception.Message)" }
}

function Show-CredentialDialog {
    $d = New-Ctl Form @{ Text = 'SQL login'; FormBorderStyle = 'FixedDialog'; StartPosition = 'CenterParent'; MaximizeBox = $false; MinimizeBox = $false; ClientSize = (New-Object System.Drawing.Size(360, 150)); Font = $script:Font; ShowInTaskbar = $false }
    $l1 = New-Ctl Label @{ Text = 'Login:'; Location = (New-Object System.Drawing.Point(14, 18)); AutoSize = $true }
    $t1 = New-Ctl TextBox @{ Location = (New-Object System.Drawing.Point(100, 15)); Width = 240; Text = $(if ($script:Cred) { $script:Cred.UserName } else { '' }) }
    $l2 = New-Ctl Label @{ Text = 'Password:'; Location = (New-Object System.Drawing.Point(14, 52)); AutoSize = $true }
    $t2 = New-Ctl TextBox @{ Location = (New-Object System.Drawing.Point(100, 49)); Width = 240; UseSystemPasswordChar = $true }
    $n = New-Ctl Label @{ Text = 'Held in memory as a SecureString for this session only.'; Location = (New-Object System.Drawing.Point(14, 82)); AutoSize = $true; ForeColor = (C '#647086') }
    $ok = New-Ctl Button @{ Text = 'OK'; DialogResult = 'OK'; Location = (New-Object System.Drawing.Point(184, 110)); Width = 75 }
    $ca = New-Ctl Button @{ Text = 'Cancel'; DialogResult = 'Cancel'; Location = (New-Object System.Drawing.Point(265, 110)); Width = 75 }
    $d.Controls.AddRange([System.Windows.Forms.Control[]]@($l1, $t1, $l2, $t2, $n, $ok, $ca)); $d.AcceptButton = $ok; $d.CancelButton = $ca
    $res = $null
    if ($d.ShowDialog($form) -eq 'OK' -and $t1.Text.Trim()) {
        $sec = New-Object System.Security.SecureString
        foreach ($ch in $t2.Text.ToCharArray()) { $sec.AppendChar($ch) }
        $sec.MakeReadOnly()
        $res = New-Object System.Management.Automation.PSCredential($t1.Text.Trim(), $sec)
    }
    $t2.Text = ''; $d.Dispose()
    return $res
}
#endregion

#region ---------- background work (keeps the window responsive) ----------
$script:Timer = New-Object System.Windows.Forms.Timer
$script:Timer.Interval = 250

function Start-UiJob {
    param([string]$Script, [hashtable]$Params, [scriptblock]$OnDone)
    $queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    $rs = [runspacefactory]::CreateRunspace((New-LSWorkerSessionState))
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('LSLogQueue', $queue)
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Script)
    foreach ($k in $Params.Keys) { [void]$ps.AddParameter($k, $Params[$k]) }
    $script:Job = @{ PS = $ps; RS = $rs; Queue = $queue; OnDone = $OnDone; Handle = $ps.BeginInvoke() }
    $script:Timer.Start()
}

$script:Timer.Add_Tick({
        $j = $script:Job
        if (-not $j) { $script:Timer.Stop(); return }
        $line = $null
        while ($j.Queue.TryDequeue([ref]$line)) { Add-Log $line; Set-Status ($line -replace '^\S+\s+\[\w+\]\s+', '') '#1D4ED8' }
        if (-not $j.Handle.IsCompleted) { return }
        $script:Timer.Stop()
        $out = $null; $err = $null
        try { $out = $j.PS.EndInvoke($j.Handle) } catch { $err = $_.Exception.GetBaseException().Message }
        while ($j.Queue.TryDequeue([ref]$line)) { Add-Log $line }
        if (-not $err -and (-not $out -or $out.Count -eq 0) -and $j.PS.Streams.Error.Count) { $err = $j.PS.Streams.Error[0].Exception.GetBaseException().Message }
        foreach ($e in $j.PS.Streams.Error) { Add-Log "ERROR: $($e.Exception.GetBaseException().Message)" }
        $j.PS.Dispose(); $j.RS.Dispose()
        $script:Job = $null
        & $j.OnDone $out $err
    })

$script:RunScript = @'
param($Opt)
$targets = @(foreach ($t in $Opt.Targets) { [pscustomobject]@{ Name = $t.Name; Mode = $t.Mode } })
$r = Invoke-LSLoginSync -ReferenceServer $Opt.Reference -ReferenceRole $Opt.ReferenceRole -Targets $targets -AgName $Opt.AgName `
    -Credential $Opt.Credential -Encrypt:$Opt.Encrypt -TrustServerCertificate:$Opt.Trust -DatabaseScope $Opt.Scope `
    -ExcludeLogin $Opt.Exclude -SkipPasswordCompare:(-not $Opt.ComparePwd) -LockTimeoutMs $Opt.LockMs -Demo:$Opt.Demo
$files = Export-LSReport -Result $r -OutputPath $Opt.OutputPath
[pscustomobject]@{ Result = $r; Files = $files }
'@

$script:DiscoverScript = @'
param($Opt)
Write-LSLog "Discovering availability groups from $($Opt.Server)"
@(Get-LSAgTopology -Server $Opt.Server -Credential $Opt.Credential -Encrypt:$Opt.Encrypt -TrustServerCertificate:$Opt.Trust -LockTimeoutMs $Opt.LockMs)
'@
#endregion

#region ---------- rendering results ----------
function Set-CellStatus {
    param($Cell, [string]$Status, $Notes)
    $st = $script:StatusStyle[$Status]; if (-not $st) { $st = @('#FFFFFF', '#1D2330') }
    $Cell.Value = $(if ($script:StatusText[$Status]) { $script:StatusText[$Status] } else { $Status })
    $Cell.Style.BackColor = C $st[0]; $Cell.Style.ForeColor = C $st[1]
    $Cell.Style.SelectionBackColor = C $st[0]; $Cell.Style.SelectionForeColor = C $st[1]
    $Cell.Style.Alignment = 'MiddleCenter'; $Cell.Style.Font = $script:FontB
    $Cell.ToolTipText = (@($Notes) -join [Environment]::NewLine)
    $Cell.Tag = $Status
}

function Get-ServerHeader { param($s) $m = if ($s.Role -eq 'Target') { $s.Mode } else { $s.Role }; return "$($s.Name)`n($m)" }

function Show-Summary {
    $R = $script:R
    $g = $ui.GridSum; $g.Columns.Clear(); $g.Rows.Clear()
    foreach ($h in @('Server', 'Role / mode', 'Status', 'In sync', 'Critical', 'High', 'Medium', 'Low', 'Orphan risk', 'Version', 'Permissions', 'ms / queries')) { [void]$g.Columns.Add(($h -replace '\W', ''), $h) }
    $g.Columns[0].FillWeight = 140; $g.Columns[2].FillWeight = 170; $g.Columns[10].FillWeight = 150
    foreach ($s in $R.Servers) {
        $isRef = ($s.Role -ne 'Target')
        $worst = if ($s.Critical) { 'Critical' } elseif ($s.High) { 'High' } elseif ($s.Medium) { 'Medium' } elseif ($s.Low) { 'Low' } else { 'None' }
        $status = if (-not $s.Connected) { "Unreachable: $($s.Error)" } elseif ($isRef) { 'Reference (source of truth)' } elseif ($worst -eq 'None') { 'In sync' } else { "Drift - worst $worst" }
        $sync = if ($isRef -or -not $s.Connected) { '' } else { "$($s.Score)%  ($($s.InSync)/$($s.InScope))" }
        $perm = @(); if ($s.CanViewDefinition) { $perm += 'VIEW ANY DEF' } else { $perm += 'no VIEW ANY DEF!' }; if ($s.CanControlServer) { $perm += 'CONTROL SERVER' } else { $perm += 'no hashes' }
        $mode = if ($isRef) { $s.Role } else { $s.Mode }
        $i = $g.Rows.Add($s.Name, $mode, $status, $sync, $s.Critical, $s.High, $s.Medium, $s.Low, $s.OrphanRisk, $s.Version, ($perm -join ', '), "$($s.ElapsedMs) / $($s.QueryCount)")
        $row = $g.Rows[$i]
        $row.Cells[0].Style.Font = $script:FontB
        $col = if (-not $s.Connected) { '#C02626' } elseif ($isRef) { '#1D4ED8' } else { $script:SevColor[$worst] }
        $row.Cells[2].Style.ForeColor = C $col; $row.Cells[2].Style.Font = $script:FontB
        foreach ($k in @(@(4, 'Critical'), @(5, 'High'), @(6, 'Medium'), @(7, 'Low'))) { if ([int]$row.Cells[$k[0]].Value -gt 0) { $row.Cells[$k[0]].Style.ForeColor = C $script:SevColor[$k[1]]; $row.Cells[$k[0]].Style.Font = $script:FontB } }
    }
    $rtb = $ui.RtbRec; $rtb.Clear()
    foreach ($r in $R.Recommendations) {
        $rtb.SelectionStart = $rtb.TextLength
        $rtb.SelectionColor = C $script:SevColor[$r.Level]; $rtb.SelectionFont = $script:FontB
        $rtb.AppendText(('{0,-11}' -f $r.Level))
        $rtb.SelectionColor = C '#1D2330'; $rtb.SelectionFont = $script:Font
        $rtb.AppendText($r.Text + "`n`n")
    }
    if ($ui.Welcome.Parent) { $tabOverview.Controls.Remove($ui.Welcome); $tabOverview.Controls.Add($split) }
    try {
        $h = $tabOverview.ClientSize.Height
        $split.SplitterDistance = [Math]::Max(90, [Math]::Min(60 + 28 * @($R.Servers).Count, [int]($h * 0.55)))
    } catch { }
}

function Show-Matrix {
    $R = $script:R; if (-not $R) { return }
    $g = $ui.GridMatrix; $g.SuspendLayout(); $g.Columns.Clear(); $g.Rows.Clear()
    [void]$g.Columns.Add('Principal', 'Principal'); [void]$g.Columns.Add('Type', 'Type')
    foreach ($s in $R.Servers) { $ix = $g.Columns.Add("srv_$($s.Name)", (Get-ServerHeader $s)); $g.Columns[$ix].MinimumWidth = 120; $g.Columns[$ix].Tag = $s.Name }
    $g.Columns[0].MinimumWidth = 200
    $q = $ui.TxtMFilter.Text.Trim(); $drift = $ui.ChkMDrift.Checked
    $tmap = @{ 1 = 'S'; 2 = 'U'; 3 = 'G'; 4 = 'R' }; $tf = $tmap[$ui.CmbMType.SelectedIndex]
    $names = @($R.Servers | ForEach-Object Name)
    foreach ($m in $R.Matrix) {
        if ($q -and $m.Name -notlike "*$q*") { continue }
        if ($tf -and $m.Type -ne $tf) { continue }
        $statuses = foreach ($n in $names) { $c = $m.Cells[$n]; if ($c) { $c.S } else { 'ABSENT' } }
        $isDrift = (-not $m.HostLocal) -and @($statuses | Where-Object { $_ -notin 'OK', 'REF', 'INFO', 'ERR' }).Count -gt 0
        if ($drift -and -not $isDrift) { continue }
        $i = $g.Rows.Add(); $row = $g.Rows[$i]
        $row.Cells[0].Value = $(if ($m.Disabled) { "$($m.Name)  (disabled)" } else { $m.Name }); $row.Cells[0].Style.Font = $script:FontB; $row.Tag = $m.Name
        $row.Cells[1].Value = ($m.TypeDesc -replace '_', ' ').ToLower(); $row.Cells[1].Style.ForeColor = C '#647086'
        for ($k = 0; $k -lt $names.Count; $k++) {
            $c = $m.Cells[$names[$k]]
            if ($c) { Set-CellStatus ($row.Cells[$k + 2]) $c.S $c.N } else { Set-CellStatus ($row.Cells[$k + 2]) 'ABSENT' @('Not present') }
        }
    }
    $g.ResumeLayout()
    $nd = @($R.Matrix | Where-Object { $mm = $_; -not $mm.HostLocal -and @($names | Where-Object { $c = $mm.Cells[$_]; -not $c -or $c.S -notin 'OK', 'REF', 'INFO', 'ERR' }).Count }).Count
    $tabMatrix.Text = "Login matrix ($nd)"
}

function Show-Issues {
    $R = $script:R; if (-not $R) { return }
    $g = $ui.GridIssues; $g.SuspendLayout(); $g.Columns.Clear(); $g.Rows.Clear()
    foreach ($h in @('Severity', 'Server', 'Mode', 'Principal', 'Category', 'Detail', 'Suggestion')) { [void]$g.Columns.Add($h, $h) }
    $g.Columns['Severity'].FillWeight = 45; $g.Columns['Server'].FillWeight = 70; $g.Columns['Mode'].FillWeight = 60; $g.Columns['Principal'].FillWeight = 90
    $g.Columns['Category'].FillWeight = 65; $g.Columns['Detail'].FillWeight = 230; $g.Columns['Suggestion'].FillWeight = 170
    $g.Columns['Detail'].DefaultCellStyle.WrapMode = 'True'; $g.Columns['Suggestion'].DefaultCellStyle.WrapMode = 'True'; $g.Columns['Principal'].DefaultCellStyle.WrapMode = 'True'
    $g.Columns['Suggestion'].DefaultCellStyle.ForeColor = C '#4B5565'
    $min = $script:SevRank[[string]$ui.CmbSev.SelectedItem]
    $srv = [string]$ui.CmbISrv.SelectedItem; if ($srv -eq 'All servers') { $srv = '' }
    $q = $ui.TxtISearch.Text.Trim()
    $shown = 0
    foreach ($i in $R.Issues) {
        if ($script:SevRank[$i.Severity] -lt $min) { continue }
        if ($srv -and $i.Target -ne $srv) { continue }
        if ($q -and "$($i.Object) $($i.Category) $($i.Detail)" -notlike "*$q*") { continue }
        $ix = $g.Rows.Add($i.Severity, $i.Target, $i.Mode, $i.Object, $i.Category, $i.Detail, $i.Suggestion)
        $g.Rows[$ix].Cells[0].Style.ForeColor = C $script:SevColor[$i.Severity]; $g.Rows[$ix].Cells[0].Style.Font = $script:FontB
        $g.Rows[$ix].Cells[3].Style.Font = $script:FontB
        $shown++
    }
    $g.ResumeLayout()
    $tabIssues.Text = "Findings ($(@($R.Issues | Where-Object Severity -ne 'Info').Count))"
}

function Show-Orphans {
    $R = $script:R; if (-not $R) { return }
    $g = $ui.GridOrph; $g.SuspendLayout(); $g.Columns.Clear(); $g.Rows.Clear()
    foreach ($h in @('Database', 'User', 'Type', 'Reference login')) { [void]$g.Columns.Add(($h -replace '\W', ''), $h) }
    foreach ($s in $R.Servers) { [void]$g.Columns.Add("o_$($s.Name)", (Get-ServerHeader $s)) }
    $names = @($R.Servers | ForEach-Object Name)
    foreach ($o in $R.Orphans) {
        if ($ui.ChkORisk.Checked -and -not $o.AtRisk) { continue }
        $ix = $g.Rows.Add($o.Database, $o.User, $o.Type, $(if ($o.Login) { $o.Login } else { '(none)' }))
        $row = $g.Rows[$ix]; $row.Cells[1].Style.Font = $script:FontB
        if (-not $o.Login) { $row.Cells[3].Style.ForeColor = C '#C02626' }
        for ($k = 0; $k -lt $names.Count; $k++) { $c = $o.Cells[$names[$k]]; if ($c) { Set-CellStatus ($row.Cells[$k + 4]) $c.S $c.N } else { Set-CellStatus ($row.Cells[$k + 4]) 'ERR' @() } }
    }
    $g.ResumeLayout()
    $ui.LblONote.Text = "$($R.Meta.DbUsersScanned) users in $($R.Meta.DatabasesScanned) reference databases, matched by SID against every server's logins. On an AG, fix the LOGIN (SID) - never ALTER USER on a secondary."
    $tabOrph.Text = "Orphan risk ($(@($R.Orphans | Where-Object AtRisk).Count))"
}

function Show-Sql {
    $R = $script:R; if (-not $R) { return }
    $k = [string]$ui.CmbSql.SelectedItem
    if (-not $k -or -not $R.Scripts.Contains($k)) { $ui.RtbSql.Text = 'No reachable target - nothing to script.'; return }
    $txt = $R.Scripts[$k]
    if (-not $ui.ChkReveal.Checked) { $txt = Get-LSRedactedScript $txt }
    $ui.RtbSql.Text = $txt
    # light syntax cue: comments grey
    $ui.RtbSql.SuspendLayout()
    foreach ($m in [regex]::Matches($txt, '(?m)--.*$|/\*[\s\S]*?\*/')) { $ui.RtbSql.Select($m.Index, $m.Length); $ui.RtbSql.SelectionColor = C '#6B7280' }
    foreach ($m in [regex]::Matches($txt, '(?m)^-- \[MANUAL\].*$')) { $ui.RtbSql.Select($m.Index, $m.Length); $ui.RtbSql.SelectionColor = C '#C2410C' }
    $ui.RtbSql.Select(0, 0); $ui.RtbSql.ResumeLayout()
}

function Get-AutomationText {
    $o = Get-Options
    $q = { param($s) '"' + $s + '"' }
    $argsList = New-Object System.Collections.Generic.List[string]
    if ($o.AgName) { $argsList.Add("-DiscoverAgFrom $(& $q $o.Reference) -AvailabilityGroup $(& $q $o.AgName)") }
    else {
        $argsList.Add("-ReferenceServer $(& $q $o.Reference)")
        $sync = @($o.Targets | Where-Object Mode -eq 'Synchronous' | ForEach-Object Name)
        $std = @($o.Targets | Where-Object Mode -eq 'Standalone' | ForEach-Object Name)
        $asy = @($o.Targets | Where-Object Mode -eq 'Asynchronous' | ForEach-Object Name)
        if ($sync.Count) { $argsList.Add("-TargetServer $((@($sync + $std) | ForEach-Object { & $q $_ }) -join ',') -TargetMode Synchronous") }
        elseif ($std.Count) { $argsList.Add("-TargetServer $(($std | ForEach-Object { & $q $_ }) -join ',')") }
        if ($asy.Count) { $argsList.Add("-AsyncTarget $(($asy | ForEach-Object { & $q $_ }) -join ',')") }
    }
    if ($o.Encrypt) { $argsList.Add('-Encrypt') }
    if ($o.Trust) { $argsList.Add('-TrustServerCertificate') }
    if ($o.Scope -ne 'Auto') { $argsList.Add("-DatabaseScope $($o.Scope)") }
    if (-not $o.ComparePwd) { $argsList.Add('-SkipPasswordCompare') }
    if (-not $ui.ChkExclude.Checked) { $argsList.Add("-ExcludeLogin @()") }
    $argsList.Add('-NoOpen -FailOnSeverity High -OutputPath "\\fileserver\dba$\LoginSync"')
    $a = $argsList -join ' '
    $core = if ($script:SelfPath) { $script:SelfPath } else { 'C:\DBA\LoginSync\LoginSync.ps1' }
    $refName = if ($o.Reference) { $o.Reference } else { '<reference>' }
    $sb = New-Object System.Text.StringBuilder
    $L = { param($s) [void]$sb.AppendLine($s) }
    & $L '# ============================================================================================='
    & $L "#  Automating the login sync check for $refName"
    & $L '#  The check is read-only, so it is safe to schedule. Remediation stays a reviewed, manual step.'
    & $L '#  Run it from a DBA utility / jump server - not on the replicas themselves.'
    & $L '#  Exit codes: 0 = clean (below threshold)  2 = drift at/above -FailOnSeverity  1 = run failed.'
    & $L '# ============================================================================================='
    & $L ''
    & $L '# --- 1. Command line ----------------------------------------------------------------------'
    & $L "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$core`" $a"
    & $L ''
    & $L '# --- 2. Windows Task Scheduler (daily 06:15 as a gMSA; the gMSA needs VIEW ANY DEFINITION,'
    & $L '#        VIEW SERVER STATE and - for password comparison - CONTROL SERVER on every replica) ----'
    & $L "`$action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File `"$core`" $($a.Replace("'", "''"))'"
    & $L "`$trigger   = New-ScheduledTaskTrigger -Daily -At '06:15'"
    & $L "`$principal = New-ScheduledTaskPrincipal -UserId 'DOMAIN\gmsa_sqlcheck`$' -LogonType Password"
    & $L "Register-ScheduledTask -TaskName 'LoginSync - $($refName.Replace("'", ''))' -Action `$action -Trigger `$trigger -Principal `$principal"
    & $L ''
    & $L '# --- 3. SQL Server Agent job step (on a DBA/utility instance) ------------------------------'
    & $L '#   Step type: Operating system (CmdExec). Process exit code of a successful command: 0.'
    & $L '#   Exit code 2 fails the step -> the job''s failure notification pages the DBA operator.'
    & $L "#   Command:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$core`" $a"
    & $L ''
    & $L '# --- 4. Also worth wiring in ------------------------------------------------------------------'
    & $L '#  * Run after every failover (alert on AG role change, error 1480) and after every security change.'
    & $L '#  * Feed result.json / findings.csv into your monitoring; the HTML is the human view.'
    & $L '#  * Keep ~30 days of report folders; the .sql files can contain password hashes - restrict the share.'
    & $L '#  * After review, dbatools can apply the fix for you:'
    & $L "#      Copy-DbaLogin -Source $refName -Destination <secondary> -Login <name> -Force   # keeps SID + hash"
    & $L "#      Sync-DbaAvailabilityGroup -Primary $refName -AvailabilityGroup <ag>   # logins + more; see Get-Help -Full"
    & $L '#    and then re-run this check as the read-only before/after proof.'
    if ($o.Credential -or $ui.RbSql.Checked) {
        & $L ''
        & $L '# NOTE: you used SQL authentication in the GUI. For scheduled runs prefer Windows auth (gMSA) so no'
        & $L '#       password has to be stored; -Credential is supported for interactive use.'
    }
    return $sb.ToString()
}

function Show-All {
    $R = $script:R
    Show-Summary
    $ui.CmbISrv.Items.Clear(); [void]$ui.CmbISrv.Items.Add('All servers')
    foreach ($s in $R.Servers) { [void]$ui.CmbISrv.Items.Add($s.Name) }
    $ui.CmbISrv.SelectedIndex = 0
    Show-Matrix; Show-Issues; Show-Orphans
    $ui.CmbSql.Items.Clear(); foreach ($k in $R.Scripts.Keys) { [void]$ui.CmbSql.Items.Add($k) }
    if ($ui.CmbSql.Items.Count) { $ui.CmbSql.SelectedIndex = 0 } else { Show-Sql }
    $ui.RtbAuto.Text = Get-AutomationText
    $ui.BtnHtml.Enabled = [bool]$script:Files; $ui.BtnFolder.Enabled = [bool]$script:Files
    $w = $R.Meta.WorstSeverity
    $col = if ($script:SevColor[$w]) { $script:SevColor[$w] } else { '#15803D' }
    $txt = if ($w -in 'None', 'Info') { "All in sync.  $($R.Meta.DurationSec)s" } else { "$(@($R.Issues | Where-Object Severity -ne 'Info').Count) finding(s), worst: $w.  $($R.Meta.DurationSec)s" }
    Set-Status $txt $col
    $form.Text = "Login Sync Check  -  $($R.Meta.ReferenceServerName)$(if ($R.Meta.AgName) { ' / ' + $R.Meta.AgName })$(if ($R.Meta.Demo) { '  [DEMO DATA]' })  -  read-only"
    $tabs.SelectedTab = $tabOverview
}

function Start-Comparison {
    param([hashtable]$Opt)
    Set-Busy $true 'Collecting (read-only)...'
    Add-Log ('-' * 80)
    Start-UiJob -Script $script:RunScript -Params @{ Opt = $Opt } -OnDone {
        param($out, $err)
        Set-Busy $false
        if ($err) { Set-Status "Failed: $err" '#C02626'; Add-Log "FAILED: $err"; Show-Message "The comparison did not complete:`n`n$err" 'Error'; return }
        $script:R = $out[0].Result; $script:Files = $out[0].Files
        Add-Log "Report: $($script:Files.Html)"
        Show-All
    }
}
#endregion

#region ---------- events ----------
$ui.RbSql.Add_CheckedChanged({ $ui.BtnCred.Enabled = $ui.RbSql.Checked })
$ui.BtnCred.Add_Click({ $c = Show-CredentialDialog; if ($c) { $script:Cred = $c; $ui.LblCred.Text = $c.UserName } })
$ui.BtnOut.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = 'Folder for reports (a sub-folder is created per run)'
        if (Test-Path $ui.TxtOut.Text) { $d.SelectedPath = $ui.TxtOut.Text }
        if ($d.ShowDialog($form) -eq 'OK') { $ui.TxtOut.Text = $d.SelectedPath }
    })
$ui.TxtRef.Add_TextChanged({ if ($script:DiscoveredRef -and $ui.TxtRef.Text.Trim() -ne $script:DiscoveredRef) { $ui.LblRefInfo.Text = 'manual reference' } })

$ui.BtnDiscover.Add_Click({
        $srv = $ui.TxtRef.Text.Trim()
        if (-not $srv) { Show-Message 'Enter any replica or the AG listener in the Reference box first.'; return }
        if ($ui.RbSql.Checked -and -not $script:Cred) { Show-Message 'Set the SQL credential first.'; return }
        Set-Busy $true "Discovering AG from $srv..."
        $opt = @{ Server = $srv; Credential = $(if ($ui.RbSql.Checked) { $script:Cred } else { $null }); Encrypt = $ui.ChkEncrypt.Checked; Trust = $ui.ChkTrust.Checked; LockMs = [int]$ui.NumLock.Value }
        Start-UiJob -Script $script:DiscoverScript -Params @{ Opt = $opt } -OnDone {
            param($out, $err)
            Set-Busy $false
            if ($err) { Set-Status "Discovery failed: $err" '#C02626'; Show-Message "AG discovery failed:`n`n$err" 'Error'; return }
            $script:Topo = @($out | ForEach-Object { $_ })
            if (-not $script:Topo.Count) { Set-Status 'No availability groups on that instance.' '#B45309'; Show-Message 'That instance hosts no availability groups. Enter targets manually (Standalone mode for a migration).'; return }
            $ags = @($script:Topo | ForEach-Object AgName | Select-Object -Unique)
            $ui.CmbAg.Items.Clear(); foreach ($a in $ags) { [void]$ui.CmbAg.Items.Add($a) }
            $ui.CmbAg.Visible = ($ags.Count -gt 1)
            $ui.CmbAg.SelectedIndex = 0   # fires SelectedIndexChanged -> Use-AgTopology fills the grid
        }
    })

function Use-AgTopology {
    param([string]$Ag)
    $reps = @($script:Topo | Where-Object AgName -eq $Ag)
    $pri = $reps | Where-Object IsPrimary | Select-Object -First 1
    if (-not $pri) { Show-Message "Could not determine the primary of $Ag (VIEW SERVER STATE needed)." 'Warning'; return }
    $script:AgName = $Ag; $script:DiscoveredRef = $pri.Replica
    $ui.TxtRef.Text = $pri.Replica
    $ui.LblRefInfo.Text = "AG $Ag  -  primary"
    $ui.GridSrv.Rows.Clear()
    foreach ($r in $reps) {
        if ($r.IsPrimary) { continue }
        $use = ($r.Mode -ne 'ConfigurationOnly')
        $mode = if ($r.Mode -in 'Synchronous', 'Asynchronous') { $r.Mode } else { 'Standalone' }
        $info = "failover $($r.FailoverMode.ToLower()), readable: $($r.Readable.ToLower())$(if ($r.Connected) { ', ' + $r.Connected.ToLower() })$(if (-not $use) { ' - configuration-only replica, nothing to compare' })"
        [void]$ui.GridSrv.Rows.Add($use, $r.Replica, $mode, $info)
    }
    Set-Status "AG $Ag : $(@($reps).Count) replicas loaded. Review and Run." '#15803D'
    Add-Log "AG $Ag primary=$($pri.Replica); secondaries: $((@($reps | Where-Object { -not $_.IsPrimary } | ForEach-Object { "$($_.Replica) [$($_.Mode)]" })) -join ', ')"
}
$ui.CmbAg.Add_SelectedIndexChanged({ if ($ui.CmbAg.SelectedItem) { Use-AgTopology ([string]$ui.CmbAg.SelectedItem) } })

$ui.BtnRun.Add_Click({
        $o = Get-Options
        if (-not $o.Reference) { Show-Message 'Enter the reference server (AG primary or migration source).'; return }
        if (-not @($o.Targets).Count) { Show-Message 'Add at least one target server (and tick "Compare").'; return }
        if (@($o.Targets | Where-Object { $_.Name -eq $o.Reference }).Count) { Show-Message 'The reference server is also listed as a target - remove it from the grid.'; return }
        if ($ui.RbSql.Checked -and -not $script:Cred) { Show-Message 'SQL authentication is selected: set the credential first.'; return }
        if (-not $o.OutputPath) { Show-Message 'Choose an output folder.'; return }
        Save-Settings
        Start-Comparison $o
    })

$ui.BtnDemo.Add_Click({
        $o = Get-Options; $o.Demo = $true
        if (-not $o.OutputPath) { $o.OutputPath = Join-Path $env:TEMP 'LoginSyncReports' }
        Start-Comparison $o
    })

$ui.BtnCancel.Add_Click({
        if ($script:Job) { Add-Log 'Cancel requested - stopping after the current query (lock/query timeouts bound the wait).'; try { [void]$script:Job.PS.BeginStop($null, $null) } catch {} }
    })
$ui.BtnHtml.Add_Click({ if ($script:Files) { Invoke-Item $script:Files.Html } })
$ui.BtnFolder.Add_Click({ if ($script:Files) { Invoke-Item $script:Files.Folder } })

$ui.TxtMFilter.Add_TextChanged({ Show-Matrix })
$ui.ChkMDrift.Add_CheckedChanged({ Show-Matrix })
$ui.CmbMType.Add_SelectedIndexChanged({ Show-Matrix })
$ui.GridMatrix.Add_CellClick({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 2 -or -not $script:R) { return }
        $g = $ui.GridMatrix
        $principal = [string]$g.Rows[$e.RowIndex].Tag
        $server = [string]$g.Columns[$e.ColumnIndex].Tag
        $cell = $g.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
        $lines = @("$principal  on  $server  :  $($cell.Value)")
        foreach ($n in ($cell.ToolTipText -split "`r?`n")) { if ($n) { $lines += "   - $n" } }
        $rel = @($script:R.Issues | Where-Object { $_.Target -eq $server -and $_.Object -eq $principal })
        if ($rel.Count) { $lines += ''; $lines += 'Suggested action:'; foreach ($i in $rel) { $lines += "   [$($i.Severity)] $($i.Suggestion)" } }
        $ui.TxtMDetail.Text = $lines -join [Environment]::NewLine
    })
$ui.CmbSev.Add_SelectedIndexChanged({ Show-Issues })
$ui.CmbISrv.Add_SelectedIndexChanged({ Show-Issues })
$ui.TxtISearch.Add_TextChanged({ Show-Issues })
$ui.ChkORisk.Add_CheckedChanged({ Show-Orphans })
$ui.CmbSql.Add_SelectedIndexChanged({ Show-Sql })
$ui.ChkReveal.Add_CheckedChanged({
        if ($ui.ChkReveal.Checked) {
            $a = [System.Windows.Forms.MessageBox]::Show($form, "Show SQL password hashes on screen?`n`nThey can be used to recreate the logins elsewhere. Make sure nobody else can see your screen.", 'Login Sync', 'YesNo', 'Warning')
            if ($a -ne 'Yes') { $ui.ChkReveal.Checked = $false; return }
        }
        Show-Sql
    })
$ui.BtnCopySql.Add_Click({ if ($ui.RtbSql.Text) { [System.Windows.Forms.Clipboard]::SetText($ui.RtbSql.Text); Set-Status ("Copied" + $(if ($ui.ChkReveal.Checked) { ' (with hashes).' } else { ' (hashes redacted).' })) '#15803D' } })
$ui.BtnSaveSql.Add_Click({
        $k = [string]$ui.CmbSql.SelectedItem
        if (-not $k -or -not $script:R) { return }
        $d = New-Object System.Windows.Forms.SaveFileDialog
        $d.Filter = 'SQL script (*.sql)|*.sql'; $d.FileName = "remediation_$($k -replace '[\\/:*?""<>|,]', '_').sql"
        if ($d.ShowDialog($form) -eq 'OK') {
            [System.IO.File]::WriteAllText($d.FileName, $script:R.Scripts[$k], (New-Object System.Text.UTF8Encoding($true)))
            Set-Status "Saved $($d.FileName) (full script incl. hashes - keep it secure)." '#B45309'
        }
    })
$ui.BtnAutoRefresh.Add_Click({ $ui.RtbAuto.Text = Get-AutomationText })
$ui.BtnAutoCopy.Add_Click({ if ($ui.RtbAuto.Text) { [System.Windows.Forms.Clipboard]::SetText($ui.RtbAuto.Text); Set-Status 'Automation snippet copied.' '#15803D' } })

$form.Add_Shown({
        Import-Settings
        $ui.RtbAuto.Text = Get-AutomationText
        Add-Log "LoginSync v$script:LSVersion  $(if ($script:SelfPath) { $script:SelfPath } else { '(loaded without a file path)' })"
        Add-Log 'Guardrails: catalog SELECTs only (guarded), LOCK_TIMEOUT, DEADLOCK_PRIORITY LOW, READ UNCOMMITTED, pooling off, one connection at a time.'
    })
$form.Add_FormClosing({
        if ($script:Job) { try { $script:Job.PS.Stop(); $script:Job.PS.Dispose(); $script:Job.RS.Dispose() } catch {} }
        $script:Timer.Stop()
    })
#endregion

[void]$form.ShowDialog()
$form.Dispose()
}

#requires -Modules dbatools
<#
    SqlPatch - SQL Server patching pipeline (dbatools-based)
    Unit of work = one SQL instance (one AG replica). An FCI's nodes are all
    patched in the same run. Compatible with Windows PowerShell 5.1 and PS 7.
#>

$script:Ctx = [ordered]@{
    ControlInstance  = $null
    DryRun           = $false
    Local            = $false
    MediaShare       = $null          # \\server\SQLPatch\Media
    NodeMediaPath    = 'C:\SQLPatch'  # where media is copied on each node
    LogPath          = Join-Path $env:ProgramData 'SqlPatch\logs'
    RunId            = $null
    FullOrDiffHours  = 26             # max age of last full/diff
    LogBackupMinutes = 60             # max age of last log backup (FULL recovery)
    MinFreeGB        = 6
    RebootTimeoutSec = 1800
    AgSyncTimeoutSec = 1800
    PendingLocalReboot = $false
}
$script:MajorMap = @{ 13 = '2016'; 14 = '2017'; 15 = '2019'; 16 = '2022'; 17 = '2025' }

#region ---------- plumbing ----------------------------------------------------
function Set-PatchContext {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ControlInstance, [switch]$DryRun, [switch]$Local,
          [string]$MediaShare, [hashtable]$Override)
    $script:Ctx.ControlInstance = $ControlInstance
    $script:Ctx.DryRun = [bool]$DryRun
    $script:Ctx.Local  = [bool]$Local
    if ($MediaShare) { $script:Ctx.MediaShare = $MediaShare }
    if ($Override) { foreach ($k in $Override.Keys) { $script:Ctx[$k] = $Override[$k] } }
    New-Item -ItemType Directory -Force -Path $script:Ctx.LogPath | Out-Null
}

function Invoke-Ctl {
    param([string]$Query, [hashtable]$P = @{})
    Invoke-DbaQuery -SqlInstance $script:Ctx.ControlInstance -Database DBAPatch -Query $Query -SqlParameter $P -EnableException
}

function Invoke-Sql {
    param([string]$SqlInstance, [string]$Query, [hashtable]$P = @{}, [string]$Database = 'master')
    Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $Query -SqlParameter $P -EnableException
}

function Write-PatchLog {
    param([string]$Stage, [string]$Message, [string]$Target = '', [ValidateSet('Info','Warn','Error')][string]$Level = 'Info')
    $prefix = if ($script:Ctx.DryRun) { '[DryRun] ' } else { '' }
    $line = '{0:u} {1,-5} {2,-10} {3,-20} {4}{5}' -f (Get-Date), $Level, $Stage, $Target, $prefix, $Message
    switch ($Level) { 'Error' { Write-Warning $line } 'Warn' { Write-Warning $line } default { Write-Host $line } }
    Add-Content -Path (Join-Path $script:Ctx.LogPath ('SqlPatch_{0:yyyyMMdd}.log' -f (Get-Date))) -Value $line
    if ($script:Ctx.RunId) {
        try {
            Invoke-Ctl 'INSERT dbo.PatchRunLog (RunId, Stage, Target, Level, Message) VALUES (@r,@s,@t,@l,@m)' `
                @{ r = $script:Ctx.RunId; s = $Stage; t = $Target; l = $Level; m = "$prefix$Message" }
        } catch { Write-Warning "Could not write to control DB: $_" }
    }
}

function Test-IsSelf([string]$ComputerName) {
    ($ComputerName -split '\.')[0] -in @($env:COMPUTERNAME, 'localhost', '.')
}

function Invoke-OnNode {
    param([string]$ComputerName, [scriptblock]$ScriptBlock, [object[]]$ArgumentList = @())
    if (Test-IsSelf $ComputerName) { & $ScriptBlock @ArgumentList }
    else { Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList -ErrorAction Stop }
}

function Get-SqlBuild([string]$SqlInstance) {
    [version](Invoke-Sql $SqlInstance "SELECT CAST(SERVERPROPERTY('ProductVersion') AS varchar(20)) AS v").v
}
#endregion

#region ---------- discovery + media -----------------------------------------
function Update-PatchInventory {
    <# Refreshes build, major version and WinRM reachability for every enabled instance. #>
    [CmdletBinding()] param()
    $instances = Invoke-Ctl 'SELECT InstanceId, SqlInstance FROM dbo.PatchInstance WHERE IsEnabled = 1'
    foreach ($i in $instances) {
        $build = $null; $major = $null; $winrm = $true
        try {
            $build = Get-SqlBuild $i.SqlInstance
            $major = $script:MajorMap[$build.Major]
        } catch { Write-PatchLog Discover "SQL unreachable: $_" $i.SqlInstance Warn }
        $nodes = Invoke-Ctl 'SELECT ComputerName FROM dbo.PatchNode WHERE InstanceId = @i AND IsEnabled = 1' @{ i = $i.InstanceId }
        foreach ($n in $nodes) {
            try { Test-WSMan -ComputerName $n.ComputerName -ErrorAction Stop | Out-Null }
            catch { $winrm = $false; Write-PatchLog Discover 'WinRM not reachable' $n.ComputerName Warn }
        }
        Invoke-Ctl @'
UPDATE dbo.PatchInstance SET CurrentBuild = @b, MajorVersion = COALESCE(@m, MajorVersion),
       WinRmOk = @w, LastCheckedAt = SYSDATETIME() WHERE InstanceId = @i
'@ @{ b = $(if ($build) { "$build" } else { [DBNull]::Value }); m = $(if ($major) { $major } else { [DBNull]::Value }); w = $winrm; i = $i.InstanceId }
        Write-PatchLog Discover "Build $build ($major), WinRM ok: $winrm" $i.SqlInstance
    }
    Invoke-Ctl 'SELECT * FROM dbo.vw_PatchCompliance ORDER BY PatchGroup, SqlInstance'
}

function Save-PatchMedia {
    <# Downloads the installer for every active baseline that has no media yet. Central node needs internet. #>
    [CmdletBinding()] param()
    if (-not $script:Ctx.MediaShare) { throw 'MediaShare not set (Set-PatchContext -MediaShare).' }
    $todo = Invoke-Ctl 'SELECT BaselineId, MajorVersion, TargetKB, MediaFile FROM dbo.PatchBaseline WHERE IsActive = 1'
    foreach ($b in $todo) {
        if ($b.MediaFile -isnot [DBNull] -and (Test-Path $b.MediaFile)) { continue }
        Write-PatchLog Media "Downloading $($b.TargetKB) for SQL $($b.MajorVersion)" $script:Ctx.MediaShare
        $file = Save-DbaKbUpdate -Name $b.TargetKB -Path $script:Ctx.MediaShare -Architecture x64 -EnableException | Select-Object -First 1
        Invoke-Ctl 'UPDATE dbo.PatchBaseline SET MediaFile = @f WHERE BaselineId = @id' @{ f = $file.FullName; id = $b.BaselineId }
        Write-PatchLog Media "Saved $($file.FullName)" $script:Ctx.MediaShare
    }
}
#endregion

#region ---------- AG helpers -------------------------------------------------
function Get-AgMembership([string]$SqlInstance) {
    # AGs this instance belongs to, with its local role and the current primary
    Invoke-Sql $SqlInstance @'
IF SERVERPROPERTY('IsHadrEnabled') = 1
SELECT ag.name AS AgName, ags.primary_replica AS PrimaryReplica, ars.role_desc AS LocalRole
FROM sys.availability_groups ag
JOIN sys.dm_hadr_availability_group_states ags ON ags.group_id = ag.group_id
JOIN sys.dm_hadr_availability_replica_states ars ON ars.group_id = ag.group_id AND ars.is_local = 1
'@
}

function Get-AgReplicaState([string]$PrimaryInstance, [string]$AgName) {
    Invoke-Sql $PrimaryInstance @'
SELECT ar.replica_server_name AS ReplicaName, ar.availability_mode_desc AS AvailabilityMode,
       ars.role_desc AS Role, ars.synchronization_health_desc AS Health, ars.connected_state_desc AS Connected
FROM sys.availability_groups ag
JOIN sys.availability_replicas ar ON ar.group_id = ag.group_id
LEFT JOIN sys.dm_hadr_availability_replica_states ars ON ars.replica_id = ar.replica_id
WHERE ag.name = @ag
'@ @{ ag = $AgName }
}

function Wait-AgHealthy([string]$AgName, [string]$AnyReplica) {
    $deadline = (Get-Date).AddSeconds($script:Ctx.AgSyncTimeoutSec)
    do {
        try {
            $primary = (Get-AgMembership $AnyReplica | Where-Object AgName -eq $AgName).PrimaryReplica
            $bad = Get-AgReplicaState $primary $AgName | Where-Object { $_.Health -ne 'HEALTHY' }
            if (-not $bad) { Write-PatchLog AgHealth "All replicas HEALTHY (primary $primary)" $AgName; return $true }
            Write-PatchLog AgHealth ("Waiting: " + (($bad | ForEach-Object { "$($_.ReplicaName)=$($_.Health)" }) -join ', ')) $AgName
        } catch { Write-PatchLog AgHealth "Not reachable yet: $_" $AgName Warn }
        Start-Sleep -Seconds 30
    } while ((Get-Date) -lt $deadline)
    Write-PatchLog AgHealth "Not healthy after $($script:Ctx.AgSyncTimeoutSec)s" $AgName Error
    return $false
}

function Invoke-AgFailoverAway([string]$AgName, [string]$FromInstance, [version]$TargetBuild) {
    # Fail over to a synchronous, HEALTHY secondary that is already at the target build
    $candidates = Get-AgReplicaState $FromInstance $AgName |
        Where-Object { $_.Role -eq 'SECONDARY' -and $_.AvailabilityMode -eq 'SYNCHRONOUS_COMMIT' -and $_.Health -eq 'HEALTHY' }
    $target = $candidates | Where-Object { (Get-SqlBuild $_.ReplicaName) -ge $TargetBuild } | Select-Object -First 1
    if (-not $target) { throw "AG $AgName has no healthy, patched synchronous secondary to fail over to." }
    Write-PatchLog Failover "Failing over $AgName from $FromInstance to $($target.ReplicaName)" $AgName
    if (-not $script:Ctx.DryRun) {
        Invoke-DbaAgFailover -SqlInstance $target.ReplicaName -AvailabilityGroup $AgName -Confirm:$false -EnableException | Out-Null
        if (-not (Wait-AgHealthy $AgName $target.ReplicaName)) { throw "AG $AgName unhealthy after failover." }
    }
    return $target.ReplicaName
}
#endregion

#region ---------- pre-checks -------------------------------------------------
function Test-PatchPrereq {
    param([Parameter(Mandatory)]$Instance, [Parameter(Mandatory)][object[]]$Nodes, [Parameter(Mandatory)][version]$TargetBuild, [bool]$AllowFailover = $true)
    $fail = New-Object System.Collections.Generic.List[string]
    $warn = New-Object System.Collections.Generic.List[string]
    $si = $Instance.SqlInstance

    if ($script:Ctx.Local -and $Instance.InstanceType -eq 'FCI') { $fail.Add('Local mode supports standalone instances only; FCI needs Remote mode.') }

    try { $build = Get-SqlBuild $si } catch { $fail.Add("Cannot connect to SQL: $_"); return [pscustomobject]@{ Pass = $false; Failures = $fail; Warnings = $warn } }
    $compliant = $build -ge $TargetBuild

    # AG role + replicas
    $ags = @(Get-AgMembership $si)
    $primaryOf = @($ags | Where-Object LocalRole -eq 'PRIMARY')
    $replicaNames = @($si)
    foreach ($ag in $ags) {
        $state = @(Get-AgReplicaState $ag.PrimaryReplica $ag.AgName)
        $replicaNames += $state.ReplicaName
        $unhealthy = $state | Where-Object Health -ne 'HEALTHY'
        if ($unhealthy) { $fail.Add("AG $($ag.AgName) not healthy: " + (($unhealthy | ForEach-Object { "$($_.ReplicaName)=$($_.Health)" }) -join ', ')) }
    }
    foreach ($ag in $primaryOf) {
        if (-not $AllowFailover) { $fail.Add("Primary of $($ag.AgName) but schedule does not allow failover."); continue }
        $others = Get-AgReplicaState $si $ag.AgName | Where-Object { $_.ReplicaName -ne $si }
        foreach ($o in $others) {
            try { if ((Get-SqlBuild $o.ReplicaName) -lt $TargetBuild) { $fail.Add("Primary of $($ag.AgName): patch replica $($o.ReplicaName) first.") } }
            catch { $fail.Add("Cannot reach replica $($o.ReplicaName): $_") }
        }
    }

    # Backups: AG backups may run on any replica, so take the newest across all replicas
    $bq = @'
SELECT d.name AS Db, d.recovery_model_desc AS Recovery,
  MAX(CASE WHEN b.type IN ('D','I') THEN b.backup_finish_date END) AS LastFullDiff,
  MAX(CASE WHEN b.type = 'L' THEN b.backup_finish_date END) AS LastLog
FROM sys.databases d LEFT JOIN msdb.dbo.backupset b ON b.database_name = d.name
WHERE d.name <> 'tempdb' AND d.state_desc = 'ONLINE' GROUP BY d.name, d.recovery_model_desc
'@
    $rows = foreach ($r in ($replicaNames | Select-Object -Unique)) { try { Invoke-Sql $r $bq } catch { $warn.Add("Backup history unreadable on ${r}: $_") } }
    foreach ($g in ($rows | Group-Object Db)) {
        $full = ($g.Group | Where-Object { $_.LastFullDiff -isnot [DBNull] } | Measure-Object LastFullDiff -Maximum).Maximum
        $log  = ($g.Group | Where-Object { $_.LastLog -isnot [DBNull] } | Measure-Object LastLog -Maximum).Maximum
        if (-not $full -or $full -lt (Get-Date).AddHours(-$script:Ctx.FullOrDiffHours)) { $fail.Add("No full/diff backup of [$($g.Name)] in $($script:Ctx.FullOrDiffHours)h") }
        if ($g.Group[0].Recovery -eq 'FULL' -and (-not $log -or $log -lt (Get-Date).AddMinutes(-$script:Ctx.LogBackupMinutes))) { $fail.Add("No log backup of [$($g.Name)] in $($script:Ctx.LogBackupMinutes)m") }
    }

    # Long open transactions: warn only
    $open = (Invoke-Sql $si "SELECT COUNT(*) AS c FROM sys.dm_tran_active_transactions WHERE transaction_begin_time < DATEADD(MINUTE,-5,GETDATE()) AND transaction_type = 1").c
    if ($open -gt 0) { $warn.Add("$open transaction(s) open > 5 min") }

    # Node checks: disk + pending reboot
    if (-not $compliant) {
        foreach ($n in $Nodes) {
            try {
                $r = Invoke-OnNode $n.ComputerName {
                    $pending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
                               (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
                    [pscustomobject]@{ FreeGB = [math]::Round((Get-PSDrive C).Free / 1GB, 1); Pending = $pending }
                }
                if ($r.FreeGB -lt $script:Ctx.MinFreeGB) { $fail.Add("$($n.ComputerName): only $($r.FreeGB) GB free on C:") }
                if ($r.Pending) { $fail.Add("$($n.ComputerName): reboot already pending - reboot first") }
            } catch { $fail.Add("$($n.ComputerName): remote check failed (WinRM?) - $_") }
        }
    }

    foreach ($f in $fail) { Write-PatchLog Precheck $f $si Error }
    foreach ($w in $warn) { Write-PatchLog Precheck $w $si Warn }
    [pscustomobject]@{
        Pass = ($fail.Count -eq 0); Failures = $fail; Warnings = $warn
        CurrentBuild = $build; Compliant = $compliant; PrimaryOf = $primaryOf
    }
}
#endregion

#region ---------- patch actions ----------------------------------------------
function Wait-NodeRestart([string]$ComputerName) {
    $deadline = (Get-Date).AddSeconds($script:Ctx.RebootTimeoutSec)
    # wait for it to go down (max 10 min), then come back
    $downBy = (Get-Date).AddMinutes(10)
    while ((Get-Date) -lt $downBy) { try { Test-WSMan $ComputerName -ErrorAction Stop | Out-Null; Start-Sleep 10 } catch { break } }
    while ((Get-Date) -lt $deadline) {
        try { Invoke-Command -ComputerName $ComputerName -ScriptBlock { 1 } -ErrorAction Stop | Out-Null; return } catch { Start-Sleep 20 }
    }
    throw "$ComputerName did not come back within $($script:Ctx.RebootTimeoutSec)s"
}

function Wait-SqlOnline([string]$SqlInstance, [int]$TimeoutSec = 1200) {
    # Polls until SQL accepts logins (script-upgrade mode blocks them) and non-AG DBs are online
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $c = (Invoke-Sql $SqlInstance "SELECT COUNT(*) AS c FROM sys.databases WHERE replica_id IS NULL AND state_desc NOT IN ('ONLINE','OFFLINE')").c
            if ($c -eq 0) { Write-PatchLog Wait 'SQL online' $SqlInstance; return }
        } catch { }
        Start-Sleep 15
    }
    throw "$SqlInstance not online after ${TimeoutSec}s"
}

function Install-PatchOnNode {
    param([Parameter(Mandatory)]$Node, [Parameter(Mandatory)][string]$MediaFile)
    $cn = $Node.ComputerName
    $leaf = Split-Path $MediaFile -Leaf
    $dest = Join-Path $script:Ctx.NodeMediaPath $leaf
    Write-PatchLog Patch "Installing $leaf" $cn
    if ($script:Ctx.DryRun) { return }

    # 1. Copy media to the node (avoids the double-hop / CredSSP problem)
    if (Test-IsSelf $cn) {
        New-Item -ItemType Directory -Force $script:Ctx.NodeMediaPath | Out-Null
        if (-not (Test-Path $dest)) { Copy-Item $MediaFile $dest }
    } else {
        $s = New-PSSession -ComputerName $cn -ErrorAction Stop
        try {
            $exists = Invoke-Command -Session $s { param($p, $d) New-Item -ItemType Directory -Force $p | Out-Null; Test-Path $d } -ArgumentList $script:Ctx.NodeMediaPath, $dest
            if (-not $exists) { Copy-Item -Path $MediaFile -Destination $dest -ToSession $s -Force }
        } finally { Remove-PSSession $s }
    }

    # 2. Drain the cluster node, run setup
    $exit = Invoke-OnNode $cn {
        param($exe, $clustered)
        if ($clustered -and (Get-Command Suspend-ClusterNode -ErrorAction SilentlyContinue)) {
            Suspend-ClusterNode -Name $env:COMPUTERNAME -Drain -Wait -ErrorAction Stop | Out-Null
        }
        $p = Start-Process -FilePath $exe -ArgumentList '/quiet', '/IAcceptSQLServerLicenseTerms', '/Action=Patch', '/AllInstances' -Wait -PassThru
        $p.ExitCode
    } -ArgumentList $dest, [bool]$Node.IsClustered
    if ($exit -notin 0, 3010) { throw "Setup on $cn failed with exit code $exit (see Setup Bootstrap\Log\Summary.txt)" }
    Write-PatchLog Patch "Setup exit code $exit, restarting" $cn

    # 3. Restart + resume cluster node
    if (Test-IsSelf $cn) {
        # Can't wait for our own reboot: orchestrator records PendingReboot and restarts at the end.
        $script:Ctx.PendingLocalReboot = $true
        return
    }
    Invoke-Command -ComputerName $cn -ScriptBlock { shutdown.exe /r /t 5 /c "SqlPatch CU install" /d p:2:17 } | Out-Null
    Wait-NodeRestart $cn
    if ($Node.IsClustered) {
        Invoke-OnNode $cn { if (Get-Command Resume-ClusterNode -ErrorAction SilentlyContinue) { Resume-ClusterNode -Name $env:COMPUTERNAME -Failback NoFailback | Out-Null } }
    }
    Write-PatchLog Patch 'Node back online' $cn
}

function Invoke-FciPatch {
    param($Instance, [object[]]$Nodes, [string]$MediaFile)
    $role = $Instance.FciRoleName
    $any = $Nodes[0].ComputerName
    $owner = Invoke-OnNode $any { param($r) (Get-ClusterGroup -Name $r).OwnerNode.Name } -ArgumentList $role
    $passive = @($Nodes | Where-Object { ($_.ComputerName -split '\.')[0] -ne $owner })
    $active  = $Nodes | Where-Object { ($_.ComputerName -split '\.')[0] -eq $owner } | Select-Object -First 1
    if (-not $passive) { throw "FCI $role has no passive node to patch first." }
    Write-PatchLog Patch "FCI role '$role' owned by $owner; passive: $($passive.ComputerName -join ', ')" $Instance.SqlInstance

    foreach ($n in $passive) { Install-PatchOnNode $n $MediaFile }

    $moveTo = ($passive[0].ComputerName -split '\.')[0]
    Write-PatchLog Patch "Moving role '$role' to patched node $moveTo" $Instance.SqlInstance
    if (-not $script:Ctx.DryRun) {
        Invoke-OnNode $passive[0].ComputerName { param($r, $t) Move-ClusterGroup -Name $r -Node $t -ErrorAction Stop | Out-Null } -ArgumentList $role, $moveTo
        Wait-SqlOnline $Instance.SqlInstance
    }
    if ($active) { Install-PatchOnNode $active $MediaFile }
}

function Test-PatchResult {
    param($Instance, [version]$TargetBuild, [string[]]$AgNames)
    $si = $Instance.SqlInstance; $ok = $true
    $build = Get-SqlBuild $si
    if ($build -lt $TargetBuild) { Write-PatchLog Validate "Build $build < target $TargetBuild" $si Error; $ok = $false }
    else { Write-PatchLog Validate "Build $build OK" $si }
    $svc = Invoke-Sql $si "SELECT servicename, status_desc FROM sys.dm_server_services WHERE status_desc <> 'Running' AND startup_type_desc = 'Automatic'"
    foreach ($s in $svc) { Write-PatchLog Validate "Service not running: $($s.servicename)" $si Error; $ok = $false }
    $dbs = Invoke-Sql $si "SELECT name, state_desc FROM sys.databases WHERE replica_id IS NULL AND state_desc NOT IN ('ONLINE','OFFLINE')"
    foreach ($d in $dbs) { Write-PatchLog Validate "DB $($d.name) is $($d.state_desc)" $si Error; $ok = $false }
    foreach ($ag in $AgNames) { if (-not (Wait-AgHealthy $ag $si)) { $ok = $false } }
    return $ok
}
#endregion

#region ---------- orchestrator -----------------------------------------------
function Invoke-PatchInstance {
    <# Patches one claimed schedule row end to end. Returns final status string. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Row)
    $si = $Row.SqlInstance
    $mode = if ($script:Ctx.Local) { 'Local' } else { 'Remote' }
    $script:Ctx.RunId = (Invoke-Ctl @'
INSERT dbo.PatchRun (ScheduleId, InstanceId, RunHost, RunAs, Mode, DryRun, FromBuild)
VALUES (@s, @i, @h, @u, @m, @d, @b); SELECT CAST(SCOPE_IDENTITY() AS int) AS Id
'@ @{ s = $Row.ScheduleId; i = $Row.InstanceId; h = $env:COMPUTERNAME; u = "$env:USERDOMAIN\$env:USERNAME"; m = $mode; d = $script:Ctx.DryRun; b = "$($Row.CurrentBuild)" }).Id
    Invoke-Ctl "UPDATE dbo.PatchSchedule SET Status = 'InProgress', LastRunId = @r WHERE ScheduleId = @s" @{ r = $script:Ctx.RunId; s = $Row.ScheduleId }
    $status = 'Failed'
    try {
        $nodes = @(Invoke-Ctl 'SELECT ComputerName, IsClustered FROM dbo.PatchNode WHERE InstanceId = @i AND IsEnabled = 1' @{ i = $Row.InstanceId })
        $major = $script:MajorMap[(Get-SqlBuild $si).Major]
        $base = Invoke-Ctl 'SELECT TargetKB, TargetBuild, MediaFile FROM dbo.PatchBaseline WHERE MajorVersion = @v AND IsActive = 1' @{ v = $major }
        if (-not $base) { throw "No active baseline for SQL $major" }
        $target = [version]$base.TargetBuild
        Write-PatchLog Start "Run $($script:Ctx.RunId): $si -> $($base.TargetKB) ($target), mode $mode" $si

        $pre = Test-PatchPrereq -Instance $Row -Nodes $nodes -TargetBuild $target -AllowFailover ([bool]$Row.AllowFailover)
        $agNames = @((Get-AgMembership $si).AgName | Where-Object { $_ })
        if ($pre.Compliant) {
            Write-PatchLog Precheck "Already at $($pre.CurrentBuild) - validating only" $si
        } elseif (-not $pre.Pass) {
            throw "Pre-checks failed ($($pre.Failures.Count))"
        } else {
            if ($base.MediaFile -is [DBNull] -or -not (Test-Path $base.MediaFile)) { throw "Media for $($base.TargetKB) not staged - run Save-PatchMedia" }
            $wasPrimaryOf = @()
            foreach ($ag in $pre.PrimaryOf) { $null = Invoke-AgFailoverAway $ag.AgName $si $target; $wasPrimaryOf += $ag.AgName }

            if ($Row.InstanceType -eq 'FCI') { Invoke-FciPatch -Instance $Row -Nodes $nodes -MediaFile $base.MediaFile }
            else {
                Install-PatchOnNode $nodes[0] $base.MediaFile
                if ($script:Ctx.PendingLocalReboot) {
                    $status = 'PendingReboot'
                    Write-PatchLog Patch 'Installed. Restarting this machine; re-run Invoke-SqlPatch.ps1 -Local after boot to validate.' $si Warn
                    return $status   # finally{} still runs and leaves the schedule InProgress
                }
                if (-not $script:Ctx.DryRun) { Wait-SqlOnline $si }
            }

            if ($Row.PreferredPrimary -and $wasPrimaryOf -and -not $script:Ctx.DryRun) {
                foreach ($ag in $wasPrimaryOf) {
                    if (Wait-AgHealthy $ag $si) {
                        Write-PatchLog Failover "Failing $ag back to preferred primary" $si
                        Invoke-DbaAgFailover -SqlInstance $si -AvailabilityGroup $ag -Confirm:$false -EnableException | Out-Null
                    }
                }
            }
        }

        if ($script:Ctx.DryRun) { $status = 'DryRun'; Write-PatchLog Done 'Dry run complete - nothing changed' $si }
        elseif (Test-PatchResult -Instance $Row -TargetBuild $target -AgNames $agNames) { $status = 'Succeeded' }
        else { $status = 'Failed' }
    } catch {
        Write-PatchLog Error "$_" $si Error
        $status = 'Failed'
    } finally {
        $schedStatus = switch ($status) { 'DryRun' { 'Planned' } 'PendingReboot' { 'InProgress' } default { $status } }
        $now = try { "$(Get-SqlBuild $si)" } catch { '' }
        Invoke-Ctl "UPDATE dbo.PatchRun SET Status = @st, EndTime = SYSDATETIME(), ToBuild = @b WHERE RunId = @r" @{ st = $status; b = $now; r = $script:Ctx.RunId }
        Invoke-Ctl "UPDATE dbo.PatchSchedule SET Status = @st WHERE ScheduleId = @s" @{ st = $schedStatus; s = $Row.ScheduleId }
        if ($now) { Invoke-Ctl "UPDATE dbo.PatchInstance SET CurrentBuild = @b, LastCheckedAt = SYSDATETIME() WHERE InstanceId = @i" @{ b = $now; i = $Row.InstanceId } }
        Write-PatchLog Done "Finished: $status" $si
        $script:Ctx.RunId = $null
        if ($script:Ctx.PendingLocalReboot) { $script:Ctx.PendingLocalReboot = $false; Restart-Computer -Force }
    }
    return $status
}
#endregion

Export-ModuleMember -Function Set-PatchContext, Update-PatchInventory, Save-PatchMedia, Test-PatchPrereq, Invoke-PatchInstance, Write-PatchLog, Invoke-Ctl

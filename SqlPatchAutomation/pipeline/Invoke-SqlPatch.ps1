<#
.SYNOPSIS
    Entry point for the SQL Server patching pipeline.

.EXAMPLE  # Refresh builds + WinRM status for the whole inventory
    .\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Action Discover

.EXAMPLE  # Download approved CUs to the share
    .\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Action Download -MediaShare \\CENTRAL01\SQLPatch\Media

.EXAMPLE  # Dry run the next due schedule row (all pre-checks, no changes)
    .\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -DryRun

.EXAMPLE  # Patch one specific instance now (from the central node)
    .\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -SqlInstance SQLAG02 -Now

.EXAMPLE  # Run ON a target where WinRM is blocked (standalone only). Re-run after reboot to validate.
    .\Invoke-SqlPatch.ps1 -ControlInstance CENTRAL01 -Local
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ControlInstance,
    [ValidateSet('Patch', 'Discover', 'Download', 'Plan')][string]$Action = 'Patch',
    [string]$SqlInstance,
    [int]$ScheduleId,
    [switch]$Local,
    [switch]$DryRun,
    [switch]$Now,              # with -SqlInstance: create an ad-hoc schedule row due immediately
    [string]$MediaShare,
    [int]$MaxRuns = 1,         # serial for now; raise later - the claim proc already enforces one-per-PatchGroup
    [switch]$TrustServerCertificate   # lab / self-signed certs: accept the SQL Server certificate without CA validation
)
$ErrorActionPreference = 'Stop'
Import-Module dbatools
if ($TrustServerCertificate) {
    # dbatools 2.x encrypts connections and validates the cert by default. Session-only change.
    Set-DbatoolsConfig -FullName sql.connection.trustcert -Value $true
}
Import-Module (Join-Path $PSScriptRoot 'module\SqlPatch.psm1') -Force
Set-PatchContext -ControlInstance $ControlInstance -DryRun:$DryRun -Local:$Local -MediaShare $MediaShare

switch ($Action) {
    'Discover' { Update-PatchInventory | Format-Table -AutoSize; return }
    'Download' { Save-PatchMedia; return }
    'Plan' {
        Invoke-Ctl @'
SELECT s.ScheduleId, i.SqlInstance, i.PatchGroup, s.PlannedStart, s.Status, c.CurrentBuild, c.TargetBuild, c.Compliance
FROM dbo.PatchSchedule s JOIN dbo.PatchInstance i ON i.InstanceId = s.InstanceId
JOIN dbo.vw_PatchCompliance c ON c.SqlInstance = i.SqlInstance
WHERE s.Status IN ('Planned','Claimed','InProgress') ORDER BY s.PlannedStart
'@ | Format-Table -AutoSize
        return
    }
}

# ---------- Patch ----------
if ($Local -and -not $SqlInstance) {
    # Resume a run interrupted by our own reboot, else find the instance hosted on this machine
    $resume = Invoke-Ctl @'
SELECT TOP (1) s.ScheduleId, i.SqlInstance FROM dbo.PatchSchedule s JOIN dbo.PatchInstance i ON i.InstanceId = s.InstanceId
WHERE s.Status = 'InProgress' AND s.ClaimedBy = @h
'@ @{ h = $env:COMPUTERNAME }
    if ($resume) {
        Write-PatchLog Resume "Resuming schedule $($resume.ScheduleId) after reboot" $resume.SqlInstance
        Invoke-Ctl "UPDATE dbo.PatchSchedule SET Status = 'Planned' WHERE ScheduleId = @s" @{ s = $resume.ScheduleId }
        $ScheduleId = $resume.ScheduleId
    } else {
        $SqlInstance = (Invoke-Ctl @'
SELECT TOP (1) i.SqlInstance FROM dbo.PatchInstance i JOIN dbo.PatchNode n ON n.InstanceId = i.InstanceId
WHERE n.ComputerName = @h OR n.ComputerName LIKE @h + '.%'
'@ @{ h = $env:COMPUTERNAME }).SqlInstance
        if (-not $SqlInstance) { throw "No inventory row has a node named $env:COMPUTERNAME" }
    }
}

if ($Now -and $SqlInstance -and -not $ScheduleId) {
    $ScheduleId = (Invoke-Ctl @'
INSERT dbo.PatchSchedule (InstanceId, PlannedStart) SELECT InstanceId, SYSDATETIME() FROM dbo.PatchInstance WHERE SqlInstance = @i;
SELECT CAST(SCOPE_IDENTITY() AS int) AS Id
'@ @{ i = $SqlInstance }).Id
    if (-not $ScheduleId) { throw "$SqlInstance is not in dbo.PatchInstance" }
}

$results = @()
for ($n = 0; $n -lt $MaxRuns; $n++) {
    $p = @{ h = $env:COMPUTERNAME; s = $(if ($ScheduleId) { $ScheduleId } else { [DBNull]::Value }); i = $(if ($SqlInstance) { $SqlInstance } else { [DBNull]::Value }) }
    $row = Invoke-Ctl 'EXEC dbo.usp_ClaimNextSchedule @RunHost = @h, @ScheduleId = @s, @SqlInstance = @i' $p
    if (-not $row) { Write-PatchLog Claim 'Nothing due (or its PatchGroup is busy).' $ControlInstance; break }
    if (-not $Local -and $row.ExecutionMode -eq 'Local') {
        Write-PatchLog Claim "$($row.SqlInstance) is flagged Local (no WinRM) - run this script on that server with -Local" $row.SqlInstance Warn
        Invoke-Ctl "UPDATE dbo.PatchSchedule SET Status = 'Planned', ClaimedBy = NULL WHERE ScheduleId = @s" @{ s = $row.ScheduleId }
        break
    }
    $results += [pscustomobject]@{ SqlInstance = $row.SqlInstance; Status = (Invoke-PatchInstance -Row $row) }
    if ($ScheduleId -or $SqlInstance) { break }
}
$results | Format-Table -AutoSize
if ($results.Status -contains 'Failed') { exit 1 }

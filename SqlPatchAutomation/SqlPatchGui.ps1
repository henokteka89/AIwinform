<#
.SYNOPSIS
    SQL Server single-server patcher (Windows GUI).

.DESCRIPTION
    1. Check     - connects to the destination, shows CU level, last reboot, uptime and
                   runs guardrail checks (PASS / WARN / FAIL). Lists every update for that
                   SQL version on the right so you can pick one.
    2. Download  - downloads the selected update into a folder ON THE CENTRAL SERVER and
                   shares that folder (temporary share) for the duration of the patching.
    3. Install   - copies the installer from the share to the destination, runs setup
                   silently with live progress, reboots Windows if setup requires it,
                   validates, and removes the share (optional).
    Extra buttons: Reboot Windows, Restart SQL service, Validate, Remove share.

    Run from: the central server, your workstation, or the destination itself.
    If WinRM to the destination is blocked, run this tool ON the destination.

.NOTES
    Requires: dbatools (Install-Module dbatools), Windows PowerShell 5.1 or PowerShell 7 on Windows.
    Run elevated, as an account that is sysadmin on SQL and local admin on the servers.
#>
#requires -Version 5.1
param()

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Crash guard written in C#, so it still works if the PowerShell engine is stopped (e.g. Ctrl+C in the console).
# A stopped pipeline closes the window cleanly instead of looping "Unhandled exception" dialogs.
if (-not ('SqlPatchGuiGuard' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System;
using System.Threading;
using System.Windows.Forms;
public static class SqlPatchGuiGuard {
    static string lastMsg = "";
    static DateTime lastTime = DateTime.MinValue;
    public static void Install() {
        try { Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException); } catch { }
        Application.ThreadException -= OnThreadException;
        Application.ThreadException += OnThreadException;
    }
    static void OnThreadException(object sender, ThreadExceptionEventArgs e) {
        for (Exception ex = e.Exception; ex != null; ex = ex.InnerException) {
            string msg = ex.Message ?? "";
            if (ex.GetType().Name == "PipelineStoppedException" || msg.IndexOf("pipeline has been stopped", StringComparison.OrdinalIgnoreCase) >= 0) {
                Application.ExitThread();   // engine is gone: close the window, no dialogs
                return;
            }
        }
        string m = e.Exception.Message ?? "unknown error";
        if (m == lastMsg && (DateTime.Now - lastTime).TotalSeconds < 30) return;   // don't repeat the same popup
        lastMsg = m; lastTime = DateTime.Now;
        MessageBox.Show(m, "SQL patcher - unexpected error (the tool keeps running)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
'@
}
[SqlPatchGuiGuard]::Install()

# Ctrl+C in the console must not kill the engine behind an open window
$script:OldCtrlC = $null
try { $script:OldCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } catch { }

# ============================================================================
#  Shared state between the window and the background worker
# ============================================================================
$sync = [hashtable]::Synchronized(@{
    Queue  = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    Busy   = $false
    State  = [hashtable]::Synchronized(@{})
    Checks = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
})
$LogDir      = Join-Path $env:ProgramData 'SqlPatchGui\logs'
$SettingsDir = Join-Path $env:APPDATA 'SqlPatchGui'
New-Item -ItemType Directory -Force -Path $LogDir, $SettingsDir | Out-Null
$LogFile      = Join-Path $LogDir ('SqlPatchGui_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
$SettingsFile = Join-Path $SettingsDir 'settings.json'

# ============================================================================
#  Helpers - defined in both the UI thread and the worker runspace
# ============================================================================
$Helpers = {
    function Log([string]$Msg, [string]$Level = 'Info') { $sync.Queue.Enqueue(@{ T = 'Log'; Msg = $Msg; Level = $Level; Time = Get-Date }) }
    function Prog([int]$Pct, [string]$Text) { $sync.Queue.Enqueue(@{ T = 'Prog'; Pct = $Pct; Text = $Text }) }
    function Emit([string]$Type, $Data) { $sync.Queue.Enqueue(@{ T = $Type; Data = $Data }) }
    function Add-Check([string]$Status, [string]$Name, [string]$Detail) {
        $row = [pscustomobject]@{ Status = $Status; Check = $Name; Detail = $Detail }
        [void]$sync.Checks.Add($row); Emit 'Check' $row
        # Only pre-install guardrails block Install; post-install validation results don't.
        if ($Status -eq 'FAIL' -and $Name -notlike 'Post:*') { $sync.State.CheckFailed = $true }
    }
    function Set-Info([string]$Name, $Value) { Emit 'InfoSet' @($Name, "$Value") }
    function Set-Phase([string]$Phase) { $sync.Phase = $Phase }
    function Get-BuildKey($v) { $x = [version]"$v"; [version]('{0}.{1}.{2}' -f $x.Major, $x.Minor, [math]::Max(0, $x.Build)) }
    function Fmt-Elapsed([datetime]$Start) { '{0:hh\:mm\:ss}' -f ((Get-Date) - $Start) }

    function Test-IsLocal([string]$Name) {
        $n = ($Name -split '[\\.,]')[0]
        $n -eq '' -or $n -in @($env:COMPUTERNAME, 'localhost', '.', '(local)')
    }

    function Get-Cim([string]$ComputerName) {
        # WinRM first, then DCOM (works when WinRM is blocked). Local = DCOM.
        $dcom = New-CimSessionOption -Protocol Dcom
        if (Test-IsLocal $ComputerName) { return New-CimSession -SessionOption $dcom -ErrorAction Stop }
        try { New-CimSession -ComputerName $ComputerName -OperationTimeoutSec 20 -ErrorAction Stop }
        catch { New-CimSession -ComputerName $ComputerName -SessionOption $dcom -OperationTimeoutSec 20 -ErrorAction Stop }
    }

    function Test-RegKey($Cim, [string]$Path) {
        $r = Invoke-CimMethod -CimSession $Cim -Namespace root/default -ClassName StdRegProv -MethodName EnumKey `
            -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = $Path }
        $r.ReturnValue -eq 0
    }

    function Invoke-Sql([string]$Query) {
        Invoke-DbaQuery -SqlInstance $sync.State.SqlInstance -Query $Query -EnableException
    }

    function Get-SetupLogRoot([switch]$Unc) {
        $p = 'C:\Program Files\Microsoft SQL Server\{0}0\Setup Bootstrap\Log' -f $sync.State.Major
        if ($Unc -and -not $sync.State.IsLocal) { return "\\$($sync.State.ComputerName)\" + $p.Replace(':', '$') }
        $p
    }

    function Get-LogTail([string]$Root, [datetime]$After) {
        try {
            $d = Get-ChildItem $Root -Directory -ErrorAction Stop | Where-Object LastWriteTime -gt $After |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if (-not $d) { return '' }
            $f = Join-Path $d.FullName 'Detail.txt'
            if (-not (Test-Path $f)) { return '' }
            $fs = [IO.File]::Open($f, 'Open', 'Read', 'ReadWrite')
            try {
                $len = [math]::Min(4096, $fs.Length); [void]$fs.Seek(-$len, 'End')
                $buf = New-Object byte[] $len; [void]$fs.Read($buf, 0, $len)
            } finally { $fs.Close() }
            $line = ([Text.Encoding]::Unicode.GetString($buf) -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 1)
            if (-not $line -or $line -match '[^\x09\x20-\x7E]') { $line = ([Text.Encoding]::UTF8.GetString($buf) -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 1) }
            $line = ($line -replace '^\(\d+\)\s+[\d\-: ]+\s+', '').Trim()
            if ($line.Length -gt 140) { $line = $line.Substring(0, 140) + '...' }
            $line
        } catch { '' }
    }

    function Copy-WithProgress([string]$Src, [string]$Dst, [int]$P0, [int]$P1) {
        $dir = Split-Path $Dst
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $srcLen = (Get-Item $Src).Length
        if ((Test-Path $Dst) -and (Get-Item $Dst).Length -eq $srcLen) { Log "Installer already on destination ($Dst)"; return }
        # write to .partial then rename, so an interrupted copy is never mistaken for a staged installer
        $tmp = "$Dst.partial"
        $in = [IO.File]::OpenRead($Src); $out = [IO.File]::Create($tmp)
        $buf = New-Object byte[] (4MB); $done = 0; $last = -1
        Set-Phase 'Copy'
        try {
            while (($n = $in.Read($buf, 0, $buf.Length)) -gt 0) {
                $out.Write($buf, 0, $n); $done += $n
                $pct = [int]($done * 100 / $srcLen)
                if ($pct -ne $last) {
                    $last = $pct
                    Prog ($P0 + [int](($P1 - $P0) * $pct / 100)) ("Copying installer {0}% ({1:N0} of {2:N0} MB)" -f $pct, ($done / 1MB), ($srcLen / 1MB))
                }
            }
        } finally { $in.Close(); $out.Close() }
        if ($done -ne $srcLen) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue; throw "Copy incomplete ($done of $srcLen bytes)." }
        Move-Item -Path $tmp -Destination $Dst -Force
        Log "Copied installer to $Dst"
    }

    function Remove-PatchShare {
        # Works in any session: the share is recognised by the description this tool gives it.
        $c = if ($sync.State.ShareCentral) { $sync.State.ShareCentral } elseif ($a) { $a.Central.Trim() } else { $null }
        $n = if ($sync.State.ShareName) { $sync.State.ShareName } elseif ($a) { $a.ShareName.Trim() } else { $null }
        if (-not $c -or -not $n) { Log 'Enter the central server and share name first.' 'Warn'; return }
        $cim = $null; $p = @{ Name = $n }
        if (-not (Test-IsLocal $c)) { $cim = Get-Cim $c; $p.CimSession = $cim }
        try {
            $sh = Get-SmbShare @p -ErrorAction SilentlyContinue
            if (-not $sh) { Log "Share \\$c\$n does not exist - nothing to remove."; $sync.State.ShareCreated = $false; Set-Info 'Share' 'none'; return }
            if ($sh.Description -ne 'Temporary - SQL patching (SqlPatchGui)' -and -not $sync.State.ShareCreated) {
                Log "Share \\$c\$n was NOT created by this tool (description: '$($sh.Description)', path: $($sh.Path)) - left in place. To remove it anyway: Remove-SmbShare -Name '$n' -CimSession $c -Force" 'Warn'
                return
            }
            Remove-SmbShare @p -Force -ErrorAction Stop
            $sync.State.ShareCreated = $false
            Log "Removed share \\$c\$n (path $($sh.Path); downloaded files stay on disk)."
            Set-Info 'Share' 'removed'
        } finally { if ($cim) { Remove-CimSession $cim } }
    }

    function Clear-DestinationInstallers {
        # Deletes staged installers (*.exe / *.partial) from C:\SQLPatch on the destination - on demand only.
        if (-not $sync.State.ComputerName) { throw 'Run Check first so the tool knows the destination.' }
        $cn = $sync.State.ComputerName
        $busy = if ($sync.State.IsLocal) { @(Get-CimInstance Win32_Process -Filter "Name='setup.exe' OR ExecutablePath LIKE 'C:\\SQLPatch\\%'") }
                else { $c = Get-Cim $cn; try { @(Get-CimInstance -CimSession $c Win32_Process -Filter "Name='setup.exe' OR ExecutablePath LIKE 'C:\\SQLPatch\\%'") } finally { Remove-CimSession $c } }
        if ($busy) { throw "Setup appears to be running on $cn ($(($busy.Name | Select-Object -Unique) -join ', ')) - not deleting installers now." }
        $dir = if ($sync.State.IsLocal) { 'C:\SQLPatch' } elseif ($sync.State.AdminShare) { "\\$cn\C$\SQLPatch" } else { $null }
        if ($dir) {
            $files = @(Get-ChildItem $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.exe', '.partial' })
            if (-not $files) { Log "No installers in C:\SQLPatch on $cn - nothing to delete."; return }
            $mb = [math]::Round(($files | Measure-Object Length -Sum).Sum / 1MB)
            $files | Remove-Item -Force
            Log "Deleted $($files.Count) installer file(s) ($mb MB) from C:\SQLPatch on ${cn}: $($files.Name -join ', ')"
        } elseif ($sync.State.WinRm) {
            $r = Invoke-Command -ComputerName $cn -ScriptBlock {
                $f = @(Get-ChildItem C:\SQLPatch -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.exe', '.partial' })
                $f | Remove-Item -Force; [pscustomobject]@{ Count = $f.Count; Names = ($f.Name -join ', ') }
            }
            if ($r.Count) { Log "Deleted $($r.Count) installer file(s) from C:\SQLPatch on ${cn}: $($r.Names)" } else { Log "No installers in C:\SQLPatch on $cn - nothing to delete." }
        } else { throw "Neither \\$cn\C$ nor WinRM is reachable - run this tool on $cn to clear installers." }
        Set-Info 'Staged on destination' 'cleared'
    }
}

# Worker-only functions (need dbatools + the UI snapshot $a)
$Worker = {
    function Init-Worker {
        if (-not (Get-Module dbatools)) { Log 'Loading dbatools (first run only)...'; Import-Module dbatools -ErrorAction Stop }
        Set-DbatoolsConfig -FullName sql.connection.trustcert -Value ([bool]$a.TrustCert)
    }

    function Get-CuList([int]$Major, [version]$Current) {
        Prog -1 'Loading update list for this SQL version...'
        try { Update-DbaBuildReference -EnableException; Log 'Build reference updated from the internet.' }
        catch { Log "Could not refresh build reference (using the copy shipped with dbatools): $($_.Exception.Message)" 'Warn' }
        $candidates = @(
            (Join-Path $env:LOCALAPPDATA 'dbatools\dbatools-buildref-index.json')
            try { Join-Path (Get-DbatoolsConfigValue -FullName 'Path.DbatoolsData') 'dbatools-buildref-index.json' } catch { }
        ) + @(Get-ChildItem (Get-Module dbatools).ModuleBase -Recurse -Filter 'dbatools-buildref-index.json' -ErrorAction SilentlyContinue | ForEach-Object FullName)
        $file = $candidates | Where-Object { $_ -and (Test-Path $_) } | Get-Item | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $file) { throw 'dbatools build reference file not found.' }
        $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $data = if ($json.Data) { $json.Data } else { $json }
        $versions = foreach ($v in $data.Version) { try { $vv = [version]$v; if ($vv.Major -eq $Major) { $v } } catch { } }
        $builds = Get-DbaBuild -Build $versions -EnableException
        $list = foreach ($b in $builds) {
            $bv = [version]"$($b.Build)"
            $kb = (@($b.KBLevel) | Where-Object { $_ }) -join ','
            $lvl = (@($b.SPLevel, $b.CULevel) | Where-Object { $_ }) -join ' '
            [pscustomobject]@{
                Build = "$bv"; Level = $lvl; Type = ''; KB = $(if ($kb) { "KB$($kb -replace '^KB','')" } else { '' })
                SupportedUntil = $(if ($b.SupportedUntil) { ([datetime]$b.SupportedUntil).ToString('yyyy-MM-dd') } else { '' })
                Status = $(if ((Get-BuildKey $bv) -eq (Get-BuildKey $Current)) { 'Installed' } elseif ((Get-BuildKey $bv) -gt (Get-BuildKey $Current)) { 'Newer' } else { 'Older' })
                Sort = $bv
            }
        }
        # Tidy labels: drop 'RTM' when a CU is present (2017+ has no SPs); within one CU,
        # the lowest build is the CU itself and higher builds are security updates on top of it.
        foreach ($g in ($list | Group-Object Level)) {
            $first = $true
            foreach ($row in ($g.Group | Sort-Object Sort)) {
                $lvl = $row.Level -replace '^RTM\s+(?=CU)', ''
                if ($lvl -eq '') { $lvl = 'RTM' }
                $type = if ($lvl -match 'CU\d+') { 'Cumulative update' } elseif ($lvl -eq 'RTM') { 'Base release' } else { 'Service pack' }
                if ($lvl -match 'CU\d+' -and -not $first) { $lvl = "$lvl + GDR"; $type = 'Security update (CU+GDR)' }
                elseif ($lvl -notmatch 'CU\d+' -and -not $first) { $lvl = "$lvl + GDR"; $type = 'Security update (GDR)' }
                $row.Level = $lvl; $row.Type = $type; $first = $false
            }
        }
        $list = @($list | Sort-Object Sort -Descending)
        Log "Found $($list.Count) builds for this version ($(@($list | Where-Object Status -eq 'Newer').Count) newer than installed)."
        Emit 'CuList' $list
    }

    function Invoke-Download {
        $sel = $a.Selected
        if (-not $sel -or -not $sel.KB) { throw 'Select an update with a KB number in the list first.' }
        $central = $a.Central.Trim(); $lp = $a.CentralPath.Trim().TrimEnd('\')
        if ($lp -notmatch '^[A-Za-z]:\\') { throw 'Download folder must be a local path on the central server, e.g. D:\SQLPatch' }
        $isLocalC = Test-IsLocal $central
        $wp = if ($isLocalC) { $lp } else { "\\$central\" + $lp.Replace(':', '$') }
        Prog 6 "Preparing $lp on $central"
        New-Item -ItemType Directory -Force -Path $wp | Out-Null

        $kbNum = ($sel.KB -split ',')[0] -replace '^KB', ''
        $file = Get-ChildItem $wp -File -Filter "*KB$kbNum*.exe" -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 1
        if ($file) { Log "Installer already downloaded: $($file.Name)" }
        else {
            $t0 = Get-Date
            Set-Phase 'Download'
            Prog -1 "Downloading KB$kbNum to $central ($lp) - can take several minutes"
            Log "Downloading KB$kbNum ..."
            $file = Save-DbaKbUpdate -Name "KB$kbNum" -Path $wp -Architecture x64 -EnableException | Select-Object -First 1
            if (-not $file) { throw "Download of KB$kbNum returned nothing." }
            Log ("Downloaded {0} ({1:N0} MB) in {2}" -f $file.Name, ($file.Length / 1MB), (Fmt-Elapsed $t0))
        }

        # Temporary share
        Prog 20 "Sharing $lp as $($a.ShareName)"
        $sp = @{ Name = $a.ShareName }; $cim = $null
        if (-not $isLocalC) { $cim = Get-Cim $central; $sp.CimSession = $cim }
        try {
            $ex = Get-SmbShare @sp -ErrorAction SilentlyContinue
            if ($ex) {
                if ($ex.Path.TrimEnd('\') -ne $lp) { throw "Share '$($a.ShareName)' already exists on $central but points to $($ex.Path). Pick another share name." }
                if ($ex.Description -eq 'Temporary - SQL patching (SqlPatchGui)') {
                    Log "Share \\$central\$($a.ShareName) is left over from an earlier run of this tool - reusing it (will be removed when finished)."
                    $sync.State.ShareCreated = $true
                } else {
                    Log "Share \\$central\$($a.ShareName) already existed (not created by this tool) - it will NOT be removed." 'Warn'
                    $sync.State.ShareCreated = $false
                }
            } else {
                New-SmbShare @sp -Path $lp -FullAccess 'BUILTIN\Administrators' -ReadAccess 'Everyone' -Description 'Temporary - SQL patching (SqlPatchGui)' -ErrorAction Stop | Out-Null
                $sync.State.ShareCreated = $true
                Log "Created temporary share \\$central\$($a.ShareName) -> $lp"
            }
        } finally { if ($cim) { Remove-CimSession $cim } }
        $sync.State.ShareCentral = $central; $sync.State.ShareName = $a.ShareName
        $sync.State.MediaFile = "\\$central\$($a.ShareName)\$($file.Name)"
        $sync.State.MediaKB = "KB$kbNum"
        if (-not (Test-Path $sync.State.MediaFile)) { throw "Cannot read $($sync.State.MediaFile) from this machine." }
        Set-Info 'Media (share)' $sync.State.MediaFile
        Set-Info 'Share' "\\$central\$($a.ShareName) ($(if ($sync.State.ShareCreated) { 'temporary' } else { 'pre-existing' }))"
    }

    function Format-Hex32([long]$Code) { '0x{0:X8}' -f ([long]$Code -band 0xFFFFFFFFL) }

    function Get-SetupExitMeaning([long]$Code) {
        $u = [long]$Code -band 0xFFFFFFFFL
        if ($u -eq 0) { return 'success' }
        if ($u -eq 3010) { return 'success, Windows reboot required to finish' }
        if (($u -band 0xFFFF) -eq 0x0BC2) { return 'blocked by setup rule: a Windows restart is pending - reboot Windows, then run again' }
        if (($u -band 0xFFFF0000L) -eq 0x84BE0000L) { return 'blocked by a setup rule check (see failed rules in Summary.txt)' }
        if ($u -band 0x80000000L) { return 'setup error (see Exit message in Summary.txt)' }
        'unexpected exit code'
    }

    function Save-SetupLogs([datetime]$After = [datetime]::MinValue, [switch]$SummaryOnly) {
        # Copies Summary.txt (+ the run's Detail/rule report unless -SummaryOnly) to the local log folder.
        # $After = MinValue -> the latest setup run on the destination, whenever it was.
        $r = @{ Summary = ''; KeyLines = @(); LocalFolder = $null; LocalSummary = $null; Final = $null }
        $cn = $sync.State.ComputerName
        $cut = if ($After -eq [datetime]::MinValue) { $After } else { $After.AddMinutes(-1) }
        $dest = Join-Path $a.LogDir ('setup\{0}_{1:yyyyMMdd_HHmmss}' -f $cn, (Get-Date))
        try {
            New-Item -ItemType Directory -Force -Path $dest | Out-Null
            $root = Get-SetupLogRoot -Unc
            if ($sync.State.IsLocal -or $sync.State.AdminShare) {
                Copy-Item (Join-Path $root 'Summary.txt') $dest -ErrorAction Stop
                $run = if ($SummaryOnly) { $null } else { Get-ChildItem $root -Directory | Where-Object LastWriteTime -gt $cut | Sort-Object LastWriteTime -Descending | Select-Object -First 1 }
                if ($run) { Get-ChildItem $run.FullName -File | Where-Object { $_.Name -match '^(Summary|Detail|SystemConfigurationCheck_Report)' -or $_.Extension -eq '.htm' } | Copy-Item -Destination $dest }
            } else {
                $files = Invoke-Command -ComputerName $cn -ScriptBlock {
                    param($root, $after, $summaryOnly)
                    $out = @(@{ Name = 'Summary.txt'; Text = Get-Content (Join-Path $root 'Summary.txt') -Raw -ErrorAction Stop })
                    if ($summaryOnly) { return $out }
                    $run = Get-ChildItem $root -Directory | Where-Object LastWriteTime -gt $after | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                    if ($run) { foreach ($f in (Get-ChildItem $run.FullName -File -Filter 'Detail.txt')) { $out += @{ Name = 'Detail.txt'; Text = Get-Content $f.FullName -Raw } } }
                    $out
                } -ArgumentList (Get-SetupLogRoot), $cut, [bool]$SummaryOnly
                foreach ($f in $files) { Set-Content -Path (Join-Path $dest $f.Name) -Value $f.Text }
            }
            $r.LocalFolder = $dest
            $r.LocalSummary = Join-Path $dest 'Summary.txt'
            $r.Summary = Get-Content $r.LocalSummary -Raw
            Log "Setup logs copied to $dest"
        } catch { Log "Could not copy setup logs from ${cn}: $($_.Exception.Message)" 'Warn' }

        # Key lines: overall result + any failed rules/features
        $text = $r.Summary
        if ($text) {
            if ($text -match 'Final result:\s*(.+)') { $r.Final = $Matches[1].Trim() }
            $keys = foreach ($pat in 'Start time:', 'End time:', 'Final result:', 'Exit code \(Decimal\):', 'Exit facility code:', 'Exit error code:', 'Exit message:', 'Requires (reboot|restart):') {
                ($text -split "`r?`n" | Where-Object { $_ -match "^\s*$pat" } | Select-Object -First 1)
            }
            # failed rules look like: "Rule name: X ... Rule result: Failed ... Message/Corrective Action: ..."
            $blocks = $text -split '(?m)^\s*Rule name:\s*'
            foreach ($bl in $blocks) {
                if ($bl -match 'Rule result:\s*Failed') {
                    $name = ($bl -split "`r?`n")[0].Trim()
                    $msg = if ($bl -match 'Message/Corrective Action:\s*(.+)') { $Matches[1].Trim() } else { '' }
                    $keys += "FAILED RULE $name - $msg"
                }
            }
            $r.KeyLines = @($keys | Where-Object { $_ } | ForEach-Object { $_.Trim() })
        }
        $r
    }

    function Get-StagedInstaller([string]$KbNum) {
        # Returns the destination-local path (C:\SQLPatch\...) of an already staged installer, or $null
        $cn = $sync.State.ComputerName
        $dir = if ($sync.State.IsLocal) { 'C:\SQLPatch' } elseif ($sync.State.AdminShare) { "\\$cn\C$\SQLPatch" } else { $null }
        if ($dir) {
            $f = Get-ChildItem $dir -File -Filter "*KB$KbNum*.exe" -ErrorAction SilentlyContinue | Sort-Object Length -Descending | Select-Object -First 1
            if ($f) { return "C:\SQLPatch\$($f.Name)" }
            return $null
        }
        if ($sync.State.WinRm) {
            $n = Invoke-Command -ComputerName $cn -ScriptBlock { param($k) (Get-ChildItem C:\SQLPatch -File -Filter "*KB$k*.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).Name } -ArgumentList $KbNum
            if ($n) { return "C:\SQLPatch\$n" }
        }
        $null
    }

    function Test-InstallerSignature([string]$ExeOnDest) {
        # Verifies the installer on the destination is intact and signed by Microsoft
        $cn = $sync.State.ComputerName
        Prog -1 'Verifying installer signature'
        $sig = if ($sync.State.IsLocal) { Get-AuthenticodeSignature $ExeOnDest }
               elseif ($sync.State.WinRm) { Invoke-Command -ComputerName $cn -ScriptBlock { param($p) $s = Get-AuthenticodeSignature $p; [pscustomobject]@{ Status = "$($s.Status)"; Subject = $s.SignerCertificate.Subject } } -ArgumentList $ExeOnDest }
               else { Get-AuthenticodeSignature ("\\$cn\" + $ExeOnDest.Replace(':', '$')) }
        $subject = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { $sig.Subject }
        if ("$($sig.Status)" -ne 'Valid' -or $subject -notmatch 'Microsoft') { throw "Installer $ExeOnDest failed signature check (status: $($sig.Status), signer: $subject). Delete it and stage again." }
        Log "Installer signature valid (Microsoft): $ExeOnDest"
    }

    function Invoke-Stage {
        # Copies the installer from the share to C:\SQLPatch on the destination. No SQL/service impact.
        if (-not $sync.State.ComputerName) { throw 'Run Check first so the tool knows the destination.' }
        if ($a.SqlInstance.Trim() -ne $sync.State.SqlInstance) { throw 'Destination changed since the last Check - run Check again.' }
        $cn = $sync.State.ComputerName
        $leaf = Split-Path $sync.State.MediaFile -Leaf
        $exe = "C:\SQLPatch\$leaf"
        if ($sync.State.IsLocal) { Copy-WithProgress $sync.State.MediaFile $exe 25 45 }
        elseif ($sync.State.AdminShare) { Copy-WithProgress $sync.State.MediaFile "\\$cn\C$\SQLPatch\$leaf" 25 45 }
        elseif ($sync.State.WinRm) {
            Prog -1 'Copying installer over WinRM (no progress available)'
            $s = New-PSSession -ComputerName $cn
            try { Invoke-Command -Session $s { New-Item -ItemType Directory -Force C:\SQLPatch | Out-Null }; Copy-Item $sync.State.MediaFile -Destination $exe -ToSession $s -Force }
            finally { Remove-PSSession $s }
            Log "Copied installer to ${cn}: $exe"
        } else { throw "Neither \\$cn\C$ nor WinRM is reachable - run this tool ON $cn to stage." }
        Test-InstallerSignature $exe
        Set-Info 'Staged on destination' "$exe ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"
        $exe
    }

    function Wait-SqlUp([int]$Minutes = 15) {
        $deadline = (Get-Date).AddMinutes($Minutes); $t0 = Get-Date; Set-Phase 'Wait'
        while ((Get-Date) -lt $deadline) {
            try { Invoke-Sql 'SELECT 1 AS ok' | Out-Null; Log "SQL Server accepting connections ($(Get-Date -Format 'HH:mm:ss'))."; return }
            catch { Prog -1 "Waiting for SQL Server to accept connections ($(Fmt-Elapsed $t0))"; Start-Sleep 10 }
        }
        throw "SQL Server not reachable after $Minutes minutes."
    }

    function Send-Restart([string]$ComputerName, [string[]]$Skip = @()) {
        # Tries each channel until one ACCEPTS the restart. Returns the method name used, or $null.
        $cn = $ComputerName; $msg = 'SqlPatchGui: SQL Server patching'
        foreach ($m in 'WinRM', 'CIM', 'WMI', 'RPC') {
            if ($m -in $Skip) { continue }
            if ($m -eq 'WinRM' -and -not $sync.State.WinRm) { continue }
            if ($m -eq 'WMI' -and $PSVersionTable.PSVersion.Major -gt 5) { continue }
            try {
                switch ($m) {
                    'WinRM' {
                        # run shutdown.exe SYNCHRONOUSLY: it registers the restart with Windows and returns at once.
                        # (A background Start-Process gets killed when the WinRM session closes.)
                        $r = Invoke-Command -ComputerName $cn -ErrorAction Stop -ScriptBlock {
                            param($c) $o = & shutdown.exe /r /t 10 /d p:2:17 /c $c 2>&1
                            [pscustomobject]@{ Code = $LASTEXITCODE; Out = "$o" }
                        } -ArgumentList $msg
                        if ($r.Code -notin 0, 1190) { throw "shutdown.exe exit $($r.Code): $($r.Out)" }   # 1190 = already scheduled
                    }
                    'CIM' {
                        $c = Get-Cim $cn
                        try {
                            $res = Get-CimInstance -CimSession $c Win32_OperatingSystem | Invoke-CimMethod -MethodName Win32Shutdown -Arguments @{ Flags = 6 } -ErrorAction Stop
                            if ($res.ReturnValue -ne 0) { throw "Win32Shutdown returned $($res.ReturnValue)" }
                        } finally { Remove-CimSession $c -ErrorAction SilentlyContinue }
                    }
                    'WMI' {
                        Restart-Computer -ComputerName $cn -Force -ErrorAction Stop
                    }
                    'RPC' {
                        $o = & shutdown.exe /r /m "\\$cn" /t 10 /d p:2:17 /c $msg 2>&1
                        if ($LASTEXITCODE -notin 0, 1190) { throw "shutdown.exe /m exit ${LASTEXITCODE}: $o" }
                    }
                }
                Log "Restart accepted via $m."
                return $m
            } catch { Log "  restart via $m failed - $($_.Exception.Message)" 'Warn' }
        }
        $null
    }

    function Invoke-Reboot([string]$Reason) {
        $cn = $sync.State.ComputerName
        $cim = Get-Cim $cn
        $before = (Get-CimInstance -CimSession $cim Win32_OperatingSystem).LastBootUpTime
        Remove-CimSession $cim -ErrorAction SilentlyContinue
        $t0 = Get-Date
        $sync.State.ChangedSinceCheck = $t0.ToString('HH:mm:ss')
        Log "REBOOTING WINDOWS on $cn - $Reason. Previous boot: $before" 'Warn'
        Set-Info 'Reboot requested' $t0.ToString('yyyy-MM-dd HH:mm:ss')
        if ($sync.State.IsLocal) {
            Log 'Destination is this machine - this window will close now. Re-open the tool and press Validate after boot.' 'Warn'
            Start-Sleep 3; Restart-Computer -Force; return
        }

        $used = @(); $m = Send-Restart $cn
        if (-not $m) { throw "Could not restart $cn by any method (WinRM, CIM, WMI, RPC). Restart it from the console/hypervisor, then press Check + Validate." }
        $used += $m

        # Watchdog: if the server hasn't gone down within 3 minutes, the restart didn't take - try the next method.
        Set-Phase 'Reboot'
        $deadline = (Get-Date).AddMinutes(30); $up = $null; $wentDown = $false; $lastTry = Get-Date
        while ((Get-Date) -lt $deadline) {
            Start-Sleep 10
            try {
                $c = Get-Cim $cn
                $b = (Get-CimInstance -CimSession $c Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
                Remove-CimSession $c
                if ($b -gt $before) { $up = $b; break }
                if (-not $wentDown -and ((Get-Date) - $lastTry).TotalSeconds -ge 180) {
                    Log "$cn has not restarted 3 minutes after the request via $($used[-1]) - retrying with another method." 'Warn'
                    $m = Send-Restart $cn -Skip $used
                    if (-not $m) { throw "$cn did not restart and no other restart method worked. Restart it manually, then press Check + Validate." }
                    $used += $m; $lastTry = Get-Date
                }
                Prog -1 ("Waiting for $cn to go down ({0}) - still up, boot time unchanged" -f (Fmt-Elapsed $t0))
            } catch {
                if ($_.Exception.Message -like '*did not restart*') { throw }
                if (-not $wentDown) { $wentDown = $true; Log "$cn is offline (restarting) at $(Get-Date -Format 'HH:mm:ss')" }
                Prog -1 "Waiting for $cn to come back ($(Fmt-Elapsed $t0))"
            }
        }
        if (-not $up) { throw "$cn did not come back within 30 minutes." }
        Log "Windows is back on $cn. New boot time: $up (reboot took $(Fmt-Elapsed $t0))"
        Set-Info 'Last boot (Windows)' $up
        Set-Info 'Reboot completed' (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Wait-SqlUp
        Set-Info 'SQL online after reboot' (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }

    function Invoke-Validate([string]$ExpectedBuild) {
        Set-Phase 'Validate'
        Prog 95 'Validating'
        $si = $sync.State.SqlInstance
        $b = [version](Invoke-Sql "SELECT CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(32)) AS v").v
        $lvl = Invoke-Sql "SELECT CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(32)) AS cu, CAST(SERVERPROPERTY('ProductUpdateReference') AS nvarchar(32)) AS kb"
        Set-Info 'Build (now)' "$b  $($lvl.cu) $($lvl.kb)"
        if ($ExpectedBuild) {
            if ((Get-BuildKey $b) -ge (Get-BuildKey $ExpectedBuild)) { Add-Check PASS 'Post: build' "$b (expected $ExpectedBuild)" }
            else { Add-Check FAIL 'Post: build' "$b is lower than expected $ExpectedBuild - the update is NOT active. Press 'Setup summary' for the setup result; if setup asked for a restart, reboot Windows and Validate again." }
        } else { Add-Check INFO 'Post: build' "$b" }
        try {
            foreach ($s in (Get-DbaService -ComputerName $sync.State.ComputerName -InstanceName $sync.State.InstanceName -EnableException)) {
                $st = if ($s.State -eq 'Running' -or $s.StartMode -ne 'Automatic') { 'PASS' } else { 'FAIL' }
                Add-Check $st "Post: service $($s.ServiceType)" "$($s.ServiceName) is $($s.State) ($($s.StartMode))"
            }
        } catch { Add-Check WARN 'Post: services' "Could not read services: $($_.Exception.Message)" }
        $bad = @(Invoke-Sql "SELECT name, state_desc FROM sys.databases WHERE state_desc NOT IN ('ONLINE','OFFLINE')")
        if ($bad) { Add-Check FAIL 'Post: databases' (($bad | ForEach-Object { "$($_.name)=$($_.state_desc)" }) -join ', ') }
        else { Add-Check PASS 'Post: databases' 'All databases ONLINE (or intentionally OFFLINE)' }
        if ($sync.State.IsHadr) {
            $ag = @(Invoke-Sql "SELECT ag.name, ars.role_desc, ars.synchronization_health_desc AS h FROM sys.dm_hadr_availability_replica_states ars JOIN sys.availability_groups ag ON ag.group_id = ars.group_id WHERE ars.is_local = 1")
            foreach ($g in $ag) {
                $st = if ($g.h -eq 'HEALTHY') { 'PASS' } else { 'WARN' }
                Add-Check $st "Post: AG $($g.name)" "$($g.role_desc), $($g.h) (may take a few minutes to resync)"
            }
        }
        $up = Invoke-Sql 'SELECT sqlserver_start_time AS t FROM sys.dm_os_sys_info'
        Set-Info 'SQL started at' $up.t
        $sync.State.Build = "$b"
        $postFails = @($sync.Checks | Where-Object { $_.Status -eq 'FAIL' -and $_.Check -like 'Post:*' })
        foreach ($x in $postFails) { Log "  POST-CHECK FAIL  $($x.Check): $($x.Detail)" 'Error' }
        Log "Validation finished for $si ($($postFails.Count) post-check FAIL)." $(if ($postFails) { 'Error' } else { 'Info' })
    }
}

# ============================================================================
#  Work items (run in the background runspace)
# ============================================================================
# Guardrail checks as a function so Install can re-run them without touching the update list
$Guardrails = {
function Invoke-Guardrails([switch]$NoCuList) {
    Set-Phase 'Check'
    $sync.Checks.Clear(); Emit 'ChecksClear' $null; $sync.State.CheckFailed = $false
    $si = $a.SqlInstance.Trim()
    if (-not $si) { throw 'Enter the destination SQL Server instance.' }
    $sync.State.SqlInstance = $si
    Prog 5 "Connecting to $si"
    $i = Invoke-Sql @'
SELECT CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS nvarchar(128)) AS ComputerName,
       CAST(ISNULL(SERVERPROPERTY('InstanceName'), 'MSSQLSERVER') AS nvarchar(128)) AS InstanceName,
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(32))       AS Build,
       CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(32))         AS ProductLevel,
       CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(32))   AS UpdateLevel,
       CAST(SERVERPROPERTY('ProductUpdateReference') AS nvarchar(32)) AS UpdateKB,
       CAST(SERVERPROPERTY('Edition') AS nvarchar(128))             AS Edition,
       CAST(ISNULL(SERVERPROPERTY('IsClustered'), 0) AS int)        AS IsClustered,
       CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS int)      AS IsHadr,
       (SELECT sqlserver_start_time FROM sys.dm_os_sys_info)        AS SqlStart
'@
    $build = [version]$i.Build
    $names = @{ 11 = '2012'; 12 = '2014'; 13 = '2016'; 14 = '2017'; 15 = '2019'; 16 = '2022'; 17 = '2025' }
    $cn = $i.ComputerName
    $sync.State.ComputerName = $cn; $sync.State.InstanceName = $i.InstanceName
    $sync.State.Build = "$build"; $sync.State.Major = $build.Major
    $sync.State.IsHadr = [bool]$i.IsHadr; $sync.State.IsLocal = Test-IsLocal $cn
    $uptime = (Get-Date) - $i.SqlStart

    $info = [System.Collections.Generic.List[object]]::new()
    $info.Add(@('SQL instance', $si)); $info.Add(@('Windows host', "$cn$(if ($sync.State.IsLocal) { ' (this machine)' })"))
    $info.Add(@('Version', "SQL Server $($names[$build.Major]) $($i.Edition)"))
    $info.Add(@('Build', "$build")); $info.Add(@('Level', "$($i.ProductLevel) $($i.UpdateLevel) $($i.UpdateKB)".Trim()))
    $info.Add(@('SQL started at', "$($i.SqlStart)  (up $([int]$uptime.TotalDays)d $($uptime.Hours)h)"))

    Prog 15 'Checking latest available build'
    try {
        $t = Test-DbaBuild -SqlInstance $si -Latest -EnableException
        $info.Add(@('Latest available', "$($t.BuildTarget)  $($t.CUTarget)"))
        if ($t.Compliant) { Add-Check PASS 'CU level' "Up to date ($($i.UpdateLevel))" }
        else { Add-Check INFO 'CU level' "Installed $($i.UpdateLevel) $build - latest is $($t.CUTarget) $($t.BuildTarget)" }
        if ($t.SupportedUntil) {
            $info.Add(@('Supported until', ([datetime]$t.SupportedUntil).ToString('yyyy-MM-dd')))
            if ([datetime]$t.SupportedUntil -lt (Get-Date)) { Add-Check WARN 'Support' "This version left support on $(([datetime]$t.SupportedUntil).ToString('yyyy-MM-dd'))" }
        }
    } catch { Add-Check WARN 'CU level' "Could not compare with latest: $($_.Exception.Message)" }

    Prog 25 "Reading Windows info from $cn"
    try {
        $cim = Get-Cim $cn
        try {
            $os = Get-CimInstance -CimSession $cim Win32_OperatingSystem
            $disk = Get-CimInstance -CimSession $cim Win32_LogicalDisk -Filter "DeviceID='C:'"
            $osUp = (Get-Date) - $os.LastBootUpTime
            $info.Add(@('Windows', $os.Caption))
            $info.Add(@('Last boot (Windows)', "$($os.LastBootUpTime)  (up $([int]$osUp.TotalDays)d $($osUp.Hours)h)"))
            $info.Add(@('Memory free', ('{0:N1} of {1:N1} GB' -f ($os.FreePhysicalMemory / 1MB), ($os.TotalVisibleMemorySize / 1MB))))
            $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
            $info.Add(@('C: free', "$freeGB GB"))
            $info.Add(@('Remote protocol', $cim.Protocol))
            if ($freeGB -lt 6) { Add-Check FAIL 'Disk C:' "Only $freeGB GB free - CU setup needs ~6 GB" } else { Add-Check PASS 'Disk C:' "$freeGB GB free" }
            $pend = @()
            if (Test-RegKey $cim 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $pend += 'Windows servicing' }
            if (Test-RegKey $cim 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $pend += 'Windows Update' }
            # SQL setup's own restart rule also looks at pending file rename operations
            try {
                $pfr = Invoke-CimMethod -CimSession $cim -Namespace root/default -ClassName StdRegProv -MethodName GetMultiStringValue `
                    -Arguments @{ hDefKey = [uint32]2147483650; sSubKeyName = 'SYSTEM\CurrentControlSet\Control\Session Manager'; sValueName = 'PendingFileRenameOperations' }
                if ($pfr.ReturnValue -eq 0 -and @($pfr.sValue | Where-Object { $_ }).Count -gt 0) { $pend += "pending file renames ($(@($pfr.sValue | Where-Object { $_ }).Count))" }
            } catch { }
            if ($pend) { Add-Check FAIL 'Pending reboot' "Reboot already pending ($($pend -join ', ')) - reboot Windows BEFORE patching" }
            else { Add-Check PASS 'Pending reboot' 'No reboot pending' }
            if ($osUp.TotalDays -gt 90) { Add-Check WARN 'Windows uptime' "Up $([int]$osUp.TotalDays) days - consider a reboot before patching" }
        } finally { Remove-CimSession $cim }
    } catch { Add-Check WARN 'Windows info' "Could not read via WinRM or DCOM: $($_.Exception.Message)" }

    Prog 40 'Checking remote access'
    if ($sync.State.IsLocal) {
        $sync.State.WinRm = $true; $sync.State.AdminShare = $true
        $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($admin) { Add-Check PASS 'Elevation' 'Running as administrator on the destination' } else { Add-Check FAIL 'Elevation' 'Destination is this machine but PowerShell is not elevated - restart it as administrator' }
    } else {
        try { Test-WSMan -ComputerName $cn -ErrorAction Stop | Out-Null; $sync.State.WinRm = $true; Add-Check PASS 'WinRM' "Reachable - setup can run remotely" }
        catch { $sync.State.WinRm = $false; Add-Check FAIL 'WinRM' "Not reachable - run this tool ON $cn to install (checks still work)" }
        $sync.State.AdminShare = Test-Path "\\$cn\C$"
        if ($sync.State.AdminShare) { Add-Check PASS 'Admin share' "\\$cn\C$ reachable (used to copy installer + read setup logs)" }
        else { Add-Check WARN 'Admin share' "\\$cn\C$ not reachable - will copy over WinRM and no live setup log" }
    }

    Prog 50 'Checking HA configuration'
    if ($i.IsClustered) { Add-Check FAIL 'Failover cluster instance' 'This is an FCI - patch with the pipeline (passive node, move role, other node).' }
    else { Add-Check PASS 'Failover cluster instance' 'Standalone instance' }
    if ($i.IsHadr) {
        $ag = @(Invoke-Sql "SELECT ag.name, ars.role_desc, ars.synchronization_health_desc AS h FROM sys.dm_hadr_availability_replica_states ars JOIN sys.availability_groups ag ON ag.group_id = ars.group_id WHERE ars.is_local = 1")
        foreach ($g in $ag) {
            $info.Add(@("AG $($g.name)", "$($g.role_desc), $($g.h)"))
            if ($g.role_desc -eq 'PRIMARY') { Add-Check FAIL "AG $($g.name)" 'This replica is PRIMARY - fail over to a patched secondary first' }
            elseif ($g.h -ne 'HEALTHY') { Add-Check FAIL "AG $($g.name)" "Replica health is $($g.h) - fix before patching" }
            else { Add-Check PASS "AG $($g.name)" 'SECONDARY and HEALTHY' }
        }
        if (-not $ag) { Add-Check INFO 'Availability groups' 'HADR enabled, no AGs on this instance' }
    }

    Prog 60 'Checking other instances on the host'
    try {
        $eng = @(Get-DbaService -ComputerName $cn -Type Engine -EnableException)
        if ($eng.Count -gt 1) { Add-Check WARN 'Other instances' "$($eng.Count) engine instances on $cn ($($eng.InstanceName -join ', ')). Only $($sync.State.InstanceName) is patched unless 'Patch all instances on host' is ticked (shared components are always updated)." }
        else { Add-Check PASS 'Other instances' 'Only this instance on the host' }
    } catch { Add-Check WARN 'Other instances' "Could not list services: $($_.Exception.Message)" }

    Prog 70 'Checking backups'
    $bk = @(Invoke-Sql @'
SELECT d.name AS Db, d.recovery_model_desc AS Recovery,
  MAX(CASE WHEN b.type IN ('D','I') THEN b.backup_finish_date END) AS LastFullDiff,
  MAX(CASE WHEN b.type = 'L' THEN b.backup_finish_date END) AS LastLog
FROM sys.databases d LEFT JOIN msdb.dbo.backupset b ON b.database_name = d.name
WHERE d.name NOT IN ('tempdb','model') AND d.state_desc = 'ONLINE'
GROUP BY d.name, d.recovery_model_desc
'@)
    $sevFull = if ($i.IsHadr) { 'WARN' } else { 'FAIL' }   # AG backups may run on another replica
    $oldFull = @($bk | Where-Object { $_.LastFullDiff -is [DBNull] -or $_.LastFullDiff -lt (Get-Date).AddHours(-26) })
    $oldLog  = @($bk | Where-Object { $_.Recovery -eq 'FULL' -and ($_.LastLog -is [DBNull] -or $_.LastLog -lt (Get-Date).AddMinutes(-60)) })
    if ($oldFull) { Add-Check $sevFull 'Full/diff backups' ("No full/diff in 26h: " + ($oldFull.Db -join ', ') + $(if ($i.IsHadr) { ' (check other replicas)' })) }
    else { Add-Check PASS 'Full/diff backups' "All $($bk.Count) databases backed up in last 26h" }
    if ($oldLog) { Add-Check WARN 'Log backups' ("FULL recovery, no log backup in 60 min: " + ($oldLog.Db -join ', ')) }
    else { Add-Check PASS 'Log backups' 'OK' }

    Prog 80 'Checking activity'
    $bad = @(Invoke-Sql "SELECT name, state_desc FROM sys.databases WHERE state_desc NOT IN ('ONLINE','OFFLINE')")
    if ($bad) { Add-Check WARN 'Database state' (($bad | ForEach-Object { "$($_.name)=$($_.state_desc)" }) -join ', ') } else { Add-Check PASS 'Database state' 'All online' }
    $open = (Invoke-Sql "SELECT COUNT(*) AS c FROM sys.dm_tran_active_transactions WHERE transaction_type = 1 AND transaction_begin_time < DATEADD(MINUTE,-5,GETDATE())").c
    if ($open -gt 0) { Add-Check WARN 'Open transactions' "$open transaction(s) open > 5 min - patching will roll them back" } else { Add-Check PASS 'Open transactions' 'None long-running' }
    $jobs = @(Invoke-Sql @'
SELECT j.name FROM msdb.dbo.sysjobactivity ja JOIN msdb.dbo.sysjobs j ON j.job_id = ja.job_id
WHERE ja.session_id = (SELECT MAX(session_id) FROM msdb.dbo.syssessions)
  AND ja.start_execution_date IS NOT NULL AND ja.stop_execution_date IS NULL
'@)
    if ($jobs) { Add-Check WARN 'Agent jobs running' ($jobs.name -join ', ') } else { Add-Check PASS 'Agent jobs running' 'None' }
    $users = (Invoke-Sql "SELECT COUNT(*) AS c FROM sys.dm_exec_sessions WHERE is_user_process = 1 AND session_id <> @@SPID").c
    $info.Add(@('User sessions', "$users"))

    Prog 85 'Reading the last setup run on the destination'
    $last = Save-SetupLogs -SummaryOnly
    if ($last.LocalSummary) {
        $when = ($last.KeyLines | Where-Object { $_ -match '^Start time:' } | Select-Object -First 1) -replace '^Start time:\s*', ''
        $msg  = ($last.KeyLines | Where-Object { $_ -match '^Exit message:' } | Select-Object -First 1) -replace '^Exit message:\s*', ''
        $info.Add(@('Last setup run', "$when  -  $($last.Final)"))
        $info.Add(@('Last setup summary (local copy)', $last.LocalSummary))
        $st = if ($last.Final -match '^Passed') { 'PASS' } else { 'INFO' }
        Add-Check $st 'Last setup run' ("{0}: {1}{2}" -f $when, $last.Final, $(if ($msg -and $st -ne 'PASS') { " - $msg" }))
        if ($st -ne 'PASS') { foreach ($l in $last.KeyLines) { Log "  last setup: $l" 'Warn' } }
        Emit 'SetupLogs' @{ Folder = $last.LocalFolder; Summary = $last.LocalSummary; Open = $false }
    } else { $info.Add(@('Last setup run', 'no setup log found on destination')) }

    Emit 'Info' $info
    if (-not $NoCuList) { Get-CuList -Major $build.Major -Current $build }
    $sync.State.CheckTime = (Get-Date).ToString('HH:mm:ss'); $sync.State.ChangedSinceCheck = $null
    $f = @($sync.Checks | Where-Object Status -eq 'FAIL').Count; $w = @($sync.Checks | Where-Object Status -eq 'WARN').Count
    foreach ($x in ($sync.Checks | Where-Object Status -eq 'FAIL')) { Log "  FAIL  $($x.Check): $($x.Detail)" 'Error' }
    Log "Check complete: $f FAIL, $w WARN." $(if ($f) { 'Error' } elseif ($w) { 'Warn' } else { 'Info' })
    Prog 100 "Check complete - $f FAIL, $w WARN"
}
}

$WorkCheck = { Invoke-Guardrails }

$WorkDownload = {
    # Daytime-safe: download -> share -> copy to destination. Does NOT touch SQL Server or services.
    Invoke-Download
    if ($sync.State.ComputerName -and $a.SqlInstance.Trim() -eq $sync.State.SqlInstance) {
        $exe = Invoke-Stage
        Log "Staged and verified on $($sync.State.ComputerName): $exe. Install can run later without the share."
        if ($a.RemoveShare) { Remove-PatchShare }
        Prog 100 'Installer staged on destination - ready to install'
    } else {
        Log 'Media is on the share. Run Check, then Download & stage again to copy it to the destination.' 'Warn'
        Prog 100 'Media ready on the share (not staged - run Check first)'
    }
}

$WorkInstall = {
    $sel = $a.Selected; $t0 = Get-Date
    if (-not $sel) { throw 'Select an update in the list first.' }
    if (-not $sync.State.Build) { throw 'Run Check first.' }
    if ($a.SqlInstance.Trim() -ne $sync.State.SqlInstance) { throw 'Destination changed since the last Check - run Check again.' }
    if ($sync.State.ChangedSinceCheck) {
        Log "Server changed at $($sync.State.ChangedSinceCheck) (after the last Check) - re-running guardrail checks now." 'Warn'
        Invoke-Guardrails -NoCuList
    }
    # Always compare against the LIVE build, not the one from the last Check
    try { $sync.State.Build = (Invoke-Sql "SELECT CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(32)) AS v").v } catch { }
    if ((Get-BuildKey $sel.Build) -eq (Get-BuildKey $sync.State.Build)) { throw "$($sel.Level) $($sel.KB) ($($sel.Build)) is already installed on $($sync.State.SqlInstance). Nothing to do." }
    if ((Get-BuildKey $sel.Build) -lt (Get-BuildKey $sync.State.Build)) { throw "Selected build $($sel.Build) is older than installed $($sync.State.Build). Setup cannot downgrade." }
    $fails = @($sync.Checks | Where-Object { $_.Status -eq 'FAIL' -and $_.Check -notlike 'Post:*' })
    if ($fails -and -not $a.Override) {
        foreach ($f in $fails) { Log "  FAIL  $($f.Check): $($f.Detail)" 'Error' }
        Log 'Setup was NOT started. Previous setup result is under "Last setup run" in Server info - press "Setup summary" to open it.' 'Warn'
        throw ("{0} guardrail check(s) failed at the last Check ({1}): {2}. Fix them and press Check again, or tick 'Override failed checks'." -f $fails.Count, $sync.State.CheckTime, (($fails | ForEach-Object Check) -join ', '))
    }
    if (-not $sync.State.IsLocal -and -not $sync.State.WinRm) { throw "WinRM to $($sync.State.ComputerName) is blocked. Run this tool ON that server." }
    if ($sync.State.CheckFailed) { Log 'Proceeding despite FAIL checks (override ticked).' 'Warn' }
    $cn = $sync.State.ComputerName
    Log "=== Installing $($sel.Level) $($sel.KB) ($($sel.Build)) on $($sync.State.SqlInstance) ==="

    # 1+2. use a pre-staged installer if present, else download/share/copy now
    $kb = ($sel.KB -split ',')[0]; $kbNum = $kb -replace '^KB', ''
    $exe = Get-StagedInstaller $kbNum
    if ($exe) {
        Log "Using installer already staged on ${cn}: $exe (download/share/copy skipped)"
        Test-InstallerSignature $exe
        Prog 45 'Using pre-staged installer'
    } else {
        if (-not $sync.State.MediaFile -or $sync.State.MediaKB -ne $kb -or -not (Test-Path $sync.State.MediaFile)) { Invoke-Download }
        $exe = Invoke-Stage
    }
    $leaf = Split-Path $exe -Leaf

    # 3. setup
    $target = if ($a.AllInstances) { '/AllInstances' } else { "/InstanceName=$($sync.State.InstanceName)" }
    $argStr = "/quiet /IAcceptSQLServerLicenseTerms /Action=Patch $target"
    $logRoot = Get-SetupLogRoot -Unc
    $tSetup = Get-Date
    Log "Running: $exe $argStr"
    Set-Info 'Setup started' $tSetup.ToString('yyyy-MM-dd HH:mm:ss')
    Set-Phase 'Setup'
    if ($sync.State.IsLocal) {
        $p = Start-Process -FilePath $exe -ArgumentList $argStr -PassThru; $null = $p.Handle
        while (-not $p.HasExited) { Start-Sleep 5; Prog -1 ("Installing ({0}) {1}" -f (Fmt-Elapsed $tSetup), (Get-LogTail $logRoot $tSetup)) }
        $exit = $p.ExitCode
    } else {
        $job = Invoke-Command -ComputerName $cn -AsJob -ScriptBlock { param($e, $s) (Start-Process -FilePath $e -ArgumentList $s -Wait -PassThru).ExitCode } -ArgumentList $exe, $argStr
        while ($job.State -eq 'Running') {
            Start-Sleep 5
            $tail = if ($sync.State.AdminShare) { Get-LogTail $logRoot $tSetup } else { '' }
            Prog -1 ("Installing ({0}) {1}" -f (Fmt-Elapsed $tSetup), $tail)
        }
        $exit = Receive-Job $job -ErrorAction Stop | Select-Object -Last 1; Remove-Job $job
    }
    $sync.State.ChangedSinceCheck = (Get-Date).ToString('HH:mm:ss')
    Set-Info 'Setup finished' ("{0} (took {1}), exit code {2}" -f (Get-Date).ToString('HH:mm:ss'), (Fmt-Elapsed $tSetup), $exit)

    Set-Phase 'PostSetup'
    # 4. bring setup logs back to this machine, parse and explain them
    $failed = $exit -notin 0, 3010
    $logs = Save-SetupLogs $tSetup
    $summary = $logs.Summary
    $final = if ($summary -match 'Final result:\s*(.+)') { $Matches[1].Trim() } else { 'unknown' }
    $needReboot = ($exit -eq 3010) -or ($summary -match 'Requires (?:reboot|restart):\s*True')
    $meaning = Get-SetupExitMeaning $exit
    Log ("Setup exit code {0} ({1}) - {2}. Final result: {3}. Reboot required: {4}" -f $exit, (Format-Hex32 $exit), $meaning, $final, $needReboot) $(if ($failed) { 'Error' } else { 'Info' })
    foreach ($l in $logs.KeyLines) { Log "  setup: $l" $(if ($failed) { 'Error' } else { 'Info' }) }
    if ($logs.LocalFolder) {
        Set-Info 'Setup logs (local copy)' $logs.LocalFolder
        Emit 'SetupLogs' @{ Folder = $logs.LocalFolder; Summary = $logs.LocalSummary; Open = $failed }
    }
    if ($failed) { throw ("Setup failed: {0} (exit {1} / {2}). Summary copied to {3}" -f $meaning, $exit, (Format-Hex32 $exit), $(if ($logs.LocalSummary) { $logs.LocalSummary } else { "$logRoot\Summary.txt on $cn" })) }
    Prog 85 'Setup finished'

    # 5. reboot
    if ($needReboot -and $a.AutoReboot) { Invoke-Reboot 'setup requires a reboot' }
    elseif ($a.AlwaysReboot) { Invoke-Reboot '"Always reboot after install" is ticked' }
    elseif ($needReboot) { Log 'Setup says a Windows reboot is REQUIRED. Use "Reboot Windows" when ready.' 'Warn'; Set-Info 'Reboot' 'REQUIRED - not done yet' }
    else { Log 'No Windows reboot required (setup already restarted the SQL services).'; Wait-SqlUp }

    # 6. validate + cleanup
    Invoke-Validate $sel.Build
    if ($a.RemoveShare) { Remove-PatchShare }
    Log "=== Finished in $(Fmt-Elapsed $t0) ==="
    Prog 100 'Patching complete'
}

$WorkReboot     = { if (-not $sync.State.ComputerName) { throw 'Run Check first.' }; Invoke-Reboot 'requested from the Reboot Windows button'; Invoke-Validate $null; Prog 100 'Reboot complete' }
$WorkSqlRestart = {
    if (-not $sync.State.ComputerName) { throw 'Run Check first.' }
    $t0 = Get-Date
    Log "Restarting SQL Server services for $($sync.State.InstanceName) on $($sync.State.ComputerName)" 'Warn'
    Prog -1 'Restarting SQL Server services'
    $sync.State.ChangedSinceCheck = $t0.ToString('HH:mm:ss')
    Restart-DbaService -ComputerName $sync.State.ComputerName -InstanceName $sync.State.InstanceName -Type Engine, Agent -Force -Confirm:$false -EnableException | Out-Null
    Wait-SqlUp
    Log "SQL services restarted in $(Fmt-Elapsed $t0)."
    Set-Info 'SQL services restarted' (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Invoke-Validate $null
    Prog 100 'SQL restart complete'
}
$WorkValidate    = { if (-not $sync.State.SqlInstance) { throw 'Run Check first.' }; Invoke-Validate $(if ($a.Selected) { $a.Selected.Build }); Prog 100 'Validation complete' }
$WorkSummary = {
    if (-not $sync.State.ComputerName) { throw 'Run Check first.' }
    Prog -1 "Fetching latest setup logs from $($sync.State.ComputerName)"
    $l = Save-SetupLogs
    if (-not $l.LocalSummary) { throw "No setup Summary.txt found on $($sync.State.ComputerName) ($(Get-SetupLogRoot))." }
    foreach ($k in $l.KeyLines) { Log "  setup: $k" $(if ($l.Final -match '^Passed') { 'Info' } else { 'Warn' }) }
    Set-Info 'Last setup summary (local copy)' $l.LocalSummary
    Emit 'SetupLogs' @{ Folder = $l.LocalFolder; Summary = $l.LocalSummary; Open = $true }
    Prog 100 'Setup summary copied locally'
}
$WorkClearExe = { Clear-DestinationInstallers; Prog 100 'Installers cleared from destination' }
$WorkRemoveShare = { Remove-PatchShare; Prog 100 'Share removed' }

# ============================================================================
#  Background runner (one persistent runspace, one task at a time)
# ============================================================================
$script:Runspace = [runspacefactory]::CreateRunspace()
$script:Runspace.ApartmentState = 'STA'
$script:Runspace.Open()
$script:Job = $null

function Start-Work([string]$Name, [scriptblock]$Work) {
    if ($sync.Busy) { return }
    $a = Get-UiArgs
    Save-Settings
    $sync.Busy = $true; $sync.Phase = 'Starting'; Set-Buttons $false
    $body = "param(`$sync, `$a)`n`$ErrorActionPreference = 'Stop'`n" + $Helpers.ToString() + "`n" + $Worker.ToString() + "`n" + $Guardrails.ToString() +
            "`ntry {`nInit-Worker`n" + $Work.ToString() + "`n} catch { Log (""FAILED: "" + `$_.Exception.Message) 'Error'; Prog 0 ""$Name failed"" }"
    $ps = [powershell]::Create(); $ps.Runspace = $script:Runspace
    [void]$ps.AddScript($body).AddArgument($sync).AddArgument($a)
    $script:Job = @{ PS = $ps; Handle = $ps.BeginInvoke(); Name = $Name; Start = Get-Date }
    Add-LogLine (Get-Date) 'Info' "---- $Name ----"
}

# Helpers in the UI thread too (Log/Remove-PatchShare used on close)
. $Helpers

# ============================================================================
#  Window
# ============================================================================
$font = New-Object Drawing.Font('Segoe UI', 9)
$mono = New-Object Drawing.Font('Consolas', 9)
$form = New-Object Windows.Forms.Form -Property @{ Text = 'SQL Server single-server patcher'; Size = New-Object Drawing.Size(1320, 860); StartPosition = 'CenterScreen'; Font = $font; MinimumSize = New-Object Drawing.Size(1000, 650) }
$tip = New-Object Windows.Forms.ToolTip -Property @{ AutoPopDelay = 20000 }

function New-Label($text) { New-Object Windows.Forms.Label -Property @{ Text = $text; AutoSize = $true; Margin = New-Object Windows.Forms.Padding(6, 9, 2, 0) } }
function New-Text($w) { New-Object Windows.Forms.TextBox -Property @{ Width = $w; Margin = New-Object Windows.Forms.Padding(2, 5, 10, 0) } }
function New-Button($text, $w = 120) { New-Object Windows.Forms.Button -Property @{ Text = $text; Width = $w; Height = 30; Margin = New-Object Windows.Forms.Padding(4, 4, 4, 4) } }
function New-Check($text, $checked = $false) { New-Object Windows.Forms.CheckBox -Property @{ Text = $text; AutoSize = $true; Checked = $checked; Margin = New-Object Windows.Forms.Padding(6, 7, 6, 0) } }
function New-ListView([string[]]$cols, [int[]]$widths) {
    $lv = New-Object Windows.Forms.ListView -Property @{ View = 'Details'; FullRowSelect = $true; GridLines = $true; Dock = 'Fill'; HideSelection = $false; MultiSelect = $false }
    for ($k = 0; $k -lt $cols.Count; $k++) { [void]$lv.Columns.Add($cols[$k], $widths[$k]) }
    $lv
}

# --- top: inputs ---
$top = New-Object Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; Height = 76; Padding = New-Object Windows.Forms.Padding(4); WrapContents = $true }
$txtDest    = New-Text 210; $txtCentral = New-Text 170; $txtPath = New-Text 170; $txtShare = New-Text 110
$chkTrust   = New-Check 'Trust server certificate' $true
$btnCheck   = New-Button 'Check' 100
$top.Controls.AddRange(@((New-Label 'Destination SQL instance:'), $txtDest, (New-Label 'Central server:'), $txtCentral,
                         (New-Label 'Download folder on central:'), $txtPath, (New-Label 'Share name:'), $txtShare))
$top.SetFlowBreak($txtShare, $true)
$top.Controls.AddRange(@($btnCheck, $chkTrust))
$tip.SetToolTip($txtDest, 'SQL instance to patch: HOST or HOST\INSTANCE. Use this machine''s name to patch locally.')
$tip.SetToolTip($txtCentral, 'Server that stores the downloaded installer. Use this machine''s name if you run the tool on the central server.')
$tip.SetToolTip($txtPath, 'Local path ON the central server, e.g. D:\SQLPatch. It is shared temporarily while patching.')
$tip.SetToolTip($txtShare, 'Temporary share name. A trailing $ hides it from browsing.')
$tip.SetToolTip($chkTrust, 'Accept self-signed SQL Server certificates (dbatools 2.x validates them by default).')

# --- middle: info/checks (left) + updates (right) ---
$split = New-Object Windows.Forms.SplitContainer -Property @{ Dock = 'Fill'; SplitterDistance = 720 }
$tabs = New-Object Windows.Forms.TabControl -Property @{ Dock = 'Fill' }
$tabInfo = New-Object Windows.Forms.TabPage -Property @{ Text = 'Server info' }
$tabChecks = New-Object Windows.Forms.TabPage -Property @{ Text = 'Guardrail checks' }
$lvInfo = New-ListView @('Item', 'Value') @(190, 500)
$lvChecks = New-ListView @('Status', 'Check', 'Detail') @(60, 170, 480)
$tabInfo.Controls.Add($lvInfo); $tabChecks.Controls.Add($lvChecks)
$tabs.TabPages.AddRange(@($tabChecks, $tabInfo))
$split.Panel1.Controls.Add($tabs)

$grpCu = New-Object Windows.Forms.GroupBox -Property @{ Text = 'Available updates for this SQL version (select one)'; Dock = 'Fill' }
$cuBar = New-Object Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; Height = 30 }
$chkNewer = New-Check 'Show only installed + newer' $true
$lblSel = New-Object Windows.Forms.Label -Property @{ Text = 'Selected: (none)'; AutoSize = $true; Margin = New-Object Windows.Forms.Padding(12, 9, 0, 0); Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold) }
$cuBar.Controls.AddRange(@($chkNewer, $lblSel))
$lvCu = New-ListView @('Build', 'Update', 'Type', 'KB', 'Supported until', 'Status') @(95, 95, 165, 85, 90, 70)
$grpCu.Controls.Add($lvCu); $grpCu.Controls.Add($cuBar); $lvCu.BringToFront()
$split.Panel2.Controls.Add($grpCu)

# --- bottom: actions, options, progress, log ---
$bottom = New-Object Windows.Forms.Panel -Property @{ Dock = 'Bottom'; Height = 300 }
$bar = New-Object Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; Height = 40 }
$btnDownload = New-Button 'Download, share && stage' 170
$btnInstall  = New-Button 'Install selected' 130
$btnReboot   = New-Button 'Reboot Windows' 120
$btnSqlRst   = New-Button 'Restart SQL service' 140
$btnValidate = New-Button 'Validate' 90
$btnRmShare  = New-Button 'Remove share' 110
$btnLogs     = New-Button 'Open log folder' 120
$btnSummary  = New-Button 'Setup summary' 115
$btnClearExe = New-Button 'Clear installers' 120
$btnInstall.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)
$bar.Controls.AddRange(@($btnDownload, $btnInstall, $btnReboot, $btnSqlRst, $btnValidate, $btnRmShare, $btnClearExe, $btnLogs, $btnSummary))
$tip.SetToolTip($btnDownload, 'Safe during business hours: downloads to the central folder, shares it, copies the installer to C:\SQLPatch on the destination and verifies its signature. SQL Server is not touched.')
$tip.SetToolTip($btnReboot, 'Restarts WINDOWS on the destination (the whole VM/server). Needed only when setup says so, or when a reboot was already pending.')
$tip.SetToolTip($btnSqlRst, 'Restarts only the SQL Server + Agent services. CU setup already does this itself; use it for troubleshooting.')

$opts = New-Object Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; Height = 30 }
$chkAll        = New-Check 'Patch all instances on host' $false
$chkAutoReboot = New-Check 'Auto-reboot Windows if setup requires it' $true
$chkAlways     = New-Check 'Always reboot Windows after install' $false
$chkRmShare    = New-Check 'Remove share when finished' $true
$chkOverride   = New-Check 'Override failed checks' $false
$chkOverride.ForeColor = [Drawing.Color]::DarkRed
$opts.Controls.AddRange(@($chkAll, $chkAutoReboot, $chkAlways, $chkRmShare, $chkOverride))

$progPanel = New-Object Windows.Forms.FlowLayoutPanel -Property @{ Dock = 'Top'; Height = 30 }
$prog = New-Object Windows.Forms.ProgressBar -Property @{ Width = 380; Height = 20; Margin = New-Object Windows.Forms.Padding(6, 5, 6, 0); MarqueeAnimationSpeed = 30 }
$lblStatus = New-Object Windows.Forms.Label -Property @{ Text = 'Ready'; AutoSize = $true; Margin = New-Object Windows.Forms.Padding(4, 7, 0, 0) }
$lblElapsed = New-Object Windows.Forms.Label -Property @{ Text = ''; AutoSize = $true; Margin = New-Object Windows.Forms.Padding(12, 7, 0, 0); ForeColor = [Drawing.Color]::DimGray }
$progPanel.Controls.AddRange(@($prog, $lblElapsed, $lblStatus))

$log = New-Object Windows.Forms.RichTextBox -Property @{ Dock = 'Fill'; ReadOnly = $true; Font = $mono; BackColor = [Drawing.Color]::White; WordWrap = $false }
$bottom.Controls.AddRange(@($log, $progPanel, $opts, $bar))
$log.BringToFront()

$form.Controls.AddRange(@($split, $bottom, $top))
$split.BringToFront()

# ============================================================================
#  UI functions
# ============================================================================
$actionButtons = @($btnCheck, $btnDownload, $btnInstall, $btnReboot, $btnSqlRst, $btnValidate, $btnRmShare, $btnSummary, $btnClearExe)
function Set-Buttons([bool]$Enabled) { foreach ($b in $actionButtons) { $b.Enabled = $Enabled } }

function Add-LogLine([datetime]$Time, [string]$Level, [string]$Msg) {
    $line = '{0:HH:mm:ss} {1,-5} {2}' -f $Time, $Level.ToUpper(), $Msg
    $log.SelectionStart = $log.TextLength; $log.SelectionLength = 0
    $log.SelectionColor = switch ($Level) { 'Error' { [Drawing.Color]::Firebrick } 'Warn' { [Drawing.Color]::DarkOrange } default { [Drawing.Color]::Black } }
    $log.AppendText($line + "`r`n"); $log.ScrollToCaret()
    Add-Content -Path $LogFile -Value ('{0:yyyy-MM-dd HH:mm:ss} {1,-5} {2}' -f $Time, $Level.ToUpper(), $Msg)
}

function Set-InfoRow([string]$Name, [string]$Value) {
    foreach ($it in $lvInfo.Items) { if ($it.Text -eq $Name) { $it.SubItems[1].Text = $Value; return } }
    $li = New-Object Windows.Forms.ListViewItem($Name); [void]$li.SubItems.Add($Value); [void]$lvInfo.Items.Add($li)
}

$script:CuAll = @(); $script:Selected = $null
function Show-CuList {
    $lvCu.BeginUpdate(); $lvCu.Items.Clear()
    foreach ($c in $script:CuAll) {
        if ($chkNewer.Checked -and $c.Status -eq 'Older') { continue }
        $li = New-Object Windows.Forms.ListViewItem($c.Build)
        foreach ($v in @($c.Level, $c.Type, $c.KB, $c.SupportedUntil, $c.Status)) { [void]$li.SubItems.Add("$v") }
        if ($c.Type -like 'Security*') { $li.ForeColor = [Drawing.Color]::MidnightBlue }
        $li.Tag = $c
        switch ($c.Status) {
            'Installed' { $li.BackColor = [Drawing.Color]::FromArgb(220, 240, 220); $li.Font = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold) }
            'Older'     { $li.ForeColor = [Drawing.Color]::Gray }
        }
        if (-not $c.KB) { $li.ForeColor = [Drawing.Color]::Gray }
        [void]$lvCu.Items.Add($li)
    }
    $lvCu.EndUpdate()
}

function Get-UiArgs {
    @{
        SqlInstance = $txtDest.Text; Central = $txtCentral.Text; CentralPath = $txtPath.Text; ShareName = $txtShare.Text
        TrustCert = $chkTrust.Checked; AllInstances = $chkAll.Checked; AutoReboot = $chkAutoReboot.Checked
        AlwaysReboot = $chkAlways.Checked; RemoveShare = $chkRmShare.Checked
        Override = $chkOverride.Checked; Selected = $script:Selected; LogDir = $LogDir
    }
}

function Save-Settings {
    try {
        @{ Dest = $txtDest.Text; Central = $txtCentral.Text; Path = $txtPath.Text; Share = $txtShare.Text; Trust = $chkTrust.Checked } |
            ConvertTo-Json | Set-Content $SettingsFile
    } catch { }
}

# load settings
$txtCentral.Text = $env:COMPUTERNAME; $txtPath.Text = 'D:\SQLPatch'; $txtShare.Text = 'SQLPatch$'
if (Test-Path $SettingsFile) {
    try {
        $s = Get-Content $SettingsFile -Raw | ConvertFrom-Json
        if ($s.Dest) { $txtDest.Text = $s.Dest }; if ($s.Central) { $txtCentral.Text = $s.Central }
        if ($s.Path) { $txtPath.Text = $s.Path }; if ($s.Share) { $txtShare.Text = $s.Share }
        $chkTrust.Checked = [bool]$s.Trust
    } catch { }
}

# ============================================================================
#  Events
# ============================================================================
$timer = New-Object Windows.Forms.Timer -Property @{ Interval = 250 }
$timer.Add_Tick({
  try {
    $item = $null
    while ($sync.Queue.TryDequeue([ref]$item)) {
        switch ($item.T) {
            'Log'  { Add-LogLine $item.Time $item.Level $item.Msg }
            'Prog' {
                if ($item.Pct -lt 0) { $prog.Style = 'Marquee' } else { $prog.Style = 'Continuous'; $prog.Value = [math]::Min(100, [math]::Max(0, $item.Pct)) }
                $lblStatus.Text = $item.Text
            }
            'Info' { $lvInfo.BeginUpdate(); $lvInfo.Items.Clear(); foreach ($r in $item.Data) { Set-InfoRow $r[0] $r[1] }; $lvInfo.EndUpdate() }
            'InfoSet' { Set-InfoRow $item.Data[0] $item.Data[1] }
            'ChecksClear' { $lvChecks.Items.Clear() }
            'Check' {
                $r = $item.Data
                $li = New-Object Windows.Forms.ListViewItem($r.Status); [void]$li.SubItems.Add($r.Check); [void]$li.SubItems.Add($r.Detail)
                $li.ForeColor = switch ($r.Status) { 'FAIL' { [Drawing.Color]::Firebrick } 'WARN' { [Drawing.Color]::DarkOrange } 'PASS' { [Drawing.Color]::DarkGreen } default { [Drawing.Color]::Black } }
                [void]$lvChecks.Items.Add($li); $tabs.SelectedTab = $tabChecks
            }
            'SetupLogs' {
                $script:LastSetupLogs = $item.Data
                if ($item.Data.Open -and $item.Data.Summary -and (Test-Path $item.Data.Summary)) { Start-Process notepad.exe $item.Data.Summary }
            }
            'CuList' { $script:CuAll = @($item.Data); $script:Selected = $null; $lblSel.Text = 'Selected: (none)'; Show-CuList }
        }
    }
    if ($script:Job) {
        $lblElapsed.Text = "[$($script:Job.Name) $('{0:hh\:mm\:ss}' -f ((Get-Date) - $script:Job.Start))]"
        if ($script:Job.Handle.IsCompleted) {
            try { [void]$script:Job.PS.EndInvoke($script:Job.Handle) } catch { Add-LogLine (Get-Date) 'Error' "Worker error: $($_.Exception.Message)" }
            foreach ($e in $script:Job.PS.Streams.Error) { Add-LogLine (Get-Date) 'Error' "$e" }
            $script:Job.PS.Dispose(); $script:Job = $null
            $sync.Busy = $false; Set-Buttons $true
            if ($prog.Style -eq 'Marquee') { $prog.Style = 'Continuous' }
        }
    }
  } catch [System.Management.Automation.PipelineStoppedException] {
    $timer.Stop(); $form.Close()
  } catch {
    # never let a UI refresh error turn into a popup loop
    try { Add-Content -Path $LogFile -Value ("{0:yyyy-MM-dd HH:mm:ss} ERROR UI refresh: {1}" -f (Get-Date), $_.Exception.Message) } catch { }
  }
})

$lvCu.Add_SelectedIndexChanged({
    if ($lvCu.SelectedItems.Count -gt 0) {
        $script:Selected = $lvCu.SelectedItems[0].Tag
        $lblSel.Text = "Selected: $($script:Selected.Level) $($script:Selected.KB) ($($script:Selected.Build))"
    }
})
$chkNewer.Add_CheckedChanged({ Show-CuList })

function Confirm([string]$Text, [string]$Title = 'Confirm') {
    [Windows.Forms.MessageBox]::Show($form, $Text, $Title, 'YesNo', 'Warning') -eq 'Yes'
}

$btnCheck.Add_Click({ Start-Work 'Check' $WorkCheck })
$btnDownload.Add_Click({ Start-Work 'Download' $WorkDownload })
$btnInstall.Add_Click({
    if (-not $script:Selected) { [void][Windows.Forms.MessageBox]::Show($form, 'Select an update in the list on the right first.', 'Install'); return }
    $fails = @($lvChecks.Items | Where-Object { $_.Text -eq 'FAIL' }).Count
    $msg = "Install $($script:Selected.Level) $($script:Selected.KB) ($($script:Selected.Build))`non $($txtDest.Text)?`n`n" +
           "SQL Server services WILL restart (short outage).`n" +
           "Windows reboot: $(if ($chkAlways.Checked) { 'always' } elseif ($chkAutoReboot.Checked) { 'automatic, only if setup requires it' } else { 'manual' })`n" +
           "Failed checks: $fails$(if ($fails -and $chkOverride.Checked) { ' (OVERRIDDEN)' })"
    if (Confirm $msg 'Install update') { Start-Work 'Install' $WorkInstall }
})
$btnReboot.Add_Click({
    $msg = "Reboot WINDOWS on $(if ($sync.State.ComputerName) { $sync.State.ComputerName } else { $txtDest.Text })?`n`n" +
           "This restarts the whole server/VM (all services, all instances).`n" +
           "Use it when setup reported 'reboot required' or a reboot was already pending.`n" +
           "For a normal CU you usually do NOT need it - setup restarts SQL itself."
    if (Confirm $msg 'Reboot Windows') { Start-Work 'Reboot' $WorkReboot }
})
$btnSqlRst.Add_Click({
    $msg = "Restart SQL Server + Agent services for $($txtDest.Text)?`n`nThis is a short outage for this instance only. Windows keeps running."
    if ($sync.State.IsHadr) { $msg += "`n`nThis instance is in an AG - make sure it is not the primary." }
    if (Confirm $msg 'Restart SQL Server') { Start-Work 'SQL restart' $WorkSqlRestart }
})
$btnValidate.Add_Click({ Start-Work 'Validate' $WorkValidate })
$btnRmShare.Add_Click({ Start-Work 'Remove share' $WorkRemoveShare })
$btnClearExe.Add_Click({
    if (Confirm "Delete all staged installers (*.exe) from C:\SQLPatch on $(if ($sync.State.ComputerName) { $sync.State.ComputerName } else { $txtDest.Text })?`n`nDo this after patching is finished. The files on the central server are not touched." 'Clear installers') { Start-Work 'Clear installers' $WorkClearExe }
})
$tip.SetToolTip($btnClearExe, 'Deletes staged installers from C:\SQLPatch on the destination. Refuses while setup is running.')
$btnLogs.Add_Click({ Start-Process explorer.exe $LogDir })

$btnSummary.Add_Click({ Start-Work 'Setup summary' $WorkSummary })
$tip.SetToolTip($btnSummary, 'Copies the latest setup Summary.txt (and that run''s Detail.txt) from the destination to this machine and opens it.')

$form.Add_FormClosing({
    param($sender, $e)
    if ($sync.Busy -and -not (Confirm $(if ($sync.Phase -eq 'Setup') { 'SQL setup is RUNNING on the server. Closing this window now can abort setup half-way. Wait until it finishes. Close anyway?' } else { "A task is still running ($($sync.Phase)). Close anyway?" }))) { $e.Cancel = $true; return }
    if ($sync.State.ShareCreated -and (Confirm "The temporary share \\$($sync.State.ShareCentral)\$($sync.State.ShareName) still exists. Remove it now?" 'Remove share')) {
        try { Remove-PatchShare } catch { [void][Windows.Forms.MessageBox]::Show("Could not remove share: $($_.Exception.Message)") }
    }
    Save-Settings
    $timer.Stop()
})

$timer.Start()
Add-LogLine (Get-Date) 'Info' "Log file: $LogFile"
Add-LogLine (Get-Date) 'Info' 'Enter the destination and press Check.'
try {
    [void]$form.ShowDialog()
} finally {
    # always clean up, however the window ended
    try { $timer.Stop(); $timer.Dispose() } catch { }
    try { if ($script:Job) { $script:Job.PS.Stop(); $script:Job.PS.Dispose() } } catch { }
    try { $script:Runspace.Close(); $script:Runspace.Dispose() } catch { }
    try { if ($null -ne $script:OldCtrlC) { [Console]::TreatControlCAsInput = $script:OldCtrlC } } catch { }
    try { $form.Dispose() } catch { }
}

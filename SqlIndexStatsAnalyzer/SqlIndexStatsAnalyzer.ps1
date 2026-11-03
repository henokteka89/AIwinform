#Requires -Version 5.1
<#
.SYNOPSIS
    SQL Server Index & Statistics Analyzer: a READ-ONLY WinForms tool.

.DESCRIPTION
    Pick a server (typed, cached, or loaded from a central inventory table), pick a database,
    search tables (fuzzy, handles tens of thousands), then review the table's indexes and
    statistics in plain English, with priorities, histograms and maintenance scripts.

    The tool never changes anything. Every statement passes a read-only guard. Maintenance
    commands are generated as text for review and are never executed. See README.md.

.EXAMPLE
    powershell.exe -STA -ExecutionPolicy Bypass -File .\SqlIndexStatsAnalyzer.ps1
.EXAMPLE
    .\SqlIndexStatsAnalyzer.ps1 -Server SQLAG01 -Database Sales
.EXAMPLE
    . .\SqlIndexStatsAnalyzer.ps1 -NoGui      # just load the read-only functions, e.g. Test-ReadOnlySql
.NOTES
    Single self-contained file: GUI + engine. Windows PowerShell 5.1, no modules to install.
#>
param(
    [string]$Server,
    [string]$Database,
    [switch]$NoGui      # load the functions only (for scripting and tests):  . .\SqlIndexStatsAnalyzer.ps1 -NoGui
)

#region ================================================================== CORE (read-only engine)
# Everything that talks to SQL Server lives in this block. It is loaded as an in-memory module
# here and again inside the background runspace, so both use the same guarded code.
$script:CoreBlock = {
    <#
        IndexStatsAnalyzer.psm1
        ---------------------------------------------------------------------------
        Read-only SQL Server index & statistics analyzer: data access, guardrails,
        scoring and script generation. Used by Start-IndexStatsAnalyzer.ps1 (GUI),
        but every function also works from a plain PowerShell prompt.

        SAFETY MODEL (read this before changing anything)
          1. Every statement goes through Invoke-SafeQuery, which runs
             Assert-ReadOnlySql first. Only SELECT / WITH / DECLARE / harmless SET
             options / DBCC SHOW_STATISTICS can pass. DML, DDL, EXEC, maintenance
             commands, SELECT INTO, OPENROWSET, xp_/sp_ calls, and so on are refused
             before any connection is opened.
          2. Each batch runs with LOCK_TIMEOUT, DEADLOCK_PRIORITY LOW and
             READ UNCOMMITTED, so this tool gives way to the workload and never waits
             behind it.
          3. Row counts and sizes come from metadata (sys.dm_db_partition_stats).
             There is no COUNT(*) and no table data is read.
          4. sys.dm_db_index_physical_stats runs per index, LIMITED or SAMPLED only
             (never DETAILED), with OPTION (MAXDOP 1). It is refused for indexes
             above a size cap and while the server is busy.
          5. Maintenance commands (REBUILD / REORGANIZE / UPDATE STATISTICS) are
             only GENERATED AS TEXT. This module has no code path that executes
             them. Generated scripts carry a server/database check, and production
             scripts carry a safety latch that must be removed by hand.
    #>

    $script:ToolName    = 'SqlIndexStatsAnalyzer'
    $script:ToolVersion = '1.0.0'

    #region ---------------------------------------------------------------- Fuzzy matcher (C#)
    if (-not ('IsaTableMatcher' -as [type])) {
        Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class IsaTableMatcher
{
    // Returns indexes into the input arrays, best match first.
    public static int[] Search(string[] schemas, string[] names, string query, int top)
    {
        if (names == null) return new int[0];
        if (top <= 0) top = 500;
        string q = Normalize(query);
        if (q.Length == 0)
        {
            int n = Math.Min(top, names.Length);
            int[] all = new int[n];
            for (int i = 0; i < n; i++) all[i] = i;
            return all;
        }

        string schemaQ = null;
        int dot = q.LastIndexOf('.');
        if (dot > 0 && dot < q.Length - 1) { schemaQ = q.Substring(0, dot); q = q.Substring(dot + 1); }
        else if (dot == q.Length - 1) { schemaQ = q.Substring(0, dot); q = ""; }

        Regex wildcard = null;
        if (q.IndexOfAny(new char[] { '*', '?', '%' }) >= 0)
        {
            string pattern = "^" + Regex.Escape(q).Replace("\\*", ".*").Replace("%", ".*").Replace("\\?", ".") + "$";
            wildcard = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        string qc = Compact(q);
        string[] tokens = q.Split(new char[] { ' ', '_', '-' }, StringSplitOptions.RemoveEmptyEntries);
        List<KeyValuePair<int, int>> hits = new List<KeyValuePair<int, int>>();

        for (int i = 0; i < names.Length; i++)
        {
            string s = schemas == null || schemas[i] == null ? "" : schemas[i].ToLowerInvariant();
            if (schemaQ != null && s.IndexOf(schemaQ, StringComparison.Ordinal) < 0) continue;
            string n = names[i] == null ? "" : names[i].ToLowerInvariant();
            int score = q.Length == 0 ? 500 : Score(n, q, qc, tokens, wildcard);
            if (score > 0) hits.Add(new KeyValuePair<int, int>(i, score));
        }

        hits.Sort(delegate (KeyValuePair<int, int> a, KeyValuePair<int, int> b)
        {
            int c = b.Value.CompareTo(a.Value);
            if (c != 0) return c;
            c = names[a.Key].Length.CompareTo(names[b.Key].Length);
            if (c != 0) return c;
            return string.Compare(names[a.Key], names[b.Key], StringComparison.OrdinalIgnoreCase);
        });

        int count = Math.Min(top, hits.Count);
        int[] result = new int[count];
        for (int i = 0; i < count; i++) result[i] = hits[i].Key;
        return result;
    }

    public static int CountMatches(string[] schemas, string[] names, string query)
    {
        return Search(schemas, names, query, int.MaxValue).Length;
    }

    public static int Score(string n, string q, string qc, string[] tokens, Regex wildcard)
    {
        if (wildcard != null) return wildcard.IsMatch(n) ? 900 : 0;
        if (n == q) return 1000;
        if (n.StartsWith(q, StringComparison.Ordinal)) return 900 - Math.Min(99, n.Length - q.Length);
        int idx = n.IndexOf(q, StringComparison.Ordinal);
        if (idx >= 0) return 800 - Math.Min(99, idx);

        string nc = Compact(n);
        if (qc.Length > 0 && nc.IndexOf(qc, StringComparison.Ordinal) >= 0) return 700;

        if (tokens.Length > 1)
        {
            bool allTokens = true;
            foreach (string t in tokens) { if (nc.IndexOf(Compact(t), StringComparison.Ordinal) < 0) { allTokens = false; break; } }
            if (allTokens) return 650;
        }

        // Typo tolerant: best edit distance of the query against any substring of the name (Sellers).
        if (qc.Length >= 3)
        {
            int d = SubstringEditDistance(qc, nc);
            double sim = 1.0 - (double)d / qc.Length;
            double needed = qc.Length <= 4 ? 0.75 : 0.70;
            if (sim >= needed) return 300 + (int)(sim * 200);
        }

        // Abbreviation: query letters appear in order (e.g. "ordhdr" -> "orderheader").
        if (qc.Length >= 3 && IsSubsequence(qc, nc)) return 200 + (int)(100.0 * qc.Length / Math.Max(1, nc.Length));

        return 0;
    }

    public static string Normalize(string s)
    {
        if (s == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (char ch in s.Trim().ToLowerInvariant())
        {
            if (ch == '[' || ch == ']' || ch == '"' || ch == '\'') continue;
            sb.Append(ch);
        }
        return sb.ToString();
    }

    public static string Compact(string s)
    {
        StringBuilder sb = new StringBuilder(s.Length);
        foreach (char ch in s) { if (char.IsLetterOrDigit(ch)) sb.Append(ch); }
        return sb.ToString();
    }

    public static bool IsSubsequence(string q, string n)
    {
        int j = 0;
        for (int i = 0; i < n.Length && j < q.Length; i++) if (n[i] == q[j]) j++;
        return j == q.Length;
    }

    public static int SubstringEditDistance(string q, string n)
    {
        int m = q.Length;
        int[] prev = new int[m + 1];
        int[] cur = new int[m + 1];
        for (int i = 0; i <= m; i++) prev[i] = i;
        int best = prev[m];
        for (int j = 1; j <= n.Length; j++)
        {
            cur[0] = 0; // the match may start anywhere in the name
            for (int i = 1; i <= m; i++)
            {
                int cost = q[i - 1] == n[j - 1] ? 0 : 1;
                int v = Math.Min(Math.Min(prev[i] + 1, cur[i - 1] + 1), prev[i - 1] + cost);
                cur[i] = v;
            }
            if (cur[m] < best) best = cur[m];
            int[] t = prev; prev = cur; cur = t;
        }
        return best;
    }
}
'@
    }
    #endregion

    #region ---------------------------------------------------------------- Settings
    function Get-AnalyzerSettingsPath {
        $base = $env:APPDATA
        if ([string]::IsNullOrEmpty($base)) { $base = Join-Path $HOME '.config' }
        $dir = Join-Path $base $script:ToolName
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Join-Path $dir 'settings.json'
    }

    function Get-AnalyzerDefaultSettings {
        [ordered]@{
            CentralServer          = ''
            CentralDatabase        = 'DBAPatch'
            CentralQuery           = 'SELECT SqlInstance AS ServerName FROM dbo.PatchInstance ORDER BY SqlInstance'
            Servers                = @()          # cached: @{ Name=''; Environment='' }
            RecentServers          = @()
            ProdNamePattern        = '(?i)(prd|prod)'
            NonProdNamePattern     = '(?i)(dev|tst|test|qa|uat|lab|sbx|sandbox|stg|stage)'
            UnknownIsProduction    = $true
            Encrypt                = $false
            TrustServerCertificate = $true
            ReadOnlyIntent         = $true
            CommandTimeoutSec      = 30
            ScanTimeoutSec         = 300
            LockTimeoutMs          = 3000
            MinPageCount           = 1000
            ReorgThreshold         = 5
            RebuildThreshold       = 30
            PageDensityThreshold   = 70
            MaxScanGBProduction    = 20
            MaxScanGBNonProduction = 200
            MaxSampledScanGB       = 2
            MaxAvgRunnableTasks    = 2.0
            DefaultMaxDop          = 4
            LowPriorityWaitMinutes = 5
            ResumableMaxMinutes    = 240
            MaxTablesDisplayed     = 2000
        }
    }

    function Get-AnalyzerSettings {
        $defaults = Get-AnalyzerDefaultSettings
        $path = Get-AnalyzerSettingsPath
        if (Test-Path -LiteralPath $path) {
            try {
                $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
                foreach ($p in $saved.PSObject.Properties) {
                    if ($defaults.Contains($p.Name)) { $defaults[$p.Name] = $p.Value }
                }
            } catch {
                Write-Warning "Settings file is unreadable, using defaults: $($_.Exception.Message)"
            }
        }
        # normalise arrays coming back from JSON
        $defaults['Servers']       = @($defaults['Servers'] | Where-Object { $_ })
        $defaults['RecentServers'] = @($defaults['RecentServers'] | Where-Object { $_ })
        $defaults
    }

    function Save-AnalyzerSettings {
        param([Parameter(Mandatory)] $Settings)
        $path = Get-AnalyzerSettingsPath
        $Settings | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
    }

    function Get-ServerEnvironment {
        <# Production unless proven otherwise (by the central list or a non-prod name pattern). #>
        param([string]$ServerName, $Settings)
        if ([string]::IsNullOrWhiteSpace($ServerName)) { return 'Production (assumed)' }
        $entry = @($Settings.Servers | Where-Object { $_ -and $_.Name -eq $ServerName }) | Select-Object -First 1
        if ($entry -and $entry.Environment) {
            $e = [string]$entry.Environment
            if ($e -match '(?i)^(prod|prd|production)') { return 'Production' }
            return $e
        }
        if ($Settings.ProdNamePattern -and $ServerName -match $Settings.ProdNamePattern) { return 'Production' }
        if ($Settings.NonProdNamePattern -and $ServerName -match $Settings.NonProdNamePattern) { return 'Non-Production' }
        if ($Settings.UnknownIsProduction) { return 'Production (assumed)' }
        return 'Unknown'
    }

    function Test-IsProductionEnvironment {
        param([string]$Environment)
        if ([string]::IsNullOrWhiteSpace($Environment)) { return $true }
        if ($Environment -match '(?i)assumed|unknown') { return $true }
        if ($Environment -match '(?i)non[\s_-]*prod') { return $false }
        return ($Environment -match '(?i)prod|prd|live|dr\b')
    }
    #endregion

    #region ---------------------------------------------------------------- Read-only guard
    function ConvertTo-SqlSkeleton {
        <#
            Removes comments, string literals and quoted identifiers so that keyword
            checks only see real T-SQL tokens. Handles nested /* */ comments and
            doubled quotes.
        #>
        param([string]$Sql)
        $sb = New-Object System.Text.StringBuilder
        $i = 0; $len = $Sql.Length
        while ($i -lt $len) {
            $c = $Sql[$i]
            $n = if ($i + 1 -lt $len) { $Sql[$i + 1] } else { [char]0 }
            if ($c -eq '-' -and $n -eq '-') {
                while ($i -lt $len -and $Sql[$i] -ne "`n") { $i++ }
                [void]$sb.Append(' ')
                continue
            }
            if ($c -eq '/' -and $n -eq '*') {
                $depth = 1; $i += 2
                while ($i -lt $len -and $depth -gt 0) {
                    if ($Sql[$i] -eq '/' -and $i + 1 -lt $len -and $Sql[$i + 1] -eq '*') { $depth++; $i += 2; continue }
                    if ($Sql[$i] -eq '*' -and $i + 1 -lt $len -and $Sql[$i + 1] -eq '/') { $depth--; $i += 2; continue }
                    $i++
                }
                [void]$sb.Append(' ')
                continue
            }
            if ($c -eq "'") {
                $i++
                while ($i -lt $len) {
                    if ($Sql[$i] -eq "'") {
                        if ($i + 1 -lt $len -and $Sql[$i + 1] -eq "'") { $i += 2; continue }
                        break
                    }
                    $i++
                }
                $i++
                [void]$sb.Append("''")
                continue
            }
            if ($c -eq '[') {
                $i++
                while ($i -lt $len) {
                    if ($Sql[$i] -eq ']') {
                        if ($i + 1 -lt $len -and $Sql[$i + 1] -eq ']') { $i += 2; continue }
                        break
                    }
                    $i++
                }
                $i++
                [void]$sb.Append('[q]')
                continue
            }
            if ($c -eq '"') {
                $i++
                while ($i -lt $len) {
                    if ($Sql[$i] -eq '"') {
                        if ($i + 1 -lt $len -and $Sql[$i + 1] -eq '"') { $i += 2; continue }
                        break
                    }
                    $i++
                }
                $i++
                [void]$sb.Append('[q]')
                continue
            }
            [void]$sb.Append($c)
            $i++
        }
        $sb.ToString()
    }

    function Test-ReadOnlySql {
        <# Returns [pscustomobject]@{ IsSafe; Reason }. Deny-by-default for anything that can change state. #>
        param([Parameter(Mandatory)][AllowEmptyString()][string]$Sql)

        if ([string]::IsNullOrWhiteSpace($Sql)) { return [pscustomobject]@{ IsSafe = $false; Reason = 'Empty statement' } }
        $s = ConvertTo-SqlSkeleton -Sql $Sql
        $w = '(?<![\w@#$])'   # token start
        $e = '(?![\w$#])'     # token end

        $forbidden = @(
            'INSERT','UPDATE','DELETE','MERGE','DROP','ALTER','CREATE','TRUNCATE','EXEC','EXECUTE',
            'GRANT','REVOKE','DENY','BACKUP','RESTORE','SHUTDOWN','KILL','RECONFIGURE','BULK',
            'OPENROWSET','OPENQUERY','OPENDATASOURCE','OPENXML','INTO','WAITFOR','USE','BEGIN','COMMIT',
            'ROLLBACK','SAVE','CHECKPOINT','SETUSER','REVERT','ENABLE','DISABLE','RECEIVE','SEND','MOVE',
            'GO','DUMP','LOAD','READTEXT','WRITETEXT','UPDATETEXT','OUTPUT','RAISERROR','THROW','PRINT',
            'CURSOR','FETCH','DEALLOCATE','REBUILD','REORGANIZE','STATISTICS'
        )
        foreach ($k in $forbidden) {
            if ($s -match ($w + $k + $e)) { return [pscustomobject]@{ IsSafe = $false; Reason = "Keyword '$k' is not allowed (read-only tool)" } }
        }
        if ($s -match ($w + '(sp|xp)_\w+')) { return [pscustomobject]@{ IsSafe = $false; Reason = "Calls to sp_/xp_ procedures are not allowed: '$($Matches[0])'" } }
        if ($s -match '::') { return [pscustomobject]@{ IsSafe = $false; Reason = "'::' scope qualifier is not allowed" } }

        # DBCC: only SHOW_STATISTICS. (The STATISTICS keyword above is matched as a whole token, so SHOW_STATISTICS still passes.)
        foreach ($m in [regex]::Matches($s, $w + 'DBCC' + $e + '\s*(\w+)', 'IgnoreCase')) {
            if ($m.Groups[1].Value -ne 'SHOW_STATISTICS') { return [pscustomobject]@{ IsSafe = $false; Reason = "DBCC $($m.Groups[1].Value) is not allowed; only DBCC SHOW_STATISTICS" } }
        }
        # SET: only session options or variables
        $allowedSet = @('NOCOUNT','LOCK_TIMEOUT','DEADLOCK_PRIORITY','TRANSACTION','ARITHABORT','ANSI_WARNINGS','ANSI_NULLS','QUOTED_IDENTIFIER','CONCAT_NULL_YIELDS_NULL','ANSI_PADDING','NUMERIC_ROUNDABORT','XACT_ABORT','DATEFORMAT','LANGUAGE','TEXTSIZE')
        foreach ($m in [regex]::Matches($s, $w + 'SET' + $e + '\s*(@?\w+)', 'IgnoreCase')) {
            $opt = $m.Groups[1].Value
            if ($opt.StartsWith('@')) { continue }
            if ($allowedSet -notcontains $opt.ToUpperInvariant()) { return [pscustomobject]@{ IsSafe = $false; Reason = "SET $opt is not allowed" } }
        }
        if ($s -match ($w + 'TRANSACTION' + $e) -and $s -notmatch ($w + 'SET\s+TRANSACTION\s+ISOLATION\s+LEVEL\s+READ\s+UNCOMMITTED')) {
            return [pscustomobject]@{ IsSafe = $false; Reason = 'Only SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED is allowed' }
        }
        if ($s -notmatch ($w + '(SELECT|DBCC)' + $e)) { return [pscustomobject]@{ IsSafe = $false; Reason = 'Statement must be a SELECT (or DBCC SHOW_STATISTICS)' } }
        [pscustomobject]@{ IsSafe = $true; Reason = 'OK' }
    }

    function Assert-ReadOnlySql {
        param([Parameter(Mandatory)][string]$Sql)
        $r = Test-ReadOnlySql -Sql $Sql
        if (-not $r.IsSafe) { throw "READ-ONLY GUARD blocked a statement: $($r.Reason)" }
    }
    #endregion

    #region ---------------------------------------------------------------- Connection & query
    function New-AnalyzerContext {
        param(
            [Parameter(Mandatory)][string]$Server,
            [string]$Database = 'master',
            [ValidateSet('Windows','Sql')][string]$Auth = 'Windows',
            [string]$UserName,
            [System.Security.SecureString]$Password,
            [bool]$Encrypt = $false,
            [bool]$TrustServerCertificate = $true,
            [bool]$ReadOnlyIntent = $true,
            [int]$CommandTimeoutSec = 30,
            [int]$LockTimeoutMs = 3000,
            $Shared
        )
        if (-not $Shared) {
            $Shared = [hashtable]::Synchronized(@{
                Log     = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
                Command = $null
                Cancel  = $false
            })
        }
        [pscustomobject]@{
            Server = $Server; Database = $Database; Auth = $Auth; UserName = $UserName; Password = $Password
            Encrypt = $Encrypt; TrustServerCertificate = $TrustServerCertificate; ReadOnlyIntent = $ReadOnlyIntent
            CommandTimeoutSec = $CommandTimeoutSec; LockTimeoutMs = $LockTimeoutMs; Shared = $Shared
        }
    }

    function Copy-AnalyzerContext {
        param([Parameter(Mandatory)]$Context, [string]$Database)
        $c = $Context.PSObject.Copy()
        if ($Database) { $c.Database = $Database }
        $c
    }

    function Get-AnalyzerConnectionString {
        param([Parameter(Mandatory)]$Context)
        $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $b['Data Source']              = $Context.Server
        $b['Initial Catalog']          = $Context.Database
        $b['Application Name']         = "$($script:ToolName) (read-only)"
        $b['Connect Timeout']          = 15
        $b['Encrypt']                  = [bool]$Context.Encrypt
        $b['TrustServerCertificate']   = [bool]$Context.TrustServerCertificate
        $b['Pooling']                  = $true
        $b['Max Pool Size']            = 2       # this tool never needs more than one connection at a time
        if ($Context.ReadOnlyIntent) { $b['ApplicationIntent'] = 'ReadOnly' }
        if ($Context.Auth -eq 'Windows') { $b['Integrated Security'] = $true }
        $b.ConnectionString
    }

    function Invoke-SafeQuery {
        <#
            The ONLY place this module talks to SQL Server.
            Returns a System.Data.DataTable (or DataSet with -AsDataSet).
        #>
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)][string]$Sql,
            [hashtable]$Parameters,
            [int]$TimeoutSec,
            [string]$Label = 'query',
            [switch]$AsDataSet
        )
        Assert-ReadOnlySql -Sql $Sql

        $lockMs = [int]$Context.LockTimeoutMs
        $prefix = "SET NOCOUNT ON; SET LOCK_TIMEOUT $lockMs; SET DEADLOCK_PRIORITY LOW; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`r`n"
        $batch = $prefix + $Sql
        Assert-ReadOnlySql -Sql $batch    # belt and braces: the exact text that is sent is checked too

        $timeout = if ($TimeoutSec -gt 0) { $TimeoutSec } else { [int]$Context.CommandTimeoutSec }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $conn = New-Object System.Data.SqlClient.SqlConnection (Get-AnalyzerConnectionString -Context $Context)
        if ($Context.Auth -eq 'Sql') {
            $pw = $Context.Password.Copy(); $pw.MakeReadOnly()
            $conn.Credential = New-Object System.Data.SqlClient.SqlCredential($Context.UserName, $pw)
        }
        $status = 'OK'; $rowsOut = 0
        try {
            $conn.Open()
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $batch
            $cmd.CommandType = [System.Data.CommandType]::Text
            $cmd.CommandTimeout = $timeout
            if ($Parameters) {
                foreach ($k in $Parameters.Keys) {
                    $v = $Parameters[$k]; if ($null -eq $v) { $v = [DBNull]::Value }
                    [void]$cmd.Parameters.AddWithValue('@' + $k.TrimStart('@'), $v)
                }
            }
            if ($null -ne $Context.Shared) { $Context.Shared.Command = $cmd }
            $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
            if ($AsDataSet) {
                $ds = New-Object System.Data.DataSet
                [void]$da.Fill($ds)
                foreach ($t in $ds.Tables) { $rowsOut += $t.Rows.Count }
                return ,$ds
            } else {
                $dt = New-Object System.Data.DataTable
                [void]$da.Fill($dt)
                $rowsOut = $dt.Rows.Count
                return ,$dt
            }
        } catch {
            $status = 'ERROR: ' + $_.Exception.Message
            throw
        } finally {
            if ($null -ne $Context.Shared) { $Context.Shared.Command = $null }
            $conn.Dispose()
            $sw.Stop()
            if ($null -ne $Context.Shared -and $null -ne $Context.Shared.Log) {
                [void]$Context.Shared.Log.Add([pscustomobject]@{
                    Time = Get-Date; Server = $Context.Server; Database = $Context.Database; Label = $Label
                    DurationMs = $sw.ElapsedMilliseconds; Rows = $rowsOut; Status = $status; Sql = $Sql
                })
            }
        }
    }

    function ConvertTo-ObjectList {
        <# DataTable -> array of PSCustomObject, DBNull -> $null #>
        param($Table)
        $out = New-Object System.Collections.Generic.List[object]
        if (-not $Table) { return }
        $cols = @($Table.Columns | ForEach-Object { $_.ColumnName })
        foreach ($r in $Table.Rows) {
            $h = [ordered]@{}
            foreach ($c in $cols) { $v = $r[$c]; if ($v -is [DBNull]) { $v = $null }; $h[$c] = $v }
            $out.Add([pscustomobject]$h)
        }
        # Emit rows one by one (callers wrap in @()). Returning ,$array here made @(f) a 1-item list holding the whole array.
        $out.ToArray()
    }
    #endregion

    #region ---------------------------------------------------------------- Naming helpers
    function ConvertTo-QuotedName {
        param([Parameter(Mandatory)][AllowEmptyString()][string]$Name)
        '[' + $Name.Replace(']', ']]') + ']'
    }
    function ConvertTo-SqlLiteral {
        param([AllowEmptyString()][AllowNull()][string]$Value)
        if ($null -eq $Value) { return 'NULL' }
        "N'" + $Value.Replace("'", "''") + "'"
    }
    function Format-Size {
        param([double]$MB)
        if ($MB -ge 1048576) { return ('{0:N2} TB' -f ($MB / 1048576)) }
        if ($MB -ge 1024)    { return ('{0:N2} GB' -f ($MB / 1024)) }
        return ('{0:N1} MB' -f $MB)
    }
    function Format-Count {
        param([double]$N)
        if ($N -ge 1e9) { return ('{0:N1} billion' -f ($N / 1e9)) }
        if ($N -ge 1e6) { return ('{0:N1} million' -f ($N / 1e6)) }
        if ($N -ge 1e4) { return ('{0:N0} thousand' -f ($N / 1e3)) }
        return ('{0:N0}' -f $N)
    }
    function Format-Age {
        param($Date)
        if (-not $Date) { return 'never' }
        $span = (Get-Date) - [datetime]$Date
        if ($span.TotalMinutes -lt 60) { return ('{0:N0} minutes ago' -f [math]::Max(0, $span.TotalMinutes)) }
        if ($span.TotalHours -lt 48)   { return ('{0:N0} hours ago' -f $span.TotalHours) }
        return ('{0:N0} days ago' -f $span.TotalDays)
    }
    #endregion

    #region ---------------------------------------------------------------- Server / database discovery
    function Get-ServerInfo {
        param([Parameter(Mandatory)]$Context)
        $ctx = Copy-AnalyzerContext -Context $Context -Database 'master'
        $sql = @"
SELECT  @@SERVERNAME                                           AS ServerName,
        CAST(SERVERPROPERTY('MachineName')       AS nvarchar(128)) AS MachineName,
        CAST(SERVERPROPERTY('ProductVersion')    AS nvarchar(128)) AS ProductVersion,
        CAST(SERVERPROPERTY('ProductLevel')      AS nvarchar(128)) AS ProductLevel,
        CAST(SERVERPROPERTY('ProductUpdateLevel') AS nvarchar(128)) AS ProductUpdateLevel,
        CAST(SERVERPROPERTY('Edition')           AS nvarchar(128)) AS Edition,
        CAST(SERVERPROPERTY('EngineEdition')     AS int)           AS EngineEdition,
        CAST(ISNULL(SERVERPROPERTY('IsHadrEnabled'), 0) AS int)    AS IsHadrEnabled,
        CAST(ISNULL(SERVERPROPERTY('IsClustered'), 0) AS int)      AS IsClustered,
        IS_SRVROLEMEMBER('sysadmin')                               AS IsSysadmin,
        HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER STATE')         AS HasViewServerState,
        HAS_PERMS_BY_NAME(NULL, NULL, 'ALTER ANY DATABASE')        AS HasAlterAnyDatabase,
        HAS_PERMS_BY_NAME(NULL, NULL, 'CONTROL SERVER')            AS HasControlServer,
        SUSER_SNAME()                                              AS LoginName;
"@
        $info = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $ctx -Sql $sql -Label 'server info'))[0]

        $v = [version]($info.ProductVersion -replace '[^\d\.]', '')
        $info | Add-Member -NotePropertyName Version -NotePropertyValue $v
        $info | Add-Member -NotePropertyName StartTime -NotePropertyValue $null
        $info | Add-Member -NotePropertyName CpuCount -NotePropertyValue $null
        if ($info.HasViewServerState -eq 1) {
            try {
                $os = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $ctx -Label 'uptime' -Sql 'SELECT sqlserver_start_time AS StartTime, cpu_count AS CpuCount FROM sys.dm_os_sys_info;'))[0]
                $info.StartTime = $os.StartTime; $info.CpuCount = $os.CpuCount
            } catch { }
        }
        $info | Add-Member -NotePropertyName Capabilities -NotePropertyValue (Get-ServerCapabilities -ServerInfo $info)
        $info
    }

    function Get-ServerCapabilities {
        <# Feature flags by version/edition. Build numbers are the first build containing the feature. #>
        param([Parameter(Mandatory)]$ServerInfo)
        $v = $ServerInfo.Version
        $major = $v.Major; $build = $v.Build
        $ee = [int]$ServerInfo.EngineEdition       # 3 = Enterprise/Developer/Evaluation, 5 = Azure SQL DB, 8 = Managed Instance
        $enterpriseLike = $ee -in 3, 5, 8
        [pscustomobject]@{
            Major                    = $major
            OnlineRebuild            = $enterpriseLike
            OnlinePartitionRebuild   = $enterpriseLike -and $major -ge 12
            WaitAtLowPriority        = $major -ge 12
            ResumableRebuild         = $enterpriseLike -and $major -ge 14
            OnlineColumnstoreRebuild = $enterpriseLike -and $major -ge 15
            IncrementalStats         = $major -ge 12
            StatsHistogramDmv        = ($major -gt 13) -or ($major -eq 13 -and $build -ge 4446)   # 2016 SP1 CU2
            PersistSamplePercent     = ($major -gt 14) -or ($major -eq 14 -and $build -ge 3006) -or ($major -eq 13 -and $build -ge 4474)  # 2017 CU1 / 2016 SP1 CU4
            UpdateStatsMaxDop        = ($major -gt 14) -or ($major -eq 14 -and $build -ge 3015) -or ($major -eq 13 -and $build -ge 5026) -or ($major -eq 12 -and $build -ge 6024)
            ColumnstoreRowGroupDmv   = $major -ge 13
            CompressAllRowGroups     = $major -ge 13
            FnHadrIsPrimary          = $major -ge 12
            AutoDropStats            = $major -ge 16
            StatsPropertiesDmv       = ($major -gt 11) -or ($major -eq 11 -and $build -ge 3000) -or ($major -eq 10 -and $v.Minor -eq 50 -and $build -ge 4000)
        }
    }

    function Get-DatabaseList {
        param([Parameter(Mandatory)]$Context, $ServerInfo)
        $ctx = Copy-AnalyzerContext -Context $Context -Database 'master'
        $hadr = $ServerInfo -and $ServerInfo.IsHadrEnabled -eq 1 -and $ServerInfo.HasViewServerState -eq 1
        $agCols = if ($hadr) { 'ars.role_desc AS AgRole, ag.name AS AgName' } else { 'CAST(NULL AS nvarchar(60)) AS AgRole, CAST(NULL AS sysname) AS AgName' }
        $agJoin = if ($hadr) {
    @"
LEFT JOIN sys.dm_hadr_database_replica_states AS drs ON drs.database_id = d.database_id AND drs.is_local = 1
LEFT JOIN sys.dm_hadr_availability_replica_states AS ars ON ars.replica_id = drs.replica_id
LEFT JOIN sys.availability_groups AS ag ON ag.group_id = drs.group_id
"@ } else { '' }
        $sql = @"
SELECT d.database_id AS DatabaseId, d.name AS DatabaseName, d.state_desc AS State, d.recovery_model_desc AS RecoveryModel,
       d.compatibility_level AS CompatLevel, d.is_read_only AS IsReadOnly, HAS_DBACCESS(d.name) AS HasAccess,
       sz.SizeGB, $agCols
FROM sys.databases AS d
LEFT JOIN (SELECT database_id, CAST(SUM(CAST(size AS bigint)) * 8 / 1048576.0 AS decimal(19,2)) AS SizeGB
           FROM sys.master_files WHERE type = 0 GROUP BY database_id) AS sz ON sz.database_id = d.database_id
$agJoin
WHERE d.database_id > 4
ORDER BY d.name;
"@
        ConvertTo-ObjectList (Invoke-SafeQuery -Context $ctx -Sql $sql -Label 'database list')
    }

    function Get-DatabaseInfo {
        param([Parameter(Mandatory)]$Context, $ServerInfo)
        $caps = if ($ServerInfo) { $ServerInfo.Capabilities } else { $null }
        $isPrimary = if ($caps -and $caps.FnHadrIsPrimary) { 'CASE WHEN d.replica_id IS NULL THEN NULL ELSE CAST(sys.fn_hadr_is_primary_replica(DB_NAME()) AS int) END' } else { 'CAST(NULL AS int)' }
        $sql = @"
SELECT DB_NAME() AS DatabaseName, d.compatibility_level AS CompatLevel, d.recovery_model_desc AS RecoveryModel,
       d.is_auto_create_stats_on AS AutoCreateStats, d.is_auto_update_stats_on AS AutoUpdateStats,
       d.is_auto_update_stats_async_on AS AutoUpdateStatsAsync, d.log_reuse_wait_desc AS LogReuseWait,
       d.is_read_only AS IsReadOnly, DATABASEPROPERTYEX(DB_NAME(), 'Updateability') AS Updateability,
       CASE WHEN d.replica_id IS NULL THEN 0 ELSE 1 END AS IsInAg, $isPrimary AS IsAgPrimary,
       (SELECT CAST(SUM(CAST(size AS bigint)) * 8 / 1024.0 AS decimal(19,1)) FROM sys.database_files WHERE type = 0) AS DataSizeMB,
       (SELECT CAST(SUM(CAST(size - FILEPROPERTY(name, 'SpaceUsed') AS bigint)) * 8 / 1024.0 AS decimal(19,1)) FROM sys.database_files WHERE type = 0) AS DataFreeMB,
       (SELECT CAST(SUM(CAST(size AS bigint)) * 8 / 1024.0 AS decimal(19,1)) FROM sys.database_files WHERE type = 1) AS LogSizeMB
FROM sys.databases AS d WHERE d.database_id = DB_ID();
"@
        $db = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Label 'database info'))[0]
        $db | Add-Member -NotePropertyName LogUsedPct -NotePropertyValue $null
        try {
            $ls = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'log space' -Sql 'SELECT CAST(used_log_space_in_percent AS decimal(5,1)) AS LogUsedPct FROM sys.dm_db_log_space_usage;'))[0]
            $db.LogUsedPct = $ls.LogUsedPct
        } catch { }
        # Is Ola Hallengren's IndexOptimize installed? (read-only metadata check)
        $ola = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'ola check' -Sql @"
SELECT CASE WHEN OBJECT_ID(N'master.dbo.IndexOptimize') IS NOT NULL THEN N'master'
            WHEN OBJECT_ID(N'dbo.IndexOptimize') IS NOT NULL THEN DB_NAME() END AS OlaDatabase;
"@))
        $db | Add-Member -NotePropertyName OlaDatabase -NotePropertyValue $ola[0].OlaDatabase
        $db
    }

    function Get-ServerLoad {
        <# Cheap check used to refuse physical-stats scans while the server is busy. #>
        param([Parameter(Mandatory)]$Context)
        $sql = @"
SELECT CAST(AVG(CAST(s.runnable_tasks_count AS decimal(9,2))) AS decimal(9,2)) AS AvgRunnableTasks,
       (SELECT COUNT(*) FROM sys.dm_exec_requests AS r WHERE r.blocking_session_id <> 0) AS BlockedRequests,
       (SELECT COUNT(*) FROM sys.dm_exec_requests AS r WHERE r.session_id > 50 AND r.status IN (N'running', N'runnable', N'suspended')) AS ActiveRequests
FROM sys.dm_os_schedulers AS s WHERE s.status = N'VISIBLE ONLINE';
"@
        @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Label 'server load'))[0]
    }

    function Get-CentralServerList {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Query)
        $rows = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $Query -Label 'central server list'))
        foreach ($r in $rows) {
            $name = $null; $envName = $null
            foreach ($p in $r.PSObject.Properties) {
                if (-not $name -and $p.Name -match '(?i)^(ServerName|SqlInstance|InstanceName|Name|Server)$') { $name = [string]$p.Value }
                if (-not $envName -and $p.Name -match '(?i)^(Environment|Env|Tier)$') { $envName = [string]$p.Value }
            }
            if (-not $name) { $name = [string]($r.PSObject.Properties | Select-Object -First 1).Value }
            if ($name) { [pscustomobject]@{ Name = $name.Trim(); Environment = $envName } }
        }
    }
    #endregion

    #region ---------------------------------------------------------------- Table list & search
    function Get-TableList {
        <# Metadata only. Fine for tens of thousands of tables. Falls back to a lighter query if DMVs are not permitted. #>
        param([Parameter(Mandatory)]$Context)
        $sql = @"
SELECT t.object_id AS ObjectId, s.name AS SchemaName, t.name AS TableName,
       ISNULL(ps.RowCnt, 0) AS [RowCount],
       CAST(ISNULL(ps.ReservedPages, 0) * 8 / 1024.0 AS decimal(19,1)) AS ReservedMB,
       ISNULL(ps.PartitionCount, 0) AS Partitions
FROM sys.tables AS t
JOIN sys.schemas AS s ON s.schema_id = t.schema_id
LEFT JOIN (SELECT object_id,
                  SUM(CASE WHEN index_id IN (0, 1) THEN row_count ELSE 0 END) AS RowCnt,
                  SUM(reserved_page_count) AS ReservedPages,
                  SUM(CASE WHEN index_id IN (0, 1) THEN 1 ELSE 0 END) AS PartitionCount
           FROM sys.dm_db_partition_stats GROUP BY object_id) AS ps ON ps.object_id = t.object_id
WHERE t.is_ms_shipped = 0
OPTION (MAXDOP 1);
"@
        try {
            Invoke-SafeQuery -Context $Context -Sql $sql -Label 'table list' -TimeoutSec ([math]::Max(60, $Context.CommandTimeoutSec))
        } catch {
            $fallback = @"
SELECT t.object_id AS ObjectId, s.name AS SchemaName, t.name AS TableName,
       ISNULL((SELECT SUM(p.rows) FROM sys.partitions AS p WHERE p.object_id = t.object_id AND p.index_id IN (0, 1)), 0) AS [RowCount],
       CAST(NULL AS decimal(19,1)) AS ReservedMB,
       (SELECT COUNT(*) FROM sys.partitions AS p WHERE p.object_id = t.object_id AND p.index_id IN (0, 1)) AS Partitions
FROM sys.tables AS t JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.is_ms_shipped = 0 OPTION (MAXDOP 1);
"@
            Invoke-SafeQuery -Context $Context -Sql $fallback -Label 'table list (fallback, no VIEW DATABASE STATE)' -TimeoutSec ([math]::Max(60, $Context.CommandTimeoutSec))
        }
    }

    function Find-Table {
        <# Fuzzy search over a cached table list: exact, prefix, contains, tokens, typos, abbreviations, wildcards (* ? %), schema.table #>
        param([AllowEmptyCollection()][string[]]$SchemaNames = @(), [AllowEmptyCollection()][string[]]$TableNames = @(), [string]$Query, [int]$Top = 500)
        [IsaTableMatcher]::Search($SchemaNames, $TableNames, $Query, $Top)
    }
    #endregion

    #region ---------------------------------------------------------------- Table analysis queries
    function Get-TableOverview {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId)
        $sql = @"
SELECT s.name AS SchemaName, t.name AS TableName, t.create_date AS Created, t.modify_date AS SchemaModified,
       t.lock_escalation_desc AS LockEscalation,
       CAST(OBJECTPROPERTY(t.object_id, 'TableHasClustIndex') AS bit) AS HasClusteredIndex,
       CAST(OBJECTPROPERTY(t.object_id, 'TableHasPrimaryKey') AS bit) AS HasPrimaryKey,
       (SELECT COUNT(*) FROM sys.columns AS c WHERE c.object_id = t.object_id) AS ColumnCount,
       (SELECT COUNT(*) FROM sys.indexes AS i WHERE i.object_id = t.object_id AND i.index_id > 0) AS IndexCount,
       (SELECT COUNT(*) FROM sys.stats AS st WHERE st.object_id = t.object_id) AS StatsCount,
       (SELECT TOP (1) ps.name FROM sys.indexes AS i JOIN sys.partition_schemes AS ps ON ps.data_space_id = i.data_space_id
         WHERE i.object_id = t.object_id AND i.index_id IN (0, 1)) AS PartitionScheme,
       (SELECT MAX(p.partition_number) FROM sys.partitions AS p WHERE p.object_id = t.object_id AND p.index_id IN (0, 1)) AS Partitions,
       (SELECT SUM(p.rows) FROM sys.partitions AS p WHERE p.object_id = t.object_id AND p.index_id IN (0, 1)) AS [RowCount],
       CAST(ISNULL(OBJECTPROPERTYEX(t.object_id, 'IsMemoryOptimized'), 0) AS bit) AS IsMemoryOptimized
FROM sys.tables AS t JOIN sys.schemas AS s ON s.schema_id = t.schema_id
WHERE t.object_id = @ObjectId;
"@
        $o = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'table overview'))[0]
        if ($o) {
            try {
                $sz = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'table size' -Parameters @{ ObjectId = $ObjectId } -Sql @"
SELECT CAST(SUM(reserved_page_count) * 8 / 1024.0 AS decimal(19,1)) AS ReservedMB,
       CAST(SUM(CASE WHEN index_id IN (0, 1) THEN reserved_page_count ELSE 0 END) * 8 / 1024.0 AS decimal(19,1)) AS DataMB,
       CAST(SUM(CASE WHEN index_id > 1 THEN reserved_page_count ELSE 0 END) * 8 / 1024.0 AS decimal(19,1)) AS IndexMB
FROM sys.dm_db_partition_stats WHERE object_id = @ObjectId;
"@))[0]
                foreach ($p in 'ReservedMB', 'DataMB', 'IndexMB') { $o | Add-Member -NotePropertyName $p -NotePropertyValue $sz.$p }
            } catch {
                foreach ($p in 'ReservedMB', 'DataMB', 'IndexMB') { $o | Add-Member -NotePropertyName $p -NotePropertyValue $null }
            }
        }
        $o
    }

    function Get-TableIndexes {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId)
        $sql = @"
DECLARE @db int = DB_ID();
WITH sz AS (
    SELECT index_id, SUM(row_count) AS row_count, SUM(reserved_page_count) AS reserved_pages,
           SUM(used_page_count) AS used_pages, SUM(in_row_data_page_count) AS in_row_pages, COUNT(*) AS partitions
    FROM sys.dm_db_partition_stats WHERE object_id = @ObjectId GROUP BY index_id),
cmp AS (
    SELECT index_id, MIN(data_compression_desc) AS min_c, MAX(data_compression_desc) AS max_c
    FROM sys.partitions WHERE object_id = @ObjectId GROUP BY index_id),
ops AS (
    SELECT index_id, SUM(leaf_allocation_count) AS page_allocations, SUM(nonleaf_allocation_count) AS nonleaf_allocations,
           SUM(row_lock_wait_in_ms) AS row_lock_wait_ms, SUM(page_lock_wait_in_ms) AS page_lock_wait_ms,
           SUM(page_latch_wait_in_ms) AS page_latch_wait_ms, SUM(page_io_latch_wait_in_ms) AS page_io_latch_wait_ms,
           SUM(forwarded_fetch_count) AS forwarded_fetches, SUM(index_lock_promotion_count) AS lock_escalations,
           SUM(leaf_insert_count + leaf_update_count + leaf_delete_count) AS leaf_writes
    FROM sys.dm_db_index_operational_stats(@db, @ObjectId, NULL, NULL) GROUP BY index_id)
SELECT i.index_id AS IndexId,
       ISNULL(i.name, N'(heap)') AS IndexName,
       i.type AS TypeCode, i.type_desc AS IndexType,
       i.is_unique AS IsUnique, i.is_primary_key AS IsPrimaryKey, i.is_unique_constraint AS IsUniqueConstraint,
       i.is_disabled AS IsDisabled, i.is_hypothetical AS IsHypothetical, i.has_filter AS HasFilter,
       i.filter_definition AS FilterDefinition, i.fill_factor AS FillFactorPct, i.is_padded AS IsPadded,
       i.allow_page_locks AS AllowPageLocks, i.allow_row_locks AS AllowRowLocks,
       STUFF((SELECT N', ' + QUOTENAME(c.name) + CASE WHEN ic.is_descending_key = 1 THEN N' DESC' ELSE N'' END
              FROM sys.index_columns AS ic JOIN sys.columns AS c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
              WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal > 0
              ORDER BY ic.key_ordinal FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, N'') AS KeyColumns,
       STUFF((SELECT N', ' + QUOTENAME(c.name)
              FROM sys.index_columns AS ic JOIN sys.columns AS c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
              WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.key_ordinal = 0
                AND (ic.is_included_column = 1 OR i.type IN (5, 6))
              ORDER BY c.name FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, N'') AS IncludedColumns,
       sz.row_count AS [RowCount], sz.partitions AS Partitions,
       CAST(sz.reserved_pages * 8 / 1024.0 AS decimal(19,2)) AS SizeMB,
       CAST(sz.used_pages * 8 / 1024.0 AS decimal(19,2)) AS UsedMB,
       sz.in_row_pages AS InRowPages,
       CASE WHEN cmp.min_c = cmp.max_c THEN cmp.min_c ELSE N'MIXED' END AS Compression,
       ISNULL(us.user_seeks, 0) AS Seeks, ISNULL(us.user_scans, 0) AS Scans, ISNULL(us.user_lookups, 0) AS Lookups,
       ISNULL(us.user_updates, 0) AS Writes,
       us.last_user_seek AS LastSeek, us.last_user_scan AS LastScan, us.last_user_lookup AS LastLookup, us.last_user_update AS LastWrite,
       (SELECT MAX(v) FROM (VALUES (us.last_user_seek), (us.last_user_scan), (us.last_user_lookup)) AS x(v)) AS LastRead,
       ops.page_allocations AS LeafPageAllocations, ops.nonleaf_allocations AS NonLeafAllocations,
       ops.row_lock_wait_ms AS RowLockWaitMs, ops.page_lock_wait_ms AS PageLockWaitMs,
       ops.page_latch_wait_ms AS PageLatchWaitMs, ops.page_io_latch_wait_ms AS PageIoLatchWaitMs,
       ops.forwarded_fetches AS ForwardedFetches, ops.lock_escalations AS LockEscalations, ops.leaf_writes AS LeafWrites,
       STATS_DATE(i.object_id, i.index_id) AS StatsLastUpdated
FROM sys.indexes AS i
LEFT JOIN sz  ON sz.index_id  = i.index_id
LEFT JOIN cmp ON cmp.index_id = i.index_id
LEFT JOIN ops ON ops.index_id = i.index_id
LEFT JOIN sys.dm_db_index_usage_stats AS us ON us.database_id = @db AND us.object_id = i.object_id AND us.index_id = i.index_id
WHERE i.object_id = @ObjectId
ORDER BY i.index_id
OPTION (MAXDOP 1);
"@
        ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'indexes')
    }

    function Get-TableStatistics {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)]$ServerInfo)
        $caps = $ServerInfo.Capabilities
        $incr    = if ($caps.IncrementalStats)     { 'st.is_incremental' }          else { 'CAST(0 AS bit)' }
        $persist = if ($caps.PersistSamplePercent) { 'sp.persisted_sample_percent' } else { 'CAST(NULL AS float)' }
        $autoDrop = if ($caps.AutoDropStats)       { 'st.auto_drop' }               else { 'CAST(0 AS bit)' }
        $sql = @"
SELECT st.stats_id AS StatsId, st.name AS StatName,
       CASE WHEN st.auto_created = 1 THEN N'Auto-created' WHEN st.user_created = 1 THEN N'User-created' ELSE N'Index' END AS StatType,
       STUFF((SELECT N', ' + c.name FROM sys.stats_columns AS sc JOIN sys.columns AS c ON c.object_id = sc.object_id AND c.column_id = sc.column_id
              WHERE sc.object_id = st.object_id AND sc.stats_id = st.stats_id ORDER BY sc.stats_column_id
              FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, N'') AS [Columns],
       lc.LeadingType, lc.LeadingIsIdentity,
       st.has_filter AS HasFilter, st.filter_definition AS FilterDefinition, st.no_recompute AS NoRecompute,
       $incr AS IsIncremental, $autoDrop AS AutoDrop, st.is_temporary AS IsTemporary,
       sp.last_updated AS LastUpdated, sp.rows AS [Rows], sp.rows_sampled AS RowsSampled, sp.steps AS Steps,
       sp.unfiltered_rows AS UnfilteredRows, sp.modification_counter AS Modifications,
       $persist AS PersistedSamplePct,
       ix.IndexSizeMB
FROM sys.stats AS st
OUTER APPLY sys.dm_db_stats_properties(st.object_id, st.stats_id) AS sp
OUTER APPLY (SELECT TOP (1) ty.name AS LeadingType, c.is_identity AS LeadingIsIdentity
             FROM sys.stats_columns AS sc
             JOIN sys.columns AS c ON c.object_id = sc.object_id AND c.column_id = sc.column_id
             JOIN sys.types AS ty ON ty.user_type_id = c.user_type_id
             WHERE sc.object_id = st.object_id AND sc.stats_id = st.stats_id ORDER BY sc.stats_column_id) AS lc
OUTER APPLY (SELECT CAST(SUM(ps.reserved_page_count) * 8 / 1024.0 AS decimal(19,2)) AS IndexSizeMB
             FROM sys.dm_db_partition_stats AS ps
             WHERE ps.object_id = st.object_id
               AND ps.index_id = CASE WHEN st.auto_created = 0 AND st.user_created = 0 THEN st.stats_id
                                      ELSE (SELECT MIN(i2.index_id) FROM sys.indexes AS i2 WHERE i2.object_id = st.object_id AND i2.index_id IN (0, 1)) END) AS ix
WHERE st.object_id = @ObjectId
ORDER BY st.stats_id
OPTION (MAXDOP 1);
"@
        ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'statistics')
    }

    function ConvertTo-DisplayValue {
        param($Value)
        if ($null -eq $Value -or $Value -is [DBNull]) { return 'NULL' }
        if ($Value -is [datetime]) {
            if ($Value.TimeOfDay.Ticks -eq 0) { return $Value.ToString('yyyy-MM-dd') }
            return $Value.ToString('yyyy-MM-dd HH:mm:ss.fff')
        }
        if ($Value -is [datetimeoffset]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss.fff zzz') }
        if ($Value -is [byte[]]) { return '0x' + (($Value | ForEach-Object { $_.ToString('X2') }) -join '') }
        [string]$Value
    }

    function Get-StatisticsDetail {
        <# Header, density vector and histogram for one statistic. Reads only the statistics object, never the table. #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)][int]$StatsId,
              [Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName, [Parameter(Mandatory)][string]$StatName,
              [Parameter(Mandatory)]$ServerInfo)

        $tbl  = ConvertTo-SqlLiteral ((ConvertTo-QuotedName $SchemaName) + '.' + (ConvertTo-QuotedName $TableName))
        $stat = ConvertTo-SqlLiteral $StatName

        $histogram = @()
        if ($ServerInfo.Capabilities.StatsHistogramDmv) {
            $h = Invoke-SafeQuery -Context $Context -Label 'histogram' -Parameters @{ ObjectId = $ObjectId; StatsId = $StatsId } -Sql @"
SELECT step_number AS Step, range_high_key AS RangeHiKey, range_rows AS RangeRows, equal_rows AS EqRows,
       distinct_range_rows AS DistinctRangeRows, average_range_rows AS AvgRangeRows
FROM sys.dm_db_stats_histogram(@ObjectId, @StatsId) ORDER BY step_number;
"@
            $histogram = foreach ($r in $h.Rows) {
                [pscustomobject]@{ Step = [int]$r.Step; RangeHiKey = (ConvertTo-DisplayValue $r.RangeHiKey); RangeRows = [double]$r.RangeRows
                                   EqRows = [double]$r.EqRows; DistinctRangeRows = [double]$r.DistinctRangeRows; AvgRangeRows = [double]$r.AvgRangeRows }
            }
        } else {
            $h = Invoke-SafeQuery -Context $Context -Label 'histogram (DBCC)' -Sql "DBCC SHOW_STATISTICS ($tbl, $stat) WITH HISTOGRAM, NO_INFOMSGS;"
            $step = 0
            $histogram = foreach ($r in $h.Rows) {
                $step++
                [pscustomobject]@{ Step = $step; RangeHiKey = (ConvertTo-DisplayValue $r['RANGE_HI_KEY']); RangeRows = [double]$r['RANGE_ROWS']
                                   EqRows = [double]$r['EQ_ROWS']; DistinctRangeRows = [double]$r['DISTINCT_RANGE_ROWS']; AvgRangeRows = [double]$r['AVG_RANGE_ROWS'] }
            }
        }

        $density = @()
        try {
            $d = Invoke-SafeQuery -Context $Context -Label 'density vector (DBCC)' -Sql "DBCC SHOW_STATISTICS ($tbl, $stat) WITH DENSITY_VECTOR, NO_INFOMSGS;"
            $density = foreach ($r in $d.Rows) {
                $dens = [double]$r['All density']
                [pscustomobject]@{
                    Columns        = [string]$r['Columns']
                    AllDensity     = $dens
                    AvgLengthBytes = [double]$r['Average Length']
                    DistinctValues = if ($dens -gt 0) { [math]::Round(1 / $dens) } else { $null }
                }
            }
        } catch { }

        [pscustomobject]@{ Histogram = @($histogram); Density = @($density) }
    }

    function Get-MissingIndexes {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId)
        $sql = @"
SELECT mid.equality_columns AS EqualityColumns, mid.inequality_columns AS InequalityColumns, mid.included_columns AS IncludedColumns,
       migs.user_seeks AS Seeks, migs.user_scans AS Scans, migs.last_user_seek AS LastSeek,
       CAST(migs.avg_total_user_cost AS decimal(19,2)) AS AvgQueryCost, CAST(migs.avg_user_impact AS decimal(5,1)) AS AvgImpactPct,
       CAST(migs.avg_total_user_cost * (migs.avg_user_impact / 100.0) * (migs.user_seeks + migs.user_scans) AS decimal(19,1)) AS ImprovementScore
FROM sys.dm_db_missing_index_details AS mid
JOIN sys.dm_db_missing_index_groups AS mig ON mig.index_handle = mid.index_handle
JOIN sys.dm_db_missing_index_group_stats AS migs ON migs.group_handle = mig.index_group_handle
WHERE mid.database_id = DB_ID() AND mid.object_id = @ObjectId
ORDER BY ImprovementScore DESC;
"@
        ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'missing indexes')
    }

    function Get-IndexFragmentation {
        <#
            sys.dm_db_index_physical_stats for ONE index with guardrails:
              - LIMITED or SAMPLED only (DETAILED is not offered)
              - refused above the size cap for the environment
              - refused while the server is busy (runnable tasks / blocking)
              - OPTION (MAXDOP 1), long but bounded timeout, cancellable
        #>
        param(
            [Parameter(Mandatory)]$Context,
            [Parameter(Mandatory)][int]$ObjectId,
            [Parameter(Mandatory)]$Index,
            [ValidateSet('LIMITED', 'SAMPLED')][string]$Mode = 'LIMITED',
            [Parameter(Mandatory)][string]$Environment,
            [Parameter(Mandatory)]$Settings,
            [switch]$SkipLoadCheck
        )
        $sizeGB = [double]$Index.SizeMB / 1024
        $isProd = Test-IsProductionEnvironment $Environment
        $cap = if ($isProd) { [double]$Settings.MaxScanGBProduction } else { [double]$Settings.MaxScanGBNonProduction }
        if ($Mode -eq 'SAMPLED') { $cap = [math]::Min($cap, [double]$Settings.MaxSampledScanGB) }

        $result = [ordered]@{ IndexId = $Index.IndexId; Mode = $Mode; Measured = $false; Skipped = $null
                              FragmentationPct = $null; PageDensityPct = $null; PageCount = $null; ForwardedRecords = $null; MeasuredAt = $null }

        if ($Index.TypeCode -in 5, 6) { $result.Skipped = 'Columnstore: uses the row group check instead'; return [pscustomobject]$result }
        if ($Index.IsDisabled)       { $result.Skipped = 'Index is disabled'; return [pscustomobject]$result }
        if ($Index.IsHypothetical)   { $result.Skipped = 'Hypothetical index (no data)'; return [pscustomobject]$result }
        if ($sizeGB -gt $cap) {
            $result.Skipped = ('Not scanned: {0:N1} GB is above the {1} GB safety cap for {2} ({3} mode). Use page splits and scan counts as a proxy, or raise the cap in Settings for a quiet window.' -f $sizeGB, $cap, $Environment, $Mode)
            return [pscustomobject]$result
        }
        if (-not $SkipLoadCheck) {
            try {
                $load = Get-ServerLoad -Context $Context
                if ([double]$load.AvgRunnableTasks -gt [double]$Settings.MaxAvgRunnableTasks -or [int]$load.BlockedRequests -gt 0) {
                    $result.Skipped = ('Not scanned: server is busy (avg runnable tasks {0}, blocked requests {1}). Try again later.' -f $load.AvgRunnableTasks, $load.BlockedRequests)
                    return [pscustomobject]$result
                }
            } catch {
                if ($isProd) { $result.Skipped = 'Not scanned: cannot verify server load (needs VIEW SERVER STATE) and this is production.'; return [pscustomobject]$result }
            }
        }

        $sql = @"
SELECT ips.partition_number AS PartitionNumber, ips.avg_fragmentation_in_percent AS FragPct, ips.page_count AS PageCount,
       ips.avg_page_space_used_in_percent AS DensityPct, ips.forwarded_record_count AS Forwarded, ips.record_count AS Records
FROM sys.dm_db_index_physical_stats(DB_ID(), @ObjectId, @IndexId, NULL, @Mode) AS ips
WHERE ips.alloc_unit_type_desc = N'IN_ROW_DATA' AND ips.index_level = 0
OPTION (MAXDOP 1);
"@
        $rows = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -TimeoutSec $Settings.ScanTimeoutSec `
                    -Parameters @{ ObjectId = $ObjectId; IndexId = [int]$Index.IndexId; Mode = $Mode } -Label "physical stats $Mode idx $($Index.IndexId)"))
        $pages = 0.0; $fragW = 0.0; $densW = 0.0; $densPages = 0.0; $fwd = $null
        foreach ($r in $rows) {
            $p = [double]$r.PageCount; $pages += $p
            if ($null -ne $r.FragPct) { $fragW += [double]$r.FragPct * $p }
            if ($null -ne $r.DensityPct) { $densW += [double]$r.DensityPct * $p; $densPages += $p }
            if ($null -ne $r.Forwarded) { $fwd = [double]$fwd + [double]$r.Forwarded }
        }
        $result.Measured = $true
        $result.MeasuredAt = Get-Date
        $result.PageCount = $pages
        $result.FragmentationPct = if ($pages -gt 0) { [math]::Round($fragW / $pages, 1) } else { 0 }
        $result.PageDensityPct = if ($densPages -gt 0) { [math]::Round($densW / $densPages, 1) } else { $null }
        $result.ForwardedRecords = $fwd
        [pscustomobject]$result
    }

    function Get-ColumnstoreHealth {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)][int]$IndexId, [Parameter(Mandatory)]$ServerInfo)
        $src = if ($ServerInfo.Capabilities.ColumnstoreRowGroupDmv) { 'sys.dm_db_column_store_row_group_physical_stats' } else { 'sys.column_store_row_groups' }
        $state = if ($ServerInfo.Capabilities.ColumnstoreRowGroupDmv) { 'state_desc' } else { 'state_description' }   # 2014 catalog view names it differently
        $sql = @"
SELECT COUNT(*) AS RowGroups,
       SUM(CASE WHEN $state = N'COMPRESSED' THEN 1 ELSE 0 END) AS Compressed,
       SUM(CASE WHEN $state IN (N'OPEN', N'CLOSED') THEN 1 ELSE 0 END) AS DeltaStores,
       SUM(CASE WHEN $state = N'TOMBSTONE' THEN 1 ELSE 0 END) AS Tombstones,
       SUM(CAST(total_rows AS bigint)) AS TotalRows, SUM(CAST(ISNULL(deleted_rows, 0) AS bigint)) AS DeletedRows,
       AVG(CASE WHEN $state = N'COMPRESSED' THEN CAST(total_rows AS bigint) END) AS AvgRowsPerCompressedGroup
FROM $src WHERE object_id = @ObjectId AND index_id = @IndexId
OPTION (MAXDOP 1);
"@
        $r = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId; IndexId = $IndexId } -Label 'columnstore row groups'))[0]
        $delPct = if ([double]$r.TotalRows -gt 0) { [math]::Round(100.0 * [double]$r.DeletedRows / [double]$r.TotalRows, 1) } else { 0 }
        $r | Add-Member -NotePropertyName DeletedPct -NotePropertyValue $delPct
        $r
    }

    function Get-TableAnalysis {
        <# Everything the GUI needs for one table, in one call (runs on the worker thread). #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)]$ServerInfo, $DatabaseInfo, $Settings, [string]$Environment)
        $warnings = New-Object System.Collections.Generic.List[string]
        $overview = Get-TableOverview -Context $Context -ObjectId $ObjectId
        if (-not $overview) { throw "Table with object_id $ObjectId no longer exists." }
        $indexes = @(); $stats = @(); $missing = @()
        try { $indexes = @(Get-TableIndexes -Context $Context -ObjectId $ObjectId) } catch { $warnings.Add("Indexes: $($_.Exception.Message)") }
        try { $stats   = @(Get-TableStatistics -Context $Context -ObjectId $ObjectId -ServerInfo $ServerInfo) } catch { $warnings.Add("Statistics: $($_.Exception.Message)") }
        try { $missing = @(Get-MissingIndexes -Context $Context -ObjectId $ObjectId) } catch { $warnings.Add("Missing indexes: $($_.Exception.Message)") }

        $columnstore = @{}
        foreach ($ix in $indexes | Where-Object { $_.TypeCode -in 5, 6 }) {
            try { $columnstore[[int]$ix.IndexId] = Get-ColumnstoreHealth -Context $Context -ObjectId $ObjectId -IndexId $ix.IndexId -ServerInfo $ServerInfo } catch { $warnings.Add("Columnstore: $($_.Exception.Message)") }
        }
        [pscustomobject]@{
            Overview = $overview; Indexes = $indexes; Statistics = $stats; Missing = $missing
            Columnstore = $columnstore; Warnings = $warnings.ToArray(); LoadedAt = Get-Date
        }
    }
    #endregion

    #region ---------------------------------------------------------------- Assessment: statistics
    function Get-StatsAutoUpdateThreshold {
        <# Modifications needed before SQL Server auto-updates the statistic. #>
        param([double]$Rows, [int]$CompatLevel = 130)
        if ($Rows -le 500) { return 500 }
        $classic = 500 + 0.20 * $Rows
        if ($CompatLevel -ge 130) { return [math]::Min($classic, [math]::Sqrt(1000 * $Rows)) }
        $classic
    }

    function Get-SuggestedSamplePercent {
        <# Heuristic: FULLSCAN up to 10M rows; above that sample about 10M rows, at least 1%. #>
        param([double]$Rows)
        if ($Rows -le 10000000) { return 100 }
        $pct = [math]::Ceiling(10000000 / $Rows * 100)
        [int][math]::Max(1, [math]::Min(100, $pct))
    }

    function Get-StatsAssessment {
        param([Parameter(Mandatory)]$Stat, [int]$CompatLevel = 130, [bool]$AutoUpdateOn = $true, $TableRows)
        $reasons = New-Object System.Collections.Generic.List[string]
        $score = 0
        $rows = if ($null -ne $Stat.Rows) { [double]$Stat.Rows } else { $null }
        $sampled = if ($null -ne $Stat.RowsSampled) { [double]$Stat.RowsSampled } else { $null }
        $mods = if ($null -ne $Stat.Modifications) { [double]$Stat.Modifications } else { 0 }
        $samplePct = if ($rows -and $rows -gt 0 -and $null -ne $sampled) { [math]::Round(100.0 * $sampled / $rows, 2) } else { $null }
        $threshold = if ($null -ne $rows) { Get-StatsAutoUpdateThreshold -Rows $rows -CompatLevel $CompatLevel } else { $null }
        $pctOfThreshold = if ($threshold) { [math]::Round(100.0 * $mods / $threshold, 0) } else { $null }
        $modPct = if ($rows -and $rows -gt 0) { [math]::Round(100.0 * $mods / $rows, 2) } else { $null }
        $ageDays = if ($Stat.LastUpdated) { [math]::Round(((Get-Date) - [datetime]$Stat.LastUpdated).TotalDays, 1) } else { $null }

        if (-not $Stat.LastUpdated) {
            if ($TableRows -and [double]$TableRows -gt 0) { $score = 90; $reasons.Add('Never built although the table has rows: the optimizer is guessing') }
            else { $reasons.Add('Empty statistic (table empty or never sampled)') }
        } else {
            if ($null -ne $pctOfThreshold) {
                if ($pctOfThreshold -ge 100)    { $score += 50; $reasons.Add("Changes since last update ($(Format-Count $mods)) passed the auto-update threshold ($(Format-Count $threshold))") }
                elseif ($pctOfThreshold -ge 50) { $score += 30; $reasons.Add("Changes are at $pctOfThreshold% of the auto-update threshold") }
                elseif ($pctOfThreshold -ge 20) { $score += 15; $reasons.Add("Changes are at $pctOfThreshold% of the auto-update threshold") }
            }
            if ($null -ne $modPct) {
                if ($modPct -ge 20)     { $score += 20; $reasons.Add("$modPct% of rows changed since the last update") }
                elseif ($modPct -ge 10) { $score += 10; $reasons.Add("$modPct% of rows changed since the last update") }
            }
            if ($null -ne $samplePct -and $rows) {
                if ($rows -gt 10000000 -and $samplePct -lt 1)    { $score += 15; $reasons.Add("Very small sample ($samplePct%) on a large table: skewed data may be estimated badly") }
                elseif ($rows -gt 1000000 -and $samplePct -lt 5) { $score += 8;  $reasons.Add("Small sample ($samplePct%)") }
            }
            if ($null -ne $ageDays -and $mods -gt 0) {
                if ($ageDays -gt 30)    { $score += 10; $reasons.Add("Last updated $([math]::Round($ageDays)) days ago and data has changed since") }
                elseif ($ageDays -gt 7) { $score += 5 }
            }
            if ($Stat.NoRecompute -and $mods -gt 0) { $score += 10; $reasons.Add('NORECOMPUTE is set: SQL Server will never refresh this automatically') }
            if (-not $AutoUpdateOn -and $mods -gt 0) { $score += 10; $reasons.Add('AUTO_UPDATE_STATISTICS is OFF for this database') }
            if ($Stat.HasFilter -and $mods -gt 0) { $reasons.Add('Filtered statistic: changes are counted for the whole table, so it can go stale quietly') }
            if ($Stat.LeadingIsIdentity -or ($Stat.LeadingType -match 'date|time')) {
                if ($mods -gt 0) { $reasons.Add('Ascending key column: rows added since the last update are above the histogram, so recent data may be under-estimated') }
            }
        }
        $score = [math]::Min(100, $score)
        $priority = if ($score -ge 60) { 'High' } elseif ($score -ge 30) { 'Medium' } elseif ($score -gt 0) { 'Low' } else { 'None' }

        $suggestPct = if ($rows) { Get-SuggestedSamplePercent -Rows $rows } else { 100 }
        $method = if ($suggestPct -ge 100) { 'FULLSCAN' } else { "SAMPLE $suggestPct PERCENT" }
        $action = if ($priority -eq 'None') { 'No action' } else { "UPDATE STATISTICS ... WITH $method" }

        [pscustomobject]@{
            StatsId = $Stat.StatsId; SamplePct = $samplePct; ModPct = $modPct; Threshold = $threshold; PctOfThreshold = $pctOfThreshold
            AgeDays = $ageDays; Score = $score; Priority = $priority; SuggestedSamplePct = $suggestPct; SuggestedMethod = $method
            Action = $action; Reasons = $reasons.ToArray()
        }
    }

    function Get-HistogramInsights {
        param([object[]]$Histogram, [double]$Rows)
        $lines = New-Object System.Collections.Generic.List[string]
        if (-not $Histogram -or $Histogram.Count -eq 0) { $lines.Add('No histogram (the statistic is empty).'); return ,$lines.ToArray() }
        $total = 0.0; $distinct = 0.0; $maxEq = 0.0; $maxKey = $null
        foreach ($h in $Histogram) {
            $total += $h.EqRows + $h.RangeRows
            $distinct += $h.DistinctRangeRows + 1
            if ($h.EqRows -gt $maxEq) { $maxEq = $h.EqRows; $maxKey = $h.RangeHiKey }
        }
        if ($total -le 0) { $lines.Add('The histogram is empty.'); return ,$lines.ToArray() }
        $lines.Add(("The histogram splits the data into {0} buckets (max 200). It describes about {1} rows and roughly {2} different values." -f $Histogram.Count, (Format-Count $total), (Format-Count $distinct)))
        $topShare = [math]::Round(100.0 * $maxEq / $total, 1)
        $top5 = ($Histogram | Sort-Object EqRows -Descending | Select-Object -First 5 | Measure-Object EqRows -Sum).Sum
        $top5Share = [math]::Round(100.0 * $top5 / $total, 1)
        if ($topShare -ge 10) {
            $lines.Add(("SKEWED: the single value '{0}' accounts for {1}% of the rows. Queries for that value behave very differently from queries for other values (a common cause of 'sometimes fast, sometimes slow', also called parameter sniffing)." -f $maxKey, $topShare))
        } elseif ($top5Share -ge 25) {
            $lines.Add(("Somewhat skewed: the 5 most common values hold {0}% of the rows." -f $top5Share))
        } else {
            $lines.Add('Values are spread fairly evenly: no single value dominates.')
        }
        $firstIdx = 0
        if ([string]$Histogram[0].RangeHiKey -eq 'NULL') {
            $lines.Add(("{0} rows ({1}%) have NULL in this column (shown as the first bucket)." -f (Format-Count $Histogram[0].EqRows), [math]::Round(100.0 * $Histogram[0].EqRows / $total, 1)))
            if ($Histogram.Count -gt 1) { $firstIdx = 1 }
        }
        $first = $Histogram[$firstIdx].RangeHiKey; $last = $Histogram[$Histogram.Count - 1].RangeHiKey
        $lines.Add(("Lowest value seen: '{0}'. Highest value seen: '{1}'. Values outside this range were added after the statistic was built and are estimated, not known." -f $first, $last))
        if ($Rows -gt 0 -and [math]::Abs($total - $Rows) / $Rows -gt 0.05) {
            $lines.Add(("Note: the histogram covers about {0} rows; the statistic header says {1}. A filtered statistic or a nullable column can cause this." -f (Format-Count $total), (Format-Count $Rows)))
        }
        ,$lines.ToArray()
    }

    function Get-StatsPlainEnglish {
        param([Parameter(Mandatory)]$Stat, [Parameter(Mandatory)]$Assessment, $Detail, [string]$TableName)
        $sb = New-Object System.Text.StringBuilder
        $cols = if ($Stat.Columns) { $Stat.Columns } else { '?' }
        [void]$sb.AppendLine("WHAT IT IS")
        [void]$sb.AppendLine("  '$($Stat.StatName)' is a summary SQL Server keeps about column(s) [$cols] of $TableName. The optimizer reads it to guess how many rows a query returns, and that guess decides the plan (seek or scan, join type, memory grant).")
        [void]$sb.AppendLine("  Type: $($Stat.StatType)$(if ($Stat.HasFilter) { "  (filtered: $($Stat.FilterDefinition))" })$(if ($Stat.IsIncremental) { '  (incremental, per partition)' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("HOW FRESH AND HOW ACCURATE")
        if (-not $Stat.LastUpdated) {
            [void]$sb.AppendLine('  This statistic has never been built.')
        } else {
            $light = if ($null -eq $Assessment.SamplePct) { '' } elseif ($Assessment.SamplePct -ge 50) { '[GREEN]' } elseif ($Assessment.SamplePct -ge 5) { '[AMBER]' } else { '[RED]' }
            [void]$sb.AppendLine(("  Built {0} ({1:yyyy-MM-dd HH:mm}) by reading {2} of {3} rows = {4}% of the table. {5}" -f (Format-Age $Stat.LastUpdated), [datetime]$Stat.LastUpdated, (Format-Count $Stat.RowsSampled), (Format-Count $Stat.Rows), $Assessment.SamplePct, $light))
            [void]$sb.AppendLine("    Think of it as a survey: 100% means everyone was asked; 1% means a small poll. Small polls are fine for evenly spread data and risky for skewed data.")
            [void]$sb.AppendLine(("  Since then about {0} rows have changed ({1}% of the table)." -f (Format-Count $Stat.Modifications), $Assessment.ModPct))
            [void]$sb.AppendLine(("  SQL Server refreshes it by itself after about {0} changes; it is at {1}% of that." -f (Format-Count $Assessment.Threshold), $Assessment.PctOfThreshold))
            if ($Assessment.PctOfThreshold -ge 100) { [void]$sb.AppendLine('    It is past the threshold. The automatic refresh only happens when a query that needs this statistic is compiled, so it can stay stale on tables that are written to but rarely queried in new ways.') }
            if ($Stat.PersistedSamplePct) { [void]$sb.AppendLine("  A persisted sample rate of $($Stat.PersistedSamplePct)% is pinned, so automatic updates reuse it.") }
            if ($Stat.NoRecompute) { [void]$sb.AppendLine('  NORECOMPUTE: automatic refresh is switched OFF for this statistic.') }
        }
        if ($Detail -and $Detail.Density -and $Detail.Density.Count -gt 0) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('HOW UNIQUE THE VALUES ARE (density)')
            foreach ($d in $Detail.Density) {
                $avgPerValue = if ($d.AllDensity -gt 0 -and $Stat.Rows) { [double]$Stat.Rows * $d.AllDensity } else { $null }
                $uniq = if ($d.AllDensity -le 0.0001) { 'very selective (good for seeks)' } elseif ($d.AllDensity -le 0.01) { 'selective' } elseif ($d.AllDensity -le 0.1) { 'not very selective' } else { 'poorly selective (few distinct values; seeks on it alone rarely help)' }
                [void]$sb.AppendLine(("  ({0}): about {1} distinct values; on average each value appears {2} times. That is {3}." -f $d.Columns, (Format-Count $d.DistinctValues), $(if ($avgPerValue) { Format-Count $avgPerValue } else { '?' }), $uniq))
            }
        }
        if ($Detail -and $Detail.Histogram) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('WHAT THE DATA LOOKS LIKE (histogram, leading column only)')
            foreach ($l in (Get-HistogramInsights -Histogram $Detail.Histogram -Rows ([double]$Stat.Rows))) { [void]$sb.AppendLine("  $l") }
            [void]$sb.AppendLine("  Chart: dark bars = rows exactly equal to the bucket's upper value (EQ_ROWS); light bars = rows between buckets (RANGE_ROWS).")
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("VERDICT: priority $($Assessment.Priority) (score $($Assessment.Score)/100)")
        foreach ($r in $Assessment.Reasons) { [void]$sb.AppendLine("  - $r") }
        if ($Assessment.Priority -ne 'None') { [void]$sb.AppendLine("  Suggested: $($Assessment.SuggestedMethod)") }
        $sb.ToString()
    }
    #endregion

    #region ---------------------------------------------------------------- Assessment: indexes
    function Find-IndexRedundancy {
        <# sp_BlitzIndex-style: exact duplicates, duplicate keys, and left-prefix (overlapping) indexes. Rowstore only. #>
        param([object[]]$Indexes)
        $notes = @{}
        $rs = @($Indexes | Where-Object { $_.TypeCode -in 1, 2 -and -not $_.IsHypothetical -and $_.KeyColumns })
        foreach ($a in $rs) {
            $list = New-Object System.Collections.Generic.List[string]
            foreach ($b in $rs) {
                if ($a.IndexId -eq $b.IndexId) { continue }
                $sameFilter = [string]$a.FilterDefinition -eq [string]$b.FilterDefinition
                if (-not $sameFilter) { continue }
                $ak = [string]$a.KeyColumns; $bk = [string]$b.KeyColumns
                $ai = [string]$a.IncludedColumns; $bi = [string]$b.IncludedColumns
                if ($ak -eq $bk -and $ai -eq $bi) { $list.Add("Exact duplicate of [$($b.IndexName)]") }
                elseif ($ak -eq $bk) { $list.Add("Same key columns as [$($b.IndexName)] (different INCLUDEs: consider merging)") }
                elseif ($bk.StartsWith($ak + ', ')) {
                    if (-not $a.IsUnique -and -not $a.IsPrimaryKey -and $a.TypeCode -ne 1) { $list.Add("Keys are a left-prefix of [$($b.IndexName)]: probably redundant (check INCLUDEs first)") }
                }
            }
            if ($list.Count) { $notes[[int]$a.IndexId] = $list.ToArray() }
        }
        $notes
    }

    function Get-IndexAssessment {
        param(
            [Parameter(Mandatory)]$Index,
            $Fragmentation,         # output of Get-IndexFragmentation or $null
            $Columnstore,           # output of Get-ColumnstoreHealth or $null
            [string[]]$Redundancy,
            [Parameter(Mandatory)]$Settings,
            [double]$UptimeDays = 30
        )
        $reasons = New-Object System.Collections.Generic.List[string]
        $score = 0; $action = 'None'
        $pages = [double]$Index.InRowPages
        if ($Fragmentation -and $Fragmentation.Measured -and $Fragmentation.PageCount) { $pages = [double]$Fragmentation.PageCount }
        $reads = [double]$Index.Seeks + [double]$Index.Scans + [double]$Index.Lookups
        $writes = [double]$Index.Writes

        if ($Index.IsHypothetical) {
            return [pscustomobject]@{ IndexId = $Index.IndexId; Score = 0; Priority = 'Review'; Action = 'Review: hypothetical index (Tuning Advisor leftover), safe candidate to drop'; Reasons = @('Hypothetical index has no data') }
        }
        if ($Index.IsDisabled) {
            return [pscustomobject]@{ IndexId = $Index.IndexId; Score = 0; Priority = 'Review'; Action = 'Review: disabled. A REBUILD would re-enable it; drop it if it is not needed'; Reasons = @('Index is disabled') }
        }

        # Columnstore
        if ($Index.TypeCode -in 5, 6) {
            if ($Columnstore) {
                if ($Columnstore.DeletedPct -ge 20) { $score += 50; $action = 'REORGANIZE'; $reasons.Add("$($Columnstore.DeletedPct)% of columnstore rows are logically deleted") }
                if ([int]$Columnstore.DeltaStores -gt 1) { $score += 15; if ($action -eq 'None') { $action = 'REORGANIZE' }; $reasons.Add("$($Columnstore.DeltaStores) open/closed delta row groups not yet compressed") }
                if ($Columnstore.AvgRowsPerCompressedGroup -and [double]$Columnstore.AvgRowsPerCompressedGroup -lt 100000 -and [int]$Columnstore.Compressed -gt 10) { $score += 15; $reasons.Add('Row groups are small on average (under 100K rows): compression and segment elimination suffer') }
            } else { $reasons.Add('Columnstore health not checked') }
        } elseif ($Index.TypeCode -eq 0) {
            # Heap
            if ($Index.ForwardedFetches -and [double]$Index.ForwardedFetches -gt 0) {
                $score += [math]::Min(60, 20 + [math]::Log10([double]$Index.ForwardedFetches + 1) * 8)
                $action = 'REBUILD'
                $reasons.Add("Heap with $(Format-Count $Index.ForwardedFetches) forwarded-record fetches since startup (extra reads per row). Consider a clustered index; ALTER TABLE REBUILD also rebuilds every nonclustered index.")
            }
            if ($Fragmentation -and $Fragmentation.Measured -and $null -ne $Fragmentation.PageDensityPct -and $Fragmentation.PageDensityPct -lt $Settings.PageDensityThreshold) {
                $score += 20; $action = 'REBUILD'; $reasons.Add("Heap pages are only $($Fragmentation.PageDensityPct)% full")
            }
        } elseif ($pages -lt [double]$Settings.MinPageCount) {
            $reasons.Add(('Small ({0:N0} pages, under {1}): maintenance makes no measurable difference' -f $pages, $Settings.MinPageCount))
        } elseif ($Fragmentation -and $Fragmentation.Measured) {
            $f = [double]$Fragmentation.FragmentationPct
            $d = $Fragmentation.PageDensityPct
            $fill = if ([int]$Index.FillFactorPct -gt 0) { [int]$Index.FillFactorPct } else { 100 }
            if ($f -ge $Settings.RebuildThreshold) { $score += 45; $action = 'REBUILD'; $reasons.Add("Fragmentation $f% (at least $($Settings.RebuildThreshold)%)") }
            elseif ($f -ge $Settings.ReorgThreshold) { $score += 20; $action = 'REORGANIZE'; $reasons.Add("Fragmentation $f% (between $($Settings.ReorgThreshold)% and $($Settings.RebuildThreshold)%)") }
            if ($null -ne $d -and $d -lt [math]::Min($Settings.PageDensityThreshold, $fill - 10)) {
                $score += 30; if ($action -ne 'REBUILD') { $action = if ($d -lt 50) { 'REBUILD' } else { 'REORGANIZE' } }
                $reasons.Add("Pages are only $d% full (fill factor $fill%): wasted memory and I/O. Microsoft's current guidance treats low page density as often more important than fragmentation.")
            }
            if ($pages -gt 1000000) { $score += 10 } elseif ($pages -gt 100000) { $score += 5 }
            if ([double]$Index.Scans -gt [double]$Index.Seeks -and [double]$Index.Scans -gt 100) { $score += 10; $reasons.Add('Mostly read by scans, which is where fragmentation and density matter most') }
            elseif ([double]$Index.Seeks -gt 0 -and [double]$Index.Scans -lt [double]$Index.Seeks / 100) { $score -= 10; $reasons.Add('Used almost only for seeks: fragmentation barely matters for single-row lookups') }
        } else {
            if ($Fragmentation -and $Fragmentation.Skipped) { $reasons.Add($Fragmentation.Skipped) }
            else { $reasons.Add('Fragmentation not measured yet (use "Check fragmentation")') }
            if ($Index.LeafPageAllocations -and $pages -gt 0) {
                $splitRatio = [double]$Index.LeafPageAllocations / $pages
                if ($splitRatio -gt 0.10) { $score += 20; $action = 'Check'; $reasons.Add(('Proxy: {0:N0} leaf page allocations since startup = {1:P0} of the index size, which suggests heavy page splitting' -f [double]$Index.LeafPageAllocations, $splitRatio)) }
            }
        }

        # Usage-based adjustments (sp_BlitzIndex ideas)
        if ($Index.TypeCode -eq 2 -and -not $Index.IsPrimaryKey -and -not $Index.IsUniqueConstraint) {
            if ($reads -eq 0 -and $writes -gt 0) {
                $score = [math]::Max(0, $score - 30)
                $txt = "Unused since startup ($([math]::Round($UptimeDays)) days): 0 reads, $(Format-Count $writes) writes. Maintaining it costs you; review whether to drop it instead of rebuilding."
                if ($UptimeDays -lt 14) { $txt += ' Uptime is short, so do not decide on this alone.' }
                $reasons.Add($txt)
                if ($action -eq 'None' -or $UptimeDays -ge 14) { $action = 'Review (unused)' }
            } elseif ($reads -gt 0 -and $writes -gt 0 -and $writes / $reads -gt 10 -and $writes -gt 100000) {
                $reasons.Add(('Written {0:N0}x more than it is read: high-cost index' -f ($writes / $reads)))
            }
        }
        if ($Redundancy) { foreach ($r in $Redundancy) { $reasons.Add($r) }; if ($action -eq 'None') { $action = 'Review (redundant)' } }
        if ($Index.RowLockWaitMs -and [double]$Index.RowLockWaitMs -gt 600000) { $reasons.Add(('Row lock waits of {0:N0} minutes since startup: a blocking hot spot' -f ([double]$Index.RowLockWaitMs / 60000))) }
        if ($Index.PageLatchWaitMs -and [double]$Index.PageLatchWaitMs -gt 600000) { $reasons.Add(('Page latch waits of {0:N0} minutes: possible last-page insert contention' -f ([double]$Index.PageLatchWaitMs / 60000))) }
        if ($Index.AllowPageLocks -eq $false -and $action -eq 'REORGANIZE') { $action = 'REBUILD'; $reasons.Add('ALLOW_PAGE_LOCKS = OFF: REORGANIZE is not possible, only REBUILD') }

        $score = [int][math]::Max(0, [math]::Min(100, $score))
        $priority = if ($action -like 'Review*') { 'Review' } elseif ($score -ge 60) { 'High' } elseif ($score -ge 30) { 'Medium' } elseif ($score -gt 0) { 'Low' } else { 'None' }
        [pscustomobject]@{ IndexId = $Index.IndexId; Score = $score; Priority = $priority; Action = $action; Reasons = $reasons.ToArray() }
    }

    function Get-IndexPlainEnglish {
        param([Parameter(Mandatory)]$Index, $Assessment, $Fragmentation, $Columnstore, [double]$UptimeDays, $StartTime)
        $sb = New-Object System.Text.StringBuilder
        $kind = switch ([int]$Index.TypeCode) {
            0 { 'a HEAP (a table without a clustered index; rows are stored in no particular order)' }
            1 { 'the CLUSTERED index (it IS the table, stored in key order)' }
            2 { 'a NONCLUSTERED index (a sorted copy of some columns that points back to the table)' }
            5 { 'a CLUSTERED COLUMNSTORE index (the table stored column by column, compressed; good for analytics)' }
            6 { 'a NONCLUSTERED COLUMNSTORE index (a compressed column-wise copy for analytics)' }
            default { "a $($Index.IndexType) index" }
        }
        [void]$sb.AppendLine("'$($Index.IndexName)' is $kind.")
        if ($Index.KeyColumns) { [void]$sb.AppendLine("  Sorted by: $($Index.KeyColumns)") }
        if ($Index.IncludedColumns) { [void]$sb.AppendLine("  Also carries: $($Index.IncludedColumns)") }
        if ($Index.HasFilter) { [void]$sb.AppendLine("  Only covers rows where: $($Index.FilterDefinition)") }
        [void]$sb.AppendLine(("  Size: {0} for {1} rows across {2} partition(s). Compression: {3}. Fill factor: {4}." -f (Format-Size ([double]$Index.SizeMB)), (Format-Count $Index.RowCount), $Index.Partitions, $Index.Compression, $(if ([int]$Index.FillFactorPct -eq 0) { '100 (default)' } else { $Index.FillFactorPct })))
        [void]$sb.AppendLine('')
        $since = if ($StartTime) { "since the server started {0:yyyy-MM-dd} ({1:N0} days ago)" -f [datetime]$StartTime, $UptimeDays } else { 'since the server started' }
        [void]$sb.AppendLine("HOW IT IS USED ($since)")
        [void]$sb.AppendLine(("  Seeks   {0,15:N0}   jumped straight to the rows it needed (the efficient case)" -f [double]$Index.Seeks))
        [void]$sb.AppendLine(("  Scans   {0,15:N0}   read through all or a large part of it" -f [double]$Index.Scans))
        [void]$sb.AppendLine(("  Lookups {0,15:N0}   went back to the table for missing columns (only on clustered/heap)" -f [double]$Index.Lookups))
        [void]$sb.AppendLine(("  Writes  {0,15:N0}   had to be updated because of INSERT/UPDATE/DELETE" -f [double]$Index.Writes))
        [void]$sb.AppendLine(("  Last read: {0}    Last write: {1}" -f $(if ($Index.LastRead) { '{0:yyyy-MM-dd HH:mm} ({1})' -f [datetime]$Index.LastRead, (Format-Age $Index.LastRead) } else { 'never (since startup)' }), $(if ($Index.LastWrite) { '{0:yyyy-MM-dd HH:mm}' -f [datetime]$Index.LastWrite } else { 'never' })))
        [void]$sb.AppendLine('  Note: these counters reset when SQL Server restarts, the database goes offline, or (on older versions) the index is rebuilt. On an AG they are per replica.')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('PHYSICAL HEALTH')
        if ($Columnstore) {
            [void]$sb.AppendLine(("  Row groups: {0} compressed, {1} delta (not yet compressed), {2} tombstone. Deleted rows: {3}%. Avg rows per compressed group: {4}." -f $Columnstore.Compressed, $Columnstore.DeltaStores, $Columnstore.Tombstones, $Columnstore.DeletedPct, (Format-Count ([double]$Columnstore.AvgRowsPerCompressedGroup))))
        } elseif ($Fragmentation -and $Fragmentation.Measured) {
            [void]$sb.AppendLine(("  Fragmentation: {0}% of pages are out of order (like a book with chapters shuffled). Measured {1:HH:mm} with {2} mode over {3:N0} pages." -f $Fragmentation.FragmentationPct, [datetime]$Fragmentation.MeasuredAt, $Fragmentation.Mode, [double]$Fragmentation.PageCount))
            if ($null -ne $Fragmentation.PageDensityPct) { [void]$sb.AppendLine(("  Page fullness: pages are {0}% full on average (like boxes only partly packed: more boxes to carry for the same goods)." -f $Fragmentation.PageDensityPct)) }
            else { [void]$sb.AppendLine('  Page fullness: not available in LIMITED mode (SAMPLED mode shows it, for indexes under the SAMPLED size cap).') }
        } elseif ($Fragmentation -and $Fragmentation.Skipped) {
            [void]$sb.AppendLine("  $($Fragmentation.Skipped)")
        } else {
            [void]$sb.AppendLine('  Not measured yet. Click "Check fragmentation" (safe LIMITED mode, size-capped).')
        }
        if ($Index.LeafPageAllocations) { [void]$sb.AppendLine(("  New leaf pages allocated since startup: {0:N0} (includes page splits)." -f [double]$Index.LeafPageAllocations)) }
        if ($Index.ForwardedFetches -and $Index.TypeCode -eq 0) { [void]$sb.AppendLine(("  Forwarded fetches: {0:N0} (rows that moved and left a pointer behind)." -f [double]$Index.ForwardedFetches)) }
        if ($Assessment) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("VERDICT: $($Assessment.Action)   priority $($Assessment.Priority) (score $($Assessment.Score)/100)")
            foreach ($r in $Assessment.Reasons) { [void]$sb.AppendLine("  - $r") }
        }
        $sb.ToString()
    }
    #endregion

    #region ---------------------------------------------------------------- Script generation (TEXT ONLY, never executed)
    function New-ScriptSafetyHeader {
        param([Parameter(Mandatory)]$ServerInfo, [Parameter(Mandatory)]$DatabaseInfo, [Parameter(Mandatory)][string]$Environment,
              [string]$Title, [string[]]$Notes, [bool]$AgDatabase)
        $srv = [string]$ServerInfo.ServerName; $db = [string]$DatabaseInfo.DatabaseName
        $isProd = Test-IsProductionEnvironment $Environment
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('/*' + ('=' * 90))
        [void]$sb.AppendLine("  $Title")
        [void]$sb.AppendLine("  Generated by $($script:ToolName) $($script:ToolVersion) (READ-ONLY TOOL) on $(Get-Date -Format 'yyyy-MM-dd HH:mm') by $env:USERNAME")
        [void]$sb.AppendLine("  Target      : $srv / $db")
        [void]$sb.AppendLine("  Environment : $Environment")
        [void]$sb.AppendLine("  Recovery    : $($DatabaseInfo.RecoveryModel)   Data free: $(Format-Size ([double]$DatabaseInfo.DataFreeMB))   Log size: $(Format-Size ([double]$DatabaseInfo.LogSizeMB)) ($($DatabaseInfo.LogUsedPct)% used)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('  THIS SCRIPT HAS NOT BEEN EXECUTED. The tool only generates text.')
        [void]$sb.AppendLine('  Review it, get change approval, and run it in a maintenance window with someone watching')
        [void]$sb.AppendLine('  blocking, log usage, AG send/redo queues, and tempdb.')
        if ($Notes) { [void]$sb.AppendLine(''); foreach ($n in $Notes) { [void]$sb.AppendLine("  * $n") } }
        [void]$sb.AppendLine(('=' * 90) + '*/')
        [void]$sb.AppendLine("USE $(ConvertTo-QuotedName $db);")
        [void]$sb.AppendLine('GO')
        [void]$sb.AppendLine('SET NOEXEC OFF;')
        [void]$sb.AppendLine('DECLARE @srv sysname = @@SERVERNAME, @db sysname = DB_NAME();')
        [void]$sb.AppendLine("IF @srv <> $(ConvertTo-SqlLiteral $srv) OR @db <> $(ConvertTo-SqlLiteral $db)")
        [void]$sb.AppendLine('BEGIN')
        [void]$sb.AppendLine("    RAISERROR(N'SAFETY STOP: connected to %s / %s, but this script was generated for $($srv.Replace("'", "''").Replace('%', '%%')) / $($db.Replace("'", "''").Replace('%', '%%')).', 16, 1, @srv, @db);")
        [void]$sb.AppendLine('    SET NOEXEC ON;')
        [void]$sb.AppendLine('END;')
        if ($AgDatabase -and $ServerInfo.Capabilities.FnHadrIsPrimary) {
            [void]$sb.AppendLine('IF sys.fn_hadr_is_primary_replica(DB_NAME()) = 0')
            [void]$sb.AppendLine('BEGIN')
            [void]$sb.AppendLine("    RAISERROR(N'SAFETY STOP: this replica is not the AG primary.', 16, 1);")
            [void]$sb.AppendLine('    SET NOEXEC ON;')
            [void]$sb.AppendLine('END;')
        }
        if ($isProd) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('-- >>> PRODUCTION SAFETY LATCH: delete the next 2 lines ONLY after the change is approved. <<<')
            [void]$sb.AppendLine("RAISERROR(N'SAFETY LATCH: production script. Remove the latch after approval.', 16, 1);")
            [void]$sb.AppendLine('SET NOEXEC ON;')
        }
        [void]$sb.AppendLine('SET LOCK_TIMEOUT 60000;   -- give up instead of queueing behind long blockers')
        [void]$sb.AppendLine('SET DEADLOCK_PRIORITY LOW;')
        [void]$sb.AppendLine('')
        $sb.ToString()
    }

    function Get-IndexMaintenanceOptionIssues {
        <# Validates requested options against version/edition/index type; returns adjusted options + warnings. #>
        param([Parameter(Mandatory)]$Index, [Parameter(Mandatory)][hashtable]$Options, [Parameter(Mandatory)]$ServerInfo)
        $caps = $ServerInfo.Capabilities
        $o = $Options.Clone()
        $w = New-Object System.Collections.Generic.List[string]
        $partitioned = $null -ne $o.Partition -and [string]$o.Partition -ne 'ALL'

        if ($o.Action -eq 'REORGANIZE') {
            if ($Index.TypeCode -eq 0) { $w.Add('A heap cannot be reorganized. Use REBUILD (ALTER TABLE ... REBUILD) instead.'); $o.Invalid = $true }
            if ($Index.AllowPageLocks -eq $false -and $Index.TypeCode -in 1, 2) { $w.Add('ALLOW_PAGE_LOCKS = OFF: REORGANIZE will fail on this index.'); $o.Invalid = $true }
            $o.Online = $false; $o.Resumable = $false; $o.WaitAtLowPriority = $false; $o.SortInTempdb = $false
            if ($o.MaxDop) { $w.Add('REORGANIZE is always single-threaded and always online; MAXDOP does not apply.'); $o.MaxDop = $null }
            return [pscustomobject]@{ Options = $o; Warnings = $w.ToArray() }
        }

        if ($Index.IsDisabled) { $w.Add('Index is DISABLED: REBUILD will ENABLE it (and on a disabled clustered index makes the whole table usable again). Confirm that this is intended.') }
        if ($o.Online) {
            if (-not $caps.OnlineRebuild) { $w.Add("ONLINE = ON needs Enterprise or Developer edition (this is $($ServerInfo.Edition)). Switched to OFFLINE: the table is locked for the whole rebuild."); $o.Online = $false }
            elseif ($Index.TypeCode -in 5, 6 -and -not $caps.OnlineColumnstoreRebuild) { $w.Add('Online columnstore rebuild needs SQL Server 2019 or later. Switched to OFFLINE; consider REORGANIZE instead.'); $o.Online = $false }
            elseif ($Index.TypeCode -in 3, 4) { $w.Add('XML and spatial indexes cannot be rebuilt online. Switched to OFFLINE.'); $o.Online = $false }
            elseif ($partitioned -and -not $caps.OnlinePartitionRebuild) { $w.Add('Online single-partition rebuild needs SQL Server 2014 or later. Switched to OFFLINE.'); $o.Online = $false }
        }
        if (-not $o.Online) {
            if ($o.Resumable) { $w.Add('RESUMABLE requires ONLINE = ON: removed.'); $o.Resumable = $false }
            if ($o.WaitAtLowPriority) { $w.Add('WAIT_AT_LOW_PRIORITY only applies to ONLINE operations here: removed.'); $o.WaitAtLowPriority = $false }
            $w.Add('OFFLINE rebuild: readers and writers are blocked for the duration (schema-modification lock).')
        }
        if ($o.Resumable) {
            if (-not $caps.ResumableRebuild) { $w.Add('RESUMABLE needs SQL Server 2017 or later (Enterprise): removed.'); $o.Resumable = $false }
            elseif ($o.SortInTempdb) { $w.Add('SORT_IN_TEMPDB = ON cannot be combined with RESUMABLE: SORT_IN_TEMPDB removed.'); $o.SortInTempdb = $false }
            elseif ($partitioned) { $w.Add('RESUMABLE with a single partition is not generated by this tool (version-dependent): RESUMABLE removed.'); $o.Resumable = $false }
            elseif ($Index.TypeCode -eq 0 -or $Index.TypeCode -in 5, 6) { $w.Add('RESUMABLE is not supported for heaps or columnstore: removed.'); $o.Resumable = $false }
        }
        if ($o.WaitAtLowPriority -and -not $caps.WaitAtLowPriority) { $w.Add('WAIT_AT_LOW_PRIORITY needs SQL Server 2014 or later: removed.'); $o.WaitAtLowPriority = $false }
        if ($Index.TypeCode -eq 0) {
            $w.Add('HEAP: ALTER TABLE ... REBUILD also rebuilds EVERY nonclustered index on the table (the row locators change). Size the window for the whole table.')
            if ($o.SortInTempdb) { $o.SortInTempdb = $false }
            if ($o.WaitAtLowPriority) { $w.Add('WAIT_AT_LOW_PRIORITY is not generated for ALTER TABLE REBUILD: removed.'); $o.WaitAtLowPriority = $false }
        }
        [pscustomobject]@{ Options = $o; Warnings = $w.ToArray() }
    }

    function Get-MaintenanceEstimate {
        param([Parameter(Mandatory)]$Index, [Parameter(Mandatory)][hashtable]$Options, $DatabaseInfo, [double]$TableReservedMB)
        $sizeMB = [double]$Index.SizeMB
        if ($Index.TypeCode -eq 0 -and $Options.Action -eq 'REBUILD' -and $TableReservedMB) { $sizeMB = $TableReservedMB }
        $lines = New-Object System.Collections.Generic.List[string]
        if ($Options.Action -eq 'REBUILD') {
            $dataNeed = $sizeMB * $(if ($Options.Online) { 1.3 } else { 1.2 })
            $lines.Add(('Needs about {0} of free space in the data files (has {1}).' -f (Format-Size $dataNeed), (Format-Size ([double]$DatabaseInfo.DataFreeMB))))
            if ($Options.SortInTempdb) { $lines.Add(('SORT_IN_TEMPDB: needs about {0} in tempdb.' -f (Format-Size $sizeMB))) }
            $logMB = if ($DatabaseInfo.RecoveryModel -eq 'FULL' -or $Options.Online) { $sizeMB * 1.1 } else { $sizeMB * 0.1 }
            $lines.Add(('Transaction log: about {0} ({1} recovery{2}). On an AG this must also be shipped to and redone on every secondary.' -f (Format-Size $logMB), $DatabaseInfo.RecoveryModel, $(if ($Options.Online) { ', online rebuilds are fully logged' } else { '' })))
            if ($dataNeed -gt [double]$DatabaseInfo.DataFreeMB) { $lines.Add('WARNING: free space in the data files looks too small: the files will grow.') }
        } else {
            $lines.Add(('REORGANIZE: online, single-threaded, can be stopped at any time without losing finished work. Log use can be up to {0} or more on a badly fragmented index.' -f (Format-Size ($sizeMB * 1.5))))
        }
        $lines.ToArray()
    }

    function New-IndexMaintenanceScript {
        <#
            Options: Action (REBUILD|REORGANIZE), Online, MaxDop, SortInTempdb, Resumable, MaxDurationMin,
                     WaitAtLowPriority, LowPriorityMinutes, Partition ('ALL' or int), LobCompaction, CompressAllRowGroups
        #>
        param(
            [Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName,
            [Parameter(Mandatory)][object[]]$Indexes, [Parameter(Mandatory)][hashtable]$Options,
            [Parameter(Mandatory)]$ServerInfo, [Parameter(Mandatory)]$DatabaseInfo, [Parameter(Mandatory)][string]$Environment,
            [double]$TableReservedMB, [switch]$IncludeHeader
        )
        $tbl = (ConvertTo-QuotedName $SchemaName) + '.' + (ConvertTo-QuotedName $TableName)
        $body = New-Object System.Text.StringBuilder
        $allNotes = New-Object System.Collections.Generic.List[string]

        foreach ($ix in $Indexes) {
            $chk = Get-IndexMaintenanceOptionIssues -Index $ix -Options $Options -ServerInfo $ServerInfo
            $o = $chk.Options
            [void]$body.AppendLine("------------------------------------------------------------------------------------------")
            [void]$body.AppendLine(("-- {0}  ({1}, {2}, {3} rows)" -f $ix.IndexName, $ix.IndexType, (Format-Size ([double]$ix.SizeMB)), (Format-Count $ix.RowCount)))
            foreach ($w in $chk.Warnings) { [void]$body.AppendLine("-- NOTE: $w") }
            foreach ($e in (Get-MaintenanceEstimate -Index $ix -Options $o -DatabaseInfo $DatabaseInfo -TableReservedMB $TableReservedMB)) { [void]$body.AppendLine("-- EST : $e") }
            if ($o.ContainsKey('Invalid') -and $o.Invalid) { [void]$body.AppendLine('-- (no statement generated)'); [void]$body.AppendLine(''); continue }
            if ($ix.IsHypothetical) { [void]$body.AppendLine('-- Hypothetical index: nothing to maintain. (no statement generated)'); [void]$body.AppendLine(''); continue }

            $part = if ($null -ne $o.Partition -and [string]$o.Partition -ne 'ALL') { " PARTITION = $([int]$o.Partition)" } else { '' }
            $with = New-Object System.Collections.Generic.List[string]
            if ($o.Action -eq 'REORGANIZE') {
                if ($ix.TypeCode -in 5, 6) {
                    if ($o.CompressAllRowGroups -and $ServerInfo.Capabilities.CompressAllRowGroups) { $with.Add('COMPRESS_ALL_ROW_GROUPS = ON') }
                } else {
                    $with.Add("LOB_COMPACTION = $(if ($o.ContainsKey('LobCompaction') -and -not $o.LobCompaction) { 'OFF' } else { 'ON' })")
                }
                $stmt = "ALTER INDEX $(ConvertTo-QuotedName $ix.IndexName) ON $tbl REORGANIZE$part"
            } else {
                if ($o.Online) {
                    if ($o.WaitAtLowPriority) { $with.Add("ONLINE = ON (WAIT_AT_LOW_PRIORITY (MAX_DURATION = $([int]$o.LowPriorityMinutes) MINUTES, ABORT_AFTER_WAIT = SELF))") }
                    else { $with.Add('ONLINE = ON') }
                } else { $with.Add('ONLINE = OFF') }
                if ($o.Resumable) { $with.Add('RESUMABLE = ON'); if ($o.MaxDurationMin) { $with.Add("MAX_DURATION = $([int]$o.MaxDurationMin) MINUTES") } }
                if ($null -ne $o.MaxDop -and [string]$o.MaxDop -ne '') { $with.Add("MAXDOP = $([int]$o.MaxDop)") }
                if ($ix.TypeCode -ne 0 -and $ix.TypeCode -notin 5, 6) { $with.Add("SORT_IN_TEMPDB = $(if ($o.SortInTempdb) { 'ON' } else { 'OFF' })") }
                if ($ix.TypeCode -eq 0) { $stmt = "ALTER TABLE $tbl REBUILD$part" }
                else { $stmt = "ALTER INDEX $(ConvertTo-QuotedName $ix.IndexName) ON $tbl REBUILD$part" }
            }
            if ($with.Count) { $stmt += "`r`nWITH (" + ($with -join ', ') + ')' }
            [void]$body.AppendLine("PRINT CONVERT(nvarchar(30), SYSDATETIME(), 121) + N'  starting $($ix.IndexName.Replace("'", "''"))';")
            [void]$body.AppendLine($stmt + ';')
            if ($o.Resumable) { [void]$body.AppendLine("-- If it is paused or times out: ALTER INDEX $(ConvertTo-QuotedName $ix.IndexName) ON $tbl RESUME;  (or ABORT;)") }
            [void]$body.AppendLine('')
        }

        $isProd = Test-IsProductionEnvironment $Environment
        if ($isProd) { $allNotes.Add('PRODUCTION: prefer ONLINE + WAIT_AT_LOW_PRIORITY (ABORT_AFTER_WAIT = SELF) + RESUMABLE, and a MAXDOP that leaves CPUs for the workload.') }
        if ($Options.Action -eq 'REBUILD') { $allNotes.Add('A REBUILD also updates that index''s statistics with FULLSCAN; do not schedule a separate stats update for the same index right after it.') }
        $allNotes.Add('Partitioned tables: rebuilding one partition (PARTITION = n) is much cheaper than PARTITION = ALL.')

        if ($IncludeHeader) {
            $ag = [int]$DatabaseInfo.IsInAg -eq 1
            (New-ScriptSafetyHeader -ServerInfo $ServerInfo -DatabaseInfo $DatabaseInfo -Environment $Environment -Title "INDEX MAINTENANCE: $tbl" -Notes $allNotes.ToArray() -AgDatabase $ag) + $body.ToString()
        } else { $body.ToString() }
    }

    function New-StatsMaintenanceScript {
        <# Options: Method (FULLSCAN|SAMPLE|RESAMPLE), SamplePercent, MaxDop, PersistSamplePercent, Partitions (int[] for incremental RESAMPLE) #>
        param(
            [Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName,
            [Parameter(Mandatory)][object[]]$Statistics, [Parameter(Mandatory)][hashtable]$Options,
            [Parameter(Mandatory)]$ServerInfo, [Parameter(Mandatory)]$DatabaseInfo, [Parameter(Mandatory)][string]$Environment,
            [double]$TableDataMB, [switch]$IncludeHeader
        )
        $caps = $ServerInfo.Capabilities
        $tbl = (ConvertTo-QuotedName $SchemaName) + '.' + (ConvertTo-QuotedName $TableName)
        $body = New-Object System.Text.StringBuilder
        $notes = New-Object System.Collections.Generic.List[string]
        $method = [string]$Options.Method
        if ($method -eq 'SAMPLE') {
            $pct = [double]$Options.SamplePercent
            if ($pct -le 0 -or $pct -gt 100) { throw 'Sample percent must be between 0 and 100.' }
            if ($pct -ge 100) { $method = 'FULLSCAN' }
        }

        foreach ($st in $Statistics) {
            $with = New-Object System.Collections.Generic.List[string]
            $local = New-Object System.Collections.Generic.List[string]
            switch ($method) {
                'FULLSCAN' { $with.Add('FULLSCAN') }
                'SAMPLE'   { $with.Add(('SAMPLE {0} PERCENT' -f ([double]$Options.SamplePercent).ToString([System.Globalization.CultureInfo]::InvariantCulture))) }
                'RESAMPLE' {
                    if ($st.IsIncremental -and $Options.Partitions) { $with.Add('RESAMPLE ON PARTITIONS (' + (($Options.Partitions | ForEach-Object { [int]$_ }) -join ', ') + ')') }
                    else { $with.Add('RESAMPLE') }
                }
            }
            if ($Options.PersistSamplePercent -and $method -ne 'RESAMPLE') {
                if ($caps.PersistSamplePercent) { $with.Add('PERSIST_SAMPLE_PERCENT = ON') } else { $local.Add('PERSIST_SAMPLE_PERCENT needs SQL Server 2016 SP1 CU4 / 2017 CU1 or later: removed.') }
            }
            if ($null -ne $Options.MaxDop -and [string]$Options.MaxDop -ne '') {
                if ($caps.UpdateStatsMaxDop) { $with.Add("MAXDOP = $([int]$Options.MaxDop)") } else { $local.Add('MAXDOP for UPDATE STATISTICS needs 2014 SP3 / 2016 SP2 / 2017 CU3 or later: removed (instance MAXDOP applies).') }
            }
            if ($st.NoRecompute) { $with.Add('NORECOMPUTE'); $local.Add('NORECOMPUTE kept: without it, this update would silently switch automatic updates back ON for this statistic.') }
            if ($st.IsIncremental -and $method -ne 'RESAMPLE') { $local.Add('Incremental statistic: RESAMPLE ON PARTITIONS (n) refreshes only the partitions that changed and is much cheaper.') }

            $readMB = if ($st.IndexSizeMB) { [double]$st.IndexSizeMB } else { $TableDataMB }
            $readPct = switch ($method) { 'FULLSCAN' { 100 } 'SAMPLE' { [double]$Options.SamplePercent } default { if ($st.Rows -and $st.RowsSampled) { 100.0 * [double]$st.RowsSampled / [double]$st.Rows } else { 100 } } }
            [void]$body.AppendLine("------------------------------------------------------------------------------------------")
            [void]$body.AppendLine(("-- {0} on ({1}): {2} rows, last updated {3}, {4} changes since" -f $st.StatName, $st.Columns, (Format-Count $st.Rows), $(if ($st.LastUpdated) { '{0:yyyy-MM-dd HH:mm}' -f [datetime]$st.LastUpdated } else { 'never' }), (Format-Count $st.Modifications)))
            if ($readMB) { [void]$body.AppendLine(("-- EST : reads about {0} ({1:N1}% of {2}). Uses a shared schema-stability lock only; does not block readers or writers, but it competes for I/O and CPU." -f (Format-Size ($readMB * $readPct / 100)), $readPct, (Format-Size $readMB))) }
            foreach ($n in $local) { [void]$body.AppendLine("-- NOTE: $n") }
            [void]$body.AppendLine("UPDATE STATISTICS $tbl ($(ConvertTo-QuotedName $st.StatName))")
            [void]$body.AppendLine('WITH ' + ($with -join ', ') + ';')
            [void]$body.AppendLine('')
        }
        $notes.Add('Updating statistics invalidates cached plans that use them, so expect a short burst of recompiles.')
        if ($method -eq 'FULLSCAN' -and $TableDataMB -gt 102400) { $notes.Add('FULLSCAN on a table over 100 GB is expensive: consider SAMPLE or RESAMPLE, or run it per statistic across several windows.') }
        if ($IncludeHeader) {
            $ag = [int]$DatabaseInfo.IsInAg -eq 1
            (New-ScriptSafetyHeader -ServerInfo $ServerInfo -DatabaseInfo $DatabaseInfo -Environment $Environment -Title "STATISTICS MAINTENANCE: $tbl" -Notes $notes.ToArray() -AgDatabase $ag) + $body.ToString()
        } else { $body.ToString() }
    }

    function New-OlaIndexOptimizeCommand {
        <# The equivalent call to Ola Hallengren's IndexOptimize, for shops that standardise on it. Text only. #>
        param(
            [Parameter(Mandatory)][string]$DatabaseName, [Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName,
            [string[]]$ObjectNames,     # index or statistics names; empty = whole table
            [ValidateSet('Index', 'Statistics')][string]$Kind = 'Index',
            [hashtable]$Options, $Settings, $ServerInfo, [string]$OlaDatabase
        )
        $names = if ($ObjectNames -and $ObjectNames.Count) { ($ObjectNames | ForEach-Object { "$DatabaseName.$SchemaName.$TableName.$_" }) -join ',' } else { "$DatabaseName.$SchemaName.$TableName" }
        $p = New-Object System.Collections.Generic.List[string]
        $p.Add("@Databases = $(ConvertTo-SqlLiteral $DatabaseName)")
        $p.Add("@Indexes = $(ConvertTo-SqlLiteral $names)")
        if ($Kind -eq 'Index') {
            $online = $Options.Online -and $ServerInfo.Capabilities.OnlineRebuild
            $high = if ($Options.Action -eq 'REORGANIZE') { 'INDEX_REORGANIZE' } elseif ($online) { 'INDEX_REBUILD_ONLINE,INDEX_REORGANIZE' } else { 'INDEX_REBUILD_OFFLINE' }
            $p.Add('@FragmentationLow = NULL')
            $p.Add("@FragmentationMedium = N'INDEX_REORGANIZE'")
            $p.Add("@FragmentationHigh = N'$high'")
            $p.Add("@FragmentationLevel1 = $([int]$Settings.ReorgThreshold)")
            $p.Add("@FragmentationLevel2 = $([int]$Settings.RebuildThreshold)")
            $p.Add("@MinNumberOfPages = $([int]$Settings.MinPageCount)")
            if ($null -ne $Options.MaxDop -and [string]$Options.MaxDop -ne '') { $p.Add("@MaxDOP = $([int]$Options.MaxDop)") }
            if ($Options.SortInTempdb) { $p.Add("@SortInTempdb = N'Y'") }
            if ($online -and $Options.WaitAtLowPriority -and $ServerInfo.Capabilities.WaitAtLowPriority) {
                $p.Add("@WaitAtLowPriorityMaxDuration = $([int]$Options.LowPriorityMinutes)")
                $p.Add("@WaitAtLowPriorityAbortAfterWait = N'SELF'")
            }
            if ($online -and $Options.Resumable -and $ServerInfo.Capabilities.ResumableRebuild) { $p.Add("@Resumable = N'Y'") }
            $p.Add('@UpdateStatistics = NULL')
        } else {
            $p.Add('@FragmentationLow = NULL'); $p.Add('@FragmentationMedium = NULL'); $p.Add('@FragmentationHigh = NULL')
            $p.Add("@UpdateStatistics = N'ALL'")
            $p.Add("@OnlyModifiedStatistics = N'Y'")
            switch ([string]$Options.Method) {
                'FULLSCAN' { $p.Add('@StatisticsSample = 100') }
                'SAMPLE'   { $p.Add("@StatisticsSample = $([int][math]::Ceiling([double]$Options.SamplePercent))") }
                'RESAMPLE' { $p.Add("@StatisticsResample = N'Y'") }
            }
            if ($null -ne $Options.MaxDop -and [string]$Options.MaxDop -ne '') { $p.Add("@MaxDOP = $([int]$Options.MaxDop)") }
        }
        $p.Add('@LockTimeout = 60')
        $p.Add("@LogToTable = N'Y'")
        $p.Add("@Execute = N'Y'")
        $where = if ($OlaDatabase) { "-- IndexOptimize was found in [$OlaDatabase]." } else { '-- IndexOptimize was NOT found in master or this database: install the Maintenance Solution first (ola.hallengren.com).' }
        $warn = if (("$SchemaName$TableName" + ($ObjectNames -join '')) -match '[\[\]%]') { "`r`n-- WARNING: a name contains [ ] or %, which IndexOptimize treats as LIKE wildcards. Check what it matches with @Execute = N'N' first." } else { '' }
        $db = if ($OlaDatabase) { ConvertTo-QuotedName $OlaDatabase } else { '[master]' }
        "-- Equivalent Ola Hallengren IndexOptimize call (text only, not executed).`r`n$where$warn`r`n-- Tip: run it once with @Execute = N'N' to print the commands it would run.`r`nEXECUTE $db.dbo.IndexOptimize`r`n    " + ($p -join ",`r`n    ") + ";`r`n"
    }

    function New-MissingIndexScript {
        param([Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName, [Parameter(Mandatory)]$Missing, [bool]$Online)
        $keys = @($Missing.EqualityColumns, $Missing.InequalityColumns | Where-Object { $_ }) -join ', '
        $nameCols = ($keys -replace '[\[\]\s]', '' -replace ',', '_')
        if ($nameCols.Length -gt 80) { $nameCols = $nameCols.Substring(0, 80) }
        $ixName = "IX_$($TableName)_$nameCols"
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("-- Missing-index suggestion (improvement score $($Missing.ImprovementScore), $($Missing.Seeks) seeks, est. impact $($Missing.AvgImpactPct)%).")
        [void]$sb.AppendLine('-- The DMV does not know about your existing indexes, column order, or write cost. Before creating:')
        [void]$sb.AppendLine('--   1) check whether an existing index can be widened instead (see the Indexes tab: duplicates and overlaps),')
        [void]$sb.AppendLine('--   2) put the most selective equality column first, 3) keep INCLUDE lists short.')
        [void]$sb.AppendLine("/* CREATE NONCLUSTERED INDEX $(ConvertTo-QuotedName $ixName)")
        [void]$sb.AppendLine("ON $(ConvertTo-QuotedName $SchemaName).$(ConvertTo-QuotedName $TableName) ($keys)")
        if ($Missing.IncludedColumns) { [void]$sb.AppendLine("INCLUDE ($($Missing.IncludedColumns))") }
        [void]$sb.AppendLine("WITH (ONLINE = $(if ($Online) { 'ON' } else { 'OFF' }), SORT_IN_TEMPDB = ON, MAXDOP = 4); */")
        $sb.ToString()
    }
    #endregion

    #region ---------------------------------------------------------------- Procedure plan analysis (opt-in, read-only)
    <#
        Reads the plan a procedure ALREADY has (plan cache, or Query Store if enabled). It never executes,
        compiles or recompiles the procedure, and never asks for an estimated or actual plan.
        From the plan it finds the statistics the optimizer used, the indexes it seeks/scans/looks up,
        plan warnings and missing-index hints. It then scores which stats updates or index maintenance
        would actually matter for THIS procedure.
    #>

    function ConvertTo-HexString {
        param($Bytes)
        if ($null -eq $Bytes -or $Bytes -is [DBNull]) { return $null }
        '0x' + (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '')
    }

    function Get-ProcedureList {
        param([Parameter(Mandatory)]$Context)
        $sql = @"
SELECT p.object_id AS ObjectId, s.name AS SchemaName, p.name AS ProcName, p.modify_date AS Modified
FROM sys.procedures AS p JOIN sys.schemas AS s ON s.schema_id = p.schema_id
WHERE p.is_ms_shipped = 0
ORDER BY s.name, p.name
OPTION (MAXDOP 1);
"@
        Invoke-SafeQuery -Context $Context -Sql $sql -Label 'procedure list' -TimeoutSec ([math]::Max(60, $Context.CommandTimeoutSec))
    }

    function Get-ProceduresReferencingTable {
        <# object_ids of procedures whose definition references the table (metadata only). #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId)
        $sql = @"
SELECT DISTINCT d.referencing_id AS ObjectId
FROM sys.sql_expression_dependencies AS d
JOIN sys.procedures AS p ON p.object_id = d.referencing_id
WHERE d.referenced_id = @ObjectId
   OR (d.referenced_id IS NULL
       AND d.referenced_entity_name = OBJECT_NAME(@ObjectId)
       AND ISNULL(d.referenced_schema_name, OBJECT_SCHEMA_NAME(@ObjectId)) = OBJECT_SCHEMA_NAME(@ObjectId))
OPTION (MAXDOP 1);
"@
        @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'procedures referencing table')) | ForEach-Object { [int]$_.ObjectId }
    }

    function Get-ProcedurePlanSources {
        <# Which plans exist for the procedure: plan cache entries (one per set-options combination) and Query Store. #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)]$ServerInfo)
        $sources = New-Object System.Collections.Generic.List[object]
        $warnings = New-Object System.Collections.Generic.List[string]
        try {
            $sql = @"
SELECT TOP (20) ps.plan_handle AS PlanHandle, ps.cached_time AS CachedTime, ps.last_execution_time AS LastExecution,
       ps.execution_count AS Executions,
       CAST(ps.total_worker_time / 1000.0 / NULLIF(ps.execution_count, 0) AS decimal(19,2)) AS AvgCpuMs,
       CAST(ps.total_elapsed_time / 1000.0 / NULLIF(ps.execution_count, 0) AS decimal(19,2)) AS AvgElapsedMs,
       CAST(ps.total_logical_reads * 1.0 / NULLIF(ps.execution_count, 0) AS decimal(19,1)) AS AvgLogicalReads
FROM sys.dm_exec_procedure_stats AS ps
WHERE ps.database_id = DB_ID() AND ps.object_id = @ObjectId
ORDER BY ps.last_execution_time DESC
OPTION (MAXDOP 1);
"@
            foreach ($r in @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'cached plans of procedure'))) {
                $hex = ConvertTo-HexString $r.PlanHandle
                $sources.Add([pscustomobject]@{
                    Key = "cache:$hex"; Source = 'Plan cache'; PlanHandle = $r.PlanHandle; LastExecution = $r.LastExecution; CachedTime = $r.CachedTime
                    Executions = $r.Executions; AvgCpuMs = $r.AvgCpuMs; AvgElapsedMs = $r.AvgElapsedMs; AvgLogicalReads = $r.AvgLogicalReads
                    Label = ('Plan cache: last run {0:yyyy-MM-dd HH:mm}, {1:N0} executions, avg {2:N1} ms CPU, {3:N0} reads (cached {4:yyyy-MM-dd HH:mm})' -f $r.LastExecution, [double]$r.Executions, [double]$r.AvgCpuMs, [double]$r.AvgLogicalReads, $r.CachedTime)
                })
            }
        } catch { $warnings.Add("Plan cache not readable: $($_.Exception.Message) (reading it needs VIEW SERVER STATE)") }

        if ($ServerInfo.Capabilities.Major -ge 13) {
            try {
                $qs = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'query store state' -Sql 'SELECT actual_state_desc AS State FROM sys.database_query_store_options;'))[0]
                if ($qs -and $qs.State -in 'READ_WRITE', 'READ_ONLY') {
                    $c = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'query store plans of procedure' -Parameters @{ ObjectId = $ObjectId } -Sql @"
SELECT COUNT(DISTINCT q.query_id) AS Queries, COUNT(*) AS Plans, MAX(p.last_execution_time) AS LastExecution
FROM sys.query_store_query AS q JOIN sys.query_store_plan AS p ON p.query_id = q.query_id
WHERE q.object_id = @ObjectId
OPTION (MAXDOP 1);
"@))[0]
                    if ($c -and [int]$c.Plans -gt 0) {
                        $sources.Add([pscustomobject]@{
                            Key = 'qs'; Source = 'Query Store'; PlanHandle = $null; LastExecution = $c.LastExecution; CachedTime = $null
                            Executions = $null; AvgCpuMs = $null; AvgElapsedMs = $null; AvgLogicalReads = $null
                            Label = ('Query Store: latest plan of each of {0} statement(s) ({1} plans recorded), last run {2:yyyy-MM-dd HH:mm}' -f $c.Queries, $c.Plans, $c.LastExecution)
                        })
                    }
                } elseif ($qs) { $warnings.Add("Query Store is $($qs.State) in this database.") }
            } catch { $warnings.Add("Query Store not readable: $($_.Exception.Message)") }
        }
        [pscustomobject]@{ Sources = $sources.ToArray(); Warnings = $warnings.ToArray() }
    }

    function Get-CachedPlanText {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][byte[]]$PlanHandle)
        $r = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'cached plan xml' -Parameters @{ PlanHandle = $PlanHandle } -Sql @"
SELECT query_plan AS PlanText FROM sys.dm_exec_text_query_plan(@PlanHandle, 0, -1);
"@))[0]
        if ($r) { [string]$r.PlanText } else { $null }
    }

    function Get-CachedStatementStats {
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][byte[]]$PlanHandle)
        $sql = @"
SELECT qs.statement_start_offset AS StartOffset, qs.query_hash AS QueryHash, qs.execution_count AS Executions,
       CAST(qs.total_worker_time / 1000.0 / NULLIF(qs.execution_count, 0) AS decimal(19,2)) AS AvgCpuMs,
       CAST(qs.total_elapsed_time / 1000.0 / NULLIF(qs.execution_count, 0) AS decimal(19,2)) AS AvgElapsedMs,
       CAST(qs.total_logical_reads * 1.0 / NULLIF(qs.execution_count, 0) AS decimal(19,1)) AS AvgLogicalReads,
       CAST(qs.total_rows * 1.0 / NULLIF(qs.execution_count, 0) AS decimal(19,1)) AS AvgRows,
       qs.last_rows AS LastRows, qs.min_rows AS MinRows, qs.max_rows AS MaxRows,
       SUBSTRING(st.text, qs.statement_start_offset / 2 + 1,
                 (CASE WHEN qs.statement_end_offset = -1 THEN DATALENGTH(st.text) ELSE qs.statement_end_offset END - qs.statement_start_offset) / 2 + 1) AS StatementText
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY sys.dm_exec_sql_text(qs.plan_handle) AS st
WHERE qs.plan_handle = @PlanHandle
ORDER BY qs.statement_start_offset
OPTION (MAXDOP 1);
"@
        foreach ($r in @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ PlanHandle = $PlanHandle } -Label 'statement stats of cached plan'))) {
            $r.QueryHash = ConvertTo-HexString $r.QueryHash
            $r
        }
    }

    function Get-QueryStoreProcedurePlans {
        <# Latest plan per statement of the procedure from Query Store, with runtime stats. #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId)
        $sql = @"
WITH lp AS (
    SELECT q.query_id, q.query_hash, p.plan_id, p.is_forced_plan, p.last_execution_time,
           ROW_NUMBER() OVER (PARTITION BY q.query_id ORDER BY p.last_execution_time DESC) AS rn
    FROM sys.query_store_query AS q JOIN sys.query_store_plan AS p ON p.query_id = q.query_id
    WHERE q.object_id = @ObjectId)
SELECT TOP (50) lp.query_id AS QueryId, lp.plan_id AS PlanId, lp.query_hash AS QueryHash, lp.is_forced_plan AS IsForced,
       lp.last_execution_time AS LastExecution,
       (SELECT COUNT(*) FROM sys.query_store_plan AS p2 WHERE p2.query_id = lp.query_id) AS PlanCount,
       rs.Executions, rs.AvgCpuMs, rs.AvgElapsedMs, rs.AvgLogicalReads, rs.AvgRows, rs.MinRows, rs.MaxRows, rs.LastRows,
       p.query_plan AS PlanText, qt.query_sql_text AS StatementText
FROM lp
JOIN sys.query_store_plan AS p ON p.plan_id = lp.plan_id
JOIN sys.query_store_query AS q ON q.query_id = lp.query_id
JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id
OUTER APPLY (SELECT SUM(r.count_executions) AS Executions,
                    CAST(SUM(r.avg_cpu_time * r.count_executions) / NULLIF(SUM(r.count_executions), 0) / 1000.0 AS decimal(19,2)) AS AvgCpuMs,
                    CAST(SUM(r.avg_duration * r.count_executions) / NULLIF(SUM(r.count_executions), 0) / 1000.0 AS decimal(19,2)) AS AvgElapsedMs,
                    CAST(SUM(r.avg_logical_io_reads * r.count_executions) / NULLIF(SUM(r.count_executions), 0) AS decimal(19,1)) AS AvgLogicalReads,
                    CAST(SUM(r.avg_rowcount * r.count_executions) / NULLIF(SUM(r.count_executions), 0) AS decimal(19,1)) AS AvgRows,
                    MIN(r.min_rowcount) AS MinRows, MAX(r.max_rowcount) AS MaxRows, MAX(r.last_rowcount) AS LastRows
             FROM sys.query_store_runtime_stats AS r WHERE r.plan_id = lp.plan_id) AS rs
WHERE lp.rn = 1
ORDER BY lp.query_id
OPTION (MAXDOP 1);
"@
        foreach ($r in @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Sql $sql -Parameters @{ ObjectId = $ObjectId } -Label 'query store plans' -TimeoutSec ([math]::Max(60, $Context.CommandTimeoutSec))))) {
            $r.QueryHash = ConvertTo-HexString $r.QueryHash
            $r
        }
    }

    function Remove-SqlBrackets {
        param([string]$Name)
        if ([string]::IsNullOrEmpty($Name)) { return $Name }
        if ($Name.StartsWith('[') -and $Name.EndsWith(']')) { return $Name.Substring(1, $Name.Length - 2).Replace(']]', ']') }
        $Name
    }

    function ConvertTo-InvariantDouble {
        param([string]$Text)
        $d = 0.0
        if ([double]::TryParse($Text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
        $null
    }

    function ConvertFrom-ShowPlanXml {
        <#
            Parses one or more showplan XML documents (cached plan = one document for the whole
            procedure; Query Store = one per statement). Returns statements, data-access operators,
            statistics used, warnings, missing indexes and compiled parameter values.
        #>
        param([Parameter(Mandatory)][string[]]$PlanText)
        $statements = New-Object System.Collections.Generic.List[object]
        $operators  = New-Object System.Collections.Generic.List[object]
        $statsUsed  = New-Object System.Collections.Generic.List[object]
        $warnings   = New-Object System.Collections.Generic.List[object]
        $missing    = New-Object System.Collections.Generic.List[object]
        $params     = New-Object System.Collections.Generic.List[object]
        $nsUri = 'http://schemas.microsoft.com/sqlserver/2004/07/showplan'
        $docIndex = -1
        foreach ($text in $PlanText) {
            $docIndex++
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $settings = New-Object System.Xml.XmlReaderSettings
            $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
            $settings.XmlResolver = $null
            $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($text)), $settings)
            $xml = New-Object System.Xml.XmlDocument
            $xml.XmlResolver = $null
            try { $xml.Load($reader) } finally { $reader.Close() }
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('p', $nsUri)

            foreach ($st in $xml.SelectNodes('//p:StmtSimple', $ns)) {
                $sid = "$docIndex." + $st.GetAttribute('StatementId')
                $qp = $st.SelectSingleNode('./p:QueryPlan', $ns)
                $stmt = [pscustomobject]@{
                    Key = $sid; DocIndex = $docIndex; StatementId = [int](ConvertTo-InvariantDouble $st.GetAttribute('StatementId'))
                    Type = $st.GetAttribute('StatementType'); Text = ($st.GetAttribute('StatementText')).Trim()
                    Cost = ConvertTo-InvariantDouble $st.GetAttribute('StatementSubTreeCost')
                    EstRows = ConvertTo-InvariantDouble $st.GetAttribute('StatementEstRows')
                    QueryHash = $st.GetAttribute('QueryHash'); OptmLevel = $st.GetAttribute('StatementOptmLevel')
                    CeVersion = $st.GetAttribute('CardinalityEstimationModelVersion')
                    HasPlan = $null -ne $qp; Tables = New-Object System.Collections.Generic.List[string]
                }
                $statements.Add($stmt)
                if (-not $qp) { continue }

                foreach ($si in $qp.SelectNodes('./p:OptimizerStatsUsage/p:StatisticsInfo', $ns)) {
                    $statsUsed.Add([pscustomobject]@{
                        StatementKey = $sid; Database = Remove-SqlBrackets $si.GetAttribute('Database'); Schema = Remove-SqlBrackets $si.GetAttribute('Schema')
                        Table = Remove-SqlBrackets $si.GetAttribute('Table'); Statistic = Remove-SqlBrackets $si.GetAttribute('Statistics')
                        ModificationCount = ConvertTo-InvariantDouble $si.GetAttribute('ModificationCount')
                        SamplingPercent = ConvertTo-InvariantDouble $si.GetAttribute('SamplingPercent')
                        LastUpdate = $(try { [datetime]::Parse($si.GetAttribute('LastUpdate'), [System.Globalization.CultureInfo]::InvariantCulture) } catch { $null })
                    })
                }
                foreach ($w in $qp.SelectNodes('./p:Warnings/*', $ns)) {
                    $detail = switch ($w.LocalName) {
                        'PlanAffectingConvert' { "$($w.GetAttribute('ConvertIssue')): $($w.GetAttribute('Expression'))" }
                        'ColumnsWithNoStatistics' { (@($w.SelectNodes('./p:ColumnReference', $ns) | ForEach-Object { '{0}.{1}.{2}' -f (Remove-SqlBrackets $_.GetAttribute('Schema')), (Remove-SqlBrackets $_.GetAttribute('Table')), (Remove-SqlBrackets $_.GetAttribute('Column')) }) -join ', ') }
                        'SpillToTempDb' { "spill level $($w.GetAttribute('SpillLevel'))" }
                        default { $w.OuterXml -replace '\s*xmlns="[^"]*"', '' }
                    }
                    $cols = @()
                    if ($w.LocalName -eq 'ColumnsWithNoStatistics') {
                        $cols = @($w.SelectNodes('./p:ColumnReference', $ns) | ForEach-Object { [pscustomobject]@{ Database = Remove-SqlBrackets $_.GetAttribute('Database'); Schema = Remove-SqlBrackets $_.GetAttribute('Schema'); Table = Remove-SqlBrackets $_.GetAttribute('Table'); Column = Remove-SqlBrackets $_.GetAttribute('Column') } })
                    }
                    $warnings.Add([pscustomobject]@{ StatementKey = $sid; Kind = $w.LocalName; Detail = $detail; Columns = $cols })
                }
                # attribute-style warnings on operators
                foreach ($w in $qp.SelectNodes('.//p:Warnings[@NoJoinPredicate="true" or @NoJoinPredicate="1"]', $ns)) { $warnings.Add([pscustomobject]@{ StatementKey = $sid; Kind = 'NoJoinPredicate'; Detail = 'A join has no join predicate (possible cartesian product).'; Columns = @() }) }
                foreach ($w in $qp.SelectNodes('.//p:RelOp/p:Warnings/p:SpillToTempDb', $ns)) { $warnings.Add([pscustomobject]@{ StatementKey = $sid; Kind = 'SpillToTempDb'; Detail = "An operator spills to tempdb (spill level $($w.GetAttribute('SpillLevel')))."; Columns = @() }) }

                foreach ($g in $qp.SelectNodes('./p:MissingIndexes/p:MissingIndexGroup', $ns)) {
                    foreach ($mi in $g.SelectNodes('./p:MissingIndex', $ns)) {
                        $grp = @{}
                        foreach ($cg in $mi.SelectNodes('./p:ColumnGroup', $ns)) {
                            $grp[$cg.GetAttribute('Usage')] = (@($cg.SelectNodes('./p:Column', $ns) | ForEach-Object { $_.GetAttribute('Name') }) -join ', ')
                        }
                        $missing.Add([pscustomobject]@{
                            StatementKey = $sid; Impact = ConvertTo-InvariantDouble $g.GetAttribute('Impact')
                            Database = Remove-SqlBrackets $mi.GetAttribute('Database'); Schema = Remove-SqlBrackets $mi.GetAttribute('Schema'); Table = Remove-SqlBrackets $mi.GetAttribute('Table')
                            EqualityColumns = $grp['EQUALITY']; InequalityColumns = $grp['INEQUALITY']; IncludedColumns = $grp['INCLUDE']
                        })
                    }
                }
                foreach ($pr in $qp.SelectNodes('./p:ParameterList/p:ColumnReference', $ns)) {
                    $params.Add([pscustomobject]@{ StatementKey = $sid; Name = $pr.GetAttribute('Column'); DataType = $pr.GetAttribute('ParameterDataType'); CompiledValue = $pr.GetAttribute('ParameterCompiledValue') })
                }
                foreach ($op in $qp.SelectNodes('.//p:RelOp', $ns)) {
                    $obj = $op.SelectSingleNode('./*/p:Object', $ns)
                    if (-not $obj) { continue }
                    $child = $obj.ParentNode
                    $phys = $op.GetAttribute('PhysicalOp')
                    $isLookup = $child.GetAttribute('Lookup') -in '1', 'true' -or $phys -eq 'RID Lookup' -or $phys -eq 'Key Lookup'
                    $access = if ($isLookup) { 'Lookup' } elseif ($phys -match 'Seek') { 'Seek' } elseif ($phys -match 'Scan') { 'Scan' } else { 'Other' }
                    $rebinds = [double](ConvertTo-InvariantDouble $op.GetAttribute('EstimateRebinds')); $rewinds = [double](ConvertTo-InvariantDouble $op.GetAttribute('EstimateRewinds'))
                    $tbl = Remove-SqlBrackets $obj.GetAttribute('Table')
                    if ($tbl -and -not $stmt.Tables.Contains($tbl)) { $stmt.Tables.Add($tbl) }
                    $operators.Add([pscustomobject]@{
                        StatementKey = $sid; NodeId = [int](ConvertTo-InvariantDouble $op.GetAttribute('NodeId')); PhysicalOp = $(if ($isLookup -and $phys -ne 'RID Lookup') { 'Key Lookup' } else { $phys })
                        Access = $access; Database = Remove-SqlBrackets $obj.GetAttribute('Database'); Schema = Remove-SqlBrackets $obj.GetAttribute('Schema')
                        Table = $tbl; Index = Remove-SqlBrackets $obj.GetAttribute('Index'); IndexKind = $obj.GetAttribute('IndexKind')
                        EstRows = ConvertTo-InvariantDouble $op.GetAttribute('EstimateRows')
                        EstRowsRead = ConvertTo-InvariantDouble $op.GetAttribute('EstimatedRowsRead')
                        TableCardinality = ConvertTo-InvariantDouble $op.GetAttribute('TableCardinality')
                        EstExecutions = 1 + $rebinds + $rewinds
                        SubtreeCost = ConvertTo-InvariantDouble $op.GetAttribute('EstimatedTotalSubtreeCost')
                        HasPredicate = $null -ne $child.SelectSingleNode('./p:Predicate', $ns)
                    })
                }
            }
        }
        [pscustomobject]@{
            Statements = $statements.ToArray(); Operators = $operators.ToArray(); StatsUsed = $statsUsed.ToArray()
            Warnings = $warnings.ToArray(); MissingIndexes = $missing.ToArray(); Parameters = $params.ToArray()
            HasStatsUsage = $statsUsed.Count -gt 0
        }
    }

    function Test-IsLocalPlanObject {
        <# True for permanent tables in the current database (not temp tables, table variables, or other databases). #>
        param([string]$Database, [string]$Table, [string]$CurrentDatabase)
        if ([string]::IsNullOrEmpty($Table)) { return $false }
        if ($Table.StartsWith('#') -or $Table.StartsWith('@')) { return $false }
        if ($Database -and $Database -ne $CurrentDatabase) { return $false }
        $true
    }

    function Invoke-ProcedureAnalysis {
        <# Runs on the worker: finds plan sources, reads the chosen plan, parses it, and loads index/stats metadata for the tables it touches. #>
        param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][int]$ObjectId, [Parameter(Mandatory)]$ServerInfo, [Parameter(Mandatory)]$DatabaseInfo,
              [string]$SourceKey, [int]$MaxTables = 30)
        $warnings = New-Object System.Collections.Generic.List[string]
        $src = Get-ProcedurePlanSources -Context $Context -ObjectId $ObjectId -ServerInfo $ServerInfo
        foreach ($w in $src.Warnings) { $warnings.Add($w) }
        $sources = @($src.Sources)
        $result = [ordered]@{ ObjectId = $ObjectId; Sources = $sources; Source = $null; PlanTexts = @(); Plan = $null; StatementStats = @(); Tables = @(); OtherObjects = @(); Warnings = $null }
        if ($sources.Count -eq 0) {
            $warnings.Add('No plan found: the procedure is not in the plan cache (not run since the last restart or cache flush, or it uses WITH RECOMPILE) and Query Store has nothing for it. Run it normally once, then analyze again. This tool will not execute it.')
            $result.Warnings = $warnings.ToArray(); return [pscustomobject]$result
        }
        $chosen = $null
        if ($SourceKey) { $chosen = $sources | Where-Object { $_.Key -eq $SourceKey } | Select-Object -First 1 }
        if (-not $chosen) { $chosen = $sources | Where-Object { $_.Source -eq 'Plan cache' } | Select-Object -First 1 }
        if (-not $chosen) { $chosen = $sources[0] }
        $result.Source = $chosen

        $stmtStats = @()
        if ($chosen.Source -eq 'Plan cache') {
            $text = Get-CachedPlanText -Context $Context -PlanHandle $chosen.PlanHandle
            if (-not $text) { $warnings.Add('The cached plan was evicted between listing and reading it. Try again.'); $result.Warnings = $warnings.ToArray(); return [pscustomobject]$result }
            $result.PlanTexts = @($text)
            try { $stmtStats = @(Get-CachedStatementStats -Context $Context -PlanHandle $chosen.PlanHandle) } catch { $warnings.Add("Statement runtime stats unavailable: $($_.Exception.Message)") }
        } else {
            $qsRows = @(Get-QueryStoreProcedurePlans -Context $Context -ObjectId $ObjectId)
            $result.PlanTexts = @($qsRows | ForEach-Object { [string]$_.PlanText })
            $stmtStats = @($qsRows | ForEach-Object { $_.PSObject.Properties.Remove('PlanText'); $_ })
        }
        $result.StatementStats = $stmtStats
        $plan = ConvertFrom-ShowPlanXml -PlanText $result.PlanTexts
        $result.Plan = $plan

        # tables this plan touches, in this database
        $db = [string]$DatabaseInfo.DatabaseName
        $names = New-Object System.Collections.Generic.List[string]
        $other = New-Object System.Collections.Generic.List[string]
        $refs = @($plan.Operators | ForEach-Object { [pscustomobject]@{ D = $_.Database; S = $_.Schema; T = $_.Table } }) +
                @($plan.StatsUsed | ForEach-Object { [pscustomobject]@{ D = $_.Database; S = $_.Schema; T = $_.Table } }) +
                @($plan.MissingIndexes | ForEach-Object { [pscustomobject]@{ D = $_.Database; S = $_.Schema; T = $_.Table } })
        foreach ($r in $refs) {
            if (Test-IsLocalPlanObject -Database $r.D -Table $r.T -CurrentDatabase $db) {
                $n = (ConvertTo-QuotedName $(if ($r.S) { $r.S } else { 'dbo' })) + '.' + (ConvertTo-QuotedName $r.T)
                if (-not $names.Contains($n)) { $names.Add($n) }
            } elseif ($r.T -and -not $r.T.StartsWith('#') -and -not $r.T.StartsWith('@')) {
                $o = "$($r.D).$($r.S).$($r.T)"; if (-not $other.Contains($o)) { $other.Add($o) }
            }
        }
        $result.OtherObjects = $other.ToArray()
        if ($names.Count -gt $MaxTables) { $warnings.Add("The plan touches $($names.Count) tables; only the first $MaxTables are analyzed."); $keep = @($names | Select-Object -First $MaxTables); $names.Clear(); foreach ($n in $keep) { $names.Add($n) } }
        $tables = New-Object System.Collections.Generic.List[object]
        if ($names.Count -gt 0) {
            $p = @{}; $vals = @()
            for ($i = 0; $i -lt $names.Count; $i++) { $p["n$i"] = $names[$i]; $vals += "(@n$i)" }
            $ids = @(ConvertTo-ObjectList (Invoke-SafeQuery -Context $Context -Label 'resolve plan tables' -Parameters $p -Sql ("SELECT v.Name, OBJECT_ID(v.Name) AS ObjectId, OBJECT_SCHEMA_NAME(OBJECT_ID(v.Name)) AS SchemaName, OBJECT_NAME(OBJECT_ID(v.Name)) AS TableName FROM (VALUES " + ($vals -join ', ') + ") AS v(Name);")))
            foreach ($t in $ids) {
                if (-not $t.ObjectId) { continue }
                $entry = [ordered]@{ ObjectId = [int]$t.ObjectId; SchemaName = $t.SchemaName; TableName = $t.TableName; Indexes = @(); Statistics = @(); RowCount = $null; ReservedMB = $null }
                try { $entry.Indexes = @(Get-TableIndexes -Context $Context -ObjectId $t.ObjectId) } catch { $warnings.Add("Indexes of $($t.Name): $($_.Exception.Message)") }
                try { $entry.Statistics = @(Get-TableStatistics -Context $Context -ObjectId $t.ObjectId -ServerInfo $ServerInfo) } catch { $warnings.Add("Statistics of $($t.Name): $($_.Exception.Message)") }
                $base = @($entry.Indexes | Where-Object { $_.IndexId -in 0, 1 })
                if ($base.Count) { $entry.RowCount = ($base | Measure-Object RowCount -Sum).Sum }
                $entry.ReservedMB = ($entry.Indexes | Measure-Object SizeMB -Sum).Sum
                $tables.Add([pscustomobject]$entry)
            }
        }
        $result.Tables = $tables.ToArray()
        $result.Warnings = $warnings.ToArray()
        [pscustomobject]$result
    }

    function Get-ProcedureFindings {
        <#
            Turns a procedure analysis into ranked findings. Pure function (no SQL).
            Frag: hashtable "<objectId>|<indexId>" -> Get-IndexFragmentation result (optional).
        #>
        param([Parameter(Mandatory)]$Analysis, [Parameter(Mandatory)]$Settings, [Parameter(Mandatory)]$DatabaseInfo, [hashtable]$Frag = @{}, [double]$UptimeDays = 30)
        $findings = New-Object System.Collections.Generic.List[object]
        $plan = $Analysis.Plan
        if (-not $plan) { return }
        $db = [string]$DatabaseInfo.DatabaseName
        $byTable = @{}
        foreach ($t in $Analysis.Tables) { $byTable[("$($t.SchemaName).$($t.TableName)").ToLowerInvariant()] = $t }
        function Get-T($s, $n) { $k = ("$(if ($s) { $s } else { 'dbo' }).$n").ToLowerInvariant(); if ($byTable.ContainsKey($k)) { $byTable[$k] } else { $null } }
        function Add-F($pri, $score, $kind, $table, $object, $finding, $action, $extra) {
            $o = [ordered]@{ Priority = $pri; Score = [int][math]::Max(0, [math]::Min(100, $score)); Kind = $kind; Table = $table; Object = $object; Finding = $finding; Action = $action }
            if ($extra) { foreach ($k in $extra.Keys) { $o[$k] = $extra[$k] } }
            $findings.Add([pscustomobject]$o)
        }
        function Get-Pri($score) { if ($score -ge 60) { 'High' } elseif ($score -ge 30) { 'Medium' } elseif ($score -gt 0) { 'Low' } else { 'None' } }

        # statement runtime vs estimate (match plan statements to runtime stats by query hash)
        $runtime = @{}
        foreach ($s in $Analysis.StatementStats) { if ($s.QueryHash -and -not $runtime.ContainsKey([string]$s.QueryHash)) { $runtime[[string]$s.QueryHash] = $s } }
        $suspectTables = @{}
        foreach ($st in $plan.Statements) {
            if (-not $st.QueryHash -or -not $runtime.ContainsKey($st.QueryHash)) { continue }
            $rt = $runtime[$st.QueryHash]
            $st | Add-Member -NotePropertyName Runtime -NotePropertyValue $rt -Force
            $act = if ($null -ne $rt.AvgRows) { [double]$rt.AvgRows } else { $null }
            $est = [double]$st.EstRows
            if ($null -ne $act -and $st.Type -match 'SELECT') {
                $hi = [math]::Max($est, $act); $lo = [math]::Max(1, [math]::Min($est, $act))
                $ratio = $hi / $lo
                $st | Add-Member -NotePropertyName MisestimateRatio -NotePropertyValue ([math]::Round($ratio, 1)) -Force
                if ($ratio -ge 10 -and $hi -ge 1000) {
                    $score = [math]::Min(80, 30 + 12 * [math]::Log10($ratio))
                    foreach ($t in $st.Tables) { $suspectTables[$t.ToLowerInvariant()] = $true }
                    $dir = if ($act -gt $est) { 'under' } else { 'over' }
                    Add-F (Get-Pri $score) $score 'Plan' (($st.Tables) -join ', ') "Statement $($st.StatementId)" ("Row estimate is {0}x off ({1}-estimated): the plan expected {2:N0} rows, executions return {3:N0} on average. Out-of-date or low-sample statistics on {4} are the first suspects." -f [math]::Round($ratio), $dir, $est, $act, $(if ($st.Tables.Count) { $st.Tables -join ', ' } else { 'the tables involved' })) 'Update the statistics this statement uses (see Statistic rows), then compare again' @{ StatementKey = $st.Key }
                }
            }
            if ($null -ne $rt.MinRows -and $null -ne $rt.MaxRows) {
                $mn = [math]::Max(1, [double]$rt.MinRows); $mx = [double]$rt.MaxRows
                if ($mx -ge 1000 -and $mx / $mn -ge 100) {
                    Add-F 'Medium' 35 'Plan' (($st.Tables) -join ', ') "Statement $($st.StatementId)" ("Parameter-sensitive: rows returned vary from {0:N0} to {1:N0} between executions, but every execution reuses one plan. Updating statistics or rebuilding will not fix this on its own." -f [double]$rt.MinRows, $mx) 'Consider OPTION (RECOMPILE) on this statement, OPTIMIZE FOR, or Parameter Sensitive Plan optimization (SQL 2022, compat 160)' @{ StatementKey = $st.Key }
                }
            }
        }

        # statistics the optimizer used (or, before SQL 2017 / 2016 SP2, all stats on the tables the plan reads)
        $seenStat = @{}
        $statCandidates = @()
        if ($plan.HasStatsUsage) {
            foreach ($u in $plan.StatsUsed) {
                if (-not (Test-IsLocalPlanObject -Database $u.Database -Table $u.Table -CurrentDatabase $db)) { continue }
                $statCandidates += [pscustomobject]@{ Schema = $u.Schema; Table = $u.Table; Statistic = $u.Statistic; Usage = $u; Direct = $true }
            }
        } else {
            foreach ($t in $Analysis.Tables) { foreach ($s in $t.Statistics) { $statCandidates += [pscustomobject]@{ Schema = $t.SchemaName; Table = $t.TableName; Statistic = $s.StatName; Usage = $null; Direct = $false } } }
        }
        foreach ($c in $statCandidates) {
            $key = ("$($c.Schema).$($c.Table).$($c.Statistic)").ToLowerInvariant()
            if ($seenStat.ContainsKey($key)) { continue }; $seenStat[$key] = $true
            $t = Get-T $c.Schema $c.Table
            if (-not $t) { continue }
            $stat = $t.Statistics | Where-Object { $_.StatName -eq $c.Statistic } | Select-Object -First 1
            if (-not $stat) { continue }
            $a = Get-StatsAssessment -Stat $stat -CompatLevel ([int]$DatabaseInfo.CompatLevel) -AutoUpdateOn ([bool]$DatabaseInfo.AutoUpdateStats) -TableRows $t.RowCount
            $score = $a.Score
            $why = New-Object System.Collections.Generic.List[string]
            if ($c.Direct) { $why.Add('The optimizer used this statistic to build this procedure''s plan.') } else { $why.Add('On a table this procedure reads (this SQL version does not record exactly which statistics a plan used).') }
            if ($c.Direct -and $score -gt 0) { $score += 10 }
            if ($suspectTables.ContainsKey(([string]$c.Table).ToLowerInvariant()) -and $score -gt 0) { $score += 20; $why.Add('A statement on this table has a large row-estimate error.') }
            if ($c.Usage -and $c.Usage.LastUpdate -and $stat.LastUpdated -and [datetime]$stat.LastUpdated -gt ([datetime]$c.Usage.LastUpdate).AddMinutes(1)) {
                $why.Add(('It was updated ({0:yyyy-MM-dd HH:mm}) after this plan was compiled ({1:yyyy-MM-dd HH:mm}); the plan will recompile on its next use.' -f [datetime]$stat.LastUpdated, [datetime]$c.Usage.LastUpdate))
            }
            foreach ($r in $a.Reasons) { $why.Add($r) }
            $pri = Get-Pri $score
            if (-not $c.Direct -and $pri -eq 'Low') { continue }   # keep the list focused when we are guessing
            $action = if ($pri -eq 'None') { 'No action' } else { "UPDATE STATISTICS ... ($($stat.StatName)) WITH $($a.SuggestedMethod)" }
            Add-F $pri $score 'Statistic' "$($t.SchemaName).$($t.TableName)" $stat.StatName (($why) -join ' ') $action @{ TableObjectId = $t.ObjectId; StatsId = $stat.StatsId; _Stat = $stat; _Assessment = $a }
        }

        # indexes the plan reads: rebuild/reorganize only matters where it scans
        $groups = @{}
        foreach ($op in $plan.Operators) {
            if (-not (Test-IsLocalPlanObject -Database $op.Database -Table $op.Table -CurrentDatabase $db)) { continue }
            $k = ("$($op.Schema).$($op.Table)|$($op.Index)").ToLowerInvariant()
            if (-not $groups.ContainsKey($k)) { $groups[$k] = New-Object System.Collections.Generic.List[object] }
            $groups[$k].Add($op)
        }
        $seekOnly = 0
        foreach ($k in $groups.Keys) {
            $ops = $groups[$k]; $first = $ops[0]
            $t = Get-T $first.Schema $first.Table
            if (-not $t) { continue }
            $ix = if ($first.Index) { $t.Indexes | Where-Object { $_.IndexName -eq $first.Index } | Select-Object -First 1 } else { $t.Indexes | Where-Object { $_.IndexId -eq 0 } | Select-Object -First 1 }
            if (-not $ix) { continue }
            $tableName = "$($t.SchemaName).$($t.TableName)"
            $scans = @($ops | Where-Object { $_.Access -eq 'Scan' })
            $lookups = @($ops | Where-Object { $_.Access -eq 'Lookup' })
            $extra = @{ TableObjectId = $t.ObjectId; IndexId = $ix.IndexId; _Index = $ix }
            if ($scans.Count -gt 0) {
                $fr = $Frag["$($t.ObjectId)|$($ix.IndexId)"]
                $ia = Get-IndexAssessment -Index $ix -Fragmentation $fr -Settings $Settings -UptimeDays $UptimeDays
                $read = ($scans | Measure-Object EstRowsRead -Maximum).Maximum; if (-not $read) { $read = ($scans | Measure-Object TableCardinality -Maximum).Maximum }
                $ret = ($scans | Measure-Object EstRows -Maximum).Maximum
                $scanTxt = ("Scanned by this procedure ({0} operator(s), reads about {1:N0} rows to return {2:N0})." -f $scans.Count, [double]$read, [double]$ret)
                if ([double]$ix.InRowPages -lt [double]$Settings.MinPageCount -and $ix.TypeCode -ne 0) {
                    # small: nothing to maintain
                } elseif ($fr -and $fr.Measured) {
                    $extra['_Assessment'] = $ia; $extra['_Frag'] = $fr
                    if ($ia.Action -in 'REBUILD', 'REORGANIZE') {
                        Add-F (Get-Pri ($ia.Score + 10)) ($ia.Score + 10) 'Index' $tableName $ix.IndexName ("$scanTxt " + ($ia.Reasons -join ' ')) $ia.Action $extra
                    } else {
                        Add-F 'None' 0 'Index' $tableName $ix.IndexName ("$scanTxt Physically healthy (fragmentation $($fr.FragmentationPct)%$(if ($null -ne $fr.PageDensityPct) { ", pages $($fr.PageDensityPct)% full" })). Rebuilding would not help this procedure.") 'No action' $extra
                    }
                } else {
                    $sc = if ([double]$ix.SizeMB -gt 1024) { 30 } else { 20 }
                    Add-F (Get-Pri $sc) $sc 'Index' $tableName $ix.IndexName ("$scanTxt For scans, fragmentation and page fullness matter. Not measured yet." + $(if ($fr -and $fr.Skipped) { " Last attempt: $($fr.Skipped)" } else { '' })) 'Check fragmentation (button below), then decide' $extra
                }
                if ($read -and $ret -ge 0 -and [double]$read -ge 100000 -and [double]$read / [math]::Max(1, [double]$ret) -ge 100) {
                    Add-F 'Review' 45 'Index' $tableName $ix.IndexName ("Reads about {0:N0} rows to return about {1:N0}: the filter is not supported by an index, so a rebuild will not make this fast. See the missing-index rows or the statement's predicate." -f [double]$read, [double]$ret) 'Review indexing for this predicate (not maintenance)' $extra
                }
            } elseif ($lookups.Count -gt 0) {
                $execs = ($lookups | ForEach-Object { [double]$_.EstExecutions * [math]::Max(1, [double]$_.EstRows) } | Measure-Object -Sum).Sum
                if ($execs -ge 1000) {
                    Add-F 'Review' 40 'Index' $tableName $ix.IndexName ("Key/RID lookups: about {0:N0} lookups per execution back into {1}. Each lookup is a separate seek; adding the needed columns as INCLUDEs to the seeking index removes them." -f $execs, $ix.IndexName) 'Consider a covering index (INCLUDE columns); maintenance will not help' $extra
                }
            } else { $seekOnly++ }
        }
        if ($seekOnly -gt 0) {
            Add-F 'None' 0 'Index' '' "$seekOnly index(es)" "$seekOnly index(es) are only used for seeks by this procedure. Fragmentation barely affects seeks, so rebuilding them would not speed this procedure up." 'No action' $null
        }

        # warnings from the plan
        foreach ($w in $plan.Warnings) {
            switch ($w.Kind) {
                'ColumnsWithNoStatistics' {
                    foreach ($c in $w.Columns) {
                        if (-not (Test-IsLocalPlanObject -Database $c.Database -Table $c.Table -CurrentDatabase $db)) { continue }
                        Add-F 'High' 70 'Statistic' "$($c.Schema).$($c.Table)" $c.Column "The optimizer had NO statistics for column $($c.Column) and had to guess (AUTO_CREATE_STATISTICS is $(if ($DatabaseInfo.AutoCreateStats) { 'ON, so this may be a read-only database, a secondary, or a column type that cannot have stats' } else { 'OFF' }))." 'Create statistics on this column (script is added as text)' @{ CreateStatsColumn = $c }
                    }
                }
                'PlanAffectingConvert' {
                    $seek = $w.Detail -match '^Seek Plan'
                    Add-F $(if ($seek) { 'Medium' } else { 'Low' }) $(if ($seek) { 50 } else { 20 }) 'Plan' '' 'Implicit conversion' ("$($w.Detail). $(if ($seek) { 'It can stop SQL Server from seeking an index. Neither a stats update nor a rebuild fixes it: make the parameter/variable type match the column type.' } else { 'It can distort row estimates.' })") 'Fix the data type mismatch in the procedure' @{ StatementKey = $w.StatementKey }
                }
                'SpillToTempDb' { Add-F 'Medium' 40 'Plan' '' 'tempdb spill' "$($w.Detail) Often caused by row under-estimates (stale or low-sample statistics) giving a too-small memory grant." 'Fix row estimates first (statistics), then re-check' @{ StatementKey = $w.StatementKey } }
                'NoJoinPredicate' { Add-F 'High' 60 'Plan' '' 'Missing join predicate' $w.Detail 'Review the query logic' @{ StatementKey = $w.StatementKey } }
                'UnmatchedIndexes' { Add-F 'Low' 25 'Plan' '' 'Unmatched filtered index' 'A filtered index could not be used because of parameterization.' 'Review the filtered index / query' @{ StatementKey = $w.StatementKey } }
                default { Add-F 'Low' 15 'Plan' '' $w.Kind $w.Detail 'Review' @{ StatementKey = $w.StatementKey } }
            }
        }

        foreach ($m in $plan.MissingIndexes) {
            if (-not (Test-IsLocalPlanObject -Database $m.Database -Table $m.Table -CurrentDatabase $db)) { continue }
            $sc = [math]::Min(70, 20 + 0.5 * [double]$m.Impact)
            $mi = [pscustomobject]@{ EqualityColumns = $m.EqualityColumns; InequalityColumns = $m.InequalityColumns; IncludedColumns = $m.IncludedColumns; ImprovementScore = $m.Impact; Seeks = 'n/a'; AvgImpactPct = $m.Impact }
            Add-F 'Review' $sc 'Index' "$($m.Schema).$($m.Table)" 'Missing index' ("The optimizer estimates an index would cut this statement's cost by {0:N0}%: keys ({1}){2}." -f [double]$m.Impact, (@($m.EqualityColumns, $m.InequalityColumns | Where-Object { $_ }) -join ', '), $(if ($m.IncludedColumns) { " INCLUDE ($($m.IncludedColumns))" } else { '' })) 'Review before creating (check existing indexes for overlap)' @{ StatementKey = $m.StatementKey; _Missing = $mi; MissingSchema = $m.Schema; MissingTable = $m.Table }
        }

        $order = @{ 'High' = 0; 'Medium' = 1; 'Review' = 2; 'Low' = 3; 'None' = 4 }
        # emit one finding at a time; callers collect with @()
        $findings | Sort-Object @{ Expression = { $order[$_.Priority] } }, @{ Expression = 'Score'; Descending = $true }
    }

    function Get-ProcedureSummaryText {
        param([Parameter(Mandatory)]$Analysis, [string]$ProcName, $Findings)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("PROCEDURE  $ProcName")
        if ($Analysis.Source) { [void]$sb.AppendLine("  Plan source: $($Analysis.Source.Label)") }
        [void]$sb.AppendLine('  Read from the plan the procedure already has. The procedure was NOT executed or compiled. Plan-cache numbers are estimates; runtime row counts come from sys.dm_exec_query_stats / Query Store.')
        if ($Analysis.Plan) {
            $p = $Analysis.Plan
            [void]$sb.AppendLine(("  {0} statement(s), {1} data-access operator(s), {2} statistic(s) recorded as used{3}." -f @($p.Statements).Count, @($p.Operators).Count, @($p.StatsUsed).Count, $(if (-not $p.HasStatsUsage) { ' (this SQL version does not record them: SQL 2017 / 2016 SP2+ does)' } else { '' })))
            [void]$sb.AppendLine("  Tables analyzed: " + $(if (@($Analysis.Tables).Count) { (@($Analysis.Tables) | ForEach-Object { "$($_.SchemaName).$($_.TableName)" }) -join ', ' } else { 'none' }))
            if (@($Analysis.OtherObjects).Count) { [void]$sb.AppendLine("  Other databases (not analyzed): " + (@($Analysis.OtherObjects) -join ', ')) }
            $pv = @($p.Parameters | Group-Object Name | ForEach-Object { $_.Group[0] })
            if ($pv.Count) { [void]$sb.AppendLine("  Compiled for parameter values: " + (($pv | ForEach-Object { "$($_.Name) = $($_.CompiledValue)" }) -join ', ') + '  (the plan was optimized for these; other values reuse it)') }
        }
        if ($Findings) {
            $hi = @($Findings | Where-Object { $_.Priority -in 'High', 'Medium' })
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("VERDICT  $($hi.Count) finding(s) need attention:  statistics $(@($hi | Where-Object { $_.Kind -eq 'Statistic' }).Count), index maintenance $(@($hi | Where-Object { $_.Kind -eq 'Index' -and $_.Action -in 'REBUILD', 'REORGANIZE' }).Count), plan/query issues $(@($hi | Where-Object { $_.Kind -eq 'Plan' }).Count). Review items: $(@($Findings | Where-Object { $_.Priority -eq 'Review' }).Count).")
        }
        foreach ($w in $Analysis.Warnings) { [void]$sb.AppendLine("  NOTE: $w") }
        $sb.ToString()
    }

    function New-CreateStatisticsScript {
        param([Parameter(Mandatory)][string]$SchemaName, [Parameter(Mandatory)][string]$TableName, [Parameter(Mandatory)][string]$ColumnName)
        $nm = "st_$($TableName)_$($ColumnName)" -replace '[^\w]', '_'
        "-- The plan reported no statistics on this column (ColumnsWithNoStatistics). Text only, review before running:`r`n" +
        "-- CREATE STATISTICS $(ConvertTo-QuotedName $nm) ON $(ConvertTo-QuotedName $SchemaName).$(ConvertTo-QuotedName $TableName) ($(ConvertTo-QuotedName $ColumnName)) WITH SAMPLE 10 PERCENT;`r`n"
    }
    #endregion


    Export-ModuleMember -Function *

}
#endregion ================================================================ end of CORE

$script:CoreModule = New-Module -Name IndexStatsAnalyzer -ScriptBlock $script:CoreBlock
Import-Module $script:CoreModule -Global -Force -DisableNameChecking
if ($NoGui) { return }

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $argList = @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    if ($Server) { $argList += @('-Server', "`"$Server`"") }
    if ($Database) { $argList += @('-Database', "`"$Database`"") }
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argList
    return
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$script:ChartAvailable = $true
try { Add-Type -AssemblyName System.Windows.Forms.DataVisualization } catch { $script:ChartAvailable = $false }
[System.Windows.Forms.Application]::EnableVisualStyles()


#region ------------------------------------------------------------------ state
$script:Settings = Get-AnalyzerSettings
$script:Shared = [hashtable]::Synchronized(@{
    Log     = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    Command = $null
    Cancel  = $false
})
$script:State = @{
    Context = $null; ServerInfo = $null; Environment = $null; Databases = @(); DbInfo = $null
    TablesMaster = $null; TableSchemas = @(); TableNames = @(); CurrentObjectId = $null; CurrentTable = $null
    Analysis = $null; IndexRows = @(); StatsRows = @(); Frag = @{}; StatDetail = @{}
    Busy = $false; Job = $null; PendingObjectId = $null
    ProcMaster = $null; ProcBase = $null; ProcSchemas = @(); ProcNames = @(); ProcRefIds = $null; ProcRefFor = $null
    ProcAnalysis = $null; ProcFindings = @(); ProcStatements = @(); ProcName = $null; ProcObjectId = $null; ProcPlanSuppress = $false
}
$script:Mono = New-Object System.Drawing.Font('Consolas', 9)
$script:UiFont = New-Object System.Drawing.Font('Segoe UI', 9)
$script:BoldFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$script:PriorityColors = @{
    'High'   = [System.Drawing.Color]::FromArgb(255, 199, 206)
    'Medium' = [System.Drawing.Color]::FromArgb(255, 235, 156)
    'Low'    = [System.Drawing.Color]::FromArgb(221, 235, 247)
    'Review' = [System.Drawing.Color]::FromArgb(230, 224, 248)
}

# Worker runspace: all SQL runs here so the window never freezes.
$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$script:Worker = [runspacefactory]::CreateRunspace($iss)
$script:Worker.Open()
$init = [powershell]::Create()
$init.Runspace = $script:Worker
[void]$init.AddScript('param($Code) $m = New-Module -Name IndexStatsAnalyzer -ScriptBlock ([scriptblock]::Create($Code)); Import-Module $m -Global -Force -DisableNameChecking')
[void]$init.AddParameter('Code', $script:CoreBlock.ToString())
[void]$init.Invoke()
$init.Dispose()
#endregion

#region ------------------------------------------------------------------ small UI helpers
function New-Ctl {
    param([string]$Type, [hashtable]$Props)
    $c = New-Object ("System.Windows.Forms.$Type")
    if ($Props) {
        # range properties must be set before Value, or WinForms throws
        $first = @('Minimum', 'Maximum', 'DecimalPlaces', 'Increment')
        foreach ($k in $first) { if ($Props.ContainsKey($k)) { $c.$k = $Props[$k] } }
        foreach ($k in $Props.Keys) { if ($first -notcontains $k -and $k -ne 'Value') { $c.$k = $Props[$k] } }
        if ($Props.ContainsKey('Value')) { $c.Value = $Props['Value'] }
    }
    $c
}

function New-Grid {
    $g = New-Ctl DataGridView @{
        Dock = 'Fill'; ReadOnly = $true; AllowUserToAddRows = $false; AllowUserToDeleteRows = $false
        SelectionMode = 'FullRowSelect'; MultiSelect = $true; AutoSizeColumnsMode = 'None'; RowHeadersVisible = $false
        BackgroundColor = [System.Drawing.SystemColors]::Window; BorderStyle = 'None'
        AllowUserToResizeRows = $false; ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'
    }
    $g.DefaultCellStyle.Font = $script:UiFont
    $g.ColumnHeadersDefaultCellStyle.Font = $script:BoldFont
    $g.RowTemplate.Height = 20
    # double buffering for smooth scrolling of large grids
    $prop = $g.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
    $prop.SetValue($g, $true, $null)
    $g
}

function Get-ColumnType {
    param($Value)
    if ($Value -is [datetime]) { return [datetime] }
    if ($Value -is [bool]) { return [bool] }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [single]) { return [double] }
    [string]
}

function New-GridTable {
    param([object[]]$Items, [string[]]$Columns)
    $dt = New-Object System.Data.DataTable
    $types = @{}
    foreach ($c in $Columns) {
        $t = [string]
        foreach ($it in $Items) { $v = $it.$c; if ($null -ne $v -and -not ($v -is [DBNull])) { $t = Get-ColumnType $v; break } }
        $types[$c] = $t
        [void]$dt.Columns.Add($c, $t)
    }
    foreach ($it in $Items) {
        $row = $dt.NewRow()
        foreach ($c in $Columns) {
            $v = $it.$c
            if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
            if ($null -eq $v -or $v -is [DBNull]) { $row[$c] = [DBNull]::Value; continue }
            switch ($types[$c].Name) {
                'Double'   { $row[$c] = [double]$v }
                'DateTime' { $row[$c] = [datetime]$v }
                'Boolean'  { $row[$c] = [bool]$v }
                default    { if ($v -is [array]) { $row[$c] = ($v -join ' | ') } else { $row[$c] = [string]$v } }
            }
        }
        $dt.Rows.Add($row)
    }
    ,$dt
}

function Set-GridFormats {
    param($Grid)
    foreach ($col in $Grid.Columns) {
        $col.SortMode = 'Automatic'
        $n = $col.Name
        if ($col.ValueType -eq [double]) {
            $col.DefaultCellStyle.Alignment = 'MiddleRight'
            $col.DefaultCellStyle.Format = if ($n -match 'Pct|MB|AgeDays|Density') { 'N1' } else { 'N0' }
        }
        if ($col.ValueType -eq [datetime]) { $col.DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm' }
        $col.Width = switch -Regex ($n) {
            '^(Priority|Score|Partitions|Steps)$' { 60; break }
            '^(Reasons|Notes|Why)$' { 600; break }
            'Columns$' { 220; break }
            'Name$' { 220; break }
            '^Action$' { 170; break }
            default { 95 }
        }
    }
}

function Add-PriorityColoring {
    param($Grid)
    $Grid.add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        if (-not $s.Columns.Contains('Priority')) { return }
        $p = [string]$s.Rows[$e.RowIndex].Cells['Priority'].Value
        if ($script:PriorityColors.ContainsKey($p)) { $e.CellStyle.BackColor = $script:PriorityColors[$p] }
    })
}

function Set-Status { param([string]$Text) $script:lblStatus.Text = $Text }

function Show-Error {
    param([string]$Title, $ErrorRecord)
    $msg = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $ErrorRecord.Exception.Message } else { [string]$ErrorRecord }
    $inner = $null
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord] -and $ErrorRecord.Exception.InnerException) { $inner = $ErrorRecord.Exception.InnerException.Message }
    if ($inner -and $msg -notlike "*$inner*") { $msg = "$msg`r`n`r`n$inner" }
    $msg = $msg -replace '^Exception calling "EndInvoke" with "1" argument\(s\): ', ''
    Set-Status "Error: $Title"
    [void][System.Windows.Forms.MessageBox]::Show($msg, $Title, 'OK', 'Error')
}

function Get-UptimeDays {
    if ($script:State.ServerInfo -and $script:State.ServerInfo.StartTime) { return [math]::Round(((Get-Date) - [datetime]$script:State.ServerInfo.StartTime).TotalDays, 1) }
    30
}
#endregion

#region ------------------------------------------------------------------ background jobs
function Start-AnalyzerJob {
    param([string]$Status, [scriptblock]$Work, [hashtable]$Arguments, [scriptblock]$OnDone, [switch]$Quiet)
    if ($script:State.Busy) { Set-Status 'Still working on the previous request... (Cancel to stop it)'; return $false }
    $script:State.Busy = $true
    $script:Shared.Cancel = $false
    $ps = [powershell]::Create()
    $ps.Runspace = $script:Worker
    [void]$ps.AddScript($Work.ToString())
    if ($Arguments) { foreach ($k in $Arguments.Keys) { [void]$ps.AddParameter($k, $Arguments[$k]) } }
    $script:State.Job = @{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $OnDone; Started = Get-Date; Status = $Status }
    Set-Busy $true
    Set-Status "$Status ..."
    $true
}

function Complete-AnalyzerJob {
    $job = $script:State.Job
    $result = $null; $err = $null
    try {
        $out = $job.PS.EndInvoke($job.Handle)
        if ($out.Count -gt 0) { $result = $out[$out.Count - 1]; if ($result -is [psobject] -and $result.BaseObject -is [hashtable]) { $result = $result.BaseObject } }
        if ($job.PS.Streams.Error.Count -gt 0 -and $null -eq $result) { $err = $job.PS.Streams.Error[0] }
    } catch { $err = $_ }
    $job.PS.Dispose()
    $script:State.Job = $null
    $script:State.Busy = $false
    Set-Busy $false
    Sync-QueryLog
    $elapsed = ((Get-Date) - $job.Started).TotalSeconds
    if ($err) {
        if ($script:Shared.Cancel) { Set-Status 'Cancelled.'; return }
        Show-Error $job.Status $err
        return
    }
    Set-Status ("{0}: done in {1:N1}s" -f $job.Status, $elapsed)
    if ($job.OnDone) {
        try { & $job.OnDone $result } catch { Show-Error "$($job.Status) (display)" $_ }
    }
}

function Stop-AnalyzerJob {
    $script:Shared.Cancel = $true
    $cmd = $script:Shared.Command
    if ($cmd) { try { $cmd.Cancel() } catch { } }
    Set-Status 'Cancelling...'
}

function Set-Busy {
    param([bool]$Busy)
    $script:progress.Style = if ($Busy) { 'Marquee' } else { 'Blocks' }
    $script:progress.Visible = $Busy
    $script:btnCancel.Enabled = $Busy
    foreach ($c in @($script:btnConnect, $script:cboDb, $script:btnFragLimited, $script:btnFragSampled, $script:btnRefreshTable, $script:btnProcLoad, $script:btnProcAnalyze, $script:btnProcFrag, $script:cboProcPlan)) { if ($c) { $c.Enabled = -not $Busy } }
    $script:form.Cursor = if ($Busy) { [System.Windows.Forms.Cursors]::AppStarting } else { [System.Windows.Forms.Cursors]::Default }
}

function Sync-QueryLog {
    $log = $script:Shared.Log
    if ($log.Count -eq 0) { return }
    $items = @($log.ToArray()); $log.Clear()
    foreach ($i in $items) {
        $r = $script:LogTable.NewRow()
        # values arrive from the worker runspace wrapped in PSObject: unwrap/cast before storing in typed DataColumns
        $r['Time'] = [datetime]$i.Time
        $r['Server'] = [string]$i.Server; $r['Database'] = [string]$i.Database; $r['What'] = [string]$i.Label
        $r['Ms'] = [double]$i.DurationMs; $r['Rows'] = [double]$i.Rows; $r['Status'] = [string]$i.Status; $r['Sql'] = [string]$i.Sql
        $script:LogTable.Rows.Add($r)
    }
    $script:tabLog.Text = "Query log ($($script:LogTable.Rows.Count))"
}
#endregion

#region ------------------------------------------------------------------ connect / database / tables
function Get-UiContext {
    param([string]$DatabaseName = 'master', [string]$ServerName)
    $srv = if ($ServerName) { $ServerName.Trim() } else { $script:cboServer.Text.Trim() }
    if (-not $srv) { throw 'Type or pick a server name first.' }
    $auth = if ($script:cboAuth.SelectedItem -eq 'SQL login') { 'Sql' } else { 'Windows' }
    $pw = $null
    if ($auth -eq 'Sql') {
        if (-not $script:txtUser.Text) { throw 'Enter the SQL login name.' }
        $pw = New-Object System.Security.SecureString
        foreach ($ch in $script:txtPwd.Text.ToCharArray()) { $pw.AppendChar($ch) }
    }
    New-AnalyzerContext -Server $srv -Database $DatabaseName -Auth $auth -UserName $script:txtUser.Text -Password $pw `
        -Encrypt ([bool]$script:chkEncrypt.Checked) -TrustServerCertificate ([bool]$script:Settings.TrustServerCertificate) `
        -ReadOnlyIntent ([bool]$script:chkReadOnlyIntent.Checked) -CommandTimeoutSec ([int]$script:Settings.CommandTimeoutSec) `
        -LockTimeoutMs ([int]$script:Settings.LockTimeoutMs) -Shared $script:Shared
}

function Update-ServerCombo {
    $current = $script:cboServer.Text
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($s in $script:Settings.RecentServers) { if ($s -and -not $names.Contains([string]$s)) { $names.Add([string]$s) } }
    foreach ($s in $script:Settings.Servers) { if ($s -and $s.Name -and -not $names.Contains([string]$s.Name)) { $names.Add([string]$s.Name) } }
    $script:cboServer.BeginUpdate()
    $script:cboServer.Items.Clear()
    foreach ($n in $names) { [void]$script:cboServer.Items.Add($n) }
    $script:cboServer.EndUpdate()
    $script:cboServer.Text = $current
}

function Update-EnvironmentBanner {
    $srv = $script:cboServer.Text.Trim()
    $envName = if ($script:State.Environment) { $script:State.Environment } else { Get-ServerEnvironment -ServerName $srv -Settings $script:Settings }
    $isProd = Test-IsProductionEnvironment $envName
    $script:lblEnv.Text = if ($isProd) { "  $($envName.ToUpper())   |   READ-ONLY: analysis queries only, maintenance is generated as text and never executed  " }
                          else { "  $envName   |   READ-ONLY: analysis queries only, maintenance is generated as text and never executed  " }
    $script:lblEnv.BackColor = if ($isProd) { [System.Drawing.Color]::FromArgb(192, 0, 0) } else { [System.Drawing.Color]::FromArgb(0, 120, 60) }
}

function Connect-Server {
    try { $ctx = Get-UiContext -DatabaseName 'master' } catch { Show-Error 'Connect' $_; return }
    $script:State.Context = $ctx
    $script:State.Environment = Get-ServerEnvironment -ServerName $ctx.Server -Settings $script:Settings
    Update-EnvironmentBanner
    Clear-TableViews
    Clear-ProcViews
    $script:cboDb.Items.Clear()
    $script:gridTables.DataSource = $null
    [void](Start-AnalyzerJob -Status "Connecting to $($ctx.Server)" -Arguments @{ Ctx = $ctx } -Work {
        param($Ctx)
        $ErrorActionPreference = 'Stop'
        $info = Get-ServerInfo -Context $Ctx
        $dbs = @(Get-DatabaseList -Context $Ctx -ServerInfo $info)
        @{ Info = $info; Databases = $dbs }
    } -OnDone {
        param($r)
        $script:State.ServerInfo = $r.Info
        $script:State.Databases = $r.Databases
        $i = $r.Info
        $up = if ($i.StartTime) { 'up {0:N0} days (usage stats since {1:yyyy-MM-dd})' -f ((Get-Date) - [datetime]$i.StartTime).TotalDays, [datetime]$i.StartTime } else { 'uptime unknown (needs VIEW SERVER STATE)' }
        $script:lblServerInfo.Text = "$($i.ServerName)  |  $($i.ProductVersion) $($i.ProductLevel) $($i.ProductUpdateLevel)  |  $($i.Edition)  |  $up  |  login: $($i.LoginName)"
        $warn = @()
        if ($i.IsSysadmin -eq 1 -or $i.HasControlServer -eq 1) { $warn += 'Connected with a sysadmin/CONTROL SERVER login. This tool only reads, but a least-privilege login is safer (see README).' }
        if ($i.HasViewServerState -ne 1) { $warn += 'No VIEW SERVER STATE: usage, uptime and server-load checks are limited, and fragmentation scans are refused on production.' }
        $script:lblWarn.Text = ($warn -join '   ')
        $script:cboDb.BeginUpdate(); $script:cboDb.Items.Clear()
        foreach ($d in $r.Databases) {
            $ok = $d.State -eq 'ONLINE' -and $d.HasAccess -eq 1
            $tag = @()
            if ($d.SizeGB) { $tag += ('{0:N1} GB' -f [double]$d.SizeGB) }
            if ($d.AgRole) { $tag += "AG $($d.AgName): $($d.AgRole)" }
            if (-not $ok) { $tag += "$($d.State)$(if ($d.HasAccess -ne 1) { ', no access' })" }
            [void]$script:cboDb.Items.Add(("{0}    ({1})" -f $d.DatabaseName, ($tag -join ', ')))
        }
        $script:cboDb.EndUpdate()
        $srv = $script:State.Context.Server
        $recent = @($srv) + @($script:Settings.RecentServers | Where-Object { $_ -and $_ -ne $srv })
        $script:Settings.RecentServers = @($recent | Select-Object -First 30)
        Save-AnalyzerSettings -Settings $script:Settings
        Update-ServerCombo
        $script:cboServer.Text = $srv
        if ($script:InitialDatabase) {
            for ($k = 0; $k -lt $script:State.Databases.Count; $k++) { if ($script:State.Databases[$k].DatabaseName -eq $script:InitialDatabase) { $script:cboDb.SelectedIndex = $k } }
            $script:InitialDatabase = $null
        } else { $script:cboDb.DroppedDown = $true }
    })
}

function Select-Database {
    $idx = $script:cboDb.SelectedIndex
    if ($idx -lt 0 -or -not $script:State.Context) { return }
    $d = $script:State.Databases[$idx]
    if ($d.State -ne 'ONLINE' -or $d.HasAccess -ne 1) { Set-Status "$($d.DatabaseName) is not accessible ($($d.State))."; return }
    $ctx = Copy-AnalyzerContext -Context $script:State.Context -Database $d.DatabaseName
    Clear-TableViews
    Clear-ProcViews
    [void](Start-AnalyzerJob -Status "Loading tables of $($d.DatabaseName)" -Arguments @{ Ctx = $ctx; Info = $script:State.ServerInfo } -Work {
        param($Ctx, $Info)
        $ErrorActionPreference = 'Stop'
        $db = Get-DatabaseInfo -Context $Ctx -ServerInfo $Info
        $tables = Get-TableList -Context $Ctx
        @{ DbInfo = $db; Tables = $tables; Ctx = $Ctx }
    } -OnDone {
        param($r)
        $script:State.DbContext = $r.Ctx
        $script:State.DbInfo = $r.DbInfo
        $master = $r.Tables
        $script:State.TablesMaster = $master
        $n = $master.Rows.Count
        $sch = New-Object string[] $n; $nm = New-Object string[] $n
        for ($i = 0; $i -lt $n; $i++) { $sch[$i] = [string]$master.Rows[$i]['SchemaName']; $nm[$i] = [string]$master.Rows[$i]['TableName'] }
        $script:State.TableSchemas = $sch; $script:State.TableNames = $nm
        $db = $r.DbInfo
        $script:lblDbInfo.Text = ("{0}: compat {1}, {2} recovery, auto-update stats {3}{4}{5}, {6:N0} tables" -f $db.DatabaseName, $db.CompatLevel, $db.RecoveryModel,
            $(if ($db.AutoUpdateStats) { 'ON' } else { 'OFF' }), $(if ($db.AutoUpdateStatsAsync) { ' (async)' } else { '' }),
            $(if ([int]$db.IsInAg -eq 1) { $(if ($db.IsAgPrimary -eq 1) { ', AG primary' } else { ', AG SECONDARY: usage stats are local to this replica' }) } else { '' }), $n)
        Update-TableFilter
        $script:txtSearch.Focus()
    })
}

function Update-TableFilter {
    $master = $script:State.TablesMaster
    if (-not $master) { return }
    if ($master.Rows.Count -eq 0) { $script:lblTableCount.Text = 'This database has no user tables.'; $script:gridTables.DataSource = $null; return }
    $q = $script:txtSearch.Text
    $max = [int]$script:Settings.MaxTablesDisplayed
    $view = $master.Clone()
    $total = $master.Rows.Count
    if ([string]::IsNullOrWhiteSpace($q)) {
        $dv = New-Object System.Data.DataView($master)
        $dv.Sort = 'ReservedMB DESC'
        $take = [math]::Min($max, $dv.Count)
        for ($i = 0; $i -lt $take; $i++) { [void]$view.ImportRow($dv[$i].Row) }
        $script:lblTableCount.Text = ("{0:N0} tables. Showing the {1:N0} largest. Type to search (typos, abbreviations, * wildcards and schema.table all work)." -f $total, $take)
    } else {
        $hits = Find-Table -SchemaNames $script:State.TableSchemas -TableNames $script:State.TableNames -Query $q -Top $max
        foreach ($i in $hits) { [void]$view.ImportRow($master.Rows[[int]$i]) }
        $script:lblTableCount.Text = ("{0:N0} matches for '{1}' (best first){2}" -f @($hits).Count, $q, $(if (@($hits).Count -ge $max) { ", capped at ${max}: refine the search" } else { '' }))
    }
    $script:gridTables.DataSource = $view
}

function Clear-TableViews {
    $script:State.Analysis = $null; $script:State.CurrentObjectId = $null; $script:State.CurrentTable = $null
    foreach ($g in @($script:gridIdx, $script:gridStats, $script:gridMissing, $script:gridPriority, $script:gridHist, $script:gridDensity)) { if ($g) { $g.DataSource = $null } }
    foreach ($t in @($script:txtOverview, $script:txtIdxDetail, $script:txtStatDetail, $script:txtMissingScript)) { if ($t) { $t.Text = '' } }
    if ($script:chart) { foreach ($s in $script:chart.Series) { $s.Points.Clear() }; $script:chart.Titles.Clear() }
    $script:lblTable.Text = 'No table selected'
}
#endregion

#region ------------------------------------------------------------------ table analysis
function Open-SelectedTable {
    $row = $script:gridTables.CurrentRow
    if (-not $row -or -not $script:State.DbContext) { return }
    $objectId = [int]$row.Cells['ObjectId'].Value
    if ($script:State.CurrentObjectId -eq $objectId -and $script:State.Analysis) { return }
    if ($script:State.Busy) { $script:State.PendingObjectId = $objectId; return }
    Invoke-TableLoad -ObjectId $objectId
}

function Invoke-TableLoad {
    param([int]$ObjectId, [switch]$Force)
    $script:State.PendingObjectId = $null
    $ctx = $script:State.DbContext
    $name = ($script:State.TablesMaster.Select("ObjectId = $ObjectId") | Select-Object -First 1)
    $label = if ($name) { "[$($name['SchemaName'])].[$($name['TableName'])]" } else { "object_id $ObjectId" }
    [void](Start-AnalyzerJob -Status "Analyzing $label" -Arguments @{ Ctx = $ctx; ObjectId = $ObjectId; Info = $script:State.ServerInfo; Db = $script:State.DbInfo; Settings = $script:Settings; Env = $script:State.Environment } -Work {
        param($Ctx, $ObjectId, $Info, $Db, $Settings, $Env)
        $ErrorActionPreference = 'Stop'
        $a = Get-TableAnalysis -Context $Ctx -ObjectId $ObjectId -ServerInfo $Info -DatabaseInfo $Db -Settings $Settings -Environment $Env
        @{ Analysis = $a; ObjectId = $ObjectId }
    } -OnDone {
        param($r)
        $script:State.Analysis = $r.Analysis
        $script:State.CurrentObjectId = $r.ObjectId
        $o = $r.Analysis.Overview
        $script:State.CurrentTable = $o
        $script:lblTable.Text = ("[{0}].[{1}]   {2} rows   {3}" -f $o.SchemaName, $o.TableName, (Format-Count ([double]$o.RowCount)), (Format-Size ([double]$o.ReservedMB)))
        Update-Assessments
        if ($r.Analysis.Warnings.Count) { $script:lblWarn.Text = 'Partial results: ' + ($r.Analysis.Warnings -join ' | ') }
        if ($script:chkProcForTable.Checked -and $script:State.ProcMaster) { Update-ProcReferences }
    })
}

function Get-FragKey { param([int]$IndexId) "$($script:State.DbInfo.DatabaseName)|$($script:State.CurrentObjectId)|$IndexId" }

function Update-Assessments {
    $a = $script:State.Analysis
    if (-not $a) { return }
    $settings = $script:Settings
    $uptime = Get-UptimeDays
    $db = $script:State.DbInfo
    $redundancy = Find-IndexRedundancy -Indexes $a.Indexes

    $idxRows = foreach ($ix in $a.Indexes) {
        $frag = $script:State.Frag[(Get-FragKey $ix.IndexId)]
        $cs = if ($a.Columnstore.ContainsKey([int]$ix.IndexId)) { $a.Columnstore[[int]$ix.IndexId] } else { $null }
        $red = if ($redundancy.ContainsKey([int]$ix.IndexId)) { $redundancy[[int]$ix.IndexId] } else { $null }
        $as = Get-IndexAssessment -Index $ix -Fragmentation $frag -Columnstore $cs -Redundancy $red -Settings $settings -UptimeDays $uptime
        $h = [ordered]@{ Priority = $as.Priority; Score = $as.Score; Action = $as.Action }
        foreach ($p in $ix.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $h['FragPct']    = if ($frag -and $frag.Measured) { $frag.FragmentationPct } else { $null }
        $h['DensityPct'] = if ($frag -and $frag.Measured) { $frag.PageDensityPct } else { $null }
        $h['FragMeasured'] = if ($frag -and $frag.Measured) { $frag.MeasuredAt } else { $null }
        $h['Reasons'] = $as.Reasons
        $h['_Assessment'] = $as; $h['_Frag'] = $frag; $h['_Columnstore'] = $cs
        [pscustomobject]$h
    }
    $script:State.IndexRows = @($idxRows | Sort-Object @{ Expression = 'Score'; Descending = $true }, IndexId)

    $statRows = foreach ($st in $a.Statistics) {
        $as = Get-StatsAssessment -Stat $st -CompatLevel ([int]$db.CompatLevel) -AutoUpdateOn ([bool]$db.AutoUpdateStats) -TableRows $a.Overview.RowCount
        $h = [ordered]@{ Priority = $as.Priority; Score = $as.Score; SuggestedMethod = $as.SuggestedMethod }
        foreach ($p in $st.PSObject.Properties) { $h[$p.Name] = $p.Value }
        $h['SamplePct'] = $as.SamplePct; $h['ModPct'] = $as.ModPct; $h['PctOfThreshold'] = $as.PctOfThreshold; $h['AgeDays'] = $as.AgeDays
        $h['Reasons'] = $as.Reasons; $h['_Assessment'] = $as
        [pscustomobject]$h
    }
    $script:State.StatsRows = @($statRows | Sort-Object @{ Expression = 'Score'; Descending = $true }, StatsId)

    $idxCols = 'Priority','Score','Action','IndexName','IndexType','KeyColumns','IncludedColumns','SizeMB','RowCount','Seeks','Scans','Lookups','Writes','LastRead','LastWrite','FragPct','DensityPct','FragMeasured','LeafPageAllocations','Compression','FillFactorPct','Partitions','RowLockWaitMs','PageLatchWaitMs','ForwardedFetches','IsUnique','IsPrimaryKey','IsDisabled','FilterDefinition','Reasons','IndexId'
    $script:gridIdx.DataSource = New-GridTable -Items $script:State.IndexRows -Columns $idxCols
    $statCols = 'Priority','Score','StatName','StatType','Columns','LastUpdated','AgeDays','Rows','RowsSampled','SamplePct','Modifications','ModPct','PctOfThreshold','Steps','SuggestedMethod','NoRecompute','IsIncremental','HasFilter','PersistedSamplePct','Reasons','StatsId'
    $script:gridStats.DataSource = New-GridTable -Items $script:State.StatsRows -Columns $statCols
    $missCols = 'ImprovementScore','EqualityColumns','InequalityColumns','IncludedColumns','Seeks','Scans','AvgQueryCost','AvgImpactPct','LastSeek','Idx'
    $k = 0
    foreach ($m in $a.Missing) { $m | Add-Member -NotePropertyName Idx -NotePropertyValue $k -Force; $k++ }
    $script:gridMissing.DataSource = New-GridTable -Items $a.Missing -Columns $missCols
    $script:tabMissing.Text = "Missing indexes ($(@($a.Missing).Count))"
    $script:tabIdx.Text = "Indexes ($(@($a.Indexes).Count))"
    $script:tabStats.Text = "Statistics ($(@($a.Statistics).Count))"
    Update-Overview
}

function Update-Overview {
    $a = $script:State.Analysis; $o = $a.Overview; $db = $script:State.DbInfo; $i = $script:State.ServerInfo
    $uptime = Get-UptimeDays
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("TABLE  [{0}].[{1}]" -f $o.SchemaName, $o.TableName))
    [void]$sb.AppendLine(("  Rows {0}   |  Total size {1} (data {2}, nonclustered indexes {3})" -f (Format-Count ([double]$o.RowCount)), (Format-Size ([double]$o.ReservedMB)), (Format-Size ([double]$o.DataMB)), (Format-Size ([double]$o.IndexMB))))
    [void]$sb.AppendLine(("  {0} indexes, {1} statistics, {2} columns  |  {3}{4}  |  partitions: {5}{6}" -f $o.IndexCount, $o.StatsCount, $o.ColumnCount,
        $(if ($o.HasClusteredIndex) { 'clustered' } else { 'HEAP (no clustered index)' }), $(if ($o.HasPrimaryKey) { ', has PK' } else { ', no PK' }), $o.Partitions, $(if ($o.PartitionScheme) { " (scheme $($o.PartitionScheme))" } else { '' })))
    [void]$sb.AppendLine(("  Created {0:yyyy-MM-dd}, schema last changed {1:yyyy-MM-dd}, lock escalation {2}{3}" -f [datetime]$o.Created, [datetime]$o.SchemaModified, $o.LockEscalation, $(if ($o.IsMemoryOptimized) { '  |  MEMORY-OPTIMIZED (rowstore maintenance does not apply)' } else { '' })))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(("DATABASE  {0}: compat {1}, {2} recovery, auto create stats {3}, auto update stats {4}{5}, data free {6}, log {7} ({8}% used)" -f $db.DatabaseName, $db.CompatLevel, $db.RecoveryModel,
        $(if ($db.AutoCreateStats) { 'ON' } else { 'OFF' }), $(if ($db.AutoUpdateStats) { 'ON' } else { 'OFF' }), $(if ($db.AutoUpdateStatsAsync) { ' (async)' } else { '' }),
        (Format-Size ([double]$db.DataFreeMB)), (Format-Size ([double]$db.LogSizeMB)), $db.LogUsedPct))
    if ([int]$db.CompatLevel -lt 130) { [void]$sb.AppendLine('  Compat level below 130: large tables wait for 20% of rows to change before auto-updating stats (unless trace flag 2371 is on).') }
    [void]$sb.AppendLine(("SERVER  {0}  |  {1}  |  usage counters cover {2:N0} days{3}" -f $i.ServerName, $i.Edition, $uptime, $(if ($uptime -lt 14) { '  <-- short: do not judge unused indexes on this alone' } else { '' })))
    [void]$sb.AppendLine(("        Online rebuild: {0}   Resumable: {1}   UPDATE STATISTICS MAXDOP: {2}   Histogram DMV: {3}" -f $i.Capabilities.OnlineRebuild, $i.Capabilities.ResumableRebuild, $i.Capabilities.UpdateStatsMaxDop, $i.Capabilities.StatsHistogramDmv))
    [void]$sb.AppendLine('')
    $hiI = @($script:State.IndexRows | Where-Object { $_.Priority -in 'High', 'Medium' }).Count
    $hiS = @($script:State.StatsRows | Where-Object { $_.Priority -in 'High', 'Medium' }).Count
    $rev = @($script:State.IndexRows | Where-Object { $_.Priority -eq 'Review' }).Count
    $unmeasured = @($script:State.IndexRows | Where-Object { $null -eq $_.FragPct -and $_.TypeCode -in 1, 2 -and $_.InRowPages -ge $script:Settings.MinPageCount }).Count
    [void]$sb.AppendLine("SUMMARY  $hiI index(es) and $hiS statistic(s) need attention (High/Medium); $rev index(es) to review (unused, duplicate, disabled).")
    if ($unmeasured) { [void]$sb.AppendLine("         $unmeasured index(es) have no fragmentation reading yet: Indexes tab > Check fragmentation (LIMITED, size-capped).") }
    if (@($a.Missing).Count) { [void]$sb.AppendLine("         SQL Server has logged $(@($a.Missing).Count) missing-index request(s) for this table: see the Missing indexes tab before adding anything.") }
    $script:txtOverview.Text = $sb.ToString()

    $pri = @()
    foreach ($r in $script:State.IndexRows) { if ($r.Priority -ne 'None') { $pri += [pscustomobject]@{ Priority = $r.Priority; Score = $r.Score; Kind = 'Index'; Name = $r.IndexName; Action = $r.Action; Why = ($r.Reasons -join ' | ') } } }
    foreach ($r in $script:State.StatsRows) { if ($r.Priority -ne 'None') { $pri += [pscustomobject]@{ Priority = $r.Priority; Score = $r.Score; Kind = 'Statistic'; Name = $r.StatName; Action = $r._Assessment.Action; Why = ($r.Reasons -join ' | ') } } }
    $order = @{ 'High' = 0; 'Medium' = 1; 'Review' = 2; 'Low' = 3 }
    $pri = @($pri | Sort-Object @{ Expression = { $order[$_.Priority] } }, @{ Expression = 'Score'; Descending = $true })
    $script:gridPriority.DataSource = New-GridTable -Items $pri -Columns 'Priority', 'Score', 'Kind', 'Name', 'Action', 'Why'
}

function Get-SelectedRows {
    param($Grid, [object[]]$Rows, [string]$IdColumn, [string]$IdProperty)
    $ids = @()
    foreach ($r in $Grid.SelectedRows) { $ids += [int]$r.Cells[$IdColumn].Value }
    if ($ids.Count -eq 0 -and $Grid.CurrentRow) { $ids += [int]$Grid.CurrentRow.Cells[$IdColumn].Value }
    @($Rows | Where-Object { $ids -contains [int]$_.$IdProperty })
}

function Show-IndexDetail {
    $sel = @(Get-SelectedRows -Grid $script:gridIdx -Rows $script:State.IndexRows -IdColumn 'IndexId' -IdProperty 'IndexId')
    if ($sel.Count -eq 0) { return }
    $r = $sel[0]
    $script:txtIdxDetail.Text = Get-IndexPlainEnglish -Index $r -Assessment $r._Assessment -Fragmentation $r._Frag -Columnstore $r._Columnstore -UptimeDays (Get-UptimeDays) -StartTime $script:State.ServerInfo.StartTime
}

function Invoke-FragmentationCheck {
    param([ValidateSet('LIMITED', 'SAMPLED')][string]$Mode)
    if (-not $script:State.Analysis) { return }
    $sel = @(Get-SelectedRows -Grid $script:gridIdx -Rows $script:State.IndexRows -IdColumn 'IndexId' -IdProperty 'IndexId')
    if ($script:gridIdx.SelectedRows.Count -le 1) {
        $answer = [System.Windows.Forms.MessageBox]::Show("Check all rowstore indexes of this table?`r`n`r`nYes = all indexes`r`nNo = only the selected index", "Check fragmentation ($Mode)", 'YesNoCancel', 'Question')
        if ($answer -eq 'Cancel') { return }
        if ($answer -eq 'Yes') { $sel = @($script:State.IndexRows) }
    }
    $sel = @($sel | Where-Object { $_.TypeCode -in 0, 1, 2 -and -not $_.IsDisabled -and -not $_.IsHypothetical })
    if ($sel.Count -eq 0) { Set-Status 'Nothing to scan (columnstore, disabled and hypothetical indexes are skipped).'; return }
    $isProd = Test-IsProductionEnvironment $script:State.Environment
    $cap = if ($isProd) { [double]$script:Settings.MaxScanGBProduction } else { [double]$script:Settings.MaxScanGBNonProduction }
    if ($Mode -eq 'SAMPLED') { $cap = [math]::Min($cap, [double]$script:Settings.MaxSampledScanGB) }
    $over = @($sel | Where-Object { [double]$_.SizeMB / 1024 -gt $cap })
    $sizeGB = ($sel | Where-Object { [double]$_.SizeMB / 1024 -le $cap } | Measure-Object SizeMB -Sum).Sum / 1024
    $msg = ("Read-only scan with sys.dm_db_index_physical_stats ({0} mode, MAXDOP 1, one index at a time).`r`n`r`nIndexes: {1}   Size to scan: {2:N1} GB`r`nSkipped (over the {3} GB cap for {4}): {5}`r`n`r`n{6}`r`nThe scan is refused automatically if the server looks busy. You can cancel at any time.`r`n`r`nContinue?" -f
        $Mode, ($sel.Count - $over.Count), $sizeGB, $cap, $script:State.Environment, $over.Count,
        $(if ($Mode -eq 'LIMITED') { 'LIMITED reads only the level above the leaf pages: the cheapest mode. It reports fragmentation but not page fullness.' } else { 'SAMPLED reads about 1% of the leaf pages and also reports page fullness (density).' }))
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Check fragmentation ($Mode)", 'YesNo', 'Question') -ne 'Yes') { return }

    $indexes = @($sel | ForEach-Object { $x = $_; $o = [ordered]@{}; foreach ($p in $x.PSObject.Properties) { if ($p.Name -notlike '_*') { $o[$p.Name] = $p.Value } }; [pscustomobject]$o })
    [void](Start-AnalyzerJob -Status "Fragmentation check ($Mode, $($indexes.Count) indexes)" -Arguments @{ Ctx = $script:State.DbContext; ObjectId = [int]$script:State.CurrentObjectId; Indexes = $indexes; Mode = $Mode; Env = $script:State.Environment; Settings = $script:Settings; Shared = $script:Shared } -Work {
        param($Ctx, $ObjectId, $Indexes, $Mode, $Env, $Settings, $Shared)
        $ErrorActionPreference = 'Stop'
        $results = @{}
        foreach ($ix in $Indexes) {
            if ($Shared.Cancel) { break }
            try { $results[[int]$ix.IndexId] = Get-IndexFragmentation -Context $Ctx -ObjectId $ObjectId -Index $ix -Mode $Mode -Environment $Env -Settings $Settings }
            catch {
                if ($Shared.Cancel) { break }
                $results[[int]$ix.IndexId] = [pscustomobject]@{ IndexId = $ix.IndexId; Mode = $Mode; Measured = $false; Skipped = "Scan failed: $($_.Exception.Message)"; FragmentationPct = $null; PageDensityPct = $null; PageCount = $null; ForwardedRecords = $null; MeasuredAt = $null }
            }
        }
        @{ Results = $results }
    } -OnDone {
        param($r)
        $skipped = 0
        foreach ($k in $r.Results.Keys) {
            $res = $r.Results[$k]
            $key = Get-FragKey $k
            $prev = $script:State.Frag[$key]
            # keep an earlier SAMPLED density if this was a LIMITED re-check
            if ($res.Measured -and $null -eq $res.PageDensityPct -and $prev -and $prev.Measured -and $null -ne $prev.PageDensityPct) { $res.PageDensityPct = $prev.PageDensityPct }
            if (-not $res.Measured) { $skipped++ ; if ($prev -and $prev.Measured) { continue } }
            $script:State.Frag[$key] = $res
        }
        Update-Assessments
        Show-IndexDetail
        if ($skipped) { Set-Status "Fragmentation check done; $skipped index(es) skipped (see the Action/Reasons columns)." }
    })
}

function Show-StatDetail {
    $sel = @(Get-SelectedRows -Grid $script:gridStats -Rows $script:State.StatsRows -IdColumn 'StatsId' -IdProperty 'StatsId')
    if ($sel.Count -eq 0) { return }
    $st = $sel[0]
    $key = "$($script:State.DbInfo.DatabaseName)|$($script:State.CurrentObjectId)|$($st.StatsId)|$($st.LastUpdated)"
    if ($script:State.StatDetail.ContainsKey($key)) { Write-StatDetail -Stat $st -Detail $script:State.StatDetail[$key]; return }
    Write-StatDetail -Stat $st -Detail $null
    if (-not $st.LastUpdated) { return }
    if ($script:State.Busy) { $script:StatTimer.Start(); return }
    $o = $script:State.CurrentTable
    [void](Start-AnalyzerJob -Status "Reading histogram of $($st.StatName)" -Arguments @{ Ctx = $script:State.DbContext; ObjectId = [int]$script:State.CurrentObjectId; StatsId = [int]$st.StatsId; Schema = $o.SchemaName; Table = $o.TableName; Stat = $st.StatName; Info = $script:State.ServerInfo; Key = $key } -Work {
        param($Ctx, $ObjectId, $StatsId, $Schema, $Table, $Stat, $Info, $Key)
        $ErrorActionPreference = 'Stop'
        $d = Get-StatisticsDetail -Context $Ctx -ObjectId $ObjectId -StatsId $StatsId -SchemaName $Schema -TableName $Table -StatName $Stat -ServerInfo $Info
        @{ Detail = $d; Key = $Key; StatsId = $StatsId }
    } -OnDone {
        param($r)
        $script:State.StatDetail[$r.Key] = $r.Detail
        $cur = @(Get-SelectedRows -Grid $script:gridStats -Rows $script:State.StatsRows -IdColumn 'StatsId' -IdProperty 'StatsId') | Select-Object -First 1
        if ($cur -and [int]$cur.StatsId -eq [int]$r.StatsId) { Write-StatDetail -Stat $cur -Detail $r.Detail }
    })
}

function Write-StatDetail {
    param($Stat, $Detail)
    $o = $script:State.CurrentTable
    $script:txtStatDetail.Text = Get-StatsPlainEnglish -Stat $Stat -Assessment $Stat._Assessment -Detail $Detail -TableName ("[{0}].[{1}]" -f $o.SchemaName, $o.TableName)
    if (-not $Detail) {
        $script:gridHist.DataSource = $null; $script:gridDensity.DataSource = $null
        if ($script:chart) { foreach ($s in $script:chart.Series) { $s.Points.Clear() } ; $script:chart.Titles.Clear() }
        return
    }
    $script:gridHist.DataSource = New-GridTable -Items $Detail.Histogram -Columns 'Step', 'RangeHiKey', 'EqRows', 'RangeRows', 'DistinctRangeRows', 'AvgRangeRows'
    $script:gridDensity.DataSource = New-GridTable -Items $Detail.Density -Columns 'Columns', 'AllDensity', 'DistinctValues', 'AvgLengthBytes'
    if ($script:gridDensity.Columns.Contains('AllDensity')) { $script:gridDensity.Columns['AllDensity'].DefaultCellStyle.Format = 'G6' }
    if ($script:chart) {
        $eq = $script:chart.Series['eq']; $rg = $script:chart.Series['range']
        $eq.Points.Clear(); $rg.Points.Clear(); $script:chart.Titles.Clear()
        $lead = ([string]$Stat.Columns).Split(',')[0].Trim()
        [void]$script:chart.Titles.Add("How the rows of '$lead' are spread over its values ($(@($Detail.Histogram).Count) buckets)")
        foreach ($h in $Detail.Histogram) {
            $lbl = ([string]$h.RangeHiKey) -replace '#', '# '
            if ($lbl.Length -gt 24) { $lbl = $lbl.Substring(0, 24) + '...' }
            # AddY + AxisLabel: AddXY would pick the (double, double) overload and fail on keys like NULL, dates or text.
            $p1 = $eq.Points.AddY([double]$h.EqRows)
            $eq.Points[$p1].AxisLabel = $lbl
            $eq.Points[$p1].ToolTip = ("Value = {0}`nRows with exactly this value: {1:N0}" -f $h.RangeHiKey, [double]$h.EqRows)
            $p2 = $rg.Points.AddY([double]$h.RangeRows)
            $rg.Points[$p2].AxisLabel = $lbl
            $rg.Points[$p2].ToolTip = ("Rows between the previous bucket and {0}: {1:N0}`nDistinct values in between: {2:N0}" -f $h.RangeHiKey, [double]$h.RangeRows, [double]$h.DistinctRangeRows)
        }
    }
}

function Show-MissingDetail {
    $row = $script:gridMissing.CurrentRow
    if (-not $row -or -not $script:State.Analysis) { return }
    $m = @($script:State.Analysis.Missing)[[int]$row.Cells['Idx'].Value]
    if (-not $m) { return }
    $o = $script:State.CurrentTable
    $script:txtMissingScript.Text = New-MissingIndexScript -SchemaName $o.SchemaName -TableName $o.TableName -Missing $m -Online ([bool]$script:State.ServerInfo.Capabilities.OnlineRebuild)
}
#endregion

#region ------------------------------------------------------------------ procedure analysis (opt-in)
# Nothing in this region runs unless the user clicks "Load procedures" / "Analyze".
# It reads the plan a procedure already has (plan cache or Query Store). It never executes or compiles it.

function Clear-ProcViews {
    $script:State.ProcMaster = $null; $script:State.ProcBase = $null; $script:State.ProcRefIds = $null; $script:State.ProcRefFor = $null
    $script:State.ProcAnalysis = $null; $script:State.ProcFindings = @(); $script:State.ProcName = $null; $script:State.ProcObjectId = $null
    if ($script:gridProcs) { $script:gridProcs.DataSource = $null }
    foreach ($g in @($script:gridProcFindings, $script:gridProcStmts)) { if ($g) { $g.DataSource = $null } }
    foreach ($t in @($script:txtProcSummary, $script:txtProcFinding, $script:txtProcStmt, $script:txtPlanXml)) { if ($t) { $t.Text = '' } }
    if ($script:cboProcPlan) { $script:State.ProcPlanSuppress = $true; $script:cboProcPlan.Items.Clear(); $script:State.ProcPlanSuppress = $false }
    if ($script:lblProcCount) { $script:lblProcCount.Text = 'Click "Load procedures" to list the procedures of this database (metadata only).' }
}

function Invoke-ProcedureListLoad {
    if (-not $script:State.DbContext) { Set-Status 'Pick a database first.'; return }
    [void](Start-AnalyzerJob -Status 'Loading procedure list' -Arguments @{ Ctx = $script:State.DbContext } -Work {
        param($Ctx)
        $ErrorActionPreference = 'Stop'
        @{ Procs = (Get-ProcedureList -Context $Ctx) }
    } -OnDone {
        param($r)
        $script:State.ProcMaster = $r.Procs
        $script:State.ProcRefIds = $null; $script:State.ProcRefFor = $null
        if ($script:chkProcForTable.Checked) { Update-ProcReferences } else { Set-ProcBase }
    })
}

function Update-ProcReferences {
    # procedures that reference the table currently open on the left (sys.sql_expression_dependencies, metadata only)
    if (-not $script:State.CurrentObjectId) {
        $script:chkProcForTable.Checked = $false
        Set-Status 'Open a table first to filter procedures by it.'
        return
    }
    if ($script:State.ProcRefFor -eq $script:State.CurrentObjectId -and $null -ne $script:State.ProcRefIds) { Set-ProcBase; return }
    $oid = [int]$script:State.CurrentObjectId
    [void](Start-AnalyzerJob -Status 'Finding procedures that use the selected table' -Arguments @{ Ctx = $script:State.DbContext; ObjectId = $oid } -Work {
        param($Ctx, $ObjectId)
        $ErrorActionPreference = 'Stop'
        @{ Ids = @(Get-ProceduresReferencingTable -Context $Ctx -ObjectId $ObjectId); ObjectId = $ObjectId }
    } -OnDone {
        param($r)
        $set = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($i in $r.Ids) { [void]$set.Add([int]$i) }
        $script:State.ProcRefIds = $set; $script:State.ProcRefFor = $r.ObjectId
        Set-ProcBase
    })
}

function Set-ProcBase {
    $master = $script:State.ProcMaster
    if (-not $master) { return }
    $base = $master
    if ($script:chkProcForTable.Checked -and $null -ne $script:State.ProcRefIds) {
        $base = $master.Clone()
        foreach ($row in $master.Rows) { if ($script:State.ProcRefIds.Contains([int]$row['ObjectId'])) { [void]$base.ImportRow($row) } }
    }
    $script:State.ProcBase = $base
    $n = $base.Rows.Count
    $sch = New-Object string[] $n; $nm = New-Object string[] $n
    for ($i = 0; $i -lt $n; $i++) { $sch[$i] = [string]$base.Rows[$i]['SchemaName']; $nm[$i] = [string]$base.Rows[$i]['ProcName'] }
    $script:State.ProcSchemas = $sch; $script:State.ProcNames = $nm
    Update-ProcFilter
}

function Update-ProcFilter {
    $base = $script:State.ProcBase
    if (-not $base) { return }
    $view = $base.Clone()
    $q = $script:txtProcSearch.Text
    $max = [int]$script:Settings.MaxTablesDisplayed
    if ($base.Rows.Count -gt 0) {
        $hits = @(Find-Table -SchemaNames $script:State.ProcSchemas -TableNames $script:State.ProcNames -Query $q -Top $max)
        foreach ($i in $hits) { [void]$view.ImportRow($base.Rows[[int]$i]) }
    }
    $scope = if ($script:chkProcForTable.Checked -and $script:State.CurrentTable) { " that reference [$($script:State.CurrentTable.SchemaName)].[$($script:State.CurrentTable.TableName)]" } else { '' }
    $script:lblProcCount.Text = ("{0:N0} procedure(s){1}; showing {2:N0}{3}." -f $base.Rows.Count, $scope, $view.Rows.Count, $(if ($q) { " matching '$q'" } else { '' }))
    $script:gridProcs.DataSource = $view
}

function Invoke-ProcAnalysis {
    param([string]$SourceKey)
    $row = $script:gridProcs.CurrentRow
    $oid = $null; $name = $null
    if ($SourceKey -and $script:State.ProcObjectId) { $oid = [int]$script:State.ProcObjectId; $name = $script:State.ProcName }
    elseif ($row) { $oid = [int]$row.Cells['ObjectId'].Value; $name = "[$($row.Cells['SchemaName'].Value)].[$($row.Cells['ProcName'].Value)]" }
    if (-not $oid) { Set-Status 'Pick a procedure in the list first.'; return }
    [void](Start-AnalyzerJob -Status "Reading the existing plan of $name" -Arguments @{ Ctx = $script:State.DbContext; ObjectId = $oid; Info = $script:State.ServerInfo; Db = $script:State.DbInfo; Key = $SourceKey; Name = $name } -Work {
        param($Ctx, $ObjectId, $Info, $Db, $Key, $Name)
        $ErrorActionPreference = 'Stop'
        @{ Analysis = (Invoke-ProcedureAnalysis -Context $Ctx -ObjectId $ObjectId -ServerInfo $Info -DatabaseInfo $Db -SourceKey $Key); Name = $Name; ObjectId = $ObjectId }
    } -OnDone {
        param($r)
        $script:State.ProcAnalysis = $r.Analysis; $script:State.ProcName = $r.Name; $script:State.ProcObjectId = $r.ObjectId
        $script:State.ProcPlanSuppress = $true
        $script:cboProcPlan.Items.Clear()
        $sel = -1; $k = 0
        foreach ($s in @($r.Analysis.Sources)) { [void]$script:cboProcPlan.Items.Add($s.Label); if ($r.Analysis.Source -and $s.Key -eq $r.Analysis.Source.Key) { $sel = $k }; $k++ }
        if ($sel -ge 0) { $script:cboProcPlan.SelectedIndex = $sel }
        $script:State.ProcPlanSuppress = $false
        Update-ProcFindings
        $script:procTabs.SelectedIndex = 0
    })
}

function Get-ProcFragLookup {
    $db = [string]$script:State.DbInfo.DatabaseName
    $h = @{}
    foreach ($k in @($script:State.Frag.Keys)) {
        $parts = ([string]$k).Split('|')
        if ($parts.Count -eq 3 -and $parts[0] -eq $db) { $h["$($parts[1])|$($parts[2])"] = $script:State.Frag[$k] }
    }
    $h
}

function Format-PlanXml {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text) -or $Text.Length -gt 4000000) { return $Text }
    try {
        $rs = New-Object System.Xml.XmlReaderSettings; $rs.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit; $rs.XmlResolver = $null
        $reader = [System.Xml.XmlReader]::Create((New-Object System.IO.StringReader($Text)), $rs)
        $doc = New-Object System.Xml.XmlDocument; $doc.XmlResolver = $null
        try { $doc.Load($reader) } finally { $reader.Close() }
        $sw = New-Object System.IO.StringWriter
        $ws = New-Object System.Xml.XmlWriterSettings; $ws.Indent = $true; $ws.OmitXmlDeclaration = $true
        $xw = [System.Xml.XmlWriter]::Create($sw, $ws); $doc.Save($xw); $xw.Close()
        $sw.ToString()
    } catch { $Text }
}

function Update-ProcFindings {
    $a = $script:State.ProcAnalysis
    if (-not $a) { return }
    $findings = @(Get-ProcedureFindings -Analysis $a -Settings $script:Settings -DatabaseInfo $script:State.DbInfo -Frag (Get-ProcFragLookup) -UptimeDays (Get-UptimeDays))
    $k = 0; foreach ($f in $findings) { $f | Add-Member -NotePropertyName Idx -NotePropertyValue $k -Force; $k++ }
    $script:State.ProcFindings = $findings
    $script:gridProcFindings.DataSource = New-GridTable -Items $findings -Columns 'Priority', 'Score', 'Kind', 'Table', 'Object', 'Action', 'Finding', 'Idx'
    $script:txtProcSummary.Text = Get-ProcedureSummaryText -Analysis $a -ProcName $script:State.ProcName -Findings $findings
    $stmts = @()
    if ($a.Plan) {
        foreach ($st in $a.Plan.Statements) {
            $rt = $null; if ($st.PSObject.Properties['Runtime']) { $rt = $st.Runtime }
            $mis = $null; if ($st.PSObject.Properties['MisestimateRatio']) { $mis = $st.MisestimateRatio }
            $txt = [string]$st.Text; $short = ($txt -replace '\s+', ' '); if ($short.Length -gt 200) { $short = $short.Substring(0, 200) + '...' }
            $stmts += [pscustomobject]@{
                Key = $st.Key; Statement = $st.StatementId; Type = $st.Type; EstRows = $st.EstRows
                AvgRows = $(if ($rt) { $rt.AvgRows }); MinRows = $(if ($rt) { $rt.MinRows }); MaxRows = $(if ($rt) { $rt.MaxRows }); EstimateOffBy = $mis
                Executions = $(if ($rt) { $rt.Executions }); AvgCpuMs = $(if ($rt) { $rt.AvgCpuMs }); AvgElapsedMs = $(if ($rt) { $rt.AvgElapsedMs }); AvgLogicalReads = $(if ($rt) { $rt.AvgLogicalReads })
                Cost = $st.Cost; Tables = (@($st.Tables) -join ', '); Text = $short; Doc = $st.DocIndex
            }
        }
    }
    $script:State.ProcStatements = $stmts
    $script:gridProcStmts.DataSource = New-GridTable -Items $stmts -Columns 'Statement', 'Type', 'EstRows', 'AvgRows', 'EstimateOffBy', 'MinRows', 'MaxRows', 'Executions', 'AvgCpuMs', 'AvgElapsedMs', 'AvgLogicalReads', 'Cost', 'Tables', 'Text', 'Key', 'Doc'
    $script:txtPlanXml.Text = if (@($a.PlanTexts).Count -gt 0) { (@($a.PlanTexts) | ForEach-Object { Format-PlanXml $_ }) -join "`r`n`r`n" } else { '(no plan)' }
    $script:tabProcFind.Text = "Findings ($(@($findings | Where-Object { $_.Priority -ne 'None' }).Count))"
    $script:tabProcStmts.Text = "Statements ($($stmts.Count))"
}

function Get-SelectedProcFindings {
    $out = @()
    foreach ($r in $script:gridProcFindings.SelectedRows) { $out += $script:State.ProcFindings[[int]$r.Cells['Idx'].Value] }
    $out
}

function Show-ProcFindingDetail {
    $row = $script:gridProcFindings.CurrentRow
    if (-not $row -or -not $script:State.ProcFindings) { return }
    $f = $script:State.ProcFindings[[int]$row.Cells['Idx'].Value]
    if (-not $f) { return }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($f.Priority) (score $($f.Score))  |  $($f.Kind)  |  $($f.Table)  |  $($f.Object)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine($f.Finding)
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("SUGGESTED: $($f.Action)")
    if ($f.PSObject.Properties['_Stat'] -and $f._Stat) {
        [void]$sb.AppendLine(''); [void]$sb.AppendLine((Get-StatsPlainEnglish -Stat $f._Stat -Assessment $f._Assessment -Detail $null -TableName $f.Table))
    }
    if ($f.PSObject.Properties['_Index'] -and $f._Index) {
        $fr = $null; if ($f.PSObject.Properties['_Frag']) { $fr = $f._Frag }
        $as = $null; if ($f.PSObject.Properties['_Assessment']) { $as = $f._Assessment }
        [void]$sb.AppendLine(''); [void]$sb.AppendLine((Get-IndexPlainEnglish -Index $f._Index -Assessment $as -Fragmentation $fr -UptimeDays (Get-UptimeDays) -StartTime $script:State.ServerInfo.StartTime))
    }
    $extra = Get-ProcFindingScript $f
    if ($extra) { [void]$sb.AppendLine(''); [void]$sb.AppendLine('SCRIPT (text only):'); [void]$sb.AppendLine($extra) }
    if ($f.PSObject.Properties['StatementKey'] -and $f.StatementKey) {
        $st = @($script:State.ProcAnalysis.Plan.Statements | Where-Object { $_.Key -eq $f.StatementKey }) | Select-Object -First 1
        if ($st) { [void]$sb.AppendLine(''); [void]$sb.AppendLine("STATEMENT $($st.StatementId):"); [void]$sb.AppendLine($st.Text) }
    }
    $script:txtProcFinding.Text = $sb.ToString()
}

function Get-ProcFindingScript {
    param($Finding)
    if ($Finding.PSObject.Properties['_Missing'] -and $Finding._Missing) {
        return (New-MissingIndexScript -SchemaName $Finding.MissingSchema -TableName $Finding.MissingTable -Missing $Finding._Missing -Online ([bool]$script:State.ServerInfo.Capabilities.OnlineRebuild))
    }
    if ($Finding.PSObject.Properties['CreateStatsColumn'] -and $Finding.CreateStatsColumn) {
        $c = $Finding.CreateStatsColumn
        return (New-CreateStatisticsScript -SchemaName $(if ($c.Schema) { $c.Schema } else { 'dbo' }) -TableName $c.Table -ColumnName $c.Column)
    }
    $null
}

function Get-ProcTable {
    param([int]$ObjectId)
    @($script:State.ProcAnalysis.Tables | Where-Object { [int]$_.ObjectId -eq $ObjectId }) | Select-Object -First 1
}

function Add-ProcStatsScript {
    $a = $script:State.ProcAnalysis
    if (-not $a) { return }
    $sel = @(Get-SelectedProcFindings | Where-Object { $_.Kind -eq 'Statistic' })
    if ($sel.Count -eq 0) { $sel = @($script:State.ProcFindings | Where-Object { $_.Kind -eq 'Statistic' -and $_.Priority -in 'High', 'Medium' }) }
    if ($sel.Count -eq 0) { Set-Status 'No statistics need updating for this procedure (select rows to script specific ones).'; return }
    $caps = $script:State.ServerInfo.Capabilities
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("-- Statistics maintenance generated from the plan of $($script:State.ProcName) ($($a.Source.Source)). Text only.")
    $first = $true
    foreach ($g in @($sel | Where-Object { $_.PSObject.Properties['_Stat'] -and $_._Stat } | Group-Object TableObjectId)) {
        $t = Get-ProcTable ([int]$g.Name)
        if (-not $t) { continue }
        $pct = [int](($g.Group | ForEach-Object { $_._Assessment.SuggestedSamplePct } | Measure-Object -Maximum).Maximum)
        $opts = @{ Method = $(if ($pct -ge 100) { 'FULLSCAN' } else { 'SAMPLE' }); SamplePercent = $pct; MaxDop = $(if ($caps.UpdateStatsMaxDop) { [int]$script:Settings.DefaultMaxDop } else { $null }) }
        $dataMB = ($t.Indexes | Where-Object { $_.IndexId -in 0, 1 } | Measure-Object SizeMB -Sum).Sum
        [void]$sb.AppendLine((New-StatsMaintenanceScript -SchemaName $t.SchemaName -TableName $t.TableName -Statistics @($g.Group | ForEach-Object { $_._Stat }) -Options $opts -ServerInfo $script:State.ServerInfo -DatabaseInfo $script:State.DbInfo -Environment $script:State.Environment -TableDataMB ([double]$dataMB) -IncludeHeader:$first))
        $first = $false
    }
    foreach ($f in @($sel | Where-Object { $_.PSObject.Properties['CreateStatsColumn'] -and $_.CreateStatsColumn })) { [void]$sb.AppendLine((Get-ProcFindingScript $f)) }
    Add-ScriptText $sb.ToString()
}

function Add-ProcIndexScript {
    $a = $script:State.ProcAnalysis
    if (-not $a) { return }
    $sel = @(Get-SelectedProcFindings | Where-Object { $_.Kind -eq 'Index' -and $_.Action -in 'REBUILD', 'REORGANIZE' })
    if ($sel.Count -eq 0) { $sel = @($script:State.ProcFindings | Where-Object { $_.Kind -eq 'Index' -and $_.Action -in 'REBUILD', 'REORGANIZE' }) }
    if ($sel.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show("No index used by this procedure is recommended for REBUILD/REORGANIZE yet.`r`n`r`nOnly indexes the plan SCANS can benefit, and only once fragmentation/page density has been measured: use 'Check fragmentation of scanned indexes' first.", 'Index maintenance', 'OK', 'Information')
        return
    }
    $caps = $script:State.ServerInfo.Capabilities
    $isProd = Test-IsProductionEnvironment $script:State.Environment
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("-- Index maintenance generated from the plan of $($script:State.ProcName) ($($a.Source.Source)). Text only. Only indexes this procedure scans are included.")
    $first = $true
    foreach ($g in @($sel | Group-Object { "$($_.TableObjectId)|$($_.Action)" })) {
        $f0 = $g.Group[0]
        $t = Get-ProcTable ([int]$f0.TableObjectId)
        if (-not $t) { continue }
        $opts = @{
            Action = $f0.Action; Online = [bool]$caps.OnlineRebuild; WaitAtLowPriority = [bool]$caps.WaitAtLowPriority; LowPriorityMinutes = [int]$script:Settings.LowPriorityWaitMinutes
            Resumable = ($isProd -and [bool]$caps.ResumableRebuild); MaxDurationMin = [int]$script:Settings.ResumableMaxMinutes; SortInTempdb = $false
            MaxDop = [int]$script:Settings.DefaultMaxDop; Partition = 'ALL'; LobCompaction = $true; CompressAllRowGroups = $true
        }
        [void]$sb.AppendLine((New-IndexMaintenanceScript -SchemaName $t.SchemaName -TableName $t.TableName -Indexes @($g.Group | ForEach-Object { $_._Index }) -Options $opts -ServerInfo $script:State.ServerInfo -DatabaseInfo $script:State.DbInfo -Environment $script:State.Environment -TableReservedMB ([double]$t.ReservedMB) -IncludeHeader:$first))
        $first = $false
    }
    Add-ScriptText $sb.ToString()
    Set-Status 'Index script added with production-safe defaults (ONLINE / low-priority wait / resumable where available). For other options use the Indexes tab dialog.'
}

function Invoke-ProcFragmentationCheck {
    $a = $script:State.ProcAnalysis
    if (-not $a) { return }
    $items = @{}
    $pool = @(Get-SelectedProcFindings | Where-Object { $_.Kind -eq 'Index' })
    if ($pool.Count -eq 0) { $pool = @($script:State.ProcFindings | Where-Object { $_.Kind -eq 'Index' }) }
    foreach ($f in $pool) {
        if (-not ($f.PSObject.Properties['_Index'] -and $f._Index)) { continue }
        if ($f.Priority -eq 'Review' -and $f.Action -notlike 'Check*') { continue }
        $ix = $f._Index
        if ($ix.TypeCode -notin 0, 1, 2 -or $ix.IsDisabled -or $ix.IsHypothetical) { continue }
        $items["$($f.TableObjectId)|$($ix.IndexId)"] = [pscustomobject]@{ ObjectId = [int]$f.TableObjectId; Index = $ix }
    }
    if ($items.Count -eq 0) { Set-Status 'This plan scans no rowstore index large enough to matter.'; return }
    $isProd = Test-IsProductionEnvironment $script:State.Environment
    $cap = if ($isProd) { [double]$script:Settings.MaxScanGBProduction } else { [double]$script:Settings.MaxScanGBNonProduction }
    $list = @($items.Values)
    $over = @($list | Where-Object { [double]$_.Index.SizeMB / 1024 -gt $cap }).Count
    $gb = (($list | Where-Object { [double]$_.Index.SizeMB / 1024 -le $cap }) | ForEach-Object { [double]$_.Index.SizeMB } | Measure-Object -Sum).Sum / 1024
    $msg = ("Measure fragmentation (LIMITED mode, MAXDOP 1, one index at a time) for the {0} index(es) this procedure scans?`r`n`r`nSize to scan: {1:N1} GB. Skipped above the {2} GB cap for {3}: {4}.`r`nRefused automatically if the server looks busy. Cancel stops it at any time." -f ($list.Count - $over), $gb, $cap, $script:State.Environment, $over)
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Check fragmentation', 'YesNo', 'Question') -ne 'Yes') { return }
    $work = @($list | ForEach-Object { $x = $_; $o = [ordered]@{}; foreach ($p in $x.Index.PSObject.Properties) { if ($p.Name -notlike '_*') { $o[$p.Name] = $p.Value } }; [pscustomobject]@{ ObjectId = $x.ObjectId; Index = [pscustomobject]$o } })
    [void](Start-AnalyzerJob -Status "Fragmentation check for $($work.Count) index(es) used by $($script:State.ProcName)" -Arguments @{ Ctx = $script:State.DbContext; Items = $work; Env = $script:State.Environment; Settings = $script:Settings; Shared = $script:Shared } -Work {
        param($Ctx, $Items, $Env, $Settings, $Shared)
        $ErrorActionPreference = 'Stop'
        $results = @{}
        foreach ($it in $Items) {
            if ($Shared.Cancel) { break }
            $key = "$($it.ObjectId)|$($it.Index.IndexId)"
            try { $results[$key] = Get-IndexFragmentation -Context $Ctx -ObjectId $it.ObjectId -Index $it.Index -Mode LIMITED -Environment $Env -Settings $Settings }
            catch {
                if ($Shared.Cancel) { break }
                $results[$key] = [pscustomobject]@{ IndexId = $it.Index.IndexId; Mode = 'LIMITED'; Measured = $false; Skipped = "Scan failed: $($_.Exception.Message)"; FragmentationPct = $null; PageDensityPct = $null; PageCount = $null; ForwardedRecords = $null; MeasuredAt = $null }
            }
        }
        @{ Results = $results }
    } -OnDone {
        param($r)
        $db = [string]$script:State.DbInfo.DatabaseName
        foreach ($k in $r.Results.Keys) {
            $res = $r.Results[$k]; $fk = "$db|$k"
            $prev = $script:State.Frag[$fk]
            if ($res.Measured -and $null -eq $res.PageDensityPct -and $prev -and $prev.Measured -and $null -ne $prev.PageDensityPct) { $res.PageDensityPct = $prev.PageDensityPct }
            if (-not $res.Measured -and $prev -and $prev.Measured) { continue }
            $script:State.Frag[$fk] = $res
        }
        Update-ProcFindings
        if ($script:State.Analysis) { Update-Assessments }
    })
}

function Save-ProcPlan {
    $a = $script:State.ProcAnalysis
    if (-not $a -or @($a.PlanTexts).Count -eq 0) { Set-Status 'No plan loaded.'; return }
    $doc = 0
    if (@($a.PlanTexts).Count -gt 1 -and $script:gridProcStmts.CurrentRow) { $doc = [int]$script:gridProcStmts.CurrentRow.Cells['Doc'].Value }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Execution plan (*.sqlplan)|*.sqlplan|XML (*.xml)|*.xml'
    $sfd.FileName = (($script:State.ProcName -replace '[\[\]\\/:*?"<>|]', '') + $(if (@($a.PlanTexts).Count -gt 1) { "_stmt$doc" } else { '' }) + '.sqlplan')
    if ($sfd.ShowDialog($script:form) -eq 'OK') {
        [System.IO.File]::WriteAllText($sfd.FileName, [string]@($a.PlanTexts)[$doc], [System.Text.Encoding]::Unicode)
        Set-Status "Saved $($sfd.FileName) (opens in SSMS / Plan Explorer)."
    }
}
#endregion

#region ------------------------------------------------------------------ maintenance dialogs (generate text only)
function Add-ScriptText {
    param([string]$Text)
    if ($script:txtScript.TextLength -gt 0) { $script:txtScript.AppendText("`r`n`r`n") }
    $script:txtScript.AppendText($Text)
    $script:tabs.SelectedTab = $script:tabScript
    Set-Status 'Script added to the Script tab. Review it there. It has NOT been executed.'
}

function Show-IndexScriptDialog {
    if (-not $script:State.Analysis) { return }
    $sel = @(Get-SelectedRows -Grid $script:gridIdx -Rows $script:State.IndexRows -IdColumn 'IndexId' -IdProperty 'IndexId')
    if ($sel.Count -eq 0) { Set-Status 'Select one or more indexes first.'; return }
    $caps = $script:State.ServerInfo.Capabilities
    $isProd = Test-IsProductionEnvironment $script:State.Environment
    $o = $script:State.CurrentTable

    $dlg = New-Ctl Form @{ Text = "Generate index maintenance script ($($sel.Count) index(es)): TEXT ONLY"; Width = 980; Height = 720; StartPosition = 'CenterParent'; Font = $script:UiFont; MinimizeBox = $false }
    $left = New-Ctl Panel @{ Dock = 'Left'; Width = 330; Padding = (New-Object System.Windows.Forms.Padding(8)) }
    $preview = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Both'; WordWrap = $false; ReadOnly = $true; Font = $script:Mono }
    $y = 8
    $script:dy = $y
    $rbRebuild = New-Ctl RadioButton @{ Text = 'REBUILD'; Left = 8; Top = $script:dy; AutoSize = $true }
    $rbReorg   = New-Ctl RadioButton @{ Text = 'REORGANIZE (always online, interruptible)'; Left = 90; Top = $script:dy; AutoSize = $true }
    $left.Controls.AddRange(@($rbRebuild, $rbReorg)); $script:dy += 30
    $recommendReorg = @($sel | Where-Object { $_.Action -eq 'REORGANIZE' }).Count -eq $sel.Count
    if ($recommendReorg) { $rbReorg.Checked = $true } else { $rbRebuild.Checked = $true }

    $chkOnline = New-Ctl CheckBox @{ Text = 'ONLINE = ON'; Left = 8; Top = $script:dy; AutoSize = $true; Checked = [bool]$caps.OnlineRebuild; Enabled = [bool]$caps.OnlineRebuild }; $script:dy += 24
    if (-not $caps.OnlineRebuild) { $chkOnline.Text = 'ONLINE = ON (needs Enterprise/Developer edition)' }
    $chkLow = New-Ctl CheckBox @{ Text = 'WAIT_AT_LOW_PRIORITY, abort self after (min):'; Left = 8; Top = $script:dy; AutoSize = $true; Checked = ($isProd -and $caps.WaitAtLowPriority); Enabled = [bool]$caps.WaitAtLowPriority }
    $numLow = New-Ctl NumericUpDown @{ Left = 270; Top = ($script:dy - 2); Width = 50; Minimum = 1; Maximum = 60; Value = [int]$script:Settings.LowPriorityWaitMinutes }; $script:dy += 26
    $chkRes = New-Ctl CheckBox @{ Text = 'RESUMABLE, MAX_DURATION (min):'; Left = 8; Top = $script:dy; AutoSize = $true; Checked = ($isProd -and $caps.ResumableRebuild); Enabled = [bool]$caps.ResumableRebuild }
    $numRes = New-Ctl NumericUpDown @{ Left = 270; Top = ($script:dy - 2); Width = 50; Minimum = 1; Maximum = 10080; Value = [int]$script:Settings.ResumableMaxMinutes }; $script:dy += 26
    $chkSort = New-Ctl CheckBox @{ Text = 'SORT_IN_TEMPDB = ON (needs tempdb space)'; Left = 8; Top = $script:dy; AutoSize = $true }; $script:dy += 26
    $lblDop = New-Ctl Label @{ Text = 'MAXDOP (0 = server default):'; Left = 8; Top = ($script:dy + 3); AutoSize = $true }
    $cpu = if ($script:State.ServerInfo.CpuCount) { [int]$script:State.ServerInfo.CpuCount } else { 8 }
    $numDop = New-Ctl NumericUpDown @{ Left = 270; Top = $script:dy; Width = 50; Minimum = 0; Maximum = 64; Value = [math]::Min([int]$script:Settings.DefaultMaxDop, [math]::Max(1, $cpu)) }; $script:dy += 30
    $lblPart = New-Ctl Label @{ Text = 'Partition:'; Left = 8; Top = ($script:dy + 3); AutoSize = $true }
    $cboPart = New-Ctl ComboBox @{ Left = 150; Top = $script:dy; Width = 170; DropDownStyle = 'DropDownList' }
    [void]$cboPart.Items.Add('ALL')
    $maxPart = [int]($sel | Measure-Object Partitions -Maximum).Maximum
    for ($p = 1; $p -le [math]::Min($maxPart, 15000); $p++) { if ($maxPart -gt 1) { [void]$cboPart.Items.Add([string]$p) } }
    $cboPart.SelectedIndex = 0; $script:dy += 30
    $chkLob = New-Ctl CheckBox @{ Text = 'LOB_COMPACTION = ON (reorganize)'; Left = 8; Top = $script:dy; AutoSize = $true; Checked = $true }; $script:dy += 24
    $chkCs = New-Ctl CheckBox @{ Text = 'COMPRESS_ALL_ROW_GROUPS (columnstore reorganize)'; Left = 8; Top = $script:dy; AutoSize = $true; Checked = $true; Enabled = [bool]$caps.CompressAllRowGroups }; $script:dy += 24
    $chkOla = New-Ctl CheckBox @{ Text = "Also add the Ola Hallengren IndexOptimize equivalent$(if ($script:State.DbInfo.OlaDatabase) { " (found in $($script:State.DbInfo.OlaDatabase))" })"; Left = 8; Top = $script:dy; Width = 310; Height = 34; Checked = [bool]$script:State.DbInfo.OlaDatabase }; $script:dy += 40
    $note = New-Ctl Label @{ Left = 8; Top = $script:dy; Width = 310; Height = 150; ForeColor = [System.Drawing.Color]::DarkRed
        Text = "This tool never executes maintenance. The script goes to the Script tab, with a server/database check$(if ($isProd) { ' and a PRODUCTION SAFETY LATCH that must be removed by hand after approval' }).`r`n`r`nOptions the server cannot do are removed or downgraded automatically, and the preview explains why." }
    $script:dy += 150
    $btnAdd = New-Ctl Button @{ Text = 'Add to Script tab'; Left = 8; Top = $script:dy; Width = 150; Height = 30 }
    $btnClose = New-Ctl Button @{ Text = 'Cancel'; Left = 170; Top = $script:dy; Width = 100; Height = 30; DialogResult = 'Cancel' }
    $left.Controls.AddRange(@($chkOnline, $chkLow, $numLow, $chkRes, $numRes, $chkSort, $lblDop, $numDop, $lblPart, $cboPart, $chkLob, $chkCs, $chkOla, $note, $btnAdd, $btnClose))
    $dlg.Controls.Add($preview); $dlg.Controls.Add($left)
    $dlg.CancelButton = $btnClose

    $getOpts = {
        @{
            Action = if ($rbReorg.Checked) { 'REORGANIZE' } else { 'REBUILD' }
            Online = $chkOnline.Checked; WaitAtLowPriority = $chkLow.Checked; LowPriorityMinutes = [int]$numLow.Value
            Resumable = $chkRes.Checked; MaxDurationMin = [int]$numRes.Value; SortInTempdb = $chkSort.Checked
            MaxDop = [int]$numDop.Value; Partition = [string]$cboPart.SelectedItem; LobCompaction = $chkLob.Checked; CompressAllRowGroups = $chkCs.Checked
        }
    }
    $build = {
        param([bool]$Header)
        $opts = & $getOpts
        $txt = New-IndexMaintenanceScript -SchemaName $o.SchemaName -TableName $o.TableName -Indexes $sel -Options $opts -ServerInfo $script:State.ServerInfo -DatabaseInfo $script:State.DbInfo -Environment $script:State.Environment -TableReservedMB ([double]$o.ReservedMB) -IncludeHeader:$Header
        if ($chkOla.Checked) {
            $names = @($sel | Where-Object { $_.TypeCode -ne 0 } | ForEach-Object { $_.IndexName })
            $txt += "`r`n" + (New-OlaIndexOptimizeCommand -DatabaseName $script:State.DbInfo.DatabaseName -SchemaName $o.SchemaName -TableName $o.TableName -ObjectNames $names -Kind Index -Options $opts -Settings $script:Settings -ServerInfo $script:State.ServerInfo -OlaDatabase $script:State.DbInfo.OlaDatabase)
        }
        $txt
    }
    $refresh = { try { $preview.Text = & $build $false } catch { $preview.Text = "Error: $($_.Exception.Message)" } }
    foreach ($c in @($rbRebuild, $rbReorg, $chkOnline, $chkLow, $chkRes, $chkSort, $chkLob, $chkCs, $chkOla)) { $c.add_CheckedChanged($refresh) }
    foreach ($c in @($numLow, $numRes, $numDop)) { $c.add_ValueChanged($refresh) }
    $cboPart.add_SelectedIndexChanged($refresh)
    $btnAdd.add_Click({ try { Add-ScriptText (& $build $true); $dlg.Close() } catch { Show-Error 'Generate script' $_ } })
    & $refresh
    [void]$dlg.ShowDialog($script:form)
    $dlg.Dispose()
}

function Show-StatsScriptDialog {
    if (-not $script:State.Analysis) { return }
    $sel = @(Get-SelectedRows -Grid $script:gridStats -Rows $script:State.StatsRows -IdColumn 'StatsId' -IdProperty 'StatsId')
    if ($sel.Count -eq 0) { Set-Status 'Select one or more statistics first.'; return }
    $caps = $script:State.ServerInfo.Capabilities
    $o = $script:State.CurrentTable
    $suggest = [int](($sel | ForEach-Object { $_._Assessment.SuggestedSamplePct } | Measure-Object -Maximum).Maximum)

    $dlg = New-Ctl Form @{ Text = "Generate UPDATE STATISTICS script ($($sel.Count) statistic(s)): TEXT ONLY"; Width = 980; Height = 640; StartPosition = 'CenterParent'; Font = $script:UiFont; MinimizeBox = $false }
    $left = New-Ctl Panel @{ Dock = 'Left'; Width = 330 }
    $preview = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Both'; WordWrap = $false; ReadOnly = $true; Font = $script:Mono }
    $rbFull = New-Ctl RadioButton @{ Text = 'FULLSCAN (most accurate, reads everything)'; Left = 8; Top = 10; AutoSize = $true }
    $rbSample = New-Ctl RadioButton @{ Text = 'SAMPLE  percent:'; Left = 8; Top = 36; AutoSize = $true }
    $numPct = New-Ctl NumericUpDown @{ Left = 150; Top = 34; Width = 70; Minimum = 0.01; Maximum = 100; DecimalPlaces = 2; Increment = 1; Value = [decimal][math]::Max(1, $suggest) }
    $rbRe = New-Ctl RadioButton @{ Text = 'RESAMPLE (reuse each statistic''s last rate)'; Left = 8; Top = 62; AutoSize = $true }
    if ($suggest -ge 100) { $rbFull.Checked = $true } else { $rbSample.Checked = $true }
    $hint = New-Ctl Label @{ Left = 24; Top = 86; Width = 300; Height = 34; ForeColor = [System.Drawing.Color]::DimGray; Text = "Suggested: $(if ($suggest -ge 100) { 'FULLSCAN' } else { "$suggest%" }) (about 10 million rows sampled; heuristic)" }
    $lblDop = New-Ctl Label @{ Text = 'MAXDOP (0 = server default):'; Left = 8; Top = 128; AutoSize = $true }
    $numDop = New-Ctl NumericUpDown @{ Left = 200; Top = 125; Width = 60; Minimum = 0; Maximum = 64; Value = [int]$script:Settings.DefaultMaxDop; Enabled = [bool]$caps.UpdateStatsMaxDop }
    if (-not $caps.UpdateStatsMaxDop) { $lblDop.Text = 'MAXDOP (not supported on this build)' }
    $chkPersist = New-Ctl CheckBox @{ Text = 'PERSIST_SAMPLE_PERCENT = ON (auto-updates keep this rate)'; Left = 8; Top = 158; Width = 320; Height = 34; Enabled = [bool]$caps.PersistSamplePercent }
    $lblParts = New-Ctl Label @{ Text = 'Incremental stats, RESAMPLE ON PARTITIONS (e.g. 3,4):'; Left = 8; Top = 198; Width = 320 }
    $txtParts = New-Ctl TextBox @{ Left = 8; Top = 218; Width = 200 }
    $chkOla = New-Ctl CheckBox @{ Text = 'Also add the Ola Hallengren IndexOptimize equivalent'; Left = 8; Top = 250; Width = 320; Checked = [bool]$script:State.DbInfo.OlaDatabase }
    $note = New-Ctl Label @{ Left = 8; Top = 282; Width = 310; Height = 140; ForeColor = [System.Drawing.Color]::DarkRed
        Text = "Never executed by this tool. Statistics with NORECOMPUTE keep it (otherwise the update would silently re-enable automatic updates).`r`n`r`nStatistics updates take only a schema-stability lock, but they read the data: size them to your I/O budget." }
    $btnAdd = New-Ctl Button @{ Text = 'Add to Script tab'; Left = 8; Top = 430; Width = 150; Height = 30 }
    $btnClose = New-Ctl Button @{ Text = 'Cancel'; Left = 170; Top = 430; Width = 100; Height = 30; DialogResult = 'Cancel' }
    $left.Controls.AddRange(@($rbFull, $rbSample, $numPct, $rbRe, $hint, $lblDop, $numDop, $chkPersist, $lblParts, $txtParts, $chkOla, $note, $btnAdd, $btnClose))
    $dlg.Controls.Add($preview); $dlg.Controls.Add($left); $dlg.CancelButton = $btnClose

    $getOpts = {
        $parts = @($txtParts.Text -split '[,\s;]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
        @{
            Method = if ($rbFull.Checked) { 'FULLSCAN' } elseif ($rbRe.Checked) { 'RESAMPLE' } else { 'SAMPLE' }
            SamplePercent = [double]$numPct.Value
            MaxDop = if ($numDop.Enabled) { [int]$numDop.Value } else { $null }
            PersistSamplePercent = $chkPersist.Checked
            Partitions = $parts
        }
    }
    $build = {
        param([bool]$Header)
        $opts = & $getOpts
        $txt = New-StatsMaintenanceScript -SchemaName $o.SchemaName -TableName $o.TableName -Statistics $sel -Options $opts -ServerInfo $script:State.ServerInfo -DatabaseInfo $script:State.DbInfo -Environment $script:State.Environment -TableDataMB ([double]$o.DataMB) -IncludeHeader:$Header
        if ($chkOla.Checked) {
            $txt += "`r`n" + (New-OlaIndexOptimizeCommand -DatabaseName $script:State.DbInfo.DatabaseName -SchemaName $o.SchemaName -TableName $o.TableName -ObjectNames @($sel | ForEach-Object { $_.StatName }) -Kind Statistics -Options $opts -Settings $script:Settings -ServerInfo $script:State.ServerInfo -OlaDatabase $script:State.DbInfo.OlaDatabase)
        }
        $txt
    }
    $refresh = { try { $preview.Text = & $build $false } catch { $preview.Text = "Error: $($_.Exception.Message)" } }
    foreach ($c in @($rbFull, $rbSample, $rbRe, $chkPersist, $chkOla)) { $c.add_CheckedChanged($refresh) }
    foreach ($c in @($numPct, $numDop)) { $c.add_ValueChanged($refresh) }
    $txtParts.add_TextChanged($refresh)
    $btnAdd.add_Click({ try { Add-ScriptText (& $build $true); $dlg.Close() } catch { Show-Error 'Generate script' $_ } })
    & $refresh
    [void]$dlg.ShowDialog($script:form)
    $dlg.Dispose()
}

function Show-ServerListDialog {
    $dlg = New-Ctl Form @{ Text = 'Server list: central inventory and local cache'; Width = 820; Height = 620; StartPosition = 'CenterParent'; Font = $script:UiFont }
    $top = New-Ctl Panel @{ Dock = 'Top'; Height = 190 }
    $top.Controls.Add((New-Ctl Label @{ Text = 'Central server:'; Left = 10; Top = 12; AutoSize = $true }))
    $txtCs = New-Ctl TextBox @{ Left = 130; Top = 9; Width = 250; Text = [string]$script:Settings.CentralServer }
    $top.Controls.Add((New-Ctl Label @{ Text = 'Database:'; Left = 400; Top = 12; AutoSize = $true }))
    $txtCd = New-Ctl TextBox @{ Left = 470; Top = 9; Width = 160; Text = [string]$script:Settings.CentralDatabase }
    $top.Controls.Add((New-Ctl Label @{ Text = 'Query (SELECT only; columns: ServerName/SqlInstance/Name [, Environment]):'; Left = 10; Top = 42; AutoSize = $true }))
    $txtCq = New-Ctl TextBox @{ Left = 10; Top = 62; Width = 780; Height = 70; Multiline = $true; ScrollBars = 'Vertical'; Font = $script:Mono; Text = [string]$script:Settings.CentralQuery }
    $btnLoad = New-Ctl Button @{ Text = 'Load from central server'; Left = 10; Top = 142; Width = 190; Height = 30 }
    $btnFile = New-Ctl Button @{ Text = 'Open settings file (thresholds, caps)'; Left = 210; Top = 142; Width = 240; Height = 30 }
    $lblMsg = New-Ctl Label @{ Left = 460; Top = 148; Width = 340; Height = 30; ForeColor = [System.Drawing.Color]::DimGray; Text = 'Edit Environment to tag servers. Anything not tagged non-production is treated as Production.' }
    $top.Controls.AddRange(@($txtCs, $txtCd, $txtCq, $btnLoad, $btnFile, $lblMsg))
    $grid = New-Grid; $grid.ReadOnly = $false; $grid.AllowUserToAddRows = $true; $grid.AllowUserToDeleteRows = $true; $grid.MultiSelect = $true
    $dt = New-Object System.Data.DataTable; [void]$dt.Columns.Add('Name', [string]); [void]$dt.Columns.Add('Environment', [string])
    foreach ($s in $script:Settings.Servers) { if ($s -and $s.Name) { [void]$dt.Rows.Add([string]$s.Name, [string]$s.Environment) } }
    $grid.DataSource = $dt
    $bottom = New-Ctl Panel @{ Dock = 'Bottom'; Height = 46 }
    $btnSave = New-Ctl Button @{ Text = 'Save'; Left = 10; Top = 8; Width = 100; Height = 30 }
    $btnCancel = New-Ctl Button @{ Text = 'Cancel'; Left = 120; Top = 8; Width = 100; Height = 30; DialogResult = 'Cancel' }
    $bottom.Controls.AddRange(@($btnSave, $btnCancel))
    $dlg.Controls.Add($grid); $dlg.Controls.Add($top); $dlg.Controls.Add($bottom); $dlg.CancelButton = $btnCancel
    $grid.add_DataBindingComplete({ param($s, $e) foreach ($c in $s.Columns) { $c.Width = 300 } })

    $btnLoad.add_Click({
        try {
            if (-not $txtCs.Text.Trim()) { throw 'Enter the central server name.' }
            $r = Test-ReadOnlySql -Sql $txtCq.Text
            if (-not $r.IsSafe) { throw "The query was refused by the read-only guard: $($r.Reason)" }
            $dlg.Cursor = 'WaitCursor'
            $cdb = if ($txtCd.Text.Trim()) { $txtCd.Text.Trim() } else { 'master' }
            $ctx = Get-UiContext -DatabaseName $cdb -ServerName $txtCs.Text
            $list = @(Get-CentralServerList -Context $ctx -Query $txtCq.Text)
            Sync-QueryLog
            $existing = @{}; foreach ($row in $dt.Rows) { $existing[[string]$row['Name']] = $row }
            $added = 0
            foreach ($s in $list) {
                if ($existing.ContainsKey($s.Name)) { if ($s.Environment) { $existing[$s.Name]['Environment'] = [string]$s.Environment } }
                else { [void]$dt.Rows.Add([string]$s.Name, [string]$s.Environment); $added++ }
            }
            $lblMsg.Text = "Loaded $($list.Count) servers ($added new). Click Save to keep them."
        } catch { Sync-QueryLog; Show-Error 'Load server list' $_ } finally { $dlg.Cursor = 'Default' }
    })
    $btnFile.add_Click({ Save-AnalyzerSettings -Settings $script:Settings; Start-Process notepad.exe -ArgumentList "`"$(Get-AnalyzerSettingsPath)`"" })
    $btnSave.add_Click({
        $grid.EndEdit()
        $script:Settings.CentralServer = $txtCs.Text.Trim()
        $script:Settings.CentralDatabase = $txtCd.Text.Trim()
        $script:Settings.CentralQuery = $txtCq.Text
        $script:Settings.Servers = @(foreach ($row in $dt.Rows) { if ($row.RowState -ne 'Deleted' -and [string]$row['Name']) { [pscustomobject]@{ Name = ([string]$row['Name']).Trim(); Environment = [string]$row['Environment'] } } })
        Save-AnalyzerSettings -Settings $script:Settings
        Update-ServerCombo
        if ($script:State.Context) { $script:State.Environment = Get-ServerEnvironment -ServerName $script:State.Context.Server -Settings $script:Settings }
        Update-EnvironmentBanner
        $dlg.Close()
    })
    [void]$dlg.ShowDialog($script:form)
    $dlg.Dispose()
}
#endregion

#region ------------------------------------------------------------------ build the main window
$script:form = New-Ctl Form @{ Text = "SQL Server Index & Statistics Analyzer (READ-ONLY)"; Width = 1500; Height = 950; StartPosition = 'CenterScreen'; Font = $script:UiFont; KeyPreview = $true }

# --- connection bar
$bar = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 38; Padding = (New-Object System.Windows.Forms.Padding(6, 6, 6, 0)); WrapContents = $false }
$bar.Controls.Add((New-Ctl Label @{ Text = 'Server'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(0, 6, 4, 0)) }))
$script:cboServer = New-Ctl ComboBox @{ Width = 250; DropDownStyle = 'DropDown'; AutoCompleteMode = 'SuggestAppend'; AutoCompleteSource = 'ListItems' }
$btnServers = New-Ctl Button @{ Text = 'Server list...'; Width = 95; Height = 25 }
$script:cboAuth = New-Ctl ComboBox @{ Width = 105; DropDownStyle = 'DropDownList' }
[void]$script:cboAuth.Items.AddRange(@('Windows', 'SQL login')); $script:cboAuth.SelectedIndex = 0
$script:txtUser = New-Ctl TextBox @{ Width = 110; Enabled = $false }
$script:txtPwd = New-Ctl TextBox @{ Width = 110; UseSystemPasswordChar = $true; Enabled = $false }
$script:chkEncrypt = New-Ctl CheckBox @{ Text = 'Encrypt'; AutoSize = $true; Checked = [bool]$script:Settings.Encrypt; Margin = (New-Object System.Windows.Forms.Padding(6, 5, 0, 0)) }
$script:chkReadOnlyIntent = New-Ctl CheckBox @{ Text = 'ReadOnly intent'; AutoSize = $true; Checked = [bool]$script:Settings.ReadOnlyIntent; Margin = (New-Object System.Windows.Forms.Padding(6, 5, 0, 0)) }
$script:btnConnect = New-Ctl Button @{ Text = 'Connect'; Width = 80; Height = 25; BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255) }
$lblDb = New-Ctl Label @{ Text = 'Database'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(14, 6, 4, 0)) }
$script:cboDb = New-Ctl ComboBox @{ Width = 330; DropDownStyle = 'DropDownList' }
$script:btnCancel = New-Ctl Button @{ Text = 'Cancel'; Width = 70; Height = 25; Enabled = $false }
$bar.Controls.AddRange(@($script:cboServer, $btnServers, $script:cboAuth, $script:txtUser, $script:txtPwd, $script:chkEncrypt, $script:chkReadOnlyIntent, $script:btnConnect, $lblDb, $script:cboDb, $script:btnCancel))
$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($script:txtUser, 'SQL login name'); $tip.SetToolTip($script:txtPwd, 'Password (kept in memory only, never saved)')
$tip.SetToolTip($script:chkReadOnlyIntent, 'ApplicationIntent=ReadOnly: through an AG listener with read-only routing you land on a readable secondary. Usage stats are then the secondary''s own.')

$script:lblEnv = New-Ctl Label @{ Dock = 'Top'; Height = 24; ForeColor = [System.Drawing.Color]::White; Font = $script:BoldFont; TextAlign = 'MiddleLeft' }
$script:lblServerInfo = New-Ctl Label @{ Dock = 'Top'; Height = 20; ForeColor = [System.Drawing.Color]::DimGray; TextAlign = 'MiddleLeft'; Padding = (New-Object System.Windows.Forms.Padding(6, 0, 0, 0)) }
$script:lblWarn = New-Ctl Label @{ Dock = 'Top'; Height = 20; ForeColor = [System.Drawing.Color]::DarkRed; TextAlign = 'MiddleLeft'; Padding = (New-Object System.Windows.Forms.Padding(6, 0, 0, 0)) }

# --- status bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblStatus.Spring = $true; $script:lblStatus.TextAlign = 'MiddleLeft'; $script:lblStatus.Text = 'Ready. Type or pick a server and click Connect.'
$script:progress = New-Object System.Windows.Forms.ToolStripProgressBar
$script:progress.Visible = $false; $script:progress.Width = 150
[void]$statusStrip.Items.AddRange(@($script:lblStatus, $script:progress))

# --- main split: tables | analysis
$split = New-Ctl SplitContainer @{ Dock = 'Fill'; FixedPanel = 'Panel1' }
$script:lblDbInfo = New-Ctl Label @{ Dock = 'Top'; Height = 36; ForeColor = [System.Drawing.Color]::DimGray; Text = 'Connect, then pick a database.' }
$script:txtSearch = New-Ctl TextBox @{ Dock = 'Top'; Font = (New-Object System.Drawing.Font('Segoe UI', 11)) }
$script:lblTableCount = New-Ctl Label @{ Dock = 'Top'; Height = 34; ForeColor = [System.Drawing.Color]::DimGray }
$script:gridTables = New-Grid; $script:gridTables.MultiSelect = $false
$script:btnRefreshTable = New-Ctl Button @{ Dock = 'Bottom'; Height = 28; Text = 'Reload selected table (F5)' }
$split.Panel1.Controls.Add($script:gridTables)
$split.Panel1.Controls.Add($script:btnRefreshTable)
$split.Panel1.Controls.Add($script:lblTableCount)
$split.Panel1.Controls.Add($script:txtSearch)
$split.Panel1.Controls.Add((New-Ctl Label @{ Dock = 'Top'; Height = 20; Text = 'Find table (partial names, typos, abbreviations, * wildcards, schema.table):' }))
$split.Panel1.Controls.Add($script:lblDbInfo)

$script:lblTable = New-Ctl Label @{ Dock = 'Top'; Height = 26; Font = (New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)); Text = 'No table selected' }
$script:tabs = New-Ctl TabControl @{ Dock = 'Fill' }
$split.Panel2.Controls.Add($script:tabs)
$split.Panel2.Controls.Add($script:lblTable)

# Overview tab
$tabOverview = New-Ctl TabPage @{ Text = 'Overview & priorities' }
$script:txtOverview = New-Ctl TextBox @{ Dock = 'Top'; Height = 230; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$script:gridPriority = New-Grid
Add-PriorityColoring $script:gridPriority
$tabOverview.Controls.Add($script:gridPriority)
$tabOverview.Controls.Add((New-Ctl Label @{ Dock = 'Top'; Height = 22; Font = $script:BoldFont; Text = 'What to do first (highest priority first). Double-click a row to jump to it.' }))
$tabOverview.Controls.Add($script:txtOverview)

# Indexes tab
$script:tabIdx = New-Ctl TabPage @{ Text = 'Indexes' }
$idxTools = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 36; Padding = (New-Object System.Windows.Forms.Padding(4)) }
$script:btnFragLimited = New-Ctl Button @{ Text = 'Check fragmentation (LIMITED)'; Width = 210; Height = 26 }
$script:btnFragSampled = New-Ctl Button @{ Text = 'Check page density (SAMPLED, small indexes)'; Width = 290; Height = 26 }
$btnIdxScript = New-Ctl Button @{ Text = 'Generate rebuild/reorganize script...'; Width = 250; Height = 26 }
$idxTools.Controls.AddRange(@($script:btnFragLimited, $script:btnFragSampled, $btnIdxScript))
$tip.SetToolTip($script:btnFragLimited, 'Reads only the level above the leaf pages (cheapest). Size-capped, refused when the server is busy, MAXDOP 1, cancellable.')
$tip.SetToolTip($script:btnFragSampled, 'Reads about 1% of leaf pages to get page fullness. Only for indexes under the SAMPLED size cap (Settings: MaxSampledScanGB).')
$idxSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$script:gridIdx = New-Grid
Add-PriorityColoring $script:gridIdx
$script:txtIdxDetail = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$idxSplit.Panel1.Controls.Add($script:gridIdx); $idxSplit.Panel2.Controls.Add($script:txtIdxDetail)
$script:tabIdx.Controls.Add($idxSplit); $script:tabIdx.Controls.Add($idxTools)

# Statistics tab
$script:tabStats = New-Ctl TabPage @{ Text = 'Statistics' }
$statTools = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 36; Padding = (New-Object System.Windows.Forms.Padding(4)) }
$btnStatScript = New-Ctl Button @{ Text = 'Generate UPDATE STATISTICS script...'; Width = 260; Height = 26 }
$statTools.Controls.Add($btnStatScript)
$statTools.Controls.Add((New-Ctl Label @{ AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(10, 8, 0, 0)); ForeColor = [System.Drawing.Color]::DimGray; Text = 'Click a statistic to see its histogram and a plain-English explanation. Reading a histogram does not touch table data.' }))
$statSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$script:gridStats = New-Grid
Add-PriorityColoring $script:gridStats
$statSplit.Panel1.Controls.Add($script:gridStats)
$statBottom = New-Ctl SplitContainer @{ Dock = 'Fill' }
if ($script:ChartAvailable) {
    $script:chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $script:chart.Dock = 'Fill'
    $area = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 'main'
    $area.AxisX.LabelStyle.Angle = -60
    $area.AxisX.LabelStyle.Font = New-Object System.Drawing.Font('Segoe UI', 7)
    $area.AxisX.MajorGrid.Enabled = $false
    $area.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::Gainsboro
    $area.AxisY.LabelStyle.Format = 'N0'
    $area.AxisY.Title = 'rows'
    [void]$script:chart.ChartAreas.Add($area)
    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend 'lg'
    $legend.Docking = 'Top'
    [void]$script:chart.Legends.Add($legend)
    $sEq = New-Object System.Windows.Forms.DataVisualization.Charting.Series 'eq'
    $sEq.ChartType = 'StackedColumn'; $sEq.Color = [System.Drawing.Color]::FromArgb(31, 78, 121); $sEq.LegendText = 'Rows exactly equal to the bucket value (EQ_ROWS)'
    $sRg = New-Object System.Windows.Forms.DataVisualization.Charting.Series 'range'
    $sRg.ChartType = 'StackedColumn'; $sRg.Color = [System.Drawing.Color]::FromArgb(157, 195, 230); $sRg.LegendText = 'Rows between buckets (RANGE_ROWS)'
    [void]$script:chart.Series.Add($sEq); [void]$script:chart.Series.Add($sRg)
    $statBottom.Panel1.Controls.Add($script:chart)
} else {
    $script:chart = $null
    $statBottom.Panel1.Controls.Add((New-Ctl Label @{ Dock = 'Fill'; Text = 'Chart control not available on this machine (System.Windows.Forms.DataVisualization). See the Histogram data tab.' }))
}
$statInfoTabs = New-Ctl TabControl @{ Dock = 'Fill' }
$tpExplain = New-Ctl TabPage @{ Text = 'Explanation' }
$script:txtStatDetail = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; WordWrap = $true; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$tpExplain.Controls.Add($script:txtStatDetail)
$tpHist = New-Ctl TabPage @{ Text = 'Histogram data' }; $script:gridHist = New-Grid; $tpHist.Controls.Add($script:gridHist)
$tpDens = New-Ctl TabPage @{ Text = 'Density vector' }; $script:gridDensity = New-Grid; $tpDens.Controls.Add($script:gridDensity)
$statInfoTabs.TabPages.AddRange(@($tpExplain, $tpHist, $tpDens))
$statBottom.Panel2.Controls.Add($statInfoTabs)
$statSplit.Panel2.Controls.Add($statBottom)
$script:tabStats.Controls.Add($statSplit); $script:tabStats.Controls.Add($statTools)

# Missing indexes tab
$script:tabMissing = New-Ctl TabPage @{ Text = 'Missing indexes' }
$missSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$script:gridMissing = New-Grid; $script:gridMissing.MultiSelect = $false
$script:txtMissingScript = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono }
$btnMissAdd = New-Ctl Button @{ Dock = 'Bottom'; Height = 28; Text = 'Add this suggestion (commented out) to the Script tab' }
$missSplit.Panel1.Controls.Add($script:gridMissing)
$missSplit.Panel2.Controls.Add($script:txtMissingScript); $missSplit.Panel2.Controls.Add($btnMissAdd)
$script:tabMissing.Controls.Add($missSplit)
$script:tabMissing.Controls.Add((New-Ctl Label @{ Dock = 'Top'; Height = 34; ForeColor = [System.Drawing.Color]::DimGray; Text = 'Requests the optimizer logged since startup (sys.dm_db_missing_index_*). They ignore existing indexes and write cost: compare with the Indexes tab before adding anything.' }))

# Script tab
$script:tabScript = New-Ctl TabPage @{ Text = 'Script (text only)' }
$scriptTools = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 36; Padding = (New-Object System.Windows.Forms.Padding(4)) }
$btnCopy = New-Ctl Button @{ Text = 'Copy'; Width = 80; Height = 26 }
$btnSave = New-Ctl Button @{ Text = 'Save as .sql...'; Width = 110; Height = 26 }
$btnClear = New-Ctl Button @{ Text = 'Clear'; Width = 80; Height = 26 }
$scriptTools.Controls.AddRange(@($btnCopy, $btnSave, $btnClear))
$scriptTools.Controls.Add((New-Ctl Label @{ AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(10, 8, 0, 0)); ForeColor = [System.Drawing.Color]::DarkRed; Font = $script:BoldFont; Text = 'This tool cannot execute scripts. Review, get approval, and run them yourself in a maintenance window.' }))
$script:txtScript = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono; MaxLength = 0 }
$script:tabScript.Controls.Add($script:txtScript); $script:tabScript.Controls.Add($scriptTools)

# Query log tab
$script:tabLog = New-Ctl TabPage @{ Text = 'Query log' }
$script:LogTable = New-Object System.Data.DataTable
foreach ($c in @(@('Time', [datetime]), @('Server', [string]), @('Database', [string]), @('What', [string]), @('Ms', [double]), @('Rows', [double]), @('Status', [string]), @('Sql', [string]))) { [void]$script:LogTable.Columns.Add($c[0], $c[1]) }
$logSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$gridLog = New-Grid; $gridLog.MultiSelect = $false; $gridLog.DataSource = $script:LogTable
$txtLogSql = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono }
$logSplit.Panel1.Controls.Add($gridLog); $logSplit.Panel2.Controls.Add($txtLogSql)
$script:tabLog.Controls.Add($logSplit)
$script:tabLog.Controls.Add((New-Ctl Label @{ Dock = 'Top'; Height = 22; ForeColor = [System.Drawing.Color]::DimGray; Text = 'Every statement this tool sent, with duration. All of them passed the read-only guard first.' }))

# Help tab
$tabHelp = New-Ctl TabPage @{ Text = 'Help' }
$txtHelp = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$txtHelp.Text = @'
SAFETY (why this tool is safe on production)
  * Read-only: every statement passes a guard that only allows SELECT, harmless SET options and
    DBCC SHOW_STATISTICS. Everything else (INSERT/UPDATE/DELETE, ALTER, CREATE, EXEC, UPDATE STATISTICS,
    REBUILD, SELECT INTO, OPENROWSET, xp_ and sp_ calls, and so on) is refused before connecting.
  * Polite: LOCK_TIMEOUT 3s, DEADLOCK_PRIORITY LOW, READ UNCOMMITTED, MAXDOP 1 on heavier reads,
    one query at a time, cancellable. Row counts and sizes come from metadata; no COUNT(*), no table data.
  * Fragmentation scans are on demand only: LIMITED or SAMPLED (never DETAILED), one index at a time,
    refused above a size cap (20 GB on production by default) and refused while the server is busy.
  * Maintenance is generated as TEXT. Scripts check @@SERVERNAME and DB_NAME(), check the AG primary,
    and on production start with a safety latch you must delete by hand after approval.
  * Unknown servers are treated as Production. Tag servers in "Server list..." to change that.

WORDS IN PLAIN ENGLISH
  Index ............ a sorted copy of some columns so SQL Server can find rows without reading everything.
  Clustered index .. the table itself, stored in key order. A table without one is a HEAP.
  Seek ............. jumped straight to the rows needed (good). Scan: read through a large part.
  Lookup ........... found the row in an index, then went to the table for more columns.
  Fragmentation .... pages out of order, like a book with shuffled chapters. Hurts big scans more than seeks.
  Page density ..... how full the 8 KB pages are. Half-empty pages = more memory and I/O for the same data.
  REORGANIZE ....... tidies pages in place. Always online, single-threaded, can stop any time.
  REBUILD .......... makes a fresh copy. Fixes everything and refreshes that index's statistics with FULLSCAN.
                     ONLINE needs Enterprise; OFFLINE blocks the table.
  Statistics ....... a summary of the values in a column. The optimizer uses it to guess row counts,
                     and the guess decides the plan.
  Histogram ........ up to 200 buckets describing how rows spread over values (the chart).
  Density .......... 1 / number of distinct values. Low density = very unique = good for seeks.
  Sample % ......... how much of the table was read to build the statistics. 100% = FULLSCAN.
  Modifications .... rows changed since the statistics were built. Past the threshold, SQL Server
                     refreshes it on the next compile that needs it.

PROCEDURE ANALYSIS (opt-in tab)
  Nothing runs until you click. "Load procedures" lists procedures (metadata only). "Analyze" reads the
  plan the procedure ALREADY has: sys.dm_exec_procedure_stats + sys.dm_exec_text_query_plan, or the latest
  plan per statement from Query Store. The procedure is never executed, compiled or recompiled.
  From the plan it lists, per statement:
  * the statistics the optimizer actually used (SQL 2017 / 2016 SP2+ record them in the plan) with their
    current freshness and sample, boosted when a statement's row estimate is far from its real row counts;
  * the indexes it seeks, scans and looks up. Rebuild/reorganize is only suggested for indexes the plan
    SCANS, and only after fragmentation is measured (button). Seek-only indexes are reported as "no action";
  * columns with no statistics, implicit conversions, tempdb spills, missing join predicates,
    missing-index hints, key lookups, and parameter sensitivity (very different row counts per execution).
  Scripts go to the Script tab as text. "Save as .sqlplan" opens the plan in SSMS or Plan Explorer.
  If the procedure has no cached plan and Query Store has none, run it normally once; the tool won't.

PRIORITY
  High / Medium / Low are scores (0-100) built from: changes vs the auto-update threshold, sample
  size, age, fragmentation, page density, size, scan-heavy use, page splits and blocking. "Review" means
  the index is unused, duplicate, disabled or hypothetical: decide on dropping, not maintaining.
  These are heuristics, not orders. Check them against what the workload actually does.
'@
$tabHelp.Controls.Add($txtHelp)

# Procedure analysis tab (opt-in: nothing runs until a button is clicked)
$script:tabProc = New-Ctl TabPage @{ Text = 'Procedure analysis' }
$procTop = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 36; Padding = (New-Object System.Windows.Forms.Padding(4)); WrapContents = $false }
$script:btnProcLoad = New-Ctl Button @{ Text = 'Load procedures'; Width = 130; Height = 26 }
$script:txtProcSearch = New-Ctl TextBox @{ Width = 240; Margin = (New-Object System.Windows.Forms.Padding(8, 4, 0, 0)) }
$script:chkProcForTable = New-Ctl CheckBox @{ Text = 'Only procedures that use the selected table'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(10, 6, 0, 0)) }
$procTop.Controls.AddRange(@($script:btnProcLoad, $script:txtProcSearch, $script:chkProcForTable))
$tip.SetToolTip($script:txtProcSearch, 'Search procedures: partial names, typos, abbreviations, * wildcards, schema.name')
$script:lblProcCount = New-Ctl Label @{ Dock = 'Top'; Height = 20; ForeColor = [System.Drawing.Color]::DimGray }
$procInfo = New-Ctl Label @{ Dock = 'Top'; Height = 34; ForeColor = [System.Drawing.Color]::DarkGreen
    Text = 'Opt-in. Reads the plan the procedure ALREADY has (plan cache, or Query Store). The procedure is never executed or compiled. From that plan it shows which statistics and indexes this procedure depends on, and whether updating or rebuilding them would help it.' }

$script:procSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; FixedPanel = 'Panel1' }
$script:gridProcs = New-Grid; $script:gridProcs.MultiSelect = $false
$script:btnProcAnalyze = New-Ctl Button @{ Dock = 'Bottom'; Height = 30; Text = 'Analyze selected procedure (reads its cached plan)'; BackColor = [System.Drawing.Color]::FromArgb(225, 238, 255) }
$script:procSplit.Panel1.Controls.Add($script:gridProcs); $script:procSplit.Panel1.Controls.Add($script:btnProcAnalyze)

$procTools = New-Ctl FlowLayoutPanel @{ Dock = 'Top'; Height = 66; Padding = (New-Object System.Windows.Forms.Padding(4)) }
$procTools.Controls.Add((New-Ctl Label @{ Text = 'Plan:'; AutoSize = $true; Margin = (New-Object System.Windows.Forms.Padding(0, 7, 2, 0)) }))
$script:cboProcPlan = New-Ctl ComboBox @{ Width = 640; DropDownStyle = 'DropDownList' }
$procTools.Controls.Add($script:cboProcPlan)
$procTools.SetFlowBreak($script:cboProcPlan, $true)
$script:btnProcFrag = New-Ctl Button @{ Text = 'Check fragmentation of scanned indexes'; Width = 260; Height = 26 }
$btnProcStats = New-Ctl Button @{ Text = 'Script statistics updates'; Width = 180; Height = 26 }
$btnProcIdx = New-Ctl Button @{ Text = 'Script index maintenance'; Width = 180; Height = 26 }
$btnProcAdd = New-Ctl Button @{ Text = 'Add selected finding''s script'; Width = 200; Height = 26 }
$procTools.Controls.AddRange(@($script:btnProcFrag, $btnProcStats, $btnProcIdx, $btnProcAdd))
$tip.SetToolTip($script:btnProcFrag, 'LIMITED-mode fragmentation scan (size-capped, refused when busy) for the indexes this procedure scans. Seeks are not affected by fragmentation.')
$tip.SetToolTip($btnProcStats, 'Selected Statistic rows, or all High/Medium ones. Goes to the Script tab as text.')
$tip.SetToolTip($btnProcIdx, 'Selected REBUILD/REORGANIZE rows, or all of them. Production-safe defaults. Text only.')
$script:txtProcSummary = New-Ctl TextBox @{ Dock = 'Top'; Height = 120; Multiline = $true; ReadOnly = $true; ScrollBars = 'Vertical'; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$script:procTabs = New-Ctl TabControl @{ Dock = 'Fill' }
$script:tabProcFind = New-Ctl TabPage @{ Text = 'Findings' }
$script:procFindSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$script:gridProcFindings = New-Grid
Add-PriorityColoring $script:gridProcFindings
$script:txtProcFinding = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$script:procFindSplit.Panel1.Controls.Add($script:gridProcFindings); $script:procFindSplit.Panel2.Controls.Add($script:txtProcFinding)
$script:tabProcFind.Controls.Add($script:procFindSplit)
$script:tabProcStmts = New-Ctl TabPage @{ Text = 'Statements' }
$script:procStmtSplit = New-Ctl SplitContainer @{ Dock = 'Fill'; Orientation = 'Horizontal' }
$script:gridProcStmts = New-Grid; $script:gridProcStmts.MultiSelect = $false
$script:txtProcStmt = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $true; Font = $script:Mono; BackColor = [System.Drawing.SystemColors]::Window }
$script:procStmtSplit.Panel1.Controls.Add($script:gridProcStmts); $script:procStmtSplit.Panel2.Controls.Add($script:txtProcStmt)
$script:tabProcStmts.Controls.Add($script:procStmtSplit)
$tabProcXml = New-Ctl TabPage @{ Text = 'Plan XML' }
$script:txtPlanXml = New-Ctl TextBox @{ Dock = 'Fill'; Multiline = $true; ReadOnly = $true; ScrollBars = 'Both'; WordWrap = $false; Font = $script:Mono; MaxLength = 0 }
$btnSavePlan = New-Ctl Button @{ Dock = 'Top'; Height = 28; Text = 'Save as .sqlplan (open in SSMS or Plan Explorer)' }
$tabProcXml.Controls.Add($script:txtPlanXml); $tabProcXml.Controls.Add($btnSavePlan)
$script:procTabs.TabPages.AddRange(@($script:tabProcFind, $script:tabProcStmts, $tabProcXml))
$script:procSplit.Panel2.Controls.Add($script:procTabs)
$script:procSplit.Panel2.Controls.Add($script:txtProcSummary)
$script:procSplit.Panel2.Controls.Add($procTools)
$script:tabProc.Controls.Add($script:procSplit)
$script:tabProc.Controls.Add($script:lblProcCount)
$script:tabProc.Controls.Add($procTop)
$script:tabProc.Controls.Add($procInfo)

$script:tabs.TabPages.AddRange(@($tabOverview, $script:tabIdx, $script:tabStats, $script:tabMissing, $script:tabProc, $script:tabScript, $script:tabLog, $tabHelp))

$script:form.Controls.Add($split)
$script:form.Controls.Add($script:lblWarn)
$script:form.Controls.Add($script:lblServerInfo)
$script:form.Controls.Add($script:lblEnv)
$script:form.Controls.Add($bar)
$script:form.Controls.Add($statusStrip)
#endregion

#region ------------------------------------------------------------------ timers & events
# SplitterDistance can only be set once a SplitContainer has a real size, so apply it lazily.
$script:Splitters = @(
    @{ C = $split; D = 400 }, @{ C = $idxSplit; D = 330 }, @{ C = $statSplit; D = 260 },
    @{ C = $script:procSplit; D = 330 }, @{ C = $script:procFindSplit; D = 0.55 }, @{ C = $script:procStmtSplit; D = 0.5 },
    @{ C = $statBottom; D = 0.55 }, @{ C = $missSplit; D = 300 }, @{ C = $logSplit; D = 330 }
)
function Set-Splitters {
    foreach ($sp in $script:Splitters) {
        if ($sp.Done) { continue }
        $c = $sp.C
        $len = if ($c.Orientation -eq 'Vertical') { $c.Width } else { $c.Height }
        $d = if ($sp.D -lt 1) { [int]($len * $sp.D) } else { [int]$sp.D }
        if ($len -gt $d + 80) { try { $c.SplitterDistance = $d; $sp.Done = $true } catch { } }
    }
}
$script:tabs.add_SelectedIndexChanged({ Set-Splitters })
$script:JobTimer = New-Object System.Windows.Forms.Timer
$script:JobTimer.Interval = 150
$script:JobTimer.add_Tick({
    $job = $script:State.Job
    if (-not $job) { return }
    if ($job.Handle.IsCompleted) {
        Complete-AnalyzerJob
        if (-not $script:State.Busy -and $script:State.PendingObjectId) { Invoke-TableLoad -ObjectId $script:State.PendingObjectId }
    } else {
        $script:lblStatus.Text = ("{0} ... {1:N0}s" -f $job.Status, ((Get-Date) - $job.Started).TotalSeconds)
    }
})
$script:JobTimer.Start()

$script:SearchTimer = New-Object System.Windows.Forms.Timer
$script:SearchTimer.Interval = 250
$script:SearchTimer.add_Tick({ $script:SearchTimer.Stop(); Update-TableFilter })

$script:TableTimer = New-Object System.Windows.Forms.Timer
$script:TableTimer.Interval = 450
$script:TableTimer.add_Tick({ $script:TableTimer.Stop(); Open-SelectedTable })

$script:StatTimer = New-Object System.Windows.Forms.Timer
$script:StatTimer.Interval = 300
$script:StatTimer.add_Tick({ $script:StatTimer.Stop(); Show-StatDetail })

$script:cboAuth.add_SelectedIndexChanged({ $sql = $script:cboAuth.SelectedItem -eq 'SQL login'; $script:txtUser.Enabled = $sql; $script:txtPwd.Enabled = $sql })
# Before connecting, the banner previews the typed server; once connected it always shows the connected server.
$script:cboServer.add_TextChanged({ if (-not $script:State.Context) { Update-EnvironmentBanner } })
$script:cboServer.add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; Connect-Server } })
$btnServers.add_Click({ Show-ServerListDialog })
$script:btnConnect.add_Click({ Connect-Server })
$script:btnCancel.add_Click({ Stop-AnalyzerJob })
$script:cboDb.add_SelectedIndexChanged({ Select-Database })

$script:txtSearch.add_TextChanged({ $script:SearchTimer.Stop(); $script:SearchTimer.Start() })
$script:txtSearch.add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq 'Down') { $e.SuppressKeyPress = $true; $script:gridTables.Focus() }
    if ($e.KeyCode -eq 'Enter') {
        $e.SuppressKeyPress = $true
        $script:SearchTimer.Stop(); Update-TableFilter
        if ($script:gridTables.Rows.Count -gt 0) { $script:gridTables.CurrentCell = $script:gridTables.Rows[0].Cells['TableName']; Open-SelectedTable }
    }
})
$script:gridTables.add_DataBindingComplete({
    param($s, $e)
    if ($s.Columns.Contains('ObjectId')) { $s.Columns['ObjectId'].Visible = $false }
    Set-GridFormats $s
    if ($s.Columns.Contains('TableName')) { $s.Columns['TableName'].Width = 170 }
    if ($s.Columns.Contains('SchemaName')) { $s.Columns['SchemaName'].Width = 60; $s.Columns['SchemaName'].HeaderText = 'Schema' }
    if ($s.Columns.Contains('RowCount')) { $s.Columns['RowCount'].HeaderText = 'Rows' }
    if ($s.Columns.Contains('ReservedMB')) { $s.Columns['ReservedMB'].HeaderText = 'Size MB' }
    if ($s.Columns.Contains('Partitions')) { $s.Columns['Partitions'].Width = 40; $s.Columns['Partitions'].HeaderText = 'Parts' }
    $s.ClearSelection()
})
$script:gridTables.add_SelectionChanged({ if ($script:gridTables.Focused) { $script:TableTimer.Stop(); $script:TableTimer.Start() } })
$script:gridTables.add_CellDoubleClick({ $script:TableTimer.Stop(); Open-SelectedTable })
$script:gridTables.add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $script:TableTimer.Stop(); Open-SelectedTable } })
$script:btnRefreshTable.add_Click({ if ($script:State.CurrentObjectId) { $script:State.Analysis = $null; Invoke-TableLoad -ObjectId $script:State.CurrentObjectId } })

foreach ($g in @($script:gridIdx, $script:gridStats, $script:gridMissing, $script:gridPriority, $script:gridHist, $script:gridDensity, $gridLog)) {
    $g.add_DataBindingComplete({ param($s, $e) Set-GridFormats $s; $s.ClearSelection() })
}
$script:gridIdx.add_SelectionChanged({ Show-IndexDetail })
$script:gridStats.add_SelectionChanged({ $script:StatTimer.Stop(); $script:StatTimer.Start() })
$script:gridMissing.add_SelectionChanged({ Show-MissingDetail })
$script:gridPriority.add_CellDoubleClick({
    param($s, $e)
    if ($e.RowIndex -lt 0) { return }
    $kind = [string]$s.Rows[$e.RowIndex].Cells['Kind'].Value; $name = [string]$s.Rows[$e.RowIndex].Cells['Name'].Value
    $grid = if ($kind -eq 'Index') { $script:gridIdx } else { $script:gridStats }
    $col = if ($kind -eq 'Index') { 'IndexName' } else { 'StatName' }
    $script:tabs.SelectedTab = if ($kind -eq 'Index') { $script:tabIdx } else { $script:tabStats }
    foreach ($row in $grid.Rows) { if ([string]$row.Cells[$col].Value -eq $name) { $grid.ClearSelection(); $row.Selected = $true; $grid.CurrentCell = $row.Cells[$col]; break } }
})
$gridLog.add_SelectionChanged({ if ($gridLog.CurrentRow) { $txtLogSql.Text = [string]$gridLog.CurrentRow.Cells['Sql'].Value } })
$gridLog.add_DataBindingComplete({ param($s, $e) if ($s.Columns.Contains('Sql')) { $s.Columns['Sql'].Visible = $false }; if ($s.Columns.Contains('Time')) { $s.Columns['Time'].DefaultCellStyle.Format = 'HH:mm:ss'; $s.Columns['Time'].Width = 70 }; if ($s.Columns.Contains('Status')) { $s.Columns['Status'].Width = 500 }; if ($s.Columns.Contains('What')) { $s.Columns['What'].Width = 220 } })

$script:btnFragLimited.add_Click({ Invoke-FragmentationCheck -Mode LIMITED })
$script:btnFragSampled.add_Click({ Invoke-FragmentationCheck -Mode SAMPLED })
$btnIdxScript.add_Click({ Show-IndexScriptDialog })
$btnStatScript.add_Click({ Show-StatsScriptDialog })
$btnMissAdd.add_Click({ if ($script:txtMissingScript.Text) { Add-ScriptText $script:txtMissingScript.Text } })
$btnCopy.add_Click({ if ($script:txtScript.Text) { [System.Windows.Forms.Clipboard]::SetText($script:txtScript.Text); Set-Status 'Script copied to the clipboard.' } })
$btnClear.add_Click({ if ([System.Windows.Forms.MessageBox]::Show('Clear the script text?', 'Clear', 'YesNo') -eq 'Yes') { $script:txtScript.Clear() } })
$btnSave.add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'SQL script (*.sql)|*.sql|All files (*.*)|*.*'
    $tblName = if ($script:State.CurrentTable) { "$($script:State.CurrentTable.SchemaName).$($script:State.CurrentTable.TableName)" } else { 'maintenance' }
    $sfd.FileName = ("{0}_{1}_{2:yyyyMMdd_HHmm}.sql" -f ($script:cboServer.Text -replace '[\\/:*?"<>|]', '_'), ($tblName -replace '[\\/:*?"<>|]', '_'), (Get-Date))
    if ($sfd.ShowDialog($script:form) -eq 'OK') { [System.IO.File]::WriteAllText($sfd.FileName, $script:txtScript.Text, (New-Object System.Text.UTF8Encoding($true))); Set-Status "Saved $($sfd.FileName)" }
})

# --- procedure analysis events
$script:ProcSearchTimer = New-Object System.Windows.Forms.Timer
$script:ProcSearchTimer.Interval = 250
$script:ProcSearchTimer.add_Tick({ $script:ProcSearchTimer.Stop(); Update-ProcFilter })
$script:btnProcLoad.add_Click({ Invoke-ProcedureListLoad })
$script:txtProcSearch.add_TextChanged({ $script:ProcSearchTimer.Stop(); $script:ProcSearchTimer.Start() })
$script:txtProcSearch.add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $script:ProcSearchTimer.Stop(); Update-ProcFilter; if ($script:gridProcs.Rows.Count -gt 0) { $script:gridProcs.CurrentCell = $script:gridProcs.Rows[0].Cells['ProcName']; Invoke-ProcAnalysis } } })
$script:chkProcForTable.add_CheckedChanged({
    if (-not $script:State.ProcMaster) { return }
    if ($script:chkProcForTable.Checked) { Update-ProcReferences } else { Set-ProcBase }
})
$script:gridProcs.add_DataBindingComplete({
    param($s, $e)
    Set-GridFormats $s
    if ($s.Columns.Contains('ObjectId')) { $s.Columns['ObjectId'].Visible = $false }
    if ($s.Columns.Contains('SchemaName')) { $s.Columns['SchemaName'].Width = 60; $s.Columns['SchemaName'].HeaderText = 'Schema' }
    if ($s.Columns.Contains('ProcName')) { $s.Columns['ProcName'].Width = 220; $s.Columns['ProcName'].HeaderText = 'Procedure' }
    $s.ClearSelection()
})
$script:gridProcs.add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { Invoke-ProcAnalysis } })
$script:btnProcAnalyze.add_Click({ Invoke-ProcAnalysis })
$script:cboProcPlan.add_SelectedIndexChanged({
    if ($script:State.ProcPlanSuppress -or -not $script:State.ProcAnalysis) { return }
    $i = $script:cboProcPlan.SelectedIndex
    $src = @($script:State.ProcAnalysis.Sources)
    if ($i -lt 0 -or $i -ge $src.Count) { return }
    if ($script:State.ProcAnalysis.Source -and $src[$i].Key -eq $script:State.ProcAnalysis.Source.Key) { return }
    Invoke-ProcAnalysis -SourceKey $src[$i].Key
})
foreach ($g in @($script:gridProcFindings, $script:gridProcStmts)) { $g.add_DataBindingComplete({ param($s, $e) Set-GridFormats $s; foreach ($h in 'Idx', 'Key', 'Doc') { if ($s.Columns.Contains($h)) { $s.Columns[$h].Visible = $false } }; if ($s.Columns.Contains('Finding')) { $s.Columns['Finding'].Width = 700 }; if ($s.Columns.Contains('Text')) { $s.Columns['Text'].Width = 600 }; if ($s.Columns.Contains('Action')) { $s.Columns['Action'].Width = 260 }; if ($s.Columns.Contains('Object')) { $s.Columns['Object'].Width = 200 }; if ($s.Columns.Contains('EstimateOffBy')) { $s.Columns['EstimateOffBy'].DefaultCellStyle.Format = 'N1'; $s.Columns['EstimateOffBy'].HeaderText = 'Off by (x)' } }) }
$script:gridProcFindings.add_SelectionChanged({ Show-ProcFindingDetail })
$script:gridProcStmts.add_SelectionChanged({
    $row = $script:gridProcStmts.CurrentRow
    if (-not $row -or -not $script:State.ProcAnalysis) { return }
    $key = [string]$row.Cells['Key'].Value
    $st = @($script:State.ProcAnalysis.Plan.Statements | Where-Object { $_.Key -eq $key }) | Select-Object -First 1
    if (-not $st) { return }
    $prm = @($script:State.ProcAnalysis.Plan.Parameters | Where-Object { $_.StatementKey -eq $key } | ForEach-Object { "$($_.Name) $($_.DataType) compiled for $($_.CompiledValue)" })
    $ops = @($script:State.ProcAnalysis.Plan.Operators | Where-Object { $_.StatementKey -eq $key } | ForEach-Object { "  {0,-22} {1}.{2} ({3})  est rows {4:N0}{5}" -f $_.PhysicalOp, $_.Table, $_.Index, $_.IndexKind, [double]$_.EstRows, $(if ($_.EstRowsRead) { ', reads {0:N0}' -f [double]$_.EstRowsRead } else { '' }) })
    $script:txtProcStmt.Text = "STATEMENT $($st.StatementId) ($($st.Type)), estimated cost $($st.Cost), estimated rows $($st.EstRows)`r`n`r`n$($st.Text)`r`n`r`nDATA ACCESS`r`n$($ops -join "`r`n")`r`n`r`nPARAMETERS`r`n  $(if ($prm.Count) { $prm -join "`r`n  " } else { '(none)' })"
})
$script:btnProcFrag.add_Click({ Invoke-ProcFragmentationCheck })
$btnProcStats.add_Click({ Add-ProcStatsScript })
$btnProcIdx.add_Click({ Add-ProcIndexScript })
$btnProcAdd.add_Click({
    $row = $script:gridProcFindings.CurrentRow
    if (-not $row) { Set-Status 'Select a finding first.'; return }
    $f = $script:State.ProcFindings[[int]$row.Cells['Idx'].Value]
    $txt = Get-ProcFindingScript $f
    if ($txt) { Add-ScriptText $txt }
    elseif ($f.Kind -eq 'Statistic') { Add-ProcStatsScript }
    elseif ($f.Action -in 'REBUILD', 'REORGANIZE') { Add-ProcIndexScript }
    else { Set-Status 'This finding has no script: it is advice about the query or indexing design.' }
})
$btnSavePlan.add_Click({ Save-ProcPlan })

$script:form.add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq 'F5' -and $script:State.CurrentObjectId -and -not $script:State.Busy) { $script:State.Analysis = $null; Invoke-TableLoad -ObjectId $script:State.CurrentObjectId }
    if ($e.Control -and $e.KeyCode -eq 'F') { $script:txtSearch.Focus(); $script:txtSearch.SelectAll() }
    if ($e.KeyCode -eq 'Escape' -and $script:State.Busy) { Stop-AnalyzerJob }
})
$script:form.add_FormClosing({
    if ($script:State.Busy) { Stop-AnalyzerJob }
    $script:JobTimer.Stop()
    try { $script:Settings.Encrypt = [bool]$script:chkEncrypt.Checked; $script:Settings.ReadOnlyIntent = [bool]$script:chkReadOnlyIntent.Checked; Save-AnalyzerSettings -Settings $script:Settings } catch { }
})
$script:form.add_Shown({
    Set-Splitters
    Update-ServerCombo
    if ($Server) { $script:cboServer.Text = $Server; $script:InitialDatabase = $Database; Connect-Server }
    Update-EnvironmentBanner
    $script:cboServer.Focus()
})
#endregion

$script:InitialDatabase = $null
try {
    [void]$script:form.ShowDialog()
} finally {
    if ($script:State.Job) { try { $script:State.Job.PS.Stop(); $script:State.Job.PS.Dispose() } catch { } }
    try { $script:Worker.Close(); $script:Worker.Dispose() } catch { }
    $script:form.Dispose()
}

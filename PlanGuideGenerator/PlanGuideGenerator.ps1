<#
.SYNOPSIS
    PlanGuideGenerator - builds a reviewed sp_create_plan_guide script from a statement plus
    a query hint or a showplan XML. WinForms GUI, Windows PowerShell 5.1.

.DESCRIPTION
    Generation is pure text: nothing is created on any server by this tool. The script it
    writes has pre-flight checks (database, @@SERVERNAME, AG primary), an optional safety
    latch, a byte-for-byte text check against the plan cache or module, the create call,
    validation with sys.fn_validate_plan_guide, and disable/drop rollback lines.

    "Start from saved plans" takes the saved good and/or bad plan, works out which is which
    (actual elapsed/CPU/reads, or your labels), extracts the T-SQL, @params or module, and
    decides: hint given -> hint; no hint + confirmed good plan -> USE PLAN with its XML.

    The optional "Check on server" button runs SELECT statements only, through a read-only
    guard, on a background runspace (the window never freezes).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\PlanGuideGenerator.ps1
.EXAMPLE
    .\PlanGuideGenerator.ps1 -Server SQLAG01 -Database Sales
.EXAMPLE
    . .\PlanGuideGenerator.ps1 -NoGui      # load the engine only, for scripting
    $r = New-PlanGuideScript -Options @{ Statement = 'SELECT ...'; Params = '@p int'; Hint = 'RECOMPILE'; Database = 'Sales' }
    $r.Findings; $r.Script | Set-Content .\guide.sql -Encoding UTF8
.EXAMPLE
    . .\PlanGuideGenerator.ps1 -NoGui
    $d = Resolve-PlanGuideFromPlans -PlanA (Get-Content -Raw .\good.sqlplan) -LabelA Good -PlanB (Get-Content -Raw .\bad.sqlplan)
    $d.Findings; $d.Report; $d.Suggestions
    if ($d.Options) { (New-PlanGuideScript -Options $d.Options).Script }
#>
[CmdletBinding()]
param(
    [string]$Server,
    [string]$Database,
    [switch]$NoGui,
    [switch]$HideConsole
)

$script:ShowPlanNs = 'http://schemas.microsoft.com/sqlserver/2004/07/showplan'
$script:Q = [char]39
$script:PlanCache = @{ Src = $null; Res = $null }

#region ---------- small helpers ----------
function Get-ShortHash([string]$Text) {
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try { $b = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([string]$Text)) } finally { $md5.Dispose() }
    (($b[0..3] | ForEach-Object { $_.ToString('X2') }) -join '')
}
function Get-Squashed([string]$s) { ([regex]::Replace([string]$s, '\s+', ' ')).Trim() }
function Add-Finding($List, [string]$Severity, [string]$Message) {
    [void]$List.Add([pscustomobject]@{ Severity = $Severity; Message = $Message })
}
function Convert-LineBreak([string]$Text, [string]$Mode) {
    if ($null -eq $Text) { return $null }
    switch ($Mode) {
        'CRLF'  { return [regex]::Replace($Text, '\r\n|\r|\n', "`r`n") }
        'LF'    { return [regex]::Replace($Text, '\r\n|\r|\n', "`n") }
        default { return $Text }
    }
}
function Get-LineBreakKind([string]$s) {
    if ([string]::IsNullOrEmpty($s)) { return 'none' }
    $crlf = [regex]::Matches($s, '\r\n').Count
    $lf   = [regex]::Matches($s, '(?<!\r)\n').Count
    $cr   = [regex]::Matches($s, '\r(?!\n)').Count
    $kinds = @(); if ($crlf) { $kinds += 'CRLF' }; if ($lf) { $kinds += 'LF' }; if ($cr) { $kinds += 'CR' }
    if ($kinds.Count -eq 0) { return 'none' }
    if ($kinds.Count -gt 1) { return 'mixed' }
    return $kinds[0]
}
function Show-Visible([string]$s) { $s.Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t') }
function Get-FirstDifference([string]$a, [string]$b) {
    $n = [Math]::Min($a.Length, $b.Length)
    for ($i = 0; $i -lt $n; $i++) { if ($a[$i] -cne $b[$i]) { return $i } }
    if ($a.Length -ne $b.Length) { return $n }
    return -1
}
function Get-DiffSnippet([string]$s, [int]$at) {
    $start = [Math]::Max(0, $at - 15)
    $len = [Math]::Min($s.Length - $start, 35)
    if ($len -le 0) { return '<end of text>' }
    (Show-Visible $s.Substring($start, $len))
}
#endregion

#region ---------- T-SQL text parsing ----------
function Read-NLiteral([string]$s, [int]$i) {
    while ($i -lt $s.Length -and [char]::IsWhiteSpace($s[$i])) { $i++ }
    if ($i -lt $s.Length -and ($s[$i] -ceq [char]'N' -or $s[$i] -ceq [char]'n')) { $i++ }
    if ($i -ge $s.Length -or $s[$i] -ne $script:Q) { return $null }
    $i++
    $sb = New-Object System.Text.StringBuilder
    while ($i -lt $s.Length) {
        if ($s[$i] -eq $script:Q) {
            if ($i + 1 -lt $s.Length -and $s[$i + 1] -eq $script:Q) { [void]$sb.Append($script:Q); $i += 2; continue }
            return [pscustomobject]@{ Value = $sb.ToString(); End = $i + 1 }
        }
        [void]$sb.Append($s[$i]); $i++
    }
    return $null
}

function ConvertFrom-SpExecuteSql([string]$Text) {
    # exec sp_executesql N'stmt', N'@p int', @p = 1   ->   Stmt + Params
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $m = [regex]::Match($Text, '^\s*(?:exec(?:ute)?\s+)?(?:\[?sys\]?\.)?\[?sp_executesql\]?\s+', 'IgnoreCase')
    if (-not $m.Success) { return $null }
    $i = $m.Length
    $n = [regex]::Match($Text.Substring($i), '^\s*@stmt\s*=\s*', 'IgnoreCase'); if ($n.Success) { $i += $n.Length }
    $a = Read-NLiteral $Text $i
    if ($null -eq $a) { return $null }
    $i = $a.End; $params = $null
    $c = [regex]::Match($Text.Substring($i), '^\s*,\s*')
    if ($c.Success) {
        $i += $c.Length
        $n = [regex]::Match($Text.Substring($i), '^@params\s*=\s*', 'IgnoreCase'); if ($n.Success) { $i += $n.Length }
        $b = Read-NLiteral $Text $i
        if ($null -ne $b) { $params = $b.Value }
    }
    [pscustomobject]@{ Stmt = $a.Value; Params = $params; Kind = 'sp_executesql' }
}

function ConvertFrom-ParamPrefix([string]$Text) {
    # plan-cache text of a parameterized statement: (@p int,@q nvarchar(10))SELECT ...
    if ([string]::IsNullOrEmpty($Text)) { return $null }
    $i = 0; while ($i -lt $Text.Length -and [char]::IsWhiteSpace($Text[$i])) { $i++ }
    if ($i -ge $Text.Length -or $Text[$i] -ne [char]'(') { return $null }
    $j = $i + 1; while ($j -lt $Text.Length -and [char]::IsWhiteSpace($Text[$j])) { $j++ }
    if ($j -ge $Text.Length -or $Text[$j] -ne [char]'@') { return $null }
    $d = 0
    for ($k = $i; $k -lt $Text.Length; $k++) {
        $ch = $Text[$k]
        if ($ch -eq [char]'(') { $d++ }
        elseif ($ch -eq [char]')') {
            $d--
            if ($d -eq 0) { return [pscustomobject]@{ Params = $Text.Substring($i + 1, $k - $i - 1); Stmt = $Text.Substring($k + 1); Kind = 'prefix' } }
        }
    }
    return $null
}

function Remove-PspHint([string]$Text) {
    # SQL Server 2022+ PSP / 2025 OPPO query variants: the plan's StatementText ends with an engine-generated
    #   option (PLAN PER VALUE(ObjectID = .., QueryVariantID = .., predicate_range(..) | optional_predicate(..)))
    # The application never sends it (and SQL Server doesn't parse it if you do), so strip it.
    $res = [pscustomobject]@{ Text = $Text; Removed = $false; VariantId = $null; ObjectId = 0; Predicate = $null }
    if ([string]::IsNullOrEmpty($Text)) { return $res }
    $code = Get-CodeOnly $Text
    $m = [regex]::Match($code, 'PLAN\s+PER\s+VALUE\s*\(', 'IgnoreCase')
    if (-not $m.Success) { return $res }
    $ppvOpen = $m.Index + $m.Length - 1
    $ppvClose = Find-CloseParen $Text $ppvOpen
    if ($ppvClose -lt 0) { return $res }
    $ppv = $Text.Substring($m.Index, $ppvClose - $m.Index + 1)
    $v = [regex]::Match($ppv, 'QueryVariantID\s*=\s*(\d+)', 'IgnoreCase'); if ($v.Success) { $res.VariantId = [int]$v.Groups[1].Value }
    $ob = [regex]::Match($ppv, 'ObjectID\s*=\s*(-?\d+)', 'IgnoreCase'); if ($ob.Success) { $res.ObjectId = [long]$ob.Groups[1].Value }
    $pr = [regex]::Match($ppv, '(predicate_range|optional_predicate)\s*\(.*\)\s*\)$', 'IgnoreCase'); if ($pr.Success) { $res.Predicate = $pr.Value.TrimEnd(')').Trim() + ')' }
    # inside an OPTION ( ... ) clause?
    $opts = [regex]::Matches($code.Substring(0, $m.Index), 'OPTION\s*\(', 'IgnoreCase')
    $handled = $false
    if ($opts.Count) {
        $o = $opts[$opts.Count - 1]
        $oOpen = $o.Index + $o.Length - 1
        $oClose = Find-CloseParen $Text $oOpen
        if ($oClose -gt $ppvClose) {
            $inner = $Text.Substring($oOpen + 1, $oClose - $oOpen - 1)
            $parts = Split-TopLevel $inner
            $keep = @($parts | Where-Object { $_ -notmatch '^PLAN\s+PER\s+VALUE\s*\(' })
            $start = $o.Index; $end = $oClose + 1
            if ($keep.Count) { $replacement = $Text.Substring($o.Index, $o.Length - 1) + '(' + ($keep -join ', ') + ')' }
            else { $replacement = ''; if ($start -gt 0 -and $Text[$start - 1] -eq [char]' ') { $start-- } }
            $res.Text = $Text.Substring(0, $start) + $replacement + $Text.Substring($end)
            $handled = $true
        }
    }
    if (-not $handled) {
        $start = $m.Index; if ($start -gt 0 -and $Text[$start - 1] -eq [char]' ') { $start-- }
        $res.Text = $Text.Substring(0, $start) + $Text.Substring($ppvClose + 1)
    }
    $res.Removed = $true
    $res
}
function Get-PlanStatementText($Stmt) { (Remove-PspHint $Stmt.GetAttribute('StatementText')).Text }

function Get-CodeOnly([string]$s) {
    # blank out comments, string literals and quoted names so keyword tests don't misfire
    [regex]::Replace([string]$s, '--[^\r\n]*|/\*[\s\S]*?\*/|N?''(?:[^'']|'''')*''|\[(?:[^\]]|\]\])*\]|"(?:[^"]|"")*"', ' ')
}
function Find-CloseParen([string]$s, [int]$open) {
    $d = 0; $inQ = $false
    for ($k = $open; $k -lt $s.Length; $k++) {
        $ch = $s[$k]
        if ($inQ) { if ($ch -eq $script:Q) { if ($k + 1 -lt $s.Length -and $s[$k + 1] -eq $script:Q) { $k++ } else { $inQ = $false } }; continue }
        if ($ch -eq $script:Q) { $inQ = $true }
        elseif ($ch -eq [char]'(') { $d++ }
        elseif ($ch -eq [char]')') { $d--; if ($d -eq 0) { return $k } }
    }
    return -1
}
function Test-Balanced([string]$s) {
    $d = 0; $inQ = $false
    for ($k = 0; $k -lt $s.Length; $k++) {
        $ch = $s[$k]
        if ($inQ) { if ($ch -eq $script:Q) { if ($k + 1 -lt $s.Length -and $s[$k + 1] -eq $script:Q) { $k++ } else { $inQ = $false } }; continue }
        if ($ch -eq $script:Q) { $inQ = $true }
        elseif ($ch -eq [char]'(') { $d++ }
        elseif ($ch -eq [char]')') { $d--; if ($d -lt 0) { return $false } }
    }
    return ($d -eq 0 -and -not $inQ)
}
function Split-TopLevel([string]$s) {
    $out = New-Object System.Collections.Generic.List[string]
    $d = 0; $inQ = $false; $start = 0
    for ($k = 0; $k -lt $s.Length; $k++) {
        $ch = $s[$k]
        if ($inQ) { if ($ch -eq $script:Q) { if ($k + 1 -lt $s.Length -and $s[$k + 1] -eq $script:Q) { $k++ } else { $inQ = $false } }; continue }
        if ($ch -eq $script:Q) { $inQ = $true }
        elseif ($ch -eq [char]'(') { $d++ }
        elseif ($ch -eq [char]')') { $d-- }
        elseif ($ch -eq [char]',' -and $d -eq 0) { $out.Add($s.Substring($start, $k - $start).Trim()); $start = $k + 1 }
    }
    $out.Add($s.Substring($start).Trim())
    ,@($out | Where-Object { $_ })
}
function Get-SqlVariables([string]$s) {
    $set = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches((Get-CodeOnly $s), '(?<![@\w])@[A-Za-z_#$][\w@#$]*')) {
        $v = $m.Value.ToLowerInvariant(); if (-not $set.Contains($v)) { $set.Add($v) }
    }
    $set.ToArray()
}
#endregion

#region ---------- hints ----------
$script:KnownHints = @(
    '^RECOMPILE$', '^MAXDOP\s+\d+$', '^OPTIMIZE\s+FOR\s+UNKNOWN$', '^OPTIMIZE\s+FOR\s*\(', '^USE\s+HINT\s*\(', '^QUERYTRACEON\s+\d+$',
    '^FORCE\s+ORDER$', '^(HASH|LOOP|MERGE)\s+JOIN$', '^(HASH|ORDER)\s+GROUP$', '^(CONCAT|HASH|MERGE)\s+UNION$', '^KEEP\s+PLAN$',
    '^KEEPFIXED\s+PLAN$', '^EXPAND\s+VIEWS$', '^FAST\s+\d+$', '^MAXRECURSION\s+\d+$', '^ROBUST\s+PLAN$', '^NO_PERFORMANCE_SPOOL$',
    '^(MAX|MIN)_GRANT_PERCENT\s*=\s*[\d.]+$', '^IGNORE_NONCLUSTERED_COLUMNSTORE_INDEX$', '^DISABLE_OPTIMIZED_PLAN_FORCING$',
    '^(FORCE|DISABLE)\s+(EXTERNALPUSHDOWN|SCALEOUTEXECUTION)$', '^PARAMETERIZATION\s+(SIMPLE|FORCED)$', '^TABLE\s+HINT\s*\(',
    '^USE\s+PLAN\s+N?''', '^LABEL\s*='
)
$script:QuickHints = @(
    'RECOMPILE', 'OPTIMIZE FOR UNKNOWN', 'MAXDOP 1', "USE HINT('DISABLE_PARAMETER_SNIFFING')",
    "USE HINT('FORCE_LEGACY_CARDINALITY_ESTIMATION')", "USE HINT('ENABLE_QUERY_OPTIMIZER_HOTFIXES')",
    'FORCE ORDER', 'KEEPFIXED PLAN', 'LOOP JOIN', 'HASH JOIN', 'MAX_GRANT_PERCENT = 10'
)

function ConvertTo-PlanGuideHint([string]$Raw, $F, [string]$Where) {
    $h = ([string]$Raw).Trim() -replace ';+\s*$', ''
    if (-not $h) { Add-Finding $F 'Error' "${Where}: enter a query hint."; return $null }
    $m = [regex]::Match($h, '^option\s*\(', 'IgnoreCase')
    if ($m.Success) {
        $close = Find-CloseParen $h ($m.Length - 1)
        if ($close -ne $h.Length - 1) { Add-Finding $F 'Error' "${Where}: the OPTION ( ... ) clause isn't one balanced set of parentheses."; return $null }
        $inner = $h.Substring($m.Length, $close - $m.Length)
    } else { $inner = $h }
    if (-not (Test-Balanced $inner)) { Add-Finding $F 'Error' "${Where}: unbalanced parentheses or quotes in the hint."; return $null }
    $inner = ([regex]::Replace($inner.Trim(), '\s*\r?\n\s*', ' '))
    if (-not $inner) { Add-Finding $F 'Error' "${Where}: the OPTION clause is empty."; return $null }
    $parts = Split-TopLevel $inner
    foreach ($p in $parts) {
        $u = ([regex]::Replace($p.ToUpperInvariant(), '\s+', ' '))
        $known = $false; foreach ($r in $script:KnownHints) { if ($u -match $r) { $known = $true; break } }
        if (-not $known) { Add-Finding $F 'Warn' "${Where}: '$p' isn't a hint this tool recognizes. Check the spelling; sp_create_plan_guide rejects invalid hints." }
    }
    [pscustomobject]@{ Hints = 'OPTION (' + ($parts -join ', ') + ')'; Parts = $parts; Inner = $inner }
}
#endregion

#region ---------- showplan XML ----------
function Read-ShowPlan([string]$Xml) {
    if ($script:PlanCache.Src -ceq $Xml) { return $script:PlanCache.Res }
    $res = [pscustomobject]@{ Error = $null; Doc = $null; Ns = $null; Statements = @(); Sig = (Get-ShortHash $Xml) }
    try {
        $settings = New-Object System.Xml.XmlReaderSettings
        $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $sr = New-Object System.IO.StringReader ($Xml.TrimStart([char]0xFEFF))
        $xr = [System.Xml.XmlReader]::Create($sr, $settings)
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        try { $doc.Load($xr) } finally { $xr.Close() }
        $root = $doc.DocumentElement
        if ($root.LocalName -ne 'ShowPlanXML') { $res.Error = 'The root element is not <ShowPlanXML>. Paste the whole showplan (in SSMS: right-click the plan > Show Execution Plan XML).' }
        elseif ($root.NamespaceURI -ne $script:ShowPlanNs) { $res.Error = "The showplan namespace is missing or different. Expected $($script:ShowPlanNs)." }
        else {
            $ns = New-Object System.Xml.XmlNamespaceManager $doc.NameTable
            $ns.AddNamespace('sp', $script:ShowPlanNs)
            $res.Doc = $doc; $res.Ns = $ns
            $res.Statements = @($doc.SelectNodes('//sp:StmtSimple[sp:QueryPlan] | //sp:StmtCursor[sp:CursorPlan]', $ns))
            if ($res.Statements.Count -eq 0) { $res.Error = 'No statement with a query plan was found in this XML.' }
        }
    } catch {
        $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
        $res.Error = 'The plan XML does not parse: ' + $e.Message
    }
    $script:PlanCache = @{ Src = $Xml; Res = $res }
    $res
}
function Get-PlanParameters($Stmt, $Ns) {
    @($Stmt.SelectNodes('sp:QueryPlan/sp:ParameterList/sp:ColumnReference', $Ns) | ForEach-Object {
        [pscustomobject]@{ Name = $_.GetAttribute('Column'); Type = $_.GetAttribute('ParameterDataType'); Compiled = $_.GetAttribute('ParameterCompiledValue') }
    })
}
function New-PlanXmlForGuide($Plan, $Stmt, [bool]$Trim, [bool]$Strip) {
    if ($Trim) {
        $nd = New-Object System.Xml.XmlDocument
        $src = $Plan.Doc.DocumentElement
        $root = $nd.CreateElement('ShowPlanXML', $script:ShowPlanNs)
        foreach ($a in 'Version', 'Build') { if ($src.HasAttribute($a)) { $root.SetAttribute($a, $src.GetAttribute($a)) } }
        [void]$nd.AppendChild($root)
        $bs = $nd.CreateElement('BatchSequence', $script:ShowPlanNs); [void]$root.AppendChild($bs)
        $b  = $nd.CreateElement('Batch', $script:ShowPlanNs);         [void]$bs.AppendChild($b)
        $ss = $nd.CreateElement('Statements', $script:ShowPlanNs);    [void]$b.AppendChild($ss)
        [void]$ss.AppendChild($nd.ImportNode($Stmt, $true))
    } else { $nd = $Plan.Doc.Clone() }
    if ($Strip) {
        $ns = New-Object System.Xml.XmlNamespaceManager $nd.NameTable; $ns.AddNamespace('sp', $script:ShowPlanNs)
        $kill = @($nd.SelectNodes('//sp:RunTimeInformation | //sp:RunTimePartitionSummary | //sp:QueryTimeStats | //sp:WaitStats | //sp:ThreadStat', $ns))
        foreach ($n in $kill) { [void]$n.ParentNode.RemoveChild($n) }
    }
    $nd.DocumentElement.OuterXml
}
function Get-PlanStatementFill([string]$Xml, [int]$Index) {
    # text + params to drop into the query boxes from a plan statement
    $p = Read-ShowPlan $Xml
    if ($p.Error -or $Index -lt 0 -or $Index -ge $p.Statements.Count) { return $null }
    $st = $p.Statements[$Index]
    $text = Get-PlanStatementText $st; $params = $null; $note = $null
    $pref = ConvertFrom-ParamPrefix $text
    if ($pref) { $text = $pref.Stmt; $params = $pref.Params }
    else {
        $pp = @(Get-PlanParameters $st $p.Ns | Where-Object { $_.Type })
        if ($pp.Count) { $params = ($pp | ForEach-Object { "$($_.Name) $($_.Type)" }) -join ', '; $note = 'Built @params from the plan''s parameter list. Check the order and spelling against what the app sends.' }
    }
    [pscustomobject]@{ Stmt = $text; Params = $params; Note = $note }
}
#endregion

#region ---------- script generation ----------
function Format-NLiteral([string]$Text, [bool]$Encode) {
    if ($null -eq $Text) { return 'NULL' }
    $parts = [regex]::Split($Text, '(\r\n|\r|\n)')
    if (-not $Encode -or $parts.Count -eq 1) { return "N'" + $Text.Replace("'", "''") + "'" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("CAST(N'" + $parts[0].Replace("'", "''") + "' AS nvarchar(max))")
    for ($k = 1; $k -lt $parts.Count; $k += 2) {
        $brk = switch -CaseSensitive ($parts[$k]) { "`r`n" { 'NCHAR(13) + NCHAR(10)' } "`n" { 'NCHAR(10)' } default { 'NCHAR(13)' } }
        $line = if ($k + 1 -lt $parts.Count) { $parts[$k + 1] } else { '' }
        [void]$sb.Append("`r`n    + $brk + N'" + $line.Replace("'", "''") + "'")
    }
    $sb.ToString()
}
function Get-PlanGuideDefaults {
    @{
        Statement = ''; Type = 'SQL'; Params = ''; UseBatch = $false; Batch = ''; Module = ''
        Parameterization = 'FORCED'; TemplateHint = $false
        Source = 'Hint'; Hint = ''; PlanXml = ''; PlanStatementIndex = -1; TrimPlan = $true; StripRuntime = $false
        Name = ''; Database = ''; ExpectedServer = ''; LineBreak = 'Keep'; EncodeLineBreaks = $true
        DropExisting = $true; SafetyLatch = $true; MatchCheck = $true; PspPlanHash = $null; PspObjectId = 0
        Method = 'PlanGuide'; QueryHash = $null
    }
}

function New-PlanGuideScript {
    [CmdletBinding()]
    param([hashtable]$Options)
    $o = Get-PlanGuideDefaults
    if ($Options) { foreach ($k in @($Options.Keys)) { $o[$k] = $Options[$k] } }

    $F = New-Object System.Collections.Generic.List[object]
    $L = New-Object System.Collections.Generic.List[string]
    function P([string]$s) { [void]$L.Add($s) }
    function PB([string]$block) { foreach ($ln in ($block -split '\r?\n')) { [void]$L.Add($ln) } }
    function Lit([string]$s) { Format-NLiteral $s ([bool]$o.EncodeLineBreaks) }
    function SqlStr([string]$s) { "N'" + ([string]$s).Replace("'", "''") + "'" }

    $type   = ([string]$o.Type).ToUpperInvariant()
    $src    = if ($type -eq 'TEMPLATE') { 'Hint' } else { [string]$o.Source }
    $raw    = [string]$o.Statement
    $q      = Convert-LineBreak $raw $o.LineBreak
    $params = if ([string]::IsNullOrWhiteSpace($o.Params)) { '' } else { [string]$o.Params }
    $useBatch = ($type -eq 'SQL' -and $o.UseBatch -and -not [string]::IsNullOrWhiteSpace($o.Batch))
    $batch  = if ($useBatch) { Convert-LineBreak ([string]$o.Batch) $o.LineBreak } else { $null }
    $module = ([string]$o.Module).Trim()
    $db     = ([string]$o.Database).Trim()
    $srv    = ([string]$o.ExpectedServer).Trim()
    $hasQ   = -not [string]::IsNullOrWhiteSpace($q)

    $result = [pscustomobject]@{
        Script = ''; Findings = $F; Name = ''; HasErrors = $false; Type = $type
        Statement = $q; Params = $params; Batch = $batch; Module = $module; Database = $db
        CachedKey = $null; PlanSig = $null; PlanLabels = @(); PlanIndex = -1; PlanMeta = ''
        Method = $(if ($o.Method -eq 'QueryStore') { 'QueryStore' } else { 'PlanGuide' }); Source = $src
        QueryHash = $null; PlanHash = $null; PspHash = $null
    }

    # ---- name ----
    $name = ([string]$o.Name).Trim()
    if (-not $name -and $hasQ) {
        $base = if ($type -eq 'OBJECT' -and $module) { ($module -replace '[\[\]]', '').Split('.')[-1] } elseif ($type -eq 'TEMPLATE') { 'TPL' } else { 'SQL' }
        $name = "PG_${base}_" + (Get-ShortHash ($raw + '|' + $params + '|' + $module))
    }
    $result.Name = $name
    if (-not $name -and $hasQ) { Add-Finding $F 'Error' 'Give the plan guide a name.' }
    if ($name.StartsWith('#')) { Add-Finding $F 'Error' 'A plan guide name can''t start with #.' }
    if ($name.Length -gt 128) { Add-Finding $F 'Error' 'The guide name is longer than 128 characters.' }

    # ---- statement ----
    if (-not $hasQ) { Add-Finding $F 'Error' 'Paste the statement the guide should match.' }
    if ($hasQ -and ((ConvertFrom-SpExecuteSql $raw) -or (ConvertFrom-ParamPrefix $raw))) {
        Add-Finding $F 'Warn' 'The query box still holds a whole sp_executesql call or a (@params) prefix. Use "Split it", or the guide will match text that never runs.'
    }
    $code = Get-CodeOnly $q
    if ($hasQ) {
        if ($code -match 'PLAN\s+PER\s+VALUE\s*\(') { Add-Finding $F 'Error' 'The statement contains OPTION (PLAN PER VALUE ...). SQL Server adds that to PSP query variants in plans; the app never sends it and SQL Server won''t parse it. Remove it (Use this statement''s text, or Analyze in section 0, strips it for you).' }
        elseif ($code -match '\bOPTION\s*\(') { Add-Finding $F 'Warn' 'The statement already has an OPTION clause. The guide''s hints are applied instead of it, so copy any hints you want to keep into the guide.' }
        if ($q -match '^\s' -or $q -match '\s$') { Add-Finding $F 'Info' 'The statement starts or ends with whitespace. It is kept, because it has to match too. Make sure it is really part of what the app sends.' }
        if ($code -match '\bCREATE\s+(PROC|PROCEDURE|FUNCTION|TRIGGER)\b') { Add-Finding $F 'Error' 'This looks like a module definition. Choose "Inside a module" and paste only the one statement from it.' }
        $kind = Get-LineBreakKind $q
        if ($kind -eq 'mixed') { Add-Finding $F 'Warn' 'The statement mixes CRLF and LF line breaks. That can be genuine, but it is usually a copy/paste artefact. Pick CRLF or LF under Line breaks unless you know it is mixed.' }
        if ($kind -ne 'none' -and -not $o.EncodeLineBreaks) { Add-Finding $F 'Info' 'Line breaks are written raw inside the string. If an editor changes them (SSMS can "normalize line endings" when it opens a file), the guide stops matching. Writing them as NCHAR() is safer.' }
    }
    $pvars = @(Get-SqlVariables $params)
    $qvars = @(Get-SqlVariables $q)

    if ($type -eq 'SQL' -and $hasQ) {
        if ($params) {
            if ($params -notmatch '@') { Add-Finding $F 'Error' '@params must declare parameters, e.g. @CustomerID int.' }
            $unused = @($pvars | Where-Object { $qvars -notcontains $_ })
            if ($unused.Count) { Add-Finding $F 'Warn' ("@params declares " + ($unused -join ', ') + ", which the statement doesn't use. That's allowed, but the string must still match the app's exactly.") }
        }
        $undeclared = @($qvars | Where-Object { $pvars -notcontains $_ })
        if ($undeclared.Count -and -not $useBatch) {
            if ($code -match '\bDECLARE\b') { Add-Finding $F 'Info' 'The statement declares its own variables. Fine for an ad hoc batch.' }
            else { Add-Finding $F 'Warn' ("The statement uses " + ($undeclared -join ', ') + ", which isn't declared in @params. If it comes from sp_executesql, put the declaration in @params. If it is a local variable in a larger batch, give the whole batch.") }
        }
        if (-not $params -and $qvars.Count -eq 0 -and ([regex]::Replace($q, '--[^\r\n]*', '') -match '(=|<>|<|>|\bIN\s*\(|\bLIKE)\s*(N?''|-?\d)')) {
            Add-Finding $F 'Info' 'This is ad hoc text with literal values. The guide matches only these exact values. If the app sends many different values, use Template (forced parameterization) plus a hint on the parameterized form.'
        }
        if ($useBatch -and $batch.IndexOf($q, [StringComparison]::Ordinal) -lt 0) { Add-Finding $F 'Error' 'The statement does not appear exactly inside the batch text.' }
        if ($useBatch) { Add-Finding $F 'Info' 'The guide matches only when that whole batch is sent unchanged.' }
        $result.CachedKey = $(if ($params) { '(' + $params + ')' } else { '' }) + $(if ($useBatch) { $batch } else { $q })
    }
    $objId = [long]$o.PspObjectId
    if ($type -ne 'OBJECT' -and $objId -gt 0) { Add-Finding $F 'Error' ("The plan's PSP clause has ObjectID = $objId, so this statement runs inside a module. Choose 'Inside a module' and enter SELECT OBJECT_SCHEMA_NAME($objId) + '.' + OBJECT_NAME($objId).") }
    if ($type -eq 'OBJECT') {
        if (-not $module) { Add-Finding $F 'Error' 'Enter the module name (schema.name).' }
        elseif ($module -notmatch '\.') { Add-Finding $F 'Warn' "Use a schema-qualified module name, e.g. dbo.$module." }
        if ($hasQ) { Add-Finding $F 'Info' 'Paste the statement exactly as it is in the module, without the surrounding procedure code. Encrypted, CLR and natively compiled modules can''t take plan guides.' }
    }
    if ($type -eq 'TEMPLATE' -and $hasQ) {
        if (-not ([regex]::IsMatch($q, '''|(?<![\w@#$.])\d'))) { Add-Finding $F 'Warn' 'No literal values found. sp_get_query_template needs a query with literals to parameterize.' }
        if ($qvars.Count) { Add-Finding $F 'Warn' ('Template guides work on queries with literal values, not variables. Paste a sample with real values in place of ' + ($qvars -join ', ') + '.') }
        if ($o.Parameterization -eq 'FORCED') { Add-Finding $F 'Info' 'A FORCED template guide only does something when the database uses SIMPLE parameterization (the default). The script shows the database setting.' }
        else { Add-Finding $F 'Info' 'A SIMPLE template guide only does something when the database uses FORCED parameterization. The script shows the database setting.' }
    }

    # ---- hint / plan ----
    $hint = $null; $hint2 = $null; $planXml = $null; $planInfo = $null
    if ($type -eq 'TEMPLATE') {
        $hint = [pscustomobject]@{ Hints = "OPTION (PARAMETERIZATION $($o.Parameterization))" }
        if ($o.TemplateHint) {
            $hint2 = ConvertTo-PlanGuideHint $o.Hint $F 'Hint for the parameterized form'
            if ($hint2 -and @($hint2.Parts | Where-Object { $_ -match '^PARAMETERIZATION' }).Count) { Add-Finding $F 'Error' 'PARAMETERIZATION belongs in the template guide itself, not the second guide.' }
            if ($hint2 -and $hint2.Inner -match 'OPTIMIZE\s+FOR\s*\(') { Add-Finding $F 'Info' 'After parameterization the parameters are named @0, @1, ... Use those names in OPTIMIZE FOR. The script prints the template so you can check.' }
        }
    }
    elseif ($src -eq 'Hint') {
        $hint = ConvertTo-PlanGuideHint $o.Hint $F 'Hint'
        if ($hint) {
            if (@($hint.Parts | Where-Object { $_ -match '^PARAMETERIZATION' }).Count) { Add-Finding $F 'Error' 'PARAMETERIZATION is only valid in a Template guide.' }
            if (@($hint.Parts | Where-Object { $_ -match '^USE\s+PLAN' }).Count) { Add-Finding $F 'Info' 'For USE PLAN, the Plan XML mode checks the plan and keeps just the one statement for you.' }
            if (@($hint.Parts | Where-Object { $_ -match '^QUERYTRACEON' }).Count) { Add-Finding $F 'Info' 'QUERYTRACEON in a plan guide works without sysadmin for the caller. Make sure you want the flag on every execution.' }
            if (@($hint.Parts | Where-Object { $_ -match '^RECOMPILE$' }).Count) { Add-Finding $F 'Info' 'RECOMPILE adds compile CPU to every execution. Fine for heavy, infrequent queries; costly for ones that run thousands of times a second.' }
            if ($type -eq 'SQL' -and -not $useBatch) {
                $hv = @(Get-SqlVariables $hint.Inner)
                $bad = @($hv | Where-Object { $pvars -notcontains $_ -and $qvars -notcontains $_ })
                if ($bad.Count) { Add-Finding $F 'Error' ('The hint refers to ' + ($bad -join ', ') + ', which neither the statement nor @params has.') }
                else {
                    $notParam = @($hv | Where-Object { $pvars -notcontains $_ })
                    if ($notParam.Count -and $code -notmatch '\bDECLARE\b') { Add-Finding $F 'Warn' ('OPTIMIZE FOR names ' + ($notParam -join ', ') + ', which isn''t in @params.') }
                }
            }
        }
    }
    else {
        $xml = ([string]$o.PlanXml).Trim()
        if (-not $xml) { Add-Finding $F 'Error' 'Paste the plan XML, or open a .sqlplan file.' }
        else {
            $plan = Read-ShowPlan $xml
            if ($plan.Error) { Add-Finding $F 'Error' $plan.Error }
            else {
                $result.PlanSig = $plan.Sig
                $labels = @(); $best = 0; $bestCost = -1.0; $matchIdx = -1
                for ($i = 0; $i -lt $plan.Statements.Count; $i++) {
                    $s = $plan.Statements[$i]
                    $txt = Get-Squashed (Get-PlanStatementText $s)
                    $cost = 0.0; [void][double]::TryParse($s.GetAttribute('StatementSubTreeCost'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$cost)
                    if ($cost -gt $bestCost) { $bestCost = $cost; $best = $i }
                    if ($hasQ -and $matchIdx -lt 0) {
                        $cmpT = Get-PlanStatementText $s; $pp0 = ConvertFrom-ParamPrefix $cmpT; if ($pp0) { $cmpT = $pp0.Stmt }
                        if ((Get-Squashed $cmpT) -ceq (Get-Squashed $q)) { $matchIdx = $i }
                    }
                    $st = $s.GetAttribute('StatementType'); if (-not $st) { $st = $s.LocalName }
                    $short = if ($txt.Length -gt 90) { $txt.Substring(0, 90) + '...' } else { $txt }
                    $labels += ('#{0}  {1}  cost {2:N3}  {3}' -f ($i + 1), $st, $cost, $short)
                }
                $idx = [int]$o.PlanStatementIndex
                if ($idx -lt 0 -or $idx -ge $plan.Statements.Count) { $idx = $(if ($matchIdx -ge 0) { $matchIdx } else { $best }) }
                $result.PlanLabels = $labels; $result.PlanIndex = $idx
                $stmt  = $plan.Statements[$idx]
                $ptext = Get-PlanStatementText $stmt
                $psp = Remove-PspHint $stmt.GetAttribute('StatementText')
                $stype = $stmt.GetAttribute('StatementType').ToUpperInvariant()
                $pp    = @(Get-PlanParameters $stmt $plan.Ns)
                $planInfo = [pscustomobject]@{ QueryHash = $stmt.GetAttribute('QueryHash'); PlanHash = $stmt.GetAttribute('QueryPlanHash') }
                $meta = @(
                    ('Statement type : ' + $(if ($stmt.GetAttribute('StatementType')) { $stmt.GetAttribute('StatementType') } else { $stmt.LocalName })),
                    ('Estimated cost : ' + $stmt.GetAttribute('StatementSubTreeCost')),
                    ('Query hash     : ' + $planInfo.QueryHash),
                    ('Plan hash      : ' + $planInfo.PlanHash),
                    ('Build          : ' + $plan.Doc.DocumentElement.GetAttribute('Build')),
                    ('Parameters     : ' + $(if ($pp.Count) { ($pp | ForEach-Object { "$($_.Name) $($_.Type) = $($_.Compiled)" }) -join ', ' } else { 'none' }))
                )
                if ($stmt.GetAttribute('PlanGuideName')) { $meta += ('Plan guide     : ' + $stmt.GetAttribute('PlanGuideName')) }
                $result.PlanMeta = $meta -join "`r`n"

                $n = $plan.Statements.Count
                if ($psp.Removed -and $psp.ObjectId -gt 0 -and $type -ne 'OBJECT' -and $objId -le 0) { Add-Finding $F 'Error' ("The plan's PSP clause has ObjectID = $($psp.ObjectId), so this statement runs inside a module. Choose 'Inside a module' and enter SELECT OBJECT_SCHEMA_NAME($($psp.ObjectId)) + '.' + OBJECT_NAME($($psp.ObjectId)).") }
                if ($psp.Removed) {
                    $planInfo | Add-Member -NotePropertyName PspVariant -NotePropertyValue $true
                    Add-Finding $F 'Warn' ("This plan is a Parameter Sensitive Plan (PSP) query variant (QueryVariantID $($psp.VariantId)). Its StatementText carries an engine-generated OPTION (PLAN PER VALUE ...) that the app never sends, so it was left out of the statement. Forcing it with a plan guide applies this one variant's plan to EVERY parameter value and replaces PSP's per-range plans; Microsoft documents forcing a single variant through Query Store instead (see the alternative at the end of the script).")
                }
                if ($n -gt 1 -and -not $o.TrimPlan) { Add-Finding $F 'Error' 'The plan holds several statements. USE PLAN needs the plan of exactly one statement, so keep "Keep only the chosen statement" ticked.' }
                if ($n -gt 1 -and $o.TrimPlan) { Add-Finding $F 'OK' "The plan has $n statements. Only #$($idx + 1) goes into the guide." }
                if ($stype -and $stype -notmatch '^(SELECT|INSERT|UPDATE|DELETE|MERGE)') { Add-Finding $F 'Warn' "USE PLAN works on SELECT, INSERT, UPDATE, DELETE and MERGE. This statement is $stype." }
                foreach ($cp in @($stmt.SelectNodes('.//sp:CursorPlan', $plan.Ns))) {
                    $ct = ($cp.GetAttribute('CursorActualType') + ' ' + $cp.GetAttribute('CursorRequestedType'))
                    if ($ct -match 'dynamic|keyset') { Add-Finding $F 'Error' 'USE PLAN can''t force plans for dynamic or keyset-driven cursors.'; break }
                }
                $ops = @($stmt.SelectNodes('.//sp:RelOp', $plan.Ns) | ForEach-Object { $_.GetAttribute('PhysicalOp') })
                if (@($ops | Where-Object { $_ -like 'Remote*' }).Count) { Add-Finding $F 'Warn' 'The plan has remote (linked-server) operators. USE PLAN often fails on distributed queries.' }
                if ($stmt.SelectNodes('.//sp:RunTimeInformation', $plan.Ns).Count) { Add-Finding $F 'Info' 'This is an actual plan. That works for USE PLAN. "Strip actual runtime stats" makes the XML smaller.' }
                if ($stmt.GetAttribute('PlanGuideName')) { Add-Finding $F 'Info' ("This plan was already compiled under plan guide " + $stmt.GetAttribute('PlanGuideName') + '.') }
                $pk = Get-LineBreakKind $ptext; $qk = Get-LineBreakKind $q
                if ($hasQ -and $pk -ne 'none' -and $qk -ne 'none' -and $pk -ne $qk) { Add-Finding $F 'Warn' "The statement text in the plan uses $pk line breaks, but the script will use $qk. Set Line breaks to match what the app sends." }
                if ($hasQ) {
                    $pref = ConvertFrom-ParamPrefix $ptext
                    $cmp = if ($pref) { $pref.Stmt } else { $ptext }
                    if ((Get-Squashed $cmp) -cne (Get-Squashed $q)) { Add-Finding $F 'Warn' 'The statement text doesn''t match the plan''s StatementText, even ignoring whitespace. Make sure the plan belongs to this statement.' }
                    elseif ($cmp -cne $q) { Add-Finding $F 'Info' 'The text matches the plan''s statement apart from whitespace. Whitespace has to match what the app sends; the text check will tell you.' }
                    else { Add-Finding $F 'OK' 'The statement text matches the plan''s StatementText exactly.' }
                }
                $planXml = New-PlanXmlForGuide $plan $stmt ([bool]$o.TrimPlan) ([bool]$o.StripRuntime)
                Add-Finding $F 'Info' ('Plan XML in the guide: {0:N1} KB.' -f ([Text.Encoding]::Unicode.GetByteCount($planXml) / 1024))
            }
        }
    }

    # ---- Query Store output checks ----
    $isQS = $result.Method -eq 'QueryStore'
    $qh = if ($planInfo -and $planInfo.QueryHash) { $planInfo.QueryHash } else { [string]$o.QueryHash }
    if ($qh -and $qh -notmatch '^0x[0-9A-Fa-f]{1,16}$') { $qh = $null }
    $pspH = if ($planInfo -and $planInfo.PSObject.Properties['PspVariant']) { $planInfo.PlanHash } elseif ($o.PspPlanHash) { [string]$o.PspPlanHash } else { $null }
    $result.QueryHash = $qh; $result.PspHash = $pspH; $result.PlanHash = $(if ($planInfo) { $planInfo.PlanHash } else { $null })
    if ($isQS) {
        if ($type -eq 'TEMPLATE') { Add-Finding $F 'Error' 'Query Store output: Template (forced parameterization) has no Query Store equivalent here. Use a plan guide.' }
        elseif ($src -eq 'Xml') {
            if (-not ($planInfo -and $planInfo.PlanHash)) { Add-Finding $F 'Error' 'Query Store output: the plan has no QueryPlanHash, so its plan_id can''t be found. Use a plan guide.' }
            else { Add-Finding $F 'Info' "Query Store output: the good plan is forced with sp_query_store_force_plan, found by its query_plan_hash $($planInfo.PlanHash). It must already be in this database's Query Store (it ran here and wasn't purged). Query Store forcing recompiles with that plan as a target; it usually gets exactly that plan, but it isn't guaranteed." }
        } else {
            if ($pspH) { Add-Finding $F 'Info' 'Query Store output: PSP query. The hint goes on the parent query (found through sys.query_store_query_variant), and every variant without its own hint inherits it.' }
            elseif ($qh) { Add-Finding $F 'Info' "Query Store output: the query_id is looked up by the plan's query_hash $qh, so whitespace differences don't matter. Every query_id with that hash gets the hint; Microsoft notes you may need exactly that when several share a hash." }
            else { Add-Finding $F 'Warn' 'Query Store output: no plan was loaded, so the query_id is looked up by exact statement text (with or without the (@params) prefix). Loading the plan gives a query_hash, which is more reliable.' }
            Add-Finding $F 'Info' 'Query Store hints need SQL Server 2022+ (or Azure SQL). They override plan guides and hints in the code, and sp_query_store_set_hints replaces any existing Query Store hint on that query_id (the script lists them first).'
            if ($code -match '(=|<>|<|>)\s*(N?''|-?\d)' -and -not $params -and $qvars.Count -eq 0) { Add-Finding $F 'Warn' 'Query Store hints aren''t supported for statements eligible for simple parameterization (ad hoc text with literal values).' }
        }
    }

    # ---- script ----
    $result.HasErrors = @($F | Where-Object { $_.Severity -eq 'Error' }).Count -gt 0
    $forces = if ($type -eq 'TEMPLATE') { $hint.Hints + $(if ($hint2) { ' + ' + $hint2.Hints } else { '' }) } elseif ($src -eq 'Xml') { 'USE PLAN (plan XML)' } elseif ($hint) { $hint.Hints } else { '?' }
    P $(if ($o.Method -eq 'QueryStore') { "/* Query Store $(if ($src -eq 'Xml') { 'forced plan' } else { 'hint' }) for the statement below" } else { "/* Plan guide: $(if ($name) { $name } else { '(unnamed)' })" })
    P ("   Type: $type" + $(if ($type -eq 'SQL' -and $params) { ' (parameterized: sp_executesql / prepared)' } elseif ($type -eq 'OBJECT') { " in $module" } else { '' }))
    P "   Forces: $forces"
    P ("   Generated {0:yyyy-MM-dd HH:mm} by PlanGuideGenerator on {1}. Review before running. */" -f (Get-Date), $env:COMPUTERNAME)
    if ($result.HasErrors) {
        P ''
        P '-- Fix the errors on the Checks tab first:'
        foreach ($f in $F) { if ($f.Severity -eq 'Error') { P ('--   ' + $f.Message) } }
        $result.Script = $L -join "`r`n"
        return $result
    }
    if ($db) { P ('USE [' + $db.Replace(']', ']]') + '];'); P 'GO' }
    P ''
    P '-- Pre-flight checks. Any failure turns on NOEXEC, so nothing below runs.'
    P 'DECLARE @db sysname = DB_NAME(), @srv sysname = @@SERVERNAME;'
    if ($db)  { P ("IF @db <> " + (SqlStr $db) + " BEGIN RAISERROR(N'Expected database " + $db.Replace("'", "''").Replace('%', '%%') + " but connected to %s.', 16, 1, @db); SET NOEXEC ON; END;") }
    if ($srv) { P ("IF @srv <> " + (SqlStr $srv) + " BEGIN RAISERROR(N'Expected server " + $srv.Replace("'", "''").Replace('%', '%%') + " but connected to %s.', 16, 1, @srv); SET NOEXEC ON; END;") }
    PB @'
IF EXISTS (SELECT 1 FROM sys.databases AS d
           JOIN sys.dm_hadr_availability_replica_states AS rs ON rs.replica_id = d.replica_id AND rs.is_local = 1
           WHERE d.database_id = DB_ID() AND rs.role_desc <> N'PRIMARY')
BEGIN RAISERROR(N'This database is an AG secondary. Plan guides live in the database: create this on the primary.', 16, 1); SET NOEXEC ON; END;
GO
'@
    if ($o.SafetyLatch) {
        PB @'

-- SAFETY LATCH: read the script, then delete the next line.
RAISERROR(N'Safety latch is on: review the script, then delete the RAISERROR / SET NOEXEC ON line.', 16, 1); SET NOEXEC ON;
GO
'@
    }
    if ($isQS) {
        Add-QueryStoreScript $L $o $type $src $q $params $module $(if ($hint) { $hint.Hints } else { $null }) $qh $(if ($planInfo) { $planInfo.PlanHash } else { $null }) $pspH
        $result.Script = $L -join "`r`n"
        return $result
    }
    P ''
    P 'SET NOCOUNT ON;'
    P ('DECLARE @name sysname = ' + (SqlStr $name) + ';')
    $dropLine = "IF EXISTS (SELECT 1 FROM sys.plan_guides WHERE name = @name) EXEC sys.sp_control_plan_guide @operation = N'DROP', @name = @name;"

    if ($type -eq 'TEMPLATE') {
        $name2 = $(if ($name.Length -gt 123) { $name.Substring(0, 123) } else { $name }) + '_hint'
        P ('DECLARE @query nvarchar(max) = ' + (Lit $q) + ';')
        P 'DECLARE @stmt nvarchar(max), @params nvarchar(max);'
        PB @'

-- Database parameterization (FORCED guides matter on SIMPLE databases, SIMPLE guides on FORCED ones)
SELECT name, CASE is_parameterization_forced WHEN 1 THEN N'FORCED' ELSE N'SIMPLE' END AS parameterization
FROM sys.databases WHERE database_id = DB_ID();

-- 1. Build the template from the sample query
EXEC sys.sp_get_query_template @querytext = @query, @templatetext = @stmt OUTPUT, @parameters = @params OUTPUT;
SELECT @stmt AS template_text, @params AS template_params;

-- 2. Create the template guide
'@
        if ($o.DropExisting) { P $dropLine }
        P ("EXEC sys.sp_create_plan_guide @name = @name, @stmt = @stmt, @type = N'TEMPLATE', @module_or_batch = NULL, @params = @params, @hints = N'" + $hint.Hints + "';")
        if ($hint2) {
            P ''
            P '-- 3. Hint for the parameterized form (a SQL guide on the template text)'
            P ('DECLARE @name2 sysname = ' + (SqlStr $name2) + ';')
            P ('DECLARE @hints2 nvarchar(max) = ' + (SqlStr $hint2.Hints) + ';')
            if ($o.DropExisting) { P "IF EXISTS (SELECT 1 FROM sys.plan_guides WHERE name = @name2) EXEC sys.sp_control_plan_guide @operation = N'DROP', @name = @name2;" }
            P "EXEC sys.sp_create_plan_guide @name = @name2, @stmt = @stmt, @type = N'SQL', @module_or_batch = NULL, @params = @params, @hints = @hints2;"
        }
        P ''
        P ("-- $(if ($hint2) { 4 } else { 3 }). Validate (msgnum NULL = valid)")
        P 'SELECT pg.plan_guide_id, pg.name, pg.scope_type_desc, pg.is_disabled, v.msgnum, v.severity, v.message'
        P 'FROM sys.plan_guides AS pg OUTER APPLY sys.fn_validate_plan_guide(pg.plan_guide_id) AS v'
        P ("WHERE pg.name IN (@name" + $(if ($hint2) { ', @name2' } else { '' }) + ');')
        P 'GO'; P 'SET NOEXEC OFF;'; P 'GO'
        Add-PlanGuideFooter $L $(if ($hint2) { @($name, $name2) } else { @($name) })
        $result.Script = $L -join "`r`n"
        return $result
    }

    P ('DECLARE @stmt nvarchar(max) = ' + (Lit $q) + ';')
    P ("DECLARE @type nvarchar(60) = N'$type';")
    P ('DECLARE @module_or_batch nvarchar(max) = ' + $(if ($type -eq 'OBJECT') { SqlStr $module } elseif ($useBatch) { Lit $batch } else { 'NULL' }) + ';')
    P ('DECLARE @params nvarchar(max) = ' + $(if ($type -eq 'SQL' -and $params) { SqlStr $params } else { 'NULL' }) + ';')
    if ($src -eq 'Xml') { P ('DECLARE @hints nvarchar(max) = ' + (SqlStr $planXml) + ';') }
    else { P ('DECLARE @hints nvarchar(max) = ' + (SqlStr $hint.Hints) + ';') }
    P ''
    $step = 1
    if ($o.MatchCheck) {
        if ($type -eq 'OBJECT') {
            P "-- $step. Does the module exist, and does it contain the statement byte-for-byte?"; $step++
            PB @'
SELECT o.type_desc,
       CASE WHEN o.type NOT IN ('P', 'FN', 'TF', 'TR') THEN N'NOT SUPPORTED: needs a T-SQL procedure, scalar function, multi-statement TVF or DML trigger'
            WHEN m.definition IS NULL THEN N'NO DEFINITION: the module is encrypted, not T-SQL, or you lack VIEW DEFINITION'
            WHEN CHARINDEX(@stmt COLLATE Latin1_General_BIN2, m.definition COLLATE Latin1_General_BIN2) > 0 THEN N'ok: statement found exactly'
            WHEN CHARINDEX(REPLACE(@stmt, NCHAR(13) + NCHAR(10), NCHAR(10)) COLLATE Latin1_General_BIN2, m.definition COLLATE Latin1_General_BIN2) > 0 THEN N'MISMATCH: the module uses LF line breaks; set Line breaks to LF'
            WHEN CHARINDEX(REPLACE(@stmt, NCHAR(10), NCHAR(13) + NCHAR(10)) COLLATE Latin1_General_BIN2, m.definition COLLATE Latin1_General_BIN2) > 0 THEN N'MISMATCH: the module uses CRLF line breaks; set Line breaks to CRLF'
            ELSE N'NOT FOUND: the text differs from the module (whitespace, case, comments?)' END AS statement_check
FROM sys.objects AS o LEFT JOIN sys.sql_modules AS m ON m.object_id = o.object_id
WHERE o.object_id = OBJECT_ID(@module_or_batch);
IF OBJECT_ID(@module_or_batch) IS NULL PRINT N'Module not found in this database.';
'@
        } else {
            P "-- $step. Is this exact text in the plan cache? (No rows can just mean the plan was evicted.)"; $step++
            PB @'
--    Binary comparison, so whitespace and line breaks must match.
DECLARE @cached nvarchar(max) = CASE WHEN @params IS NULL THEN ISNULL(@module_or_batch, @stmt)
                                     ELSE N'(' + @params + N')' + ISNULL(@module_or_batch, @stmt) END;
SELECT TOP (5) qs.execution_count, qs.last_execution_time, qs.query_hash, qs.query_plan_hash, DB_NAME(st.dbid) AS db
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS st
WHERE DATALENGTH(st.text) = DATALENGTH(@cached)
  AND st.text = @cached COLLATE Latin1_General_BIN2
ORDER BY qs.last_execution_time DESC
OPTION (MAXDOP 1);
'@
            if ($planInfo -and $planInfo.QueryHash) {
                P '--    No rows? See what the cache holds for this query hash and compare the text:'
                P ("--    SELECT TOP (5) st.text FROM sys.dm_exec_query_stats AS qs CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS st WHERE qs.query_hash = $($planInfo.QueryHash) OPTION (MAXDOP 1);")
            }
        }
        P ''
    }
    P "-- $step. Create the guide"; $step++
    if ($o.DropExisting) { P $dropLine }
    PB @'
EXEC sys.sp_create_plan_guide
     @name            = @name,
     @stmt            = @stmt,
     @type            = @type,
     @module_or_batch = @module_or_batch,
     @params          = @params,
     @hints           = @hints;

'@
    P "-- $step. Validate (msgnum NULL = valid)"
    PB @'
SELECT pg.plan_guide_id, pg.name, pg.scope_type_desc, pg.is_disabled, pg.create_date, v.msgnum, v.severity, v.message
FROM sys.plan_guides AS pg OUTER APPLY sys.fn_validate_plan_guide(pg.plan_guide_id) AS v
WHERE pg.name = @name;
GO
SET NOEXEC OFF;
GO
'@
    if ($src -eq 'Xml' -and $planInfo -and $planInfo.PlanHash) {
        P ''
        P '/* Alternative: build the guide straight from the cached plan. The statement text is then exact by construction.'
        P 'DECLARE @plan_handle varbinary(64), @offset int;'
        P 'SELECT TOP (1) @plan_handle = qs.plan_handle, @offset = qs.statement_start_offset'
        P ("FROM sys.dm_exec_query_stats AS qs WHERE qs.query_plan_hash = $($planInfo.PlanHash)")
        P 'ORDER BY qs.last_execution_time DESC OPTION (MAXDOP 1);'
        P "IF @plan_handle IS NULL RAISERROR(N'That plan is not in the cache.', 16, 1);"
        P ("ELSE EXEC sys.sp_create_plan_guide_from_handle @name = " + (SqlStr $name) + ", @plan_handle = @plan_handle, @statement_start_offset = @offset;")
        P '*/'
    }
    $pspHash = if ($planInfo -and ($planInfo.PSObject.Properties['PspVariant'])) { $planInfo.PlanHash } elseif ($o.PspPlanHash) { [string]$o.PspPlanHash } else { $null }
    if ($pspHash) {
        P ''
        P '/* PSP query (SQL Server 2022+, compatibility level 160+): the documented, per-variant alternative is Query Store.'
        P '   Find the parent query and the variant that owns this plan:'
        P '   SELECT v.parent_query_id, v.query_variant_query_id, p.plan_id, p.plan_type_desc, v.dispatcher_plan_id'
        P '   FROM sys.query_store_plan AS p'
        P '   JOIN sys.query_store_query_variant AS v ON v.query_variant_query_id = p.query_id'
        P ("   WHERE p.query_plan_hash = $pspHash;")
        if ($src -eq 'Xml') {
            P '   Force only the good variant (other parameter ranges keep their own plans):'
            P '   EXEC sys.sp_query_store_force_plan @query_id = <query_variant_query_id>, @plan_id = <plan_id>;'
        } else {
            P '   (the hash above is the bad plan''s when it was loaded, so the variant row is the one to fix)'
            P '   Hint the parent query; variants without their own Query Store hint inherit it:'
            P ("   EXEC sys.sp_query_store_set_hints @query_id = <parent_query_id>, @query_hints = " + (SqlStr $hint.Hints) + ';')
            P '   Or hint just the bad variant: same call with @query_id = <query_variant_query_id>.'
        }
        P '   Note: RECOMPILE and DISABLE_PARAMETER_SNIFFING turn PSP off for the query. */'
    }
    Add-PlanGuideFooter $L @($name)
    $result.Script = $L -join "`r`n"
    $result
}

function Add-QueryStoreScript($L, $o, [string]$Type, [string]$Src, [string]$Stmt, [string]$Params, [string]$Module, [string]$Hints, [string]$QueryHash, [string]$PlanHash, [string]$PspHash) {
    function Q([string]$x) { [void]$L.Add($x) }
    function QB([string]$block) { foreach ($ln in ($block -split '\r?\n')) { [void]$L.Add($ln) } }
    $sq = { param($x) "N'" + ([string]$x).Replace("'", "''") + "'" }
    $force = $Src -eq 'Xml'
    Q ''
    Q '-- Query Store must be on and writable here.'
    Q "IF NOT EXISTS (SELECT 1 FROM sys.database_query_store_options WHERE actual_state_desc = N'READ_WRITE')"
    Q "BEGIN RAISERROR(N'Query Store is not READ_WRITE in this database.', 16, 1); SET NOEXEC ON; END;"
    if (-not $force -or $PspHash) {
        Q "IF CAST(SERVERPROPERTY('EngineEdition') AS int) NOT IN (5, 8) AND CAST(SERVERPROPERTY('ProductMajorVersion') AS int) < 16"
        Q "BEGIN RAISERROR(N'$(if ($force) { 'PSP query variants' } else { 'Query Store hints' }) need SQL Server 2022 or later.', 16, 1); SET NOEXEC ON; END;"
    }
    Q 'GO'
    Q ''
    Q 'SET NOCOUNT ON;'
    Q "IF OBJECT_ID(N'tempdb..#pg_target') IS NOT NULL DROP TABLE #pg_target;"
    Q 'CREATE TABLE #pg_target (query_id bigint NOT NULL, plan_id bigint NULL, PRIMARY KEY (query_id));'
    Q ''
    if ($force) {
        Q "-- 1. Find the good plan in Query Store by its query_plan_hash ($PlanHash): newest plan_id per query_id"
        Q 'INSERT #pg_target (query_id, plan_id)'
        Q 'SELECT p.query_id, MAX(p.plan_id)'
        Q 'FROM sys.query_store_plan AS p'
        if ($QueryHash -and -not $PspHash) { Q 'JOIN sys.query_store_query AS q ON q.query_id = p.query_id'; Q "WHERE p.query_plan_hash = $PlanHash AND q.query_hash = $QueryHash" }
        else { Q "WHERE p.query_plan_hash = $PlanHash" }
        Q 'GROUP BY p.query_id;'
        Q ''
        Q 'SELECT t.query_id, t.plan_id, p.is_forced_plan, p.last_execution_time, q.object_id, OBJECT_NAME(q.object_id) AS module, LEFT(qt.query_sql_text, 300) AS query_text'
        Q 'FROM #pg_target AS t'
        Q 'JOIN sys.query_store_plan AS p ON p.plan_id = t.plan_id'
        Q 'JOIN sys.query_store_query AS q ON q.query_id = t.query_id'
        Q 'JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id;'
        Q "IF NOT EXISTS (SELECT 1 FROM #pg_target) BEGIN RAISERROR(N'The good plan is not in this database''s Query Store (other server, never ran here, or purged). Use the plan guide output instead.', 16, 1); SET NOEXEC ON; END;"
        Q 'GO'
        Q ''
        Q '-- 2. Force it'
        QB @'
DECLARE @query_id bigint, @plan_id bigint;
DECLARE pg_cur CURSOR LOCAL FAST_FORWARD FOR SELECT query_id, plan_id FROM #pg_target;
OPEN pg_cur;
FETCH NEXT FROM pg_cur INTO @query_id, @plan_id;
WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC sys.sp_query_store_force_plan @query_id = @query_id, @plan_id = @plan_id;
    FETCH NEXT FROM pg_cur INTO @query_id, @plan_id;
END;
CLOSE pg_cur; DEALLOCATE pg_cur;

-- 3. Verify (force_failure_count / last_force_failure_reason_desc show why forcing failed later)
SELECT p.query_id, p.plan_id, p.is_forced_plan, p.force_failure_count, p.last_force_failure_reason_desc
FROM sys.query_store_plan AS p JOIN #pg_target AS t ON t.plan_id = p.plan_id;
GO
SET NOEXEC OFF;
GO

/* Roll back: for each query_id / plan_id listed above
     EXEC sys.sp_query_store_unforce_plan @query_id = <query_id>, @plan_id = <plan_id>;
*/
'@
    } else {
        if ($PspHash) {
            Q "-- 1. PSP query: find the parent query through the variant that owns plan $PspHash"
            Q 'INSERT #pg_target (query_id)'
            Q 'SELECT DISTINCT v.parent_query_id'
            Q 'FROM sys.query_store_plan AS p'
            Q 'JOIN sys.query_store_query_variant AS v ON v.query_variant_query_id = p.query_id'
            Q "WHERE p.query_plan_hash = $PspHash;"
            Q '--    To hint only that variant instead, use v.query_variant_query_id above.'
        } elseif ($QueryHash) {
            Q "-- 1. Find the query by the plan's query_hash ($QueryHash)"
            Q 'INSERT #pg_target (query_id)'
            Q 'SELECT q.query_id FROM sys.query_store_query AS q'
            Q ("WHERE q.query_hash = $QueryHash" + $(if ($Type -eq 'OBJECT' -and $Module) { ' AND q.object_id = OBJECT_ID(' + (& $sq $Module) + ')' } else { '' }) + ';')
        } else {
            Q ('DECLARE @stmt nvarchar(max) = ' + (Format-NLiteral $Stmt ([bool]$o.EncodeLineBreaks)) + ';')
            Q ('DECLARE @params nvarchar(max) = ' + $(if ($Params) { & $sq $Params } else { 'NULL' }) + ';')
            Q '-- 1. Find the query by its exact text (with or without the (@params) prefix)'
            Q 'INSERT #pg_target (query_id)'
            Q 'SELECT q.query_id FROM sys.query_store_query AS q'
            Q 'JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id'
            Q ("WHERE (qt.query_sql_text = @stmt COLLATE Latin1_General_BIN2 OR qt.query_sql_text = N'(' + @params + N')' + @stmt COLLATE Latin1_General_BIN2)" + $(if ($Type -eq 'OBJECT' -and $Module) { "`r`n  AND q.object_id = OBJECT_ID(" + (& $sq $Module) + ')' } else { '' }) + ';')
        }
        QB @'

-- What will be hinted, and any Query Store hint it already has (it gets replaced)
SELECT t.query_id, q.object_id, OBJECT_NAME(q.object_id) AS module, q.last_execution_time,
       h.query_hint_text AS existing_hint, LEFT(qt.query_sql_text, 300) AS query_text
FROM #pg_target AS t
JOIN sys.query_store_query AS q ON q.query_id = t.query_id
JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id
LEFT JOIN sys.query_store_query_hints AS h ON h.query_id = t.query_id;
IF NOT EXISTS (SELECT 1 FROM #pg_target) BEGIN RAISERROR(N'No matching query in this database''s Query Store. Let it run once so Query Store captures it, then run this again.', 16, 1); SET NOEXEC ON; END;
GO

-- 2. Apply the hint
'@
        Q ('DECLARE @query_hints nvarchar(max) = ' + (& $sq $Hints) + ';')
        QB @'
DECLARE @query_id bigint;
DECLARE pg_cur CURSOR LOCAL FAST_FORWARD FOR SELECT query_id FROM #pg_target;
OPEN pg_cur;
FETCH NEXT FROM pg_cur INTO @query_id;
WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC sys.sp_query_store_set_hints @query_id = @query_id, @query_hints = @query_hints;
    FETCH NEXT FROM pg_cur INTO @query_id;
END;
CLOSE pg_cur; DEALLOCATE pg_cur;

-- 3. Verify (a failing hint is ignored, not fatal: check last_query_hint_failure_reason_desc after the next run)
SELECT h.query_id, h.query_hint_text, h.query_hint_failure_count, h.last_query_hint_failure_reason_desc
FROM sys.query_store_query_hints AS h JOIN #pg_target AS t ON t.query_id = h.query_id;
GO
SET NOEXEC OFF;
GO

/* Check it is used: the actual plan's statement shows QueryStoreStatementHintText / QueryStoreStatementHintId.
   Roll back: for each query_id listed above
     EXEC sys.sp_query_store_clear_hints @query_id = <query_id>;
*/
'@
    }
}

function Add-PlanGuideFooter($L, [string[]]$Names) {
    $lines = @(
        '',
        '/* Check that it is used',
        '   - Run the query with "Include Actual Execution Plan". The statement should show PlanGuideDB and PlanGuideName.',
        '   - Or watch it with Extended Events:',
        '       CREATE EVENT SESSION PlanGuideCheck ON SERVER',
        '         ADD EVENT sqlserver.plan_guide_successful, ADD EVENT sqlserver.plan_guide_unsuccessful',
        '         ADD TARGET package0.ring_buffer;',
        '   Roll back'
    )
    foreach ($n in $Names) {
        $s = "N'" + $n.Replace("'", "''") + "'"
        $lines += "     EXEC sys.sp_control_plan_guide @operation = N'DISABLE', @name = $s;   -- keep it, turn it off"
        $lines += "     EXEC sys.sp_control_plan_guide @operation = N'DROP',    @name = $s;   -- remove it"
    }
    $lines += '*/'
    foreach ($x in $lines) { [void]$L.Add($x) }
}
#endregion

#region ---------- start from saved plans: analysis and good/bad decision ----------
function ConvertTo-Num([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $d = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    return $null
}
function Get-SumAttr($Nodes, [string]$Attr) {
    $sum = 0.0; $any = $false
    foreach ($n in $Nodes) { $v = ConvertTo-Num $n.GetAttribute($Attr); if ($null -ne $v) { $sum += $v; $any = $true } }
    if ($any) { return $sum } else { return $null }
}
function ConvertFrom-CompiledValue([string]$v) {
    # showplan writes (42) for numbers, N'abc' for strings
    $t = ([string]$v).Trim()
    if ($t.StartsWith('(') -and $t.EndsWith(')') -and (Find-CloseParen $t 0) -eq $t.Length - 1) { $t = $t.Substring(1, $t.Length - 2).Trim() }
    $t
}

function Measure-PlanStatement($Plan, [int]$Index) {
    $ns = $Plan.Ns
    $s  = $Plan.Statements[$Index]
    $qp = $s.SelectSingleNode('sp:QueryPlan', $ns)
    $m = [ordered]@{
        Index = $Index; Node = $s; Text = (Get-PlanStatementText $s); Psp = (Remove-PspHint $s.GetAttribute('StatementText')); IsDispatcher = [bool]$s.SelectSingleNode('.//sp:Dispatcher', $Plan.Ns); StmtType = $s.GetAttribute('StatementType')
        QueryHash = $s.GetAttribute('QueryHash'); PlanHash = $s.GetAttribute('QueryPlanHash')
        Cost = ConvertTo-Num $s.GetAttribute('StatementSubTreeCost'); CE = $s.GetAttribute('CardinalityEstimationModelVersion')
        ParameterizedText = $s.GetAttribute('ParameterizedText')
        Dop = 1; IsActual = $false; ElapsedMs = $null; CpuMs = $null; LogicalReads = $null
        GrantKB = $null; UsedKB = $null; Spills = 0; WorstMisestimate = $null; WorstOp = ''
        Flags = New-Object System.Collections.Generic.List[string]
        Params = @(); Module = $null; ArithAbort = $null; Joins = @(); Shape = ''
    }
    if ($qp) {
        $dop = ConvertTo-Num $qp.GetAttribute('DegreeOfParallelism'); if ($dop) { $m.Dop = [int]$dop }
        $qts = $qp.SelectSingleNode('sp:QueryTimeStats', $ns)
        if ($qts) { $m.ElapsedMs = ConvertTo-Num $qts.GetAttribute('ElapsedTime'); $m.CpuMs = ConvertTo-Num $qts.GetAttribute('CpuTime') }
        $mg = $qp.SelectSingleNode('sp:MemoryGrantInfo', $ns)
        if ($mg) { $m.GrantKB = ConvertTo-Num $mg.GetAttribute('GrantedMemory'); $m.UsedKB = ConvertTo-Num $mg.GetAttribute('MaxUsedMemory') }
        $m.Params = @($qp.SelectNodes('sp:ParameterList/sp:ColumnReference', $ns) | ForEach-Object {
            [pscustomobject]@{ Name = $_.GetAttribute('Column'); Type = $_.GetAttribute('ParameterDataType')
                               Compiled = $_.GetAttribute('ParameterCompiledValue'); Runtime = $_.GetAttribute('ParameterRuntimeValue') } })
    }
    $threads = @($s.SelectNodes('.//sp:RunTimeCountersPerThread', $ns))
    $m.IsActual = $threads.Count -gt 0 -or $null -ne $m.ElapsedMs
    if ($threads.Count) {
        $m.LogicalReads = Get-SumAttr $threads 'ActualLogicalReads'
        if ($null -eq $m.ElapsedMs -and $qp) {
            $root = $qp.SelectSingleNode('sp:RelOp', $ns)
            if ($root) {
                $max = $null
                foreach ($t in @($root.SelectNodes('sp:RunTimeInformation/sp:RunTimeCountersPerThread', $ns))) { $v = ConvertTo-Num $t.GetAttribute('ActualElapsedms'); if ($null -ne $v -and ($null -eq $max -or $v -gt $max)) { $max = $v } }
                $m.ElapsedMs = $max
            }
        }
    }
    # misestimates (actual plans): rows per execution vs estimate, only where it matters (>= 1000 rows)
    $ops = @{}
    foreach ($r in @($s.SelectNodes('.//sp:RelOp', $ns))) {
        $po = $r.GetAttribute('PhysicalOp'); if ($po) { $ops[$po] = 1 + [int]$ops[$po] }
        $rt = @($r.SelectNodes('sp:RunTimeInformation/sp:RunTimeCountersPerThread', $ns))
        if (-not $rt.Count) { continue }
        $rows = Get-SumAttr $rt 'ActualRows'; $execs = Get-SumAttr $rt 'ActualExecutions'; $est = ConvertTo-Num $r.GetAttribute('EstimateRows')
        if ($null -eq $rows -or $null -eq $est) { continue }
        $a = $rows / [Math]::Max([double]$execs, 1.0)
        $big = [Math]::Max($a, $est)
        if ($big -lt 1000) { continue }
        $ratio = $big / [Math]::Max([Math]::Min($a, $est), 1.0)
        if ($null -eq $m.WorstMisestimate -or $ratio -gt $m.WorstMisestimate) {
            $m.WorstMisestimate = $ratio
            $m.WorstOp = ('{0} (node {1}): estimated {2:N0}, actual {3:N0} rows per execution' -f $po, $r.GetAttribute('NodeId'), $est, $a)
        }
    }
    $m.Joins = @($ops.Keys | Where-Object { $_ -in 'Nested Loops', 'Hash Match', 'Merge Join', 'Adaptive Join' })
    $m.Shape = (($ops.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 7 | ForEach-Object { "$($_.Key) x$($_.Value)" }) -join ', ')
    $m.Spills = @($s.SelectNodes('.//sp:SpillToTempDb | .//sp:SortSpillDetails | .//sp:HashSpillDetails | .//sp:ExchangeSpillDetails', $ns)).Count
    if ($m.Spills) { $m.Flags.Add("$($m.Spills) tempdb spill warning(s)") }
    if ($null -ne $m.WorstMisestimate -and $m.WorstMisestimate -ge 100) { $m.Flags.Add(('row estimate off by {0:N0}x at {1}' -f $m.WorstMisestimate, $m.WorstOp)) }
    if (@($s.SelectNodes('.//sp:PlanAffectingConvert[@ConvertIssue="Seek Plan"]', $ns)).Count) { $m.Flags.Add('an implicit conversion prevents a seek') }
    if (@($s.SelectNodes('.//sp:Warnings[@NoJoinPredicate="true" or @NoJoinPredicate="1"]', $ns)).Count) { $m.Flags.Add('a join has no join predicate') }
    if (@($s.SelectNodes('.//sp:MemoryGrantWarning', $ns)).Count) { $m.Flags.Add('memory grant warning') }
    $kl = @($s.SelectNodes('.//sp:RelOp[@PhysicalOp="Key Lookup" or @PhysicalOp="RID Lookup"]', $ns))
    foreach ($k in $kl) { $ex = Get-SumAttr @($k.SelectNodes('sp:RunTimeInformation/sp:RunTimeCountersPerThread', $ns)) 'ActualExecutions'; if ($ex -ge 10000) { $m.Flags.Add(('{0:N0} key/RID lookups' -f $ex)); break } }
    $mi = @($s.SelectNodes('.//sp:MissingIndexGroup', $ns)).Count
    if ($mi) { $m.Flags.Add("$mi missing-index suggestion(s)") }
    $so = $s.SelectSingleNode('sp:StatementSetOptions', $ns); if ($so) { $m.ArithAbort = $so.GetAttribute('ARITHABORT') }
    # inside a module? (actual plan of EXEC dbo.proc nests statements under StoredProc)
    $p = $s.ParentNode
    while ($p -and $p.NodeType -eq 'Element') {
        if ($p.LocalName -eq 'StoredProc') {
            $parts = @([regex]::Matches($p.GetAttribute('ProcName'), '\[(?:[^\]]|\]\])*\]|[^.\[\]]+') | ForEach-Object { $_.Value })
            $m.Module = if ($parts.Count -ge 2) { ($parts[-2..-1]) -join '.' } else { $p.GetAttribute('ProcName') }
            break
        }
        $p = $p.ParentNode
    }
    [pscustomobject]$m
}

function Select-PlanStatement($Plan, [int]$Index, [string]$QueryHash, [string]$Text) {
    # explicit index > same query hash > same text > most expensive (actual elapsed first)
    $n = $Plan.Statements.Count
    if ($Index -ge 0 -and $Index -lt $n) { return $Index }
    if ($QueryHash) { for ($i = 0; $i -lt $n; $i++) { if ($Plan.Statements[$i].GetAttribute('QueryHash') -eq $QueryHash) { return $i } } }
    if ($Text) { $t = Get-Squashed $Text; for ($i = 0; $i -lt $n; $i++) { if ((Get-Squashed (Get-PlanStatementText $Plan.Statements[$i])) -ceq $t) { return $i } } }
    if ($QueryHash -or $Text) { return -1 }
    $best = 0; $bestV = -1.0
    for ($i = 0; $i -lt $n; $i++) {
        $st = $Plan.Statements[$i]
        $qts = $st.SelectSingleNode('sp:QueryPlan/sp:QueryTimeStats', $Plan.Ns)
        $v = if ($qts) { 1e9 + [double](ConvertTo-Num $qts.GetAttribute('ElapsedTime')) } else { [double](ConvertTo-Num $st.GetAttribute('StatementSubTreeCost')) }
        if ($v -gt $bestV) { $bestV = $v; $best = $i }
    }
    $best
}

function Compare-PlanEvidence($A, $B) {
    # returns Winner ('A'|'B'|'Similar') and Strength ('Strong'|'Weak') with the reason
    if ($null -ne $A.ElapsedMs -and $null -ne $B.ElapsedMs) {
        $fa = [Math]::Max([double]$A.ElapsedMs, 1); $fb = [Math]::Max([double]$B.ElapsedMs, 1)
        $ratio = [Math]::Max($fa, $fb) / [Math]::Min($fa, $fb)
        if ($ratio -ge 1.5) { return [pscustomobject]@{ Winner = $(if ($fa -lt $fb) { 'A' } else { 'B' }); Strength = 'Strong'; Reason = ('actual elapsed time {0:N0} ms vs {1:N0} ms ({2:N1}x)' -f $A.ElapsedMs, $B.ElapsedMs, $ratio) } }
        $ca = $A.CpuMs; $cb = $B.CpuMs
        if ($null -ne $ca -and $null -ne $cb) {
            $ra = [Math]::Max([double]$ca, 1); $rb = [Math]::Max([double]$cb, 1); $r2 = [Math]::Max($ra, $rb) / [Math]::Min($ra, $rb)
            if ($r2 -ge 2) { return [pscustomobject]@{ Winner = $(if ($ra -lt $rb) { 'A' } else { 'B' }); Strength = 'Strong'; Reason = ('similar elapsed time, but CPU {0:N0} ms vs {1:N0} ms ({2:N1}x)' -f $ca, $cb, $r2) } }
        }
        return [pscustomobject]@{ Winner = 'Similar'; Strength = 'Strong'; Reason = ('actual elapsed time is close: {0:N0} ms vs {1:N0} ms' -f $A.ElapsedMs, $B.ElapsedMs) }
    }
    if ($null -ne $A.LogicalReads -and $null -ne $B.LogicalReads -and ($A.LogicalReads + $B.LogicalReads) -gt 0) {
        $ra = [Math]::Max([double]$A.LogicalReads, 1); $rb = [Math]::Max([double]$B.LogicalReads, 1); $r = [Math]::Max($ra, $rb) / [Math]::Min($ra, $rb)
        if ($r -ge 2) { return [pscustomobject]@{ Winner = $(if ($ra -lt $rb) { 'A' } else { 'B' }); Strength = 'Strong'; Reason = ('actual logical reads {0:N0} vs {1:N0} ({2:N1}x)' -f $A.LogicalReads, $B.LogicalReads, $r) } }
        return [pscustomobject]@{ Winner = 'Similar'; Strength = 'Strong'; Reason = ('actual logical reads are close: {0:N0} vs {1:N0}' -f $A.LogicalReads, $B.LogicalReads) }
    }
    if ($null -ne $A.Cost -and $null -ne $B.Cost) {
        $fa = [Math]::Max([double]$A.Cost, 0.0001); $fb = [Math]::Max([double]$B.Cost, 0.0001); $r = [Math]::Max($fa, $fb) / [Math]::Min($fa, $fb)
        $w = if ($r -lt 1.2) { 'Similar' } elseif ($fa -lt $fb) { 'A' } else { 'B' }
        return [pscustomobject]@{ Winner = $w; Strength = 'Weak'; Reason = ('estimated cost only: {0:N3} vs {1:N3}. Cost is the optimizer''s guess, and a bad plan usually looked cheap to it' -f $A.Cost, $B.Cost) }
    }
    [pscustomobject]@{ Winner = 'Similar'; Strength = 'Weak'; Reason = 'no runtime numbers or costs to compare' }
}

function Get-HintSuggestions($Good, $Bad) {
    $out = New-Object System.Collections.Generic.List[object]
    $add = { param($h, $why) if (-not @($out | Where-Object { $_.Hint -eq $h }).Count) { $out.Add([pscustomobject]@{ Hint = $h; Why = $why }) } }
    $src = if ($Bad) { $Bad } else { $Good }
    $sniffed = @($src.Params | Where-Object { $_.Compiled })
    if ($Good -and $Bad) {
        $diff = @()
        foreach ($gp in @($Good.Params | Where-Object { $_.Compiled })) {
            $bp = @($Bad.Params | Where-Object { $_.Name -eq $gp.Name })[0]
            if ($bp -and $bp.Compiled -and $bp.Compiled -cne $gp.Compiled) { $diff += $gp }
        }
        if ($diff.Count) {
            $list = ($diff | ForEach-Object { "$($_.Name) = $(ConvertFrom-CompiledValue $_.Compiled)" }) -join ', '
            & $add "OPTIMIZE FOR ($list)" 'The two plans were compiled for different parameter values (parameter sniffing). This compiles every time for the good plan''s values.'
        }
        if ($Good.Dop -le 1 -and $Bad.Dop -gt 1) { & $add 'MAXDOP 1' "The good plan is serial; the bad one runs at DOP $($Bad.Dop)." }
        if ($Good.Dop -gt 1 -and $Bad.Dop -le 1) { & $add "USE HINT('ENABLE_PARALLEL_PLAN_PREFERENCE')" "The good plan is parallel (DOP $($Good.Dop)); the bad one is serial. SQL Server 2016 SP1+." }
        if ($Good.CE -and $Bad.CE -and $Good.CE -ne $Bad.CE) {
            if ([int]$Good.CE -eq 70) { & $add "USE HINT('FORCE_LEGACY_CARDINALITY_ESTIMATION')" 'The good plan used the legacy cardinality estimator (CE 70).' }
            else { & $add "USE HINT('FORCE_DEFAULT_CARDINALITY_ESTIMATION')" "The good plan used CE $($Good.CE); the bad one CE $($Bad.CE)." }
        }
        if ($Good.Joins.Count -eq 1 -and ($Bad.Joins | Where-Object { $_ -ne $Good.Joins[0] })) {
            $j = @{ 'Nested Loops' = 'LOOP JOIN'; 'Hash Match' = 'HASH JOIN'; 'Merge Join' = 'MERGE JOIN' }[$Good.Joins[0]]
            if ($j) { & $add $j "Every join in the good plan is $($Good.Joins[0]). Careful: the hint applies to every join in the statement." }
        }
    }
    if ($sniffed.Count) {
        & $add 'OPTIMIZE FOR UNKNOWN' 'Compile for average density instead of the sniffed value. Stable, but not tuned to any value.'
        & $add 'RECOMPILE' 'Compile for the actual values on every execution. Best plan each time, at the cost of compile CPU.'
    }
    ,$out.ToArray()
}

function Format-PlanReport($Cols) {
    # $Cols: array of @{ Title; M } side by side
    $rows = [ordered]@{
        'Verdict'           = { param($c) $c.Verdict }
        'Statement #'       = { param($c) $c.M.Index + 1 }
        'Actual plan'       = { param($c) if ($c.M.IsActual) { 'yes' } else { 'no (estimated)' } }
        'Elapsed ms'        = { param($c) if ($null -ne $c.M.ElapsedMs) { '{0:N0}' -f $c.M.ElapsedMs } else { '-' } }
        'CPU ms'            = { param($c) if ($null -ne $c.M.CpuMs) { '{0:N0}' -f $c.M.CpuMs } else { '-' } }
        'Logical reads'     = { param($c) if ($null -ne $c.M.LogicalReads) { '{0:N0}' -f $c.M.LogicalReads } else { '-' } }
        'Estimated cost'    = { param($c) if ($null -ne $c.M.Cost) { '{0:N3}' -f $c.M.Cost } else { '-' } }
        'DOP'               = { param($c) $c.M.Dop }
        'CE model'          = { param($c) if ($c.M.CE) { $c.M.CE } else { '-' } }
        'Memory grant KB'   = { param($c) if ($null -ne $c.M.GrantKB) { '{0:N0} (used {1:N0})' -f $c.M.GrantKB, $c.M.UsedKB } else { '-' } }
        'Spills'            = { param($c) $c.M.Spills }
        'Worst misestimate' = { param($c) if ($null -ne $c.M.WorstMisestimate) { '{0:N0}x' -f $c.M.WorstMisestimate } else { '-' } }
        'Query plan hash'   = { param($c) $c.M.PlanHash }
        'Compiled values'   = { param($c) (@($c.M.Params | Where-Object { $_.Compiled } | ForEach-Object { "$($_.Name)=$(ConvertFrom-CompiledValue $_.Compiled)" }) -join ', ') }
        'ARITHABORT'        = { param($c) $c.M.ArithAbort }
    }
    $w = 26
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(''.PadRight(19) + (($Cols | ForEach-Object { ([string]$_.Title).PadRight($w) }) -join ''))
    foreach ($k in $rows.Keys) {
        $line = $k.PadRight(19)
        foreach ($c in $Cols) { $v = [string](& $rows[$k] $c); if ($v.Length -gt $w - 2) { $v = $v.Substring(0, $w - 3) + '~' }; $line += $v.PadRight($w) }
        [void]$sb.AppendLine($line.TrimEnd())
    }
    foreach ($c in $Cols) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$($c.Title) shape: $($c.M.Shape)")
        if ($c.M.Flags.Count) { foreach ($f in $c.M.Flags) { [void]$sb.AppendLine("  ! $f") } } else { [void]$sb.AppendLine('  no warnings found') }
    }
    $sb.ToString()
}

function Get-StatementFields($M) {
    # statement / @params / module from a measured plan statement, without judging anything
    $text = $M.Text; $params = ''; $type = 'SQL'; $module = ''
    $pref = ConvertFrom-ParamPrefix $text
    if ($pref) { $text = $pref.Stmt; $params = $pref.Params }
    $obj = if ($M.Psp.Removed) { [long]$M.Psp.ObjectId } else { 0 }
    if ($M.Module) { $type = 'OBJECT'; $module = $M.Module; $params = '' }
    elseif ($obj -gt 0) { $type = 'OBJECT'; $params = '' }
    elseif (-not $params) {
        $real = @($M.Params | Where-Object { $_.Compiled -and $_.Type })
        if ($real.Count) { $params = ($real | ForEach-Object { "$($_.Name) $($_.Type)" }) -join ',' }
    }
    @{ Statement = $text; Params = $params; Type = $type; Module = $module; PspObjectId = $obj
       PspPlanHash = $(if ($M.Psp.Removed) { $M.PlanHash } else { $null }); QueryHash = $M.QueryHash }
}

function Resolve-PlanGuideFromPlans {
    <#
      Decides how to build the guide from saved plans.
        Hint given                      -> query hint (statement taken from the plan)
        No hint + confirmed good plan   -> USE PLAN with the good plan's XML
        No hint + only a bad plan       -> nothing to force: suggestions only
      Labels: 'Auto' (let the tool judge), 'Good', 'Bad'.
    #>
    [CmdletBinding()]
    param([string]$PlanA, [string]$LabelA = 'Auto', [string]$PlanB, [string]$LabelB = 'Auto',
          [string]$Hint, [int]$IndexA = -1, [switch]$TrustLabels)
    $F = New-Object System.Collections.Generic.List[object]
    $res = [pscustomobject]@{ Findings = $F; Decision = 'None'; Options = $null; Report = ''; Suggestions = @(); Good = $null; Bad = $null; GoodXml = $null; GoodIndex = -1; Primary = $null }
    if ([string]::IsNullOrWhiteSpace($PlanA)) { Add-Finding $F 'Error' 'Plans: load at least one saved plan.'; return $res }
    $pa = Read-ShowPlan $PlanA.Trim()
    if ($pa.Error) { Add-Finding $F 'Error' ('Plan 1: ' + $pa.Error); return $res }
    $ia = Select-PlanStatement $pa $IndexA $null $null
    $A = Measure-PlanStatement $pa $ia
    $res.Primary = $A
    $haveB = -not [string]::IsNullOrWhiteSpace($PlanB)
    $B = $null
    if ($haveB) {
        $pb = Read-ShowPlan $PlanB.Trim()
        if ($pb.Error) { Add-Finding $F 'Error' ('Plan 2: ' + $pb.Error); return $res }
        $ib = Select-PlanStatement $pb -1 $A.QueryHash $null
        if ($ib -lt 0) { $ib = Select-PlanStatement $pb -1 $null $A.Text }
        if ($ib -lt 0) { Add-Finding $F 'Error' 'Plans: plan 2 does not contain the same statement as plan 1 (no matching query hash or text). Pick the same query in both, or check that you saved the right plans.'; return $res }
        $B = Measure-PlanStatement $pb $ib
        if ((Get-Squashed $A.Text) -cne (Get-Squashed $B.Text)) { Add-Finding $F 'Warn' 'Plans: same query hash, but the statement text differs (literals or whitespace). The text comes from the good plan; check it matches what the app sends.' }
    }
    $hasHint = -not [string]::IsNullOrWhiteSpace($Hint)
    $pspPlans = @(@($A, $B) | Where-Object { $_ -and $_.Psp.Removed })
    if ($pspPlans.Count) {
        $ids = ($pspPlans | ForEach-Object { $_.Psp.VariantId }) -join ' and '
        Add-Finding $F 'Warn' ("Plans: Parameter Sensitive Plan optimization is on for this query (query variant $ids). The engine-generated OPTION (PLAN PER VALUE ...) was removed from the statement, because the app sends the parent query without it.")
        if (@($pspPlans | Where-Object { $_.Psp.Predicate -like 'optional_predicate*' }).Count) {
            Add-Finding $F 'Warn' 'Plans: this is an Optional Parameter Plan (OPPO, SQL Server 2025) variant. OPPO can rewrite the query (for example drop a WHERE branch for @p IS NULL), so this text may not be what the app sends. Replace it with the real text: run "Check on server", then "Use cached text".'
        }
        if ($B -and $A.Psp.Removed -and $B.Psp.Removed -and $A.Psp.VariantId -ne $B.Psp.VariantId) {
            Add-Finding $F 'Info' 'Plans: the two plans are different PSP variants of the same parent query, so SQL Server already keeps a separate plan per parameter range. A plan guide replaces that with one plan (or one hint) for every value. If only one range is slow, fixing just that variant in Query Store is the narrower fix (the script ends with the commands).'
        }
    }
    if (@(@($A, $B) | Where-Object { $_ -and $_.IsDispatcher }).Count) { Add-Finding $F 'Info' 'Plans: this is a PSP dispatcher plan. It only routes to variants; save the actual plan of a run to get the variant that really executed.' }

    # ---- decide which plan is good ----
    $goodSide = $null; $verdict = @{ A = ''; B = '' }
    if ($haveB) {
        if ($LabelA -eq $LabelB -and $LabelA -ne 'Auto') { Add-Finding $F 'Error' "Plans: both plans are labelled $LabelA. Label one Good and the other Bad, or let the tool judge."; return $res }
        $ev = Compare-PlanEvidence $A $B
        $labelSide = if ($LabelA -eq 'Good' -or $LabelB -eq 'Bad') { 'A' } elseif ($LabelB -eq 'Good' -or $LabelA -eq 'Bad') { 'B' } else { $null }
        if ($A.PlanHash -and $A.PlanHash -eq $B.PlanHash) {
            Add-Finding $F 'Warn' 'Plans: both plans have the same query plan hash, so they are the same plan shape. Forcing it changes nothing; the difference is more likely parameter values, blocking, a cold cache or SET options.'
        }
        if ($labelSide) {
            if ($ev.Winner -in 'A', 'B' -and $ev.Winner -ne $labelSide -and $ev.Strength -eq 'Strong') {
                $msg = "Plans: your labels say plan $(if ($labelSide -eq 'A') {1} else {2}) is good, but the numbers say the opposite: $($ev.Reason)."
                if ($TrustLabels) { Add-Finding $F 'Warn' ($msg + ' Using your labels, as asked.') }
                else { Add-Finding $F 'Error' ($msg + ' Check you saved the right files, or tick "Use my labels even if the numbers disagree".'); $res.Report = Format-PlanReport @(@{ Title = 'Plan 1'; M = $A; Verdict = $LabelA }, @{ Title = 'Plan 2'; M = $B; Verdict = $LabelB }); return $res }
            } elseif ($ev.Winner -eq $labelSide) { Add-Finding $F 'OK' "Plans: the numbers agree with your labels ($($ev.Reason))." }
            else { Add-Finding $F 'Info' "Plans: using your labels. Evidence: $($ev.Reason)." }
            $goodSide = $labelSide
        } else {
            if ($ev.Winner -eq 'Similar') { Add-Finding $F 'Error' "Plans: can't tell which plan is better ($($ev.Reason)). Label them Good and Bad."; }
            elseif ($ev.Strength -eq 'Weak') { Add-Finding $F 'Error' "Plans: only estimates to go on ($($ev.Reason)). Save actual execution plans, or label them Good and Bad." }
            else { $goodSide = $ev.Winner; Add-Finding $F 'OK' "Plans: plan $(if ($goodSide -eq 'A') {1} else {2}) is the good one: $($ev.Reason)." }
        }
        if ($goodSide -eq 'A') { $res.Good = $A; $res.Bad = $B; $res.GoodXml = $PlanA; $verdict.A = 'GOOD'; $verdict.B = 'bad' }
        elseif ($goodSide -eq 'B') { $res.Good = $B; $res.Bad = $A; $res.GoodXml = $PlanB; $verdict.A = 'bad'; $verdict.B = 'GOOD' }
        else { $verdict.A = '?'; $verdict.B = '?' }
        $res.Report = Format-PlanReport @(@{ Title = 'Plan 1'; M = $A; Verdict = $verdict.A }, @{ Title = 'Plan 2'; M = $B; Verdict = $verdict.B })
    } else {
        $flags = $A.Flags.Count
        switch ($LabelA) {
            'Good' {
                $res.Good = $A; $res.GoodXml = $PlanA; $verdict.A = 'GOOD (your label)'
                if ($flags) { Add-Finding $F 'Warn' ('Plans: you marked this plan Good, but it has warnings: ' + ($A.Flags -join '; ') + '. Make sure it really is the plan you want to force.') }
                if (-not $A.IsActual) { Add-Finding $F 'Info' 'Plans: this is an estimated plan, so nothing proves it runs well. An actual plan from a good run is safer evidence.' }
            }
            'Bad' { $res.Bad = $A; $verdict.A = 'bad (your label)' }
            default {
                $verdict.A = if ($flags) { "looks bad ($flags flag(s))" } else { 'no red flags' }
                if (-not $hasHint) { Add-Finding $F 'Error' ('Plans: one plan on its own can''t prove it is the good one' + $(if ($flags) { ' (and this one has warnings: ' + ($A.Flags -join '; ') + ')' } else { '' }) + '. Label it Good or Bad, add the other plan, or give a hint.') }
            }
        }
        $res.Report = Format-PlanReport @(@{ Title = 'Plan 1'; M = $A; Verdict = $verdict.A })
    }

    $res.Suggestions = Get-HintSuggestions $res.Good $res.Bad
    if ($res.Good) { $res.Primary = $res.Good } elseif ($res.Bad) { $res.Primary = $res.Bad }
    if (-not $res.Good -and -not $res.Bad -and -not $hasHint) { return $res }

    # ---- where the T-SQL comes from: good plan if known ----
    $src = if ($res.Good) { $res.Good } elseif ($res.Bad) { $res.Bad } else { $A }
    $text = $src.Text; $params = $null; $type = 'SQL'; $module = ''
    $pref = ConvertFrom-ParamPrefix $text
    if ($pref) { $text = $pref.Stmt; $params = $pref.Params }
    $pspObj = if ($src.Psp.Removed) { [long]$src.Psp.ObjectId } else { 0 }
    if (-not $src.Module -and $pspObj -gt 0) {
        Add-Finding $F 'Warn' ("Plans: the PSP clause says ObjectID = $pspObj, which means this statement runs inside a module (procedure, function or trigger), not as ad hoc or sp_executesql SQL. A SQL guide would never match. Find the module with SELECT OBJECT_SCHEMA_NAME($pspObj) + '.' + OBJECT_NAME($pspObj) in that database, then enter it under 'Inside a module' (the server check confirms the statement is in it).")
        $type = 'OBJECT'
    }
    if ($type -eq 'OBJECT' -and -not $src.Module) { }
    elseif ($src.Module) {
        $type = 'OBJECT'; $module = $src.Module; $params = $null
        Add-Finding $F 'Info' "Plans: the statement runs inside $module, so this is an OBJECT guide on that module."
    } elseif (-not $params) {
        $real  = @($src.Params | Where-Object { $_.Compiled })
        $local = @($src.Params | Where-Object { -not $_.Compiled -and $_.Runtime })
        if ($real.Count) {
            if (@($real | Where-Object { -not $_.Type }).Count) { Add-Finding $F 'Warn' 'Plans: the plan lists parameters without data types (older build). Fill in @params by hand, exactly as the app declares them.' }
            else { $params = ($real | ForEach-Object { "$($_.Name) $($_.Type)" }) -join ','; Add-Finding $F 'Warn' "Plans: @params was rebuilt from the plan as '$params'. It must match the app's declaration string exactly (order, spacing, nvarchar lengths). Confirm it with Profiler/XE or the server check." }
        }
        if ($local.Count) {
            Add-Finding $F 'Warn' ('Plans: ' + (($local | ForEach-Object { $_.Name }) -join ', ') + ' look like local variables (a runtime value but no compiled value), so this plan probably came from a test batch in SSMS. The app most likely sends them as parameters, which compile differently, and the guide only matches the app''s exact text. Use the app''s statement and @params.')
        }
    }
    if ($src.ParameterizedText) { Add-Finding $F 'Info' 'Plans: SQL Server auto-parameterized this statement. The app sends literals, so a Template guide may fit better than a SQL guide.' }
    if ($text.Length -ge 4000) { Add-Finding $F 'Warn' 'Plans: the statement text is very long, and plan files can truncate it. Run the server check to confirm the text.' }
    if ($res.Good -and $res.Bad -and $res.Good.ArithAbort -and $res.Bad.ArithAbort -and $res.Good.ArithAbort -ne $res.Bad.ArithAbort) {
        Add-Finding $F 'Info' "Plans: ARITHABORT differs (good: $($res.Good.ArithAbort), bad: $($res.Bad.ArithAbort)). That is the classic 'fast in SSMS, slow in the app' split: two cache entries for the same text. A plan guide applies to both."
    }

    # ---- the rule ----
    if ($hasHint) {
        $res.Decision = 'Hint'
        Add-Finding $F 'OK' 'Plans: a hint was given, so the guide uses the hint. The statement text comes from the plan.'
    } elseif ($res.Good) {
        if ($res.Bad -and $res.Good.PlanHash -and $res.Good.PlanHash -eq $res.Bad.PlanHash) { Add-Finding $F 'Error' 'Plans: forcing the good plan would force the same plan shape as the bad one. Use a hint instead.'; return $res }
        $res.Decision = 'Xml'; $res.GoodIndex = $res.Good.Index
        Add-Finding $F 'OK' 'Plans: no hint was given and the good plan is confirmed, so the guide forces the good plan''s XML (USE PLAN).'
    } else {
        Add-Finding $F 'Error' ('Plans: only a bad plan, and no hint. A bad plan can''t be forced. Add the good plan, or pick a hint' + $(if ($res.Suggestions.Count) { ' (suggestions are listed)' } else { '' }) + '.')
        return $res
    }
    $res.Options = @{
        Statement = $text; Params = $(if ($params) { $params } else { '' }); Type = $type; Module = $module
        Source = $res.Decision; Hint = $(if ($hasHint) { $Hint } else { '' })
        PlanXml = $(if ($res.Decision -eq 'Xml') { $res.GoodXml } else { '' }); PlanStatementIndex = $res.GoodIndex
        LineBreak = 'Keep'
        PspObjectId = $pspObj; QueryHash = $src.QueryHash
        PspPlanHash = $(if ($res.Decision -eq 'Hint' -and $res.Bad -and $res.Bad.Psp.Removed) { $res.Bad.PlanHash } elseif ($src.Psp.Removed) { $src.PlanHash } else { $null })
    }
    Add-Finding $F 'Info' 'Plans: run "Check on server" to confirm this text matches what the app really sends.'
    $res
}
#endregion

#region ---------- read-only server check ----------
function Test-ReadOnlySql([string]$Sql) {
    $code = Get-CodeOnly $Sql
    $bad = '\b(INSERT|UPDATE|DELETE|MERGE|ALTER|CREATE|DROP|TRUNCATE|EXEC|EXECUTE|GRANT|REVOKE|DENY|BACKUP|RESTORE|KILL|SHUTDOWN|WAITFOR|BEGIN|COMMIT|ROLLBACK|SAVE|DBCC|OPENROWSET|OPENQUERY|OPENDATASOURCE|OPENXML|INTO|RECONFIGURE|USE|GOTO)\b|\b(sp|xp)_\w+'
    -not [regex]::IsMatch($code, $bad, 'IgnoreCase')
}

function Get-ServerCheckQueries($R) {
    # $R = result of New-PlanGuideScript
    $qs = New-Object System.Collections.Generic.List[object]
    $qs.Add(@{ Key = 'info'; Params = @{ '@module' = $R.Module }; Sql = @'
SELECT @@SERVERNAME AS server_name,
       CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128)) AS version,
       CAST(SERVERPROPERTY('Edition') AS nvarchar(128)) AS edition,
       DB_NAME() AS db_name,
       (SELECT is_parameterization_forced FROM sys.databases WHERE database_id = DB_ID()) AS forced,
       (SELECT TOP (1) rs.role_desc FROM sys.databases AS d
          JOIN sys.dm_hadr_availability_replica_states AS rs ON rs.replica_id = d.replica_id AND rs.is_local = 1
         WHERE d.database_id = DB_ID()) AS ag_role,
       HAS_PERMS_BY_NAME(DB_NAME(), 'DATABASE', 'ALTER') AS can_alter_db,
       HAS_PERMS_BY_NAME(NULL, NULL, 'VIEW SERVER STATE') AS can_view_state
'@ })
    $qs.Add(@{ Key = 'busy'; Params = @{}; Sql = @'
SELECT AVG(CAST(runnable_tasks_count AS float)) AS avg_runnable,
       (SELECT COUNT(*) FROM sys.dm_exec_requests WHERE blocking_session_id <> 0) AS blocked
FROM sys.dm_os_schedulers WHERE status = N'VISIBLE ONLINE'
'@ })
    if ($R.Method -eq 'QueryStore') {
        $qs.Add(@{ Key = 'qsstate'; Params = @{}; Sql = 'SELECT actual_state_desc, desired_state_desc, readonly_reason FROM sys.database_query_store_options' })
        if ($R.Source -eq 'Xml' -and $R.PlanHash) {
            $qs.Add(@{ Key = 'qsplan'; Params = @{ '@ph' = $R.PlanHash }; Sql = 'SELECT p.query_id, p.plan_id, p.is_forced_plan, p.last_execution_time FROM sys.query_store_plan AS p WHERE p.query_plan_hash = CONVERT(binary(8), @ph, 1)' })
        } elseif ($R.PspHash) {
            $qs.Add(@{ Key = 'qsparent'; Params = @{ '@ph' = $R.PspHash }; Sql = 'SELECT DISTINCT v.parent_query_id AS query_id FROM sys.query_store_plan AS p JOIN sys.query_store_query_variant AS v ON v.query_variant_query_id = p.query_id WHERE p.query_plan_hash = CONVERT(binary(8), @ph, 1)' })
        } elseif ($R.QueryHash) {
            $qs.Add(@{ Key = 'qsmatch'; Params = @{ '@qh' = $R.QueryHash }; Sql = 'SELECT q.query_id, q.object_id, q.last_execution_time FROM sys.query_store_query AS q WHERE q.query_hash = CONVERT(binary(8), @qh, 1)' })
            $qs.Add(@{ Key = 'qshints'; Params = @{ '@qh' = $R.QueryHash }; Sql = 'SELECT h.query_id, h.query_hint_text FROM sys.query_store_query_hints AS h JOIN sys.query_store_query AS q ON q.query_id = h.query_id WHERE q.query_hash = CONVERT(binary(8), @qh, 1)' })
        } else {
            $qs.Add(@{ Key = 'qsmatch'; SkipIfBusy = $true; Params = @{ '@stmt' = $R.Statement; '@params' = $R.Params }; Sql = @'
SELECT q.query_id, q.object_id, q.last_execution_time
FROM sys.query_store_query AS q JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id
WHERE qt.query_sql_text = @stmt COLLATE Latin1_General_BIN2
   OR qt.query_sql_text = N'(' + ISNULL(@params, N'') + N')' + @stmt COLLATE Latin1_General_BIN2
'@ })
        }
        return $qs
    }
    $qs.Add(@{ Key = 'byname'; Params = @{ '@name' = $R.Name }; Sql = @'
SELECT name, scope_type_desc, is_disabled, create_date, modify_date FROM sys.plan_guides WHERE name = @name
'@ })
    if ($R.Type -ne 'TEMPLATE') {
        $qs.Add(@{ Key = 'bytext'; Params = @{ '@stmt' = $R.Statement; '@name' = $R.Name }; Sql = @'
SELECT name, scope_type_desc, is_disabled
FROM sys.plan_guides
WHERE DATALENGTH(query_text) = DATALENGTH(@stmt) AND query_text = @stmt COLLATE Latin1_General_BIN2 AND name <> @name
'@ })
    }
    if ($R.Type -eq 'SQL') {
        $qs.Add(@{ Key = 'cache'; SkipIfBusy = $true; Params = @{ '@cached' = $R.CachedKey }; Sql = @'
SELECT TOP (5) qs.execution_count, qs.last_execution_time, DB_NAME(st.dbid) AS db
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS st
WHERE DATALENGTH(st.text) = DATALENGTH(@cached) AND st.text = @cached COLLATE Latin1_General_BIN2
ORDER BY qs.last_execution_time DESC
OPTION (MAXDOP 1)
'@ })
        # loose pattern: the first tokens, any whitespace between them
        $tokens = @(($R.Statement -split '\s+') | Where-Object { $_ } | Select-Object -First 8)
        $esc = $tokens | ForEach-Object { $_.Replace('\', '\\').Replace('%', '\%').Replace('_', '\_').Replace('[', '\[') }
        $pattern = '%' + ($esc -join '%') + '%'
        $qs.Add(@{ Key = 'near'; SkipIfBusy = $true; OnlyIfEmpty = 'cache'; Params = @{ '@pattern' = $pattern }; Sql = @'
SELECT DISTINCT TOP (5) st.text, DB_NAME(st.dbid) AS db
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS st
WHERE st.text LIKE @pattern ESCAPE N'\'
OPTION (MAXDOP 1)
'@ })
    }
    if ($R.Type -eq 'OBJECT') {
        $qs.Add(@{ Key = 'module'; Params = @{ '@module' = $R.Module }; Sql = @'
SELECT o.type, o.type_desc, m.definition, HAS_PERMS_BY_NAME(@module, 'OBJECT', 'ALTER') AS can_alter
FROM sys.objects AS o LEFT JOIN sys.sql_modules AS m ON m.object_id = o.object_id
WHERE o.object_id = OBJECT_ID(@module)
'@ })
    }
    $qs
}

# Runs on a background runspace: data access only, no dependencies on the functions above.
$script:ServerCheckWorker = {
    param([string]$ConnectionString, $Queries)
    $out = @{}
    $prefix = "SET LOCK_TIMEOUT 3000; SET DEADLOCK_PRIORITY LOW; SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;`r`n"
    $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    try {
        $conn.Open()
        $busy = $false
        foreach ($q in $Queries) {
            if ($q.SkipIfBusy -and $busy) { $out[$q.Key] = 'skipped: the server is busy'; continue }
            if ($q.OnlyIfEmpty) {
                $prev = $out[$q.OnlyIfEmpty]
                if (-not ($prev -is [System.Data.DataTable]) -or $prev.Rows.Count -gt 0) { continue }
            }
            try {
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = $prefix + $q.Sql
                $cmd.CommandTimeout = 30
                foreach ($k in $q.Params.Keys) {
                    $p = $cmd.Parameters.Add($k, [System.Data.SqlDbType]::NVarChar, -1)
                    $p.Value = if ($null -eq $q.Params[$k]) { [DBNull]::Value } else { [string]$q.Params[$k] }
                }
                $dt = New-Object System.Data.DataTable
                $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
                [void]$da.Fill($dt)
                $out[$q.Key] = $dt
                if ($q.Key -eq 'busy' -and $dt.Rows.Count) {
                    $ar = $dt.Rows[0]['avg_runnable']; $bl = $dt.Rows[0]['blocked']
                    if (($ar -isnot [DBNull] -and [double]$ar -gt 2) -or ($bl -isnot [DBNull] -and [int]$bl -gt 0)) { $busy = $true }
                }
            } catch { $out[$q.Key] = 'error: ' + $_.Exception.Message }
        }
    } catch { $out['_connect'] = $_.Exception.Message }
    finally { $conn.Dispose() }
    $out
}

function ConvertTo-ServerFindings($Out, $R, [bool]$DropExisting) {
    $F = New-Object System.Collections.Generic.List[object]
    $tag = 'Server: '
    if ($Out.ContainsKey('_connect')) { Add-Finding $F 'Error' ($tag + 'could not connect. ' + $Out['_connect']); return ,$F }
    $isDt = { param($v) $v -is [System.Data.DataTable] }

    $info = $Out['info']
    if (& $isDt $info) {
        $row = $info.Rows[0]
        Add-Finding $F 'OK' ("{0}connected to {1} ({2}, {3}), database {4}." -f $tag, $row.server_name, $row.version, $row.edition, $row.db_name)
        if ($row.ag_role -isnot [DBNull] -and $row.ag_role -ne 'PRIMARY') { Add-Finding $F 'Error' ($tag + "this database is an AG $($row.ag_role). Create the guide on the primary.") }
        if ($R.Type -ne 'OBJECT' -and $row.can_alter_db -isnot [DBNull] -and [int]$row.can_alter_db -eq 0) { Add-Finding $F 'Warn' ($tag + 'your login lacks ALTER on the database. Creating SQL and TEMPLATE guides needs it, so whoever runs the script will need more rights.') }
        if ($row.can_view_state -isnot [DBNull] -and [int]$row.can_view_state -eq 0) { Add-Finding $F 'Warn' ($tag + 'no VIEW SERVER STATE, so the plan-cache check cannot see anything.') }
        if ($R.Type -eq 'TEMPLATE') {
            $forced = ($row.forced -isnot [DBNull] -and [bool]$row.forced)
            $want = if ($R.Script -match 'PARAMETERIZATION FORCED') { 'FORCED' } else { 'SIMPLE' }
            if ($want -eq 'FORCED' -and $forced) { Add-Finding $F 'Warn' ($tag + 'the database already uses FORCED parameterization, so a FORCED template guide changes nothing.') }
            elseif ($want -eq 'SIMPLE' -and -not $forced) { Add-Finding $F 'Warn' ($tag + 'the database uses SIMPLE parameterization, so a SIMPLE template guide changes nothing.') }
            else { Add-Finding $F 'OK' ($tag + "the database parameterization is $(if ($forced) { 'FORCED' } else { 'SIMPLE' }), so the template guide has an effect.") }
        }
    } elseif ($info) { Add-Finding $F 'Warn' ($tag + "server info: $info") }

    if ($R.Method -eq 'QueryStore') {
        $st = $Out['qsstate']
        if ((& $isDt $st) -and $st.Rows.Count) {
            $s0 = $st.Rows[0]
            if ($s0.actual_state_desc -ne 'READ_WRITE') { Add-Finding $F 'Error' ($tag + "Query Store is $($s0.actual_state_desc) (desired $($s0.desired_state_desc)). It must be READ_WRITE to force plans or set hints.") }
            else { Add-Finding $F 'OK' ($tag + 'Query Store is READ_WRITE.') }
        } elseif ($st) { Add-Finding $F 'Warn' ($tag + "Query Store state: $st") }
        $pl = $Out['qsplan']
        if (& $isDt $pl) {
            if ($pl.Rows.Count) { $ids = ($pl.Rows | ForEach-Object { "query_id $($_.query_id) / plan_id $($_.plan_id)$(if ([bool]$_.is_forced_plan) { ' (already forced)' })" }) -join ', '; Add-Finding $F 'OK' ($tag + "the good plan is in Query Store: $ids.") }
            else { Add-Finding $F 'Error' ($tag + 'the good plan is not in this database''s Query Store, so it can''t be forced there. Use the plan guide output, or let the good plan run here first.') }
        } elseif ($pl) { Add-Finding $F 'Warn' ($tag + "plan lookup: $pl") }
        foreach ($k in 'qsmatch', 'qsparent') {
            $m2 = $Out[$k]
            if (& $isDt $m2) {
                if ($m2.Rows.Count) { Add-Finding $F 'OK' ($tag + "$(if ($k -eq 'qsparent') { 'PSP parent query' } else { 'matching query' })_id(s) in Query Store: " + (($m2.Rows | ForEach-Object { $_.query_id }) -join ', ') + '.') }
                else { Add-Finding $F 'Warn' ($tag + 'no matching query in Query Store yet. It is captured after it runs (subject to the capture mode); run the script after that.') }
            } elseif ($m2) {
                if ($m2 -match 'Invalid object name') { Add-Finding $F 'Error' ($tag + 'this server has no sys.query_store_query_variant / hints views: Query Store hints and PSP need SQL Server 2022+.') }
                else { Add-Finding $F 'Warn' ($tag + "$k`: $m2") }
            }
        }
        $h2 = $Out['qshints']
        if ((& $isDt $h2) -and $h2.Rows.Count) { Add-Finding $F 'Warn' ($tag + 'existing Query Store hint(s) will be replaced: ' + (($h2.Rows | ForEach-Object { "query_id $($_.query_id): $($_.query_hint_text)" }) -join '; ') + '. Combine them into one OPTION (...) if you want to keep them.') }
        elseif ($h2 -is [string] -and $h2 -match 'Invalid object name') { Add-Finding $F 'Error' ($tag + 'sys.query_store_query_hints does not exist: Query Store hints need SQL Server 2022+.') }
        return ,$F
    }
    $bn = $Out['byname']
    if ((& $isDt $bn) -and $bn.Rows.Count) {
        $row = $bn.Rows[0]
        if ($DropExisting) { Add-Finding $F 'Info' ($tag + "a guide named $($R.Name) already exists ($($row.scope_type_desc), created $($row.create_date)). The script replaces it.") }
        else { Add-Finding $F 'Error' ($tag + "a guide named $($R.Name) already exists. Tick ""Replace a guide with the same name"" or pick another name.") }
    } elseif (-not (& $isDt $bn) -and $bn) { Add-Finding $F 'Warn' ($tag + "name check: $bn") }

    $bt = $Out['bytext']
    if ((& $isDt $bt) -and $bt.Rows.Count) {
        foreach ($row in $bt.Rows) {
            $state = if ([bool]$row.is_disabled) { 'disabled' } else { 'ENABLED' }
            Add-Finding $F $(if ([bool]$row.is_disabled) { 'Info' } else { 'Warn' }) ($tag + "guide $($row.name) ($($row.scope_type_desc), $state) already targets this exact statement text. SQL Server refuses a second enabled guide on the same statement and scope, so disable or drop it first.")
        }
    }

    $c = $Out['cache']
    if ($R.Type -eq 'SQL') {
        if (& $isDt $c) {
            if ($c.Rows.Count) {
                $row = $c.Rows[0]
                Add-Finding $F 'OK' ($tag + "exact text found in the plan cache ($($row.execution_count) executions, last $($row.last_execution_time)). The guide will match.")
                $dbs = @($c.Rows | ForEach-Object { $_.db } | Where-Object { $_ -isnot [DBNull] -and $_ } | Select-Object -Unique)
                if ($R.Database -and $dbs.Count -and ($dbs -notcontains $R.Database)) {
                    Add-Finding $F 'Warn' ($tag + "the cached query ran in database " + ($dbs -join ', ') + ". A plan guide only applies to queries run in the database where it is created.")
                }
            } else {
                $n = $Out['near']
                if ((& $isDt $n) -and $n.Rows.Count) {
                    $k = $R.CachedKey
                    foreach ($row in $n.Rows) {
                        $t = [string]$row.text
                        $at = Get-FirstDifference $k $t
                        if ((Convert-LineBreak $t 'LF') -ceq (Convert-LineBreak $k 'LF')) {
                            $kind = Get-LineBreakKind $t
                            Add-Finding $F 'Warn' ($tag + "the cache has this statement with $kind line breaks. Set Line breaks to $kind.")
                        } elseif ((Get-Squashed $t) -ceq (Get-Squashed $k)) {
                            Add-Finding $F 'Warn' ($tag + "the cache has this statement with different whitespace. First difference at character $($at + 1): yours  [" + (Get-DiffSnippet $k $at) + "]   cache  [" + (Get-DiffSnippet $t $at) + "]. Use ""Use cached text"" or fix it by hand.")
                        } else {
                            Add-Finding $F 'Warn' ($tag + "similar text in the cache, first difference at character $($at + 1): yours  [" + (Get-DiffSnippet $k $at) + "]   cache  [" + (Get-DiffSnippet $t $at) + "].")
                        }
                    }
                    $script:LastNearText = [string]$n.Rows[0].text
                } elseif ($n -is [string] -and $n) {
                    Add-Finding $F 'Warn' ($tag + "near-match search: $n")
                } else {
                    Add-Finding $F 'Info' ($tag + 'the text is not in the plan cache, and nothing similar is. That may be fine if the query has not run since the last restart or cache flush; otherwise compare it with what the app sends (Profiler / XE rpc_completed).')
                }
            }
        } elseif ($c) { Add-Finding $F 'Info' ($tag + "plan-cache check $c.") }
    }

    $m = $Out['module']
    if ($R.Type -eq 'OBJECT') {
        if (& $isDt $m) {
            if (-not $m.Rows.Count) { Add-Finding $F 'Error' ($tag + "module $($R.Module) was not found in this database.") }
            else {
                $row = $m.Rows[0]
                if (@('P', 'FN', 'TF', 'TR') -notcontains ([string]$row.type).Trim()) { Add-Finding $F 'Error' ($tag + "$($R.Module) is a $($row.type_desc). Plan guides need a T-SQL procedure, scalar function, multi-statement TVF or DML trigger.") }
                elseif ($row.definition -is [DBNull]) { Add-Finding $F 'Error' ($tag + 'no definition visible: the module is encrypted, not T-SQL, or you lack VIEW DEFINITION.') }
                else {
                    $def = [string]$row.definition
                    $pos = $def.IndexOf($R.Statement, [StringComparison]::Ordinal)
                    if ($pos -ge 0) {
                        $line = ([regex]::Matches($def.Substring(0, $pos), '\n')).Count + 1
                        Add-Finding $F 'OK' ($tag + "statement found exactly in $($R.Module), line $line.")
                    } else {
                        $lfS = Convert-LineBreak $R.Statement 'LF'; $lfD = Convert-LineBreak $def 'LF'
                        if ($lfD.IndexOf($lfS, [StringComparison]::Ordinal) -ge 0) {
                            Add-Finding $F 'Warn' ($tag + "the statement is in the module, but with $(Get-LineBreakKind $def) line breaks. Set Line breaks to match.")
                        } else {
                            # find where the first line starts and show the first difference from there
                            $first = (($R.Statement -split '\r?\n')[0]).Trim()
                            $at0 = if ($first.Length -ge 6) { $def.IndexOf($first, [StringComparison]::OrdinalIgnoreCase) } else { -1 }
                            if ($at0 -ge 0) {
                                $lead = $R.Statement.IndexOf($first, [StringComparison]::Ordinal); if ($lead -lt 0) { $lead = 0 }
                                $startAt = [Math]::Max(0, $at0 - $lead)
                                $seg = $def.Substring($startAt, [Math]::Min($R.Statement.Length + 40, $def.Length - $startAt))
                                $d = Get-FirstDifference $R.Statement $seg
                                Add-Finding $F 'Error' ($tag + "not an exact match. First difference at character $($d + 1): yours  [" + (Get-DiffSnippet $R.Statement $d) + "]   module  [" + (Get-DiffSnippet $seg $d) + "].")
                            } else {
                                Add-Finding $F 'Error' ($tag + "the statement was not found in $($R.Module). Copy it from the module definition (sp_helptext or Script As > ALTER).")
                            }
                        }
                    }
                    if ($row.can_alter -isnot [DBNull] -and [int]$row.can_alter -eq 0) { Add-Finding $F 'Warn' ($tag + 'your login lacks ALTER on the module, which creating an OBJECT guide needs.') }
                }
            }
        } elseif ($m) { Add-Finding $F 'Warn' ($tag + "module check: $m") }
    }
    ,$F
}
#endregion

if ($NoGui) { return }

#region ---------- GUI ----------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if ($HideConsole) {
    try {
        Add-Type -Name ConsoleWin -Namespace PlanGuideGen -MemberDefinition '[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow(); [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int n);'
        [void][PlanGuideGen.ConsoleWin]::ShowWindow([PlanGuideGen.ConsoleWin]::GetConsoleWindow(), 0)
    } catch { }
}

$mono   = New-Object System.Drawing.Font('Consolas', 9.5)
$ui     = New-Object System.Drawing.Font('Segoe UI', 9)
$uiBold = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grey   = [System.Drawing.Color]::FromArgb(95, 107, 120)

function New-Ctl([string]$Type, $Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text = '') {
    $c = New-Object "System.Windows.Forms.$Type"
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    if ($W -gt 0) { $c.Width = $W }; if ($H -gt 0) { $c.Height = $H }
    if ($Text) { $c.Text = $Text }
    if ($Parent) { [void]$Parent.Controls.Add($c) }
    $c
}
function New-Note($Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Text) {
    $l = New-Ctl Label $Parent $X $Y $W $H $Text; $l.ForeColor = $grey; $l
}
function New-Box($Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [bool]$Multi = $false) {
    $t = New-Ctl TextBox $Parent $X $Y $W $H
    $t.Font = $mono
    if ($Multi) {
        $t.Multiline = $true; $t.AcceptsReturn = $true; $t.AcceptsTab = $true; $t.WordWrap = $false
        $t.ScrollBars = 'Both'; $t.Height = $H
        $t.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $this.SelectAll(); $_.SuppressKeyPress = $true } })
    }
    $t
}
function New-Group($Parent, [string]$Text) {
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text; $g.Font = $uiBold; $g.Width = 590
    $g.AutoSize = $true; $g.AutoSizeMode = 'GrowAndShrink'
    $g.Margin = New-Object System.Windows.Forms.Padding(6, 4, 6, 6)
    $g.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 8)
    [void]$Parent.Controls.Add($g)
    $g
}
function New-SubPanel($Parent, [int]$Y) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Location = New-Object System.Drawing.Point(8, $Y)
    $p.AutoSize = $true; $p.AutoSizeMode = 'GrowAndShrink'; $p.Font = $ui
    [void]$Parent.Controls.Add($p)
    $p
}

$tip = New-Object System.Windows.Forms.ToolTip
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Plan Guide Generator'
$form.Size = New-Object System.Drawing.Size(1320, 900)
$form.StartPosition = 'CenterScreen'
$form.Font = $ui
$form.MinimumSize = New-Object System.Drawing.Size(1000, 600)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
[void]$form.Controls.Add($split)

$left = New-Object System.Windows.Forms.FlowLayoutPanel
$left.Dock = 'Fill'; $left.FlowDirection = 'TopDown'; $left.WrapContents = $false; $left.AutoScroll = $true
$left.Padding = New-Object System.Windows.Forms.Padding(4)
[void]$split.Panel1.Controls.Add($left)

# ---- 0. Start from saved plans ----
$g0 = New-Group $left '0. Start from saved plans (optional)'
$p0 = New-SubPanel $g0 20
[void](New-Note $p0 0 0 565 32 'Load the saved .sqlplan of the good run and/or the bad run. The tool checks which is which, takes the T-SQL from the plan and fills the form below.')
$script:PlanFiles = @{ A = $null; B = $null }
$plRows = @{}
foreach ($k in 'A', 'B') {
    $y = if ($k -eq 'A') { 36 } else { 68 }
    [void](New-Ctl Label $p0 0 ($y + 5) 85 18 $(if ($k -eq 'A') { 'Plan 1' } else { 'Plan 2 (opt.)' }))
    $bo = New-Ctl Button $p0 88 $y 62 26 'Open...'
    $bp = New-Ctl Button $p0 154 $y 52 26 'Paste'
    $bx = New-Ctl Button $p0 210 $y 26 26 'x'
    $lf = New-Note $p0 242 ($y + 5) 180 18 '(none)'
    $cl = New-Ctl ComboBox $p0 428 ($y + 2) 137 23; $cl.DropDownStyle = 'DropDownList'
    [void]$cl.Items.AddRange(@('Let the tool judge', 'Good', 'Bad')); $cl.SelectedIndex = 0
    $plRows[$k] = @{ Open = $bo; Paste = $bp; Clear = $bx; File = $lf; Label = $cl }
}
[void](New-Note $p0 0 102 565 18 'Hint (optional). Empty = force the good plan''s XML. Filled = use this hint.')
$txtPlanHint = New-Box $p0 0 122 565 40 $true
$cmbSuggest = New-Ctl ComboBox $p0 0 168 565 23; $cmbSuggest.DropDownStyle = 'DropDownList'
[void]$cmbSuggest.Items.Add('Suggested hints appear here after Analyze...'); $cmbSuggest.SelectedIndex = 0
$chkTrustLabels = New-Ctl CheckBox $p0 0 196 400 22 'Use my labels even if the numbers disagree'
$btnAnalyze = New-Ctl Button $p0 0 222 200 30 'Analyze and fill the form'
$btnAnalyze.Font = $uiBold
$lblAnalyze = New-Note $p0 210 222 355 34 ''
$txtReport = New-Box $p0 0 258 565 230 $true; $txtReport.ReadOnly = $true; $txtReport.Font = New-Object System.Drawing.Font('Consolas', 8.5); $txtReport.Visible = $false
$tip.SetToolTip($chkTrustLabels, 'When your Good/Bad labels contradict the runtime numbers, the tool stops unless this is ticked.')

# ---- 1. Statement ----
$g1 = New-Group $left '1. Statement'
$p1 = New-SubPanel $g1 20
[void](New-Note $p1 0 0 565 32 'Exactly as the application sends it. You can paste an exec sp_executesql N''...'', N''@p int'', ... call, or plan-cache text that starts with (@p int).')
$txtQuery = New-Box $p1 0 34 565 180 $true
$lblStats = New-Note $p1 0 218 330 18 ''
$btnExample = New-Ctl Button $p1 345 214 105 25 'Demo query'
$btnClear   = New-Ctl Button $p1 460 214 105 25 'Clear'
$pnlDetect = New-Ctl Panel $p1 0 244 565 32
$pnlDetect.BackColor = [System.Drawing.Color]::FromArgb(232, 239, 252); $pnlDetect.Visible = $false
$lblDetect = New-Ctl Label $pnlDetect 6 8 440 18 ''
$btnSplit  = New-Ctl Button $pnlDetect 455 3 104 26 'Split it'

# ---- 2. Where it runs ----
$g2 = New-Group $left '2. Where the statement runs'
$p2 = New-SubPanel $g2 20
$rbSql = New-Ctl RadioButton $p2 0 0 175 22 'Ad hoc / sp_executesql'; $rbSql.Checked = $true
$rbObj = New-Ctl RadioButton $p2 180 0 140 22 'Inside a module'
$rbTpl = New-Ctl RadioButton $p2 325 0 230 22 'Template (parameterization)'
$lblTypeHelp = New-Note $p2 0 24 565 34 ''
$pnlSql = New-SubPanel $p2 60; $pnlSql.Location = New-Object System.Drawing.Point(0, 60)
[void](New-Note $pnlSql 0 0 565 18 '@params: the exact declaration string, e.g. @CustomerID int. Empty for a plain ad hoc batch.')
$txtParams = New-Box $pnlSql 0 20 565 23
$chkBatch = New-Ctl CheckBox $pnlSql 0 48 400 22 'The statement is one part of a larger batch'
$txtBatch = New-Box $pnlSql 0 72 565 100 $true; $txtBatch.Visible = $false
$pnlObj = New-SubPanel $p2 60; $pnlObj.Location = New-Object System.Drawing.Point(0, 60)
[void](New-Note $pnlObj 0 0 565 18 'Procedure, scalar function, multi-statement TVF or DML trigger, e.g. dbo.usp_GetOrders')
$txtModule = New-Box $pnlObj 0 20 565 23
$pnlTpl = New-SubPanel $p2 60; $pnlTpl.Location = New-Object System.Drawing.Point(0, 60)
[void](New-Ctl Label $pnlTpl 0 4 110 18 'Parameterization')
$cmbParamz = New-Ctl ComboBox $pnlTpl 115 0 120 23; $cmbParamz.DropDownStyle = 'DropDownList'; [void]$cmbParamz.Items.AddRange(@('FORCED', 'SIMPLE')); $cmbParamz.SelectedIndex = 0
$chkTplHint = New-Ctl CheckBox $pnlTpl 0 30 450 22 'Also apply a query hint to the parameterized form'

# ---- 3. What to force ----
$g3 = New-Group $left '3. What to force'
$p3 = New-SubPanel $g3 20
$pnlSrc = New-Ctl Panel $p3 0 0 565 26
$rbHint = New-Ctl RadioButton $pnlSrc 0 2 120 22 'Query hint'; $rbHint.Checked = $true
$rbXml  = New-Ctl RadioButton $pnlSrc 125 2 200 22 'Plan XML (USE PLAN)'
$lblTplNote = New-Note $p3 0 0 565 30 'Template guides only take PARAMETERIZATION. Tick "Also apply a query hint" above to add a hint for the parameterized form.'
$pnlHint = New-SubPanel $p3 32; $pnlHint.Location = New-Object System.Drawing.Point(0, 32)
[void](New-Note $pnlHint 0 0 565 18 'With or without the OPTION ( ) wrapper.')
$txtHint = New-Box $pnlHint 0 20 565 48 $true
$cmbQuick = New-Ctl ComboBox $pnlHint 0 74 300 23; $cmbQuick.DropDownStyle = 'DropDownList'
[void]$cmbQuick.Items.Add('Add a common hint...'); [void]$cmbQuick.Items.AddRange($script:QuickHints); $cmbQuick.SelectedIndex = 0
$pnlXml = New-SubPanel $p3 32; $pnlXml.Location = New-Object System.Drawing.Point(0, 32)
[void](New-Note $pnlXml 0 6 420 18 'Estimated or actual showplan XML.')
$btnOpenPlan = New-Ctl Button $pnlXml 445 0 120 25 'Open .sqlplan...'
$txtPlan = New-Box $pnlXml 0 30 565 130 $true
$chkTrim  = New-Ctl CheckBox $pnlXml 0 164 250 22 'Keep only the chosen statement'; $chkTrim.Checked = $true
$chkStrip = New-Ctl CheckBox $pnlXml 260 164 250 22 'Strip actual runtime stats'
$lblPick = New-Ctl Label $pnlXml 0 192 200 18 'Statement in the plan'
$cmbPick = New-Ctl ComboBox $pnlXml 0 212 565 23; $cmbPick.DropDownStyle = 'DropDownList'; $cmbPick.Font = $mono
$btnUsePlan = New-Ctl Button $pnlXml 0 240 260 25 'Use this statement''s text as the query'
$lblMeta = New-Ctl Label $pnlXml 0 270 565 105 ''; $lblMeta.Font = $mono; $lblMeta.ForeColor = $grey

# ---- 4. Script options ----
$g4 = New-Group $left '4. Script options'
$p4 = New-SubPanel $g4 20
[void](New-Ctl Label $p4 0 0 180 18 'Guide name')
$txtName = New-Box $p4 0 20 185 23
[void](New-Ctl Label $p4 195 0 180 18 'Database (USE + check)')
$txtDb = New-Box $p4 195 20 180 23
[void](New-Ctl Label $p4 385 0 180 18 'Expected @@SERVERNAME')
$txtSrvName = New-Box $p4 385 20 180 23
[void](New-Ctl Label $p4 0 52 200 18 'Line breaks in the statement')
$cmbEol = New-Ctl ComboBox $p4 0 72 300 23; $cmbEol.DropDownStyle = 'DropDownList'
[void]$cmbEol.Items.AddRange(@('As pasted', 'CRLF (Windows apps, SSMS)', 'LF only')); $cmbEol.SelectedIndex = 0
$chkEncode = New-Ctl CheckBox $p4 0 102 560 22 'Write line breaks as NCHAR() (survives copy, paste and editors)'; $chkEncode.Checked = $true
$chkDrop   = New-Ctl CheckBox $p4 0 126 280 22 'Replace a guide with the same name'; $chkDrop.Checked = $true
$chkLatch  = New-Ctl CheckBox $p4 290 126 270 22 'Safety latch'; $chkLatch.Checked = $true
$chkMatch  = New-Ctl CheckBox $p4 0 150 560 22 'Include the text check against the plan cache / module'; $chkMatch.Checked = $true
[void](New-Ctl Label $p4 0 180 70 18 'Generate')
$pnlMethod = New-Ctl Panel $p4 70 176 495 26
$rbMethodPG = New-Ctl RadioButton $pnlMethod 0 2 110 22 'Plan guide'; $rbMethodPG.Checked = $true
$rbMethodQS = New-Ctl RadioButton $pnlMethod 115 2 380 22 'Query Store (hint, or force the good plan)'

# ---- 5. Server check ----
$g5 = New-Group $left '5. Check on a server (optional, read-only)'
$p5 = New-SubPanel $g5 20
[void](New-Ctl Label $p5 0 0 180 18 'Server')
$txtConn = New-Box $p5 0 20 250 23
[void](New-Ctl Label $p5 260 0 120 18 'Authentication')
$cmbAuth = New-Ctl ComboBox $p5 260 20 140 23; $cmbAuth.DropDownStyle = 'DropDownList'; [void]$cmbAuth.Items.AddRange(@('Windows', 'SQL login')); $cmbAuth.SelectedIndex = 0
$chkEncrypt = New-Ctl CheckBox $p5 410 10 150 20 'Encrypt'
$chkTrust   = New-Ctl CheckBox $p5 410 30 155 20 'Trust server certificate'
[void](New-Ctl Label $p5 0 50 120 18 'Login')
$txtUser = New-Box $p5 0 70 180 23; $txtUser.Enabled = $false
[void](New-Ctl Label $p5 190 50 120 18 'Password')
$txtPwd = New-Ctl TextBox $p5 190 70 180 23; $txtPwd.UseSystemPasswordChar = $true; $txtPwd.Enabled = $false
$btnCheck = New-Ctl Button $p5 0 102 150 28 'Check on server'
$btnUseCached = New-Ctl Button $p5 158 102 130 28 'Use cached text'; $btnUseCached.Enabled = $false
$lblCheck = New-Note $p5 296 108 270 18 ''
$tip.SetToolTip($btnUseCached, 'Replace the query (and @params) with the closest text found in the plan cache, so the guide matches what really runs.')
[void](New-Note $p5 0 136 565 46 'Runs SELECTs only, through a read-only guard: the exact text in the plan cache (skipped while the server is busy) or the module, guides with the same name or text, AG role, parameterization and your rights. Uses the Database from Script options.')

# ---- right side ----
$top = New-Object System.Windows.Forms.Panel
$top.Dock = 'Top'; $top.Height = 40
$btnCopy = New-Ctl Button $top 6 7 110 27 'Copy script'
$btnSave = New-Ctl Button $top 122 7 110 27 'Save .sql...'
$lblSummary = New-Ctl Label $top 245 12 700 20 ''
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabScript = New-Object System.Windows.Forms.TabPage 'Script'
$tabChecks = New-Object System.Windows.Forms.TabPage 'Checks'
[void]$tabs.TabPages.Add($tabScript); [void]$tabs.TabPages.Add($tabChecks)
$txtScript = New-Object System.Windows.Forms.TextBox
$txtScript.Dock = 'Fill'; $txtScript.Multiline = $true; $txtScript.ReadOnly = $true; $txtScript.WordWrap = $false
$txtScript.ScrollBars = 'Both'; $txtScript.Font = $mono; $txtScript.BackColor = [System.Drawing.Color]::White
$txtScript.Add_KeyDown({ if ($_.Control -and $_.KeyCode -eq 'A') { $this.SelectAll(); $_.SuppressKeyPress = $true } })
[void]$tabScript.Controls.Add($txtScript)
$rtbChecks = New-Object System.Windows.Forms.RichTextBox
$rtbChecks.Dock = 'Fill'; $rtbChecks.ReadOnly = $true; $rtbChecks.BackColor = [System.Drawing.Color]::White
$rtbChecks.Font = New-Object System.Drawing.Font('Segoe UI', 9.5); $rtbChecks.BorderStyle = 'None'
[void]$tabChecks.Controls.Add($rtbChecks)
[void]$split.Panel2.Controls.Add($tabs)
[void]$split.Panel2.Controls.Add($top)

$tip.SetToolTip($chkEncode, 'Writes each line break as NCHAR(13)+NCHAR(10) (or NCHAR(10)), so editors and clipboards cannot change it.')
$tip.SetToolTip($chkLatch, 'Adds a RAISERROR + SET NOEXEC ON line that must be deleted by hand before the script does anything.')
$tip.SetToolTip($cmbEol, 'Plan guides match character for character, line breaks included. "As pasted" keeps whatever is in the box.')

# ---- state ----
$script:ManualName = $false
$script:SettingName = $false
$script:PendingSplit = $null
$script:PlanSigShown = $null
$script:UpdatingPick = $false
$script:Last = $null
$script:ServerFindings = $null
$script:ServerSig = $null
$script:Bg = $null
$script:LastNearText = $null

function Get-Options {
    @{
        Statement = $txtQuery.Text
        Type = $(if ($rbObj.Checked) { 'OBJECT' } elseif ($rbTpl.Checked) { 'TEMPLATE' } else { 'SQL' })
        Params = $txtParams.Text; UseBatch = $chkBatch.Checked; Batch = $txtBatch.Text; Module = $txtModule.Text
        Parameterization = [string]$cmbParamz.SelectedItem; TemplateHint = $chkTplHint.Checked
        Source = $(if ($rbXml.Checked) { 'Xml' } else { 'Hint' }); Hint = $txtHint.Text; PlanXml = $txtPlan.Text
        PlanStatementIndex = $(if ($cmbPick.Items.Count) { $cmbPick.SelectedIndex } else { -1 })
        TrimPlan = $chkTrim.Checked; StripRuntime = $chkStrip.Checked
        Name = $(if ($script:ManualName) { $txtName.Text } else { '' })
        Database = $txtDb.Text; ExpectedServer = $txtSrvName.Text
        LineBreak = @('Keep', 'CRLF', 'LF')[$cmbEol.SelectedIndex]
        EncodeLineBreaks = $chkEncode.Checked; DropExisting = $chkDrop.Checked; SafetyLatch = $chkLatch.Checked; MatchCheck = $chkMatch.Checked
        PspPlanHash = $script:PspPlanHash; PspObjectId = $script:PspObjectId
        Method = $(if ($rbMethodQS.Checked) { 'QueryStore' } else { 'PlanGuide' }); QueryHash = $script:PlanQueryHash
    }
}

function Update-Visibility {
    $form.SuspendLayout()
    $isTpl = $rbTpl.Checked
    $pnlSql.Visible = $rbSql.Checked; $pnlObj.Visible = $rbObj.Checked; $pnlTpl.Visible = $isTpl
    $txtBatch.Visible = $chkBatch.Checked
    $pnlSrc.Visible = -not $isTpl
    $lblTplNote.Visible = $isTpl -and -not $chkTplHint.Checked
    $pnlHint.Visible = ($isTpl -and $chkTplHint.Checked) -or (-not $isTpl -and $rbHint.Checked)
    $pnlXml.Visible = (-not $isTpl) -and $rbXml.Checked
    $hasPlan = $cmbPick.Items.Count -gt 0
    $lblPick.Visible = $hasPlan; $cmbPick.Visible = $hasPlan; $btnUsePlan.Visible = $hasPlan; $lblMeta.Visible = $hasPlan
    $lblTypeHelp.Text = $(if ($rbObj.Checked) { 'The statement is inside a stored procedure, function or trigger. Paste just that one statement, exactly as it appears in the module.' }
        elseif ($isTpl) { 'Paste one example of the query with literal values. The script calls sp_get_query_template to build the template, then forces (or blocks) parameterization for every query of that shape.' }
        else { 'The statement arrives as an ad hoc batch, or through sp_executesql / a prepared call (then fill in @params).' })
    $txtUser.Enabled = $cmbAuth.SelectedIndex -eq 1; $txtPwd.Enabled = $cmbAuth.SelectedIndex -eq 1
    $form.ResumeLayout()
}

function Update-Detect {
    $t = $txtQuery.Text
    $ex = ConvertFrom-SpExecuteSql $t
    $pp = if (-not $ex) { ConvertFrom-ParamPrefix $t } else { $null }
    $script:PendingSplit = if ($ex) { $ex } else { $pp }
    $pnlDetect.Visible = [bool]$script:PendingSplit
    if ($ex) { $lblDetect.Text = 'This is an sp_executesql call. Split it into the statement and @params?' }
    elseif ($pp) { $s = $pp.Params; if ($s.Length -gt 50) { $s = $s.Substring(0, 50) + '...' }; $lblDetect.Text = "Starts with a parameter list. Move ($s) into @params?" }
    $q = $t
    if ($q) {
        $bits = @("$($q.Length) chars", "$(([regex]::Matches($q, '\r\n|\r|\n')).Count + 1) lines", "line breaks: $(Get-LineBreakKind $q)")
        if ($q -match "`t") { $bits += 'tabs' }
        if ($q -match '[ \t]+(\r?\n|$)') { $bits += 'trailing spaces' }
        if ($q -match ';\s*$') { $bits += 'ends with ;' }
        $lblStats.Text = $bits -join '  |  '
    } else { $lblStats.Text = '' }
}

function Get-ServerSig($R) { Get-ShortHash ($R.Method + '|' + $R.Source + '|' + $R.Type + '|' + $R.Statement + '|' + $R.Params + '|' + $R.Batch + '|' + $R.Module + '|' + $R.Name + '|' + $R.Database + '|' + $txtConn.Text) }

function Show-Findings($List) {
    $order = @{ Error = 0; Warn = 1; Info = 2; OK = 3 }
    $sorted = @($List | Sort-Object { $order[$_.Severity] })
    if (-not $sorted.Count) { $sorted = @([pscustomobject]@{ Severity = 'OK'; Message = 'No problems found.' }) }
    $colors = @{ Error = [System.Drawing.Color]::FromArgb(180, 35, 24); Warn = [System.Drawing.Color]::FromArgb(150, 95, 0); Info = [System.Drawing.Color]::FromArgb(36, 87, 197); OK = [System.Drawing.Color]::FromArgb(31, 122, 63) }
    $labels = @{ Error = 'ERROR'; Warn = 'CHECK'; Info = 'NOTE'; OK = 'OK' }
    $rtbChecks.SuspendLayout()
    $rtbChecks.Clear()
    $bold = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    $norm = New-Object System.Drawing.Font('Segoe UI', 9.5)
    foreach ($f in $sorted) {
        $rtbChecks.SelectionStart = $rtbChecks.TextLength; $rtbChecks.SelectionLength = 0
        $rtbChecks.SelectionFont = $bold; $rtbChecks.SelectionColor = $colors[$f.Severity]
        $rtbChecks.AppendText($labels[$f.Severity].PadRight(7))
        $rtbChecks.SelectionFont = $norm; $rtbChecks.SelectionColor = [System.Drawing.Color]::Black
        $rtbChecks.AppendText($f.Message + "`r`n`r`n")
    }
    $rtbChecks.SelectionStart = 0
    $rtbChecks.ResumeLayout()
    $e = @($List | Where-Object { $_.Severity -eq 'Error' }).Count
    $w = @($List | Where-Object { $_.Severity -eq 'Warn' }).Count
    $tabChecks.Text = 'Checks' + $(if ($e -or $w) { " ($e error$(if ($e -ne 1) {'s'}), $w to check)" } else { '' })
    $lblSummary.Text = $(if ($e) { "$e error(s): the script is not complete. See Checks." } elseif ($w) { "Script ready. $w item(s) to check first." } else { 'Script ready.' })
    $lblSummary.ForeColor = $(if ($e) { $colors.Error } elseif ($w) { $colors.Warn } else { $colors.OK })
}

function Update-Output {
    try {
        Update-Visibility; Update-Detect
        $r = New-PlanGuideScript -Options (Get-Options)
        $script:Last = $r
        if (-not $script:ManualName) { $script:SettingName = $true; $txtName.Text = $r.Name; $script:SettingName = $false }
        if ($r.PlanSig -and $r.PlanSig -ne $script:PlanSigShown) {
            $script:UpdatingPick = $true
            $cmbPick.Items.Clear(); [void]$cmbPick.Items.AddRange([object[]]$r.PlanLabels)
            if ($r.PlanIndex -ge 0) { $cmbPick.SelectedIndex = $r.PlanIndex }
            $script:PlanSigShown = $r.PlanSig; $script:UpdatingPick = $false
            if ([string]::IsNullOrWhiteSpace($txtQuery.Text) -or $script:ExampleLoaded) {
                $pl = Read-ShowPlan $txtPlan.Text.Trim()
                if (-not $pl.Error -and $r.PlanIndex -ge 0) {
                    Set-FormFromFields (Get-StatementFields (Measure-PlanStatement $pl $r.PlanIndex))
                    $lblSummary.Text = 'Statement filled from the plan.'
                    $debounce.Stop(); $debounce.Start()
                }
            }
        } elseif (-not $r.PlanSig -and $cmbPick.Items.Count -and [string]::IsNullOrWhiteSpace($txtPlan.Text)) { $cmbPick.Items.Clear(); $script:PlanSigShown = $null }
        $lblMeta.Text = $r.PlanMeta
        Update-Visibility
        $first = $txtScript.GetCharIndexFromPosition((New-Object System.Drawing.Point(2, 2)))
        $txtScript.Text = $r.Script
        if ($first -gt 0 -and $first -lt $txtScript.TextLength) { $txtScript.SelectionStart = $first; $txtScript.ScrollToCaret(); $txtScript.SelectionLength = 0 }
        $all = New-Object System.Collections.Generic.List[object]
        foreach ($f in $r.Findings) { $all.Add($f) }
        if ($script:PlanFindings) { foreach ($f in $script:PlanFindings) { $all.Add($f) } }
        if ($script:ServerFindings -and $script:ServerSig -eq (Get-ServerSig $r)) { foreach ($f in $script:ServerFindings) { $all.Add($f) } }
        elseif ($script:ServerFindings) { $script:ServerFindings = $null; $lblCheck.Text = 'Inputs changed: run the server check again.' }
        Show-Findings $all
    } catch {
        $txtScript.Text = "-- Internal error while building the script:`r`n-- $($_.Exception.Message)`r`n-- $($_.InvocationInfo.PositionMessage)"
    }
}

$script:LogPath = Join-Path ([System.IO.Path]::GetTempPath()) 'PlanGuideGenerator.log'
function Show-GuiError([string]$Where, $Err) {
    $msg = "$Where failed: $($Err.Exception.Message)`r`n$($Err.InvocationInfo.PositionMessage)"
    try { Add-Content -Path $script:LogPath -Value ("{0:yyyy-MM-dd HH:mm:ss}  {1}`r`n{2}`r`n" -f (Get-Date), $msg, $Err.ScriptStackTrace) } catch { }
    [void][System.Windows.Forms.MessageBox]::Show($form, $msg + "`r`n`r`nDetails were written to $($script:LogPath)", 'Plan Guide Generator', 'OK', 'Error')
}
function Set-FormFromFields($f) {
    # fill the statement part of the form from Get-StatementFields output
    $txtQuery.Text = $f.Statement
    $txtParams.Text = [string]$f.Params
    $txtModule.Text = [string]$f.Module
    if ($f.Type -eq 'OBJECT') { $rbObj.Checked = $true } else { $rbSql.Checked = $true }
    $chkBatch.Checked = $false
    $cmbEol.SelectedIndex = 0
    $script:PspPlanHash = $f.PspPlanHash; $script:PspObjectId = [long]$f.PspObjectId; $script:PlanQueryHash = $f.QueryHash
    if ($script:ExampleLoaded -and $txtHint.Text -eq $script:ExampleHint) { $txtHint.Text = '' }
    $script:ExampleLoaded = $false
    $script:ManualName = $false; $script:SettingName = $true; $txtName.Text = ''; $script:SettingName = $false
}
$script:ExampleLoaded = $false
$script:ExampleHint = 'OPTION (OPTIMIZE FOR (@CustomerID UNKNOWN))'
$debounce = New-Object System.Windows.Forms.Timer
$debounce.Interval = 250
$debounce.Add_Tick({ $debounce.Stop(); Update-Output })
$request = { $debounce.Stop(); $debounce.Start() }
$requestNow = { Update-Visibility; $debounce.Stop(); $debounce.Start() }

foreach ($c in @($txtQuery, $txtParams, $txtBatch, $txtModule, $txtHint, $txtPlan, $txtDb, $txtSrvName)) { $c.Add_TextChanged($request) }
foreach ($c in @($rbMethodPG, $rbMethodQS, $rbSql, $rbObj, $rbTpl, $rbHint, $rbXml, $chkBatch, $chkTplHint, $chkTrim, $chkStrip, $chkEncode, $chkDrop, $chkLatch, $chkMatch)) { $c.Add_CheckedChanged($requestNow) }
foreach ($c in @($cmbParamz, $cmbEol, $cmbAuth)) { $c.Add_SelectedIndexChanged($requestNow) }
$cmbPick.Add_SelectedIndexChanged({ if (-not $script:UpdatingPick) { & $request } })
$txtName.Add_TextChanged({ if (-not $script:SettingName) { $script:ManualName = -not [string]::IsNullOrWhiteSpace($txtName.Text); & $request } })

$btnSplit.Add_Click({
    try {
        if (-not $script:PendingSplit) { return }
        $s = $script:PendingSplit
        $txtQuery.Text = $s.Stmt
        if ($null -ne $s.Params) { $txtParams.Text = $s.Params; $rbSql.Checked = $true }
        Update-Output
    } catch { Show-GuiError 'btnSplit' $_ }
})
$cmbQuick.Add_SelectionChangeCommitted({
    if ($cmbQuick.SelectedIndex -le 0) { return }
    $h = [string]$cmbQuick.SelectedItem
    $cur = $txtHint.Text.Trim()
    $m = [regex]::Match($cur, '^option\s*\(([\s\S]*)\)\s*;?$', 'IgnoreCase')
    $inner = if ($m.Success) { $m.Groups[1].Value.Trim() } else { $cur }
    $txtHint.Text = 'OPTION (' + $(if ($inner) { $inner + ', ' } else { '' }) + $h + ')'
    $cmbQuick.SelectedIndex = 0
})
$btnOpenPlan.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Execution plans (*.sqlplan;*.xml)|*.sqlplan;*.xml|All files (*.*)|*.*'
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $txtPlan.Text = [System.IO.File]::ReadAllText($dlg.FileName)
        $rbXml.Checked = $true
    }
})
$btnUsePlan.Add_Click({
    try {
        $pl = Read-ShowPlan $txtPlan.Text.Trim()
        if ($pl.Error) { $lblSummary.Text = $pl.Error; return }
        $idx = [Math]::Max(0, $cmbPick.SelectedIndex)
        Set-FormFromFields (Get-StatementFields (Measure-PlanStatement $pl $idx))
        Update-Output
    } catch { Show-GuiError 'btnUsePlan' $_ }
})
$btnExample.Add_Click({
    $txtQuery.Text = "exec sp_executesql N'SELECT o.OrderID, o.OrderDate, c.Name`r`nFROM dbo.Orders AS o`r`nJOIN dbo.Customers AS c ON c.CustomerID = o.CustomerID`r`nWHERE o.CustomerID = @CustomerID',N'@CustomerID int',@CustomerID=42"
    $txtHint.Text = $script:ExampleHint
    if (-not $txtDb.Text) { $txtDb.Text = 'Sales' }
    $rbSql.Checked = $true; $rbHint.Checked = $true
    $script:ManualName = $false
    Update-Output
    $script:ExampleLoaded = $true
})
$btnClear.Add_Click({
    foreach ($c in @($txtQuery, $txtParams, $txtBatch, $txtModule, $txtHint, $txtPlan)) { $c.Text = '' }
    $script:ManualName = $false; $script:SettingName = $true; $txtName.Text = ''; $script:SettingName = $false
    $cmbPick.Items.Clear(); $script:PlanSigShown = $null; $script:ServerFindings = $null; $lblCheck.Text = ''
    Set-PlanFile 'A' $null ''; Set-PlanFile 'B' $null ''; $txtPlanHint.Text = ''; $txtReport.Visible = $false; $script:SuggestList = @()
    Update-Output
})
$btnCopy.Add_Click({
    if ($txtScript.Text) { [System.Windows.Forms.Clipboard]::SetText($txtScript.Text); $lblSummary.Text = 'Script copied to the clipboard.' }
})
$btnSave.Add_Click({
    if (-not $script:Last) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'SQL script (*.sql)|*.sql'; $dlg.FileName = ($script:Last.Name -replace '[\\/:*?"<>|]', '_') + '.sql'
    if ($dlg.ShowDialog($form) -eq 'OK') {
        [System.IO.File]::WriteAllText($dlg.FileName, $txtScript.Text, (New-Object System.Text.UTF8Encoding $true))
        $lblSummary.Text = "Saved $($dlg.FileName)"
    }
})

# ---- server check (background) ----
$poll = New-Object System.Windows.Forms.Timer
$poll.Interval = 200
$poll.Add_Tick({
    $bg = $script:Bg
    if (-not $bg) { $poll.Stop(); return }
    if (-not $bg.Handle.IsCompleted) { $lblCheck.Text = 'Checking... {0:N0} s  (click Cancel to stop)' -f ((Get-Date) - $bg.Started).TotalSeconds; return }
    $poll.Stop()
    try {
        $res = $bg.PS.EndInvoke($bg.Handle)
        $out = if ($res -and $res.Count) { $res[0].psobject.BaseObject } else { @{ '_connect' = 'no result returned' } }
        $script:LastNearText = $null
        $script:ServerFindings = ConvertTo-ServerFindings $out $bg.R $chkDrop.Checked
        $script:ServerSig = $bg.Sig
        $e = @($script:ServerFindings | Where-Object { $_.Severity -eq 'Error' }).Count
        $lblCheck.Text = ('Checked at {0:HH:mm:ss}. {1}' -f (Get-Date), $(if ($e) { "$e problem(s): see Checks." } else { 'Results are on the Checks tab.' }))
        $btnUseCached.Enabled = [bool]$script:LastNearText
        $tabs.SelectedTab = $tabChecks
    } catch {
        $lblCheck.Text = $(if ($bg.Cancelled) { 'Cancelled.' } else { 'Check failed: ' + $_.Exception.Message })
    } finally {
        $bg.PS.Dispose(); $script:Bg = $null; $btnCheck.Text = 'Check on server'
        Update-Output
    }
})
$btnCheck.Add_Click({
    try {
        if ($script:Bg) { $script:Bg.Cancelled = $true; [void]$script:Bg.PS.BeginStop($null, $null); $lblCheck.Text = 'Cancelling...'; return }
        Update-Output
        $r = $script:Last
        if (-not $txtConn.Text.Trim()) { $lblCheck.Text = 'Enter the server to connect to.'; return }
        if (-not $txtDb.Text.Trim()) { $lblCheck.Text = 'Enter the Database under Script options.'; return }
        if (-not $r -or [string]::IsNullOrWhiteSpace($r.Statement)) { $lblCheck.Text = 'Paste the statement first.'; return }
        $queries = @(Get-ServerCheckQueries $r)
        foreach ($q in $queries) { if (-not (Test-ReadOnlySql $q.Sql)) { $lblCheck.Text = "Read-only guard refused query '$($q.Key)'. Nothing was sent."; return } }
        $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $b['Data Source'] = $txtConn.Text.Trim(); $b['Initial Catalog'] = $txtDb.Text.Trim()
        $b['Application Name'] = 'PlanGuideGenerator (read-only)'; $b['Connect Timeout'] = 15; $b['Max Pool Size'] = 2
        if ($chkEncrypt.Checked) { $b['Encrypt'] = $true }
        if ($chkTrust.Checked) { $b['TrustServerCertificate'] = $true }
        if ($cmbAuth.SelectedIndex -eq 1) { $b['User ID'] = $txtUser.Text; $b['Password'] = $txtPwd.Text } else { $b['Integrated Security'] = $true }
        $ps = [PowerShell]::Create()
        [void]$ps.AddScript($script:ServerCheckWorker.ToString()).AddArgument($b.ConnectionString).AddArgument($queries)
        $script:Bg = @{ PS = $ps; Handle = $ps.BeginInvoke(); Started = Get-Date; R = $r; Sig = (Get-ServerSig $r); Cancelled = $false }
        $btnCheck.Text = 'Cancel'; $lblCheck.Text = 'Connecting...'
        $poll.Start()
    } catch { Show-GuiError 'btnCheck' $_ }
})

$btnUseCached.Add_Click({
    try {
        if (-not $script:LastNearText) { return }
        $t = $script:LastNearText
        $pp = ConvertFrom-ParamPrefix $t
        if ($pp) { $txtParams.Text = $pp.Params; $t = $pp.Stmt } else { $txtParams.Text = '' }
        $txtQuery.Text = $t; $cmbEol.SelectedIndex = 0; $btnUseCached.Enabled = $false
        $lblCheck.Text = 'Query replaced with the cached text. Run the check again to confirm.'
        Update-Output
    } catch { Show-GuiError 'btnUseCached' $_ }
})

# ---- start from saved plans ----
$script:PlanFindings = $null
$script:PspPlanHash = $null
$script:PspObjectId = 0
$script:PlanQueryHash = $null
$script:SuggestList = @()
function Set-PlanFile([string]$Key, [string]$Xml, [string]$Label) {
    $script:PlanFiles[$Key] = $Xml
    $script:PlanFindings = $null; $script:PspPlanHash = $null; $script:PspObjectId = 0; $script:PlanQueryHash = $null
    $plRows[$Key].File.Text = if ($Xml) { '{0} ({1:N0} KB)' -f $Label, ([Text.Encoding]::Unicode.GetByteCount($Xml) / 1024) } else { '(none)' }
    $lblAnalyze.Text = ''
}
foreach ($k in 'A', 'B') {
    $row = $plRows[$k]
    $row.Open.Tag = $k; $row.Paste.Tag = $k; $row.Clear.Tag = $k
    $row.Open.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'Execution plans (*.sqlplan;*.xml)|*.sqlplan;*.xml|All files (*.*)|*.*'
        if ($dlg.ShowDialog($form) -eq 'OK') { Set-PlanFile $this.Tag ([System.IO.File]::ReadAllText($dlg.FileName)) ([System.IO.Path]::GetFileName($dlg.FileName)) }
    })
    $row.Paste.Add_Click({
        $t = [System.Windows.Forms.Clipboard]::GetText()
        if ($t -match '<ShowPlanXML') { Set-PlanFile $this.Tag $t 'pasted plan' } else { $lblAnalyze.Text = 'The clipboard does not hold showplan XML.' }
    })
    $row.Clear.Add_Click({ Set-PlanFile $this.Tag $null '' })
}
$cmbSuggest.Add_SelectionChangeCommitted({
    $i = $cmbSuggest.SelectedIndex - 1
    if ($i -ge 0 -and $i -lt $script:SuggestList.Count) {
        $txtPlanHint.Text = 'OPTION (' + $script:SuggestList[$i].Hint + ')'
        Invoke-Analyze
    }
})
function Invoke-Analyze {
    try {
        $labelOf = { param($c) @('Auto', 'Good', 'Bad')[$c.SelectedIndex] }
        $pa = $script:PlanFiles.A; $pb = $script:PlanFiles.B
        if (-not $pa -and -not $pb -and $txtPlan.Text.Trim()) { $pa = $txtPlan.Text; $plRows.A.File.Text = '(from section 3)' }
        if (-not $pa -and $pb) { $pa = $pb; $pb = $null; $la = & $labelOf $plRows.B.Label; $lb = 'Auto' }
        else { $la = & $labelOf $plRows.A.Label; $lb = & $labelOf $plRows.B.Label }
        if (-not $pa) { $lblAnalyze.Text = 'Load a plan first (Open... or Paste).'; return }
        try {
            $d = Resolve-PlanGuideFromPlans -PlanA $pa -LabelA $la -PlanB $pb -LabelB $lb -Hint $txtPlanHint.Text -TrustLabels:$chkTrustLabels.Checked
        } catch {
            $lblAnalyze.Text = 'Analysis failed. See the message.'
            [void][System.Windows.Forms.MessageBox]::Show($form, "The plan could not be analyzed:`r`n`r`n$($_.Exception.Message)`r`n`r`n$($_.InvocationInfo.PositionMessage)", 'Plan Guide Generator', 'OK', 'Error')
            return
        }
        $script:PlanFindings = $d.Findings
        $txtReport.Text = ($d.Report -replace '(?<!\r)\n', "`r`n"); $txtReport.Visible = [bool]$d.Report
        $script:SuggestList = @($d.Suggestions)
        $cmbSuggest.Items.Clear()
        [void]$cmbSuggest.Items.Add($(if ($script:SuggestList.Count) { "$($script:SuggestList.Count) suggested hint(s): pick one to use it..." } else { 'No hint suggestions for these plans.' }))
        foreach ($sg in $script:SuggestList) { [void]$cmbSuggest.Items.Add($sg.Hint + '   -- ' + $sg.Why) }
        $cmbSuggest.SelectedIndex = 0
        $script:PspPlanHash = $null
        $filledOnly = $false
        if (-not $d.Options -and $d.Primary) {
            # no decision yet: still load the statement so it can be reviewed
            $d.Options = Get-StatementFields $d.Primary
            $filledOnly = $true
        }
        if ($d.Options) {
            $o = $d.Options
            if ($o.PspObjectId -gt 0 -and $o.Type -ne 'OBJECT') { $o.Type = 'OBJECT' }
            Set-FormFromFields $o
            if ($filledOnly) { }
            elseif ($o.Source -eq 'Xml') {
                $rbXml.Checked = $true
                $script:PlanSigShown = $null
                $txtPlan.Text = $o.PlanXml
                Update-Output
                if ($o.PlanStatementIndex -ge 0 -and $o.PlanStatementIndex -lt $cmbPick.Items.Count) { $cmbPick.SelectedIndex = $o.PlanStatementIndex }
            } else {
                $rbHint.Checked = $true
                $txtHint.Text = $o.Hint
            }
            $lblAnalyze.Text = $(if ($filledOnly) { $e1 = @($d.Findings | Where-Object { $_.Severity -eq 'Error' })[0]; 'Statement loaded. Still needed: ' + $(if ($e1) { ($e1.Message -replace '^Plans: ', '') } else { 'see Checks' }) } elseif ($o.Source -eq 'Xml') { 'Filled: forcing the good plan (USE PLAN).' } else { 'Filled: using the hint.' })
        } else {
            $e1 = @($d.Findings | Where-Object { $_.Severity -eq 'Error' })[0]
            $lblAnalyze.Text = 'Nothing loaded: ' + $(if ($e1) { $e1.Message -replace '^Plans: ', '' } else { 'see the Checks tab.' })
        }
        $tip.SetToolTip($lblAnalyze, $lblAnalyze.Text)
        Update-Output
        $tabs.SelectedTab = $tabChecks
    } catch { Show-GuiError 'Analyze and fill' $_ }
}
$btnAnalyze.Add_Click({ Invoke-Analyze })

$form.Add_Shown({
    $split.SplitterDistance = 640
    if ($Server)   { $txtConn.Text = $Server }
    if ($Database) { $txtDb.Text = $Database }
    Update-Output
    $txtQuery.Focus()
})
$form.Add_FormClosing({ if ($script:Bg) { try { [void]$script:Bg.PS.BeginStop($null, $null) } catch { } } })

[void]$form.ShowDialog()
$form.Dispose()
#endregion

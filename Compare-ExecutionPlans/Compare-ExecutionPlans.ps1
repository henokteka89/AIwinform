<#
.SYNOPSIS
    Compare-ExecutionPlans.ps1 - WinForms tool that loads two or more SQL Server execution plans,
    analyzes and compares them, and builds an AI-ready Markdown prompt for deeper analysis.

.DESCRIPTION
    Plan sources:
      * .sqlplan / .xml files (button or drag & drop onto the list)
      * Plan XML on the clipboard (SSMS: right-click plan > Show Execution Plan XML > Ctrl+A, Ctrl+C)
      * SQL Server: Query Store (by query_id, plan_id or text) or the plan cache (by text or query_hash)

    Output:
      * Comparison  - side-by-side metrics, rows that differ are highlighted
      * Findings    - cross-plan differences + per-plan heuristics (spills, skew, lookups, conversions, grants...)
      * Operators   - every operator from every plan in one sortable grid
      * Plan trees  - indented operator trees with estimates/actuals/predicates/warnings
      * AI prompt   - Markdown package (instructions + facts + trees [+ optional raw XML]) to paste into an AI
      * JSON export - structured data for other tooling

.NOTES
    Windows PowerShell 5.1 or PowerShell 7+ (Windows). Run:
        powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\Compare-ExecutionPlans.ps1
    Server capture needs VIEW SERVER STATE (plan cache) or VIEW DATABASE STATE (Query Store).
    Estimated plans have no runtime numbers; actual plans (or Query Store runtime stats) give far better analysis.
#>
[CmdletBinding()]
param()

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

# WinForms needs STA; relaunch if required
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName Microsoft.VisualBasic
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ShowplanNs = 'http://schemas.microsoft.com/sqlserver/2004/07/showplan'
$script:Inv        = [Globalization.CultureInfo]::InvariantCulture
$script:Plans      = New-Object System.Collections.ArrayList
$script:CrossFindings = @()
$script:MaxTreeLines  = 400

#region ---------- Helpers ----------

function Get-Attr {
    param($Node, [string]$Name)
    if ($null -eq $Node -or $null -eq $Node.Attributes) { return $null }
    $at = $Node.Attributes[$Name]
    if ($at) { return $at.Value }
    return $null
}

function ConvertTo-Num {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $d = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, $script:Inv, [ref]$d)) { return $d }
    return $null
}

function Format-Num {
    param($Value, [int]$Decimals = 0)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    return ([double]$Value).ToString("N$Decimals", $script:Inv)
}

function Limit-Text {
    param([string]$Text, [int]$Max = 200)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { return $t.Substring(0, $Max) + '...' }
    return $t
}

function Get-CleanName {
    param([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return $null }
    return $Name.Trim([char[]]'[]')
}

function Get-Sum {
    param($Items, [string]$Prop)
    $t = 0.0; $any = $false
    foreach ($i in $Items) { $v = $i.$Prop; if ($null -ne $v) { $t += [double]$v; $any = $true } }
    if ($any) { return $t } else { return $null }
}

function Get-Max {
    param($Items, [string]$Prop)
    $m = $null
    foreach ($i in $Items) { $v = $i.$Prop; if ($null -ne $v -and ($null -eq $m -or [double]$v -gt $m)) { $m = [double]$v } }
    return $m
}

function ConvertTo-MdCell {
    param($Text)
    if ($null -eq $Text) { return '' }
    return ([string]$Text -replace '\|', '\|' -replace '\r?\n', ' ')
}

function ConvertTo-DbValue {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return [DBNull]::Value }
    return $Value
}

function Get-PlanTag { param([int]$Index) return "P$($Index + 1)" }

#endregion

#region ---------- Plan parsing ----------

function Get-WarningTexts {
    param($WarningsNode, $Ns)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $WarningsNode) { return @() }
    foreach ($at in $WarningsNode.Attributes) {
        if ($at.Value -eq 'true' -or $at.Value -eq '1') { [void]$out.Add($at.Name) }
    }
    foreach ($c in $WarningsNode.ChildNodes) {
        if ($c.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $bits = @()
        foreach ($at in $c.Attributes) { $bits += "$($at.Name)=$(Limit-Text $at.Value 150)" }
        $cols = @()
        foreach ($cr in $c.SelectNodes('.//p:ColumnReference', $Ns)) {
            $cols += ((@((Get-CleanName (Get-Attr $cr 'Table')), (Get-CleanName (Get-Attr $cr 'Column'))) | Where-Object { $_ }) -join '.')
        }
        $t = $c.LocalName
        if ($bits.Count) { $t += ' (' + ($bits -join ', ') + ')' }
        if ($cols.Count) { $t += ' [' + ((@($cols | Select-Object -Unique)) -join ', ') + ']' }
        [void]$out.Add($t)
    }
    return @($out)
}

function Format-OpLine {
    param($o)
    $s = '[' + $o.NodeId + '] ' + $o.PhysicalOp
    if ($o.LogicalOp -and $o.LogicalOp -ne $o.PhysicalOp) { $s += ' (' + $o.LogicalOp + ')' }
    if ($o.Object) { $s += ' {' + $o.Object + '}' }
    $s += ' est=' + (Format-Num $o.EstRows 0)
    if ($null -ne $o.ActRows) {
        $s += ' act=' + (Format-Num $o.ActRows 0) + ' execs=' + (Format-Num $o.ActExecs 0)
    } elseif ($o.EstExecs -gt 1) {
        $s += ' estExecs=' + (Format-Num $o.EstExecs 0)
    }
    $s += ' cost=' + $o.CostPct + '%'
    if ($null -ne $o.ActElapsedMs) { $s += ' time=' + (Format-Num $o.ActElapsedMs 0) + 'ms' }
    if ($o.ExecMode -eq 'Batch') { $s += ' [batch]' }
    if ($o.Parallel) { $s += ' [parallel]' }
    if ($null -ne $o.Skew -and $o.Skew -ge 10) { $s += ' SKEW x' + $o.Skew + ' (' + $o.SkewDir + ')' }
    if ($o.SeekPredicate) { $s += ' | seek: ' + (Limit-Text $o.SeekPredicate 160) }
    if ($o.Predicate) { $s += ' | pred: ' + (Limit-Text $o.Predicate 160) }
    if ($o.Udf) { $s += ' | UDF: ' + $o.Udf }
    if (@($o.Warnings).Count) { $s += ' | !WARN: ' + (@($o.Warnings) -join '; ') }
    return $s
}

function Add-TreeLine {
    param($RelOp, [int]$Depth, $Lines, $OpIndex, $Ns, [int]$Max)
    if ($Lines.Count -ge $Max) {
        if ($Lines.Count -eq $Max) { [void]$Lines.Add('... (tree truncated)') }
        return
    }
    $o = $OpIndex[[string](Get-Attr $RelOp 'NodeId')]
    if ($o) { $text = Format-OpLine $o } else { $text = '[' + (Get-Attr $RelOp 'NodeId') + '] ' + (Get-Attr $RelOp 'PhysicalOp') }
    [void]$Lines.Add(('   ' * $Depth) + $text)
    foreach ($c in $RelOp.SelectNodes('./*/p:RelOp', $Ns)) {
        Add-TreeLine -RelOp $c -Depth ($Depth + 1) -Lines $Lines -OpIndex $OpIndex -Ns $Ns -Max $Max
    }
}

function Get-PlanAnalysis {
    param([Parameter(Mandatory = $true)][string]$PlanXml)

    $doc = New-Object System.Xml.XmlDocument
    $doc.XmlResolver = $null
    $doc.LoadXml($PlanXml.Trim())
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('p', $script:ShowplanNs)
    $root = $doc.DocumentElement
    if ($root.LocalName -ne 'ShowPlanXML') { throw "Not a showplan document (root element is <$($root.LocalName)>)." }

    $stmts  = New-Object System.Collections.ArrayList
    $ops    = New-Object System.Collections.ArrayList
    $warns  = New-Object System.Collections.ArrayList
    $mis    = New-Object System.Collections.ArrayList
    $params = New-Object System.Collections.ArrayList
    $stats  = New-Object System.Collections.ArrayList
    $waits  = New-Object System.Collections.ArrayList
    $flags  = New-Object System.Collections.ArrayList
    $isActual = $false

    $stmtNodes = $doc.SelectNodes('//p:StmtSimple[p:QueryPlan]', $ns)
    if ($stmtNodes.Count -eq 0) { throw 'No statements with a query plan were found.' }

    $sIdx = 0
    foreach ($st in $stmtNodes) {
        $sIdx++
        $qp  = $st.SelectSingleNode('p:QueryPlan', $ns)
        $mg  = $qp.SelectSingleNode('p:MemoryGrantInfo', $ns)
        $qts = $qp.SelectSingleNode('p:QueryTimeStats', $ns)
        $hw  = $qp.SelectSingleNode('p:OptimizerHardwareDependentProperties', $ns)
        $sCost = ConvertTo-Num (Get-Attr $st 'StatementSubTreeCost')

        $stmt = [pscustomobject]@{
            Index              = $sIdx
            Text               = (Get-Attr $st 'StatementText')
            Type               = (Get-Attr $st 'StatementType')
            Cost               = $sCost
            EstRows            = (ConvertTo-Num (Get-Attr $st 'StatementEstRows'))
            OptmLevel          = (Get-Attr $st 'StatementOptmLevel')
            EarlyAbort         = (Get-Attr $st 'StatementOptmEarlyAbortReason')
            CEVersion          = (Get-Attr $st 'CardinalityEstimationModelVersion')
            QueryHash          = (Get-Attr $st 'QueryHash')
            QueryPlanHash      = (Get-Attr $st 'QueryPlanHash')
            DOP                = (ConvertTo-Num (Get-Attr $qp 'DegreeOfParallelism'))
            NonParallelReason  = (Get-Attr $qp 'NonParallelPlanReason')
            CachedPlanSizeKB   = (ConvertTo-Num (Get-Attr $qp 'CachedPlanSize'))
            CompileTimeMs      = (ConvertTo-Num (Get-Attr $qp 'CompileTime'))
            CompileCpuMs       = (ConvertTo-Num (Get-Attr $qp 'CompileCPU'))
            CompileMemoryKB    = (ConvertTo-Num (Get-Attr $qp 'CompileMemory'))
            RequestedMemKB     = (ConvertTo-Num (Get-Attr $mg 'RequestedMemory'))
            GrantedMemKB       = (ConvertTo-Num (Get-Attr $mg 'GrantedMemory'))
            MaxUsedMemKB       = (ConvertTo-Num (Get-Attr $mg 'MaxUsedMemory'))
            GrantWaitMs        = (ConvertTo-Num (Get-Attr $mg 'GrantWaitTime'))
            ElapsedMs          = (ConvertTo-Num (Get-Attr $qts 'ElapsedTime'))
            CpuMs              = (ConvertTo-Num (Get-Attr $qts 'CpuTime'))
            EstAvailMemGrantKB = (ConvertTo-Num (Get-Attr $hw 'EstimatedAvailableMemoryGrant'))
            EstAvailDOP        = (ConvertTo-Num (Get-Attr $hw 'EstimatedAvailableDegreeOfParallelism'))
            Tree               = ''
        }

        # ---- Operators ----
        $opIndex = @{}
        foreach ($r in $qp.SelectNodes('.//p:RelOp', $ns)) {
            $sub = ConvertTo-Num (Get-Attr $r 'EstimatedTotalSubtreeCost')
            $childSum = 0.0
            foreach ($c in $r.SelectNodes('./*/p:RelOp', $ns)) { $childSum += [double](ConvertTo-Num (Get-Attr $c 'EstimatedTotalSubtreeCost')) }
            $own = [math]::Max(0.0, [double]$sub - $childSum)
            $pct = 0
            if ($sCost -gt 0) { $pct = [math]::Round(100.0 * $own / $sCost, 1) }

            $phys = Get-Attr $r 'PhysicalOp'
            $log  = Get-Attr $r 'LogicalOp'
            $isx  = $r.SelectSingleNode('./p:IndexScan', $ns)
            $lk   = Get-Attr $isx 'Lookup'
            if (($lk -eq '1' -or $lk -eq 'true') -and $phys -like 'Clustered Index*') { $phys = 'Key Lookup' }

            # object
            $obj = $r.SelectSingleNode('./*/p:Object', $ns)
            $table = $null; $index = $null; $objText = ''
            if ($obj) {
                $table = (@((Get-CleanName (Get-Attr $obj 'Schema')), (Get-CleanName (Get-Attr $obj 'Table'))) | Where-Object { $_ }) -join '.'
                $index = Get-CleanName (Get-Attr $obj 'Index')
                $objText = $table
                if ($index) { $objText += '.' + $index }
                $alias = Get-CleanName (Get-Attr $obj 'Alias')
                if ($alias) { $objText += ' AS ' + $alias }
            }

            $estRows = ConvertTo-Num (Get-Attr $r 'EstimateRows')
            $estExecs = 1.0 + [double](ConvertTo-Num (Get-Attr $r 'EstimateRebinds')) + [double](ConvertTo-Num (Get-Attr $r 'EstimateRewinds'))

            # runtime counters (actual plans)
            $actRows = $null; $actExec = $null; $lr = $null; $prd = $null; $el = $null; $cpu = $null
            $rtNodes = $r.SelectNodes('p:RunTimeInformation/p:RunTimeCountersPerThread', $ns)
            if ($rtNodes.Count -gt 0) {
                $isActual = $true
                $actRows = 0.0; $actExec = 0.0; $lr = 0.0; $prd = 0.0
                $hasTiming = $null -ne $rtNodes[0].Attributes['ActualElapsedms']
                if ($hasTiming) { $el = 0.0; $cpu = 0.0 }
                foreach ($t in $rtNodes) {
                    $actRows += [double](ConvertTo-Num (Get-Attr $t 'ActualRows'))
                    $actExec += [double](ConvertTo-Num (Get-Attr $t 'ActualExecutions'))
                    $lr      += [double](ConvertTo-Num (Get-Attr $t 'ActualLogicalReads'))
                    $prd     += [double](ConvertTo-Num (Get-Attr $t 'ActualPhysicalReads'))
                    if ($hasTiming) {
                        $e = [double](ConvertTo-Num (Get-Attr $t 'ActualElapsedms'))
                        if ($e -gt $el) { $el = $e }
                        $cpu += [double](ConvertTo-Num (Get-Attr $t 'ActualCPUms'))
                    }
                }
            }

            $skew = $null; $skewDir = $null; $perExec = $null
            if ($null -ne $actRows -and $actExec -gt 0) {
                $perExec = $actRows / $actExec
                $e = [double]$estRows
                $hi = [math]::Max($perExec, $e)
                $lo = [math]::Max([math]::Min($perExec, $e), 1.0)
                $skew = [math]::Round($hi / $lo, 1)
                if ($perExec -gt $e) { $skewDir = 'under-estimated' } else { $skewDir = 'over-estimated' }
            }

            # predicates
            $pred = Get-Attr ($r.SelectSingleNode('./*/p:Predicate/p:ScalarOperator', $ns)) 'ScalarString'
            $seekParts = @()
            foreach ($k in $r.SelectNodes('./*/p:SeekPredicates//p:SeekKeys/*', $ns)) {
                $cols  = @($k.SelectNodes('p:RangeColumns/p:ColumnReference', $ns) | ForEach-Object { Get-CleanName (Get-Attr $_ 'Column') })
                $exprs = @($k.SelectNodes('p:RangeExpressions/p:ScalarOperator', $ns) | ForEach-Object { Get-Attr $_ 'ScalarString' })
                $seekParts += (($cols -join ',') + ' ' + (Get-Attr $k 'ScanType') + ' ' + ($exprs -join ','))
            }
            $udfs = @($r.SelectNodes('./*/p:DefinedValues//p:UserDefinedFunction', $ns) | ForEach-Object { Get-Attr $_ 'FunctionName' } | Select-Object -Unique)

            $opWarn = @(Get-WarningTexts ($r.SelectSingleNode('p:Warnings', $ns)) $ns)
            $nodeId = [int](Get-Attr $r 'NodeId')
            foreach ($w in $opWarn) {
                [void]$warns.Add([pscustomobject]@{ Stmt = $sIdx; NodeId = $nodeId; Text = $w; Where = "Stmt $sIdx Node $nodeId ($phys)" })
            }
            $par = Get-Attr $r 'Parallel'

            $op = [pscustomobject]@{
                Stmt           = $sIdx
                NodeId         = $nodeId
                PhysicalOp     = $phys
                LogicalOp      = $log
                Object         = $objText
                Table          = $table
                Index          = $index
                Parallel       = ($par -eq '1' -or $par -eq 'true')
                ExecMode       = (Get-Attr $r 'EstimatedExecutionMode')
                EstRows        = $estRows
                EstExecs       = $estExecs
                EstRowsRead    = (ConvertTo-Num (Get-Attr $r 'EstimatedRowsRead'))
                TableCardinality = (ConvertTo-Num (Get-Attr $r 'TableCardinality'))
                RowGoalEst     = (ConvertTo-Num (Get-Attr $r 'EstimateRowsWithoutRowGoal'))
                SubtreeCost    = $sub
                OwnCost        = $own
                CostPct        = $pct
                ActRows        = $actRows
                ActRowsPerExec = $perExec
                ActExecs       = $actExec
                ActLogicalReads  = $lr
                ActPhysicalReads = $prd
                ActElapsedMs   = $el
                ActCpuMs       = $cpu
                Skew           = $skew
                SkewDir        = $skewDir
                SeekPredicate  = ($seekParts -join ' AND ')
                Predicate      = $pred
                Udf            = ($udfs -join ', ')
                Warnings       = $opWarn
            }
            [void]$ops.Add($op)
            $opIndex[[string]$nodeId] = $op
        }

        # ---- Statement-level warnings ----
        foreach ($w in @(Get-WarningTexts ($qp.SelectSingleNode('p:Warnings', $ns)) $ns)) {
            [void]$warns.Add([pscustomobject]@{ Stmt = $sIdx; NodeId = $null; Text = $w; Where = "Stmt $sIdx" })
        }

        # ---- Missing indexes ----
        foreach ($g in $qp.SelectNodes('p:MissingIndexes/p:MissingIndexGroup', $ns)) {
            $impact = ConvertTo-Num (Get-Attr $g 'Impact')
            foreach ($mi in $g.SelectNodes('p:MissingIndex', $ns)) {
                $eq = @(); $ineq = @(); $inc = @()
                foreach ($cg in $mi.SelectNodes('p:ColumnGroup', $ns)) {
                    $cols = @($cg.SelectNodes('p:Column', $ns) | ForEach-Object { Get-Attr $_ 'Name' })
                    switch (Get-Attr $cg 'Usage') {
                        'EQUALITY'   { $eq = $cols }
                        'INEQUALITY' { $ineq = $cols }
                        'INCLUDE'    { $inc = $cols }
                    }
                }
                $keys = @($eq) + @($ineq)
                $ddl = 'CREATE NONCLUSTERED INDEX [IX_suggested] ON ' + (Get-Attr $mi 'Schema') + '.' + (Get-Attr $mi 'Table') + ' (' + ($keys -join ', ') + ')'
                if ($inc.Count) { $ddl += ' INCLUDE (' + ($inc -join ', ') + ')' }
                [void]$mis.Add([pscustomobject]@{
                    Stmt = $sIdx; Impact = $impact
                    Table = ((Get-CleanName (Get-Attr $mi 'Schema')) + '.' + (Get-CleanName (Get-Attr $mi 'Table')))
                    Equality = ($eq -join ', '); Inequality = ($ineq -join ', '); Include = ($inc -join ', '); Ddl = $ddl
                })
            }
        }

        # ---- Parameters ----
        foreach ($pn in $qp.SelectNodes('p:ParameterList/p:ColumnReference', $ns)) {
            [void]$params.Add([pscustomobject]@{
                Stmt = $sIdx; Name = (Get-Attr $pn 'Column'); DataType = (Get-Attr $pn 'ParameterDataType')
                Compiled = (Get-Attr $pn 'ParameterCompiledValue'); Runtime = (Get-Attr $pn 'ParameterRuntimeValue')
            })
        }

        # ---- Statistics used ----
        foreach ($si in $qp.SelectNodes('p:OptimizerStatsUsage/p:StatisticsInfo', $ns)) {
            [void]$stats.Add([pscustomobject]@{
                Stmt = $sIdx
                Table = ((Get-CleanName (Get-Attr $si 'Schema')) + '.' + (Get-CleanName (Get-Attr $si 'Table')))
                Statistics = (Get-CleanName (Get-Attr $si 'Statistics'))
                LastUpdate = (Get-Attr $si 'LastUpdate')
                ModificationCount = (ConvertTo-Num (Get-Attr $si 'ModificationCount'))
                SamplingPercent = (ConvertTo-Num (Get-Attr $si 'SamplingPercent'))
            })
        }

        # ---- Waits (actual plans) ----
        foreach ($wn in $qp.SelectNodes('p:WaitStats/p:Wait', $ns)) {
            [void]$waits.Add([pscustomobject]@{
                Stmt = $sIdx; WaitType = (Get-Attr $wn 'WaitType')
                WaitMs = (ConvertTo-Num (Get-Attr $wn 'WaitTimeMs')); WaitCount = (ConvertTo-Num (Get-Attr $wn 'WaitCount'))
            })
        }

        # ---- Trace flags ----
        foreach ($tf in $qp.SelectNodes('p:TraceFlags/p:TraceFlag', $ns)) {
            [void]$flags.Add(('{0}({1})' -f (Get-Attr $tf 'Value'), (Get-Attr $tf 'Scope')))
        }

        # ---- Tree ----
        $lines = New-Object System.Collections.ArrayList
        $top = $qp.SelectSingleNode('p:RelOp', $ns)
        if ($top) { Add-TreeLine -RelOp $top -Depth 0 -Lines $lines -OpIndex $opIndex -Ns $ns -Max $script:MaxTreeLines }
        $stmt.Tree = ($lines -join "`r`n")

        [void]$stmts.Add($stmt)
    }

    $main = $stmts | Sort-Object -Property Cost -Descending | Select-Object -First 1

    $a = [pscustomobject]@{
        IsActual       = $isActual
        Build          = (Get-Attr $root 'Build')
        Version        = (Get-Attr $root 'Version')
        Statements     = $stmts.ToArray()
        Main           = $main
        Operators      = $ops.ToArray()
        Warnings       = $warns.ToArray()
        MissingIndexes = $mis.ToArray()
        Parameters     = $params.ToArray()
        Stats          = $stats.ToArray()
        Waits          = $waits.ToArray()
        TraceFlags     = (@($flags | Select-Object -Unique) -join ', ')
        TotalCost      = (Get-Sum $stmts 'Cost')
        ElapsedMs      = (Get-Sum $stmts 'ElapsedMs')
        CpuMs          = (Get-Sum $stmts 'CpuMs')
        GrantedMemKB   = (Get-Sum $stmts 'GrantedMemKB')
        MaxUsedMemKB   = (Get-Sum $stmts 'MaxUsedMemKB')
        MaxDOP         = (Get-Max $stmts 'DOP')
        LogicalReads   = (Get-Sum $ops 'ActLogicalReads')
        Summary        = $null
    }
    $a.Summary = New-PlanSummary $a
    return $a
}

function New-PlanSummary {
    param($A)
    $st = $A.Statements; $ops = $A.Operators; $m = $A.Main
    $count = { param($sb) @($ops | Where-Object $sb).Count }

    $s = [ordered]@{}
    if ($A.IsActual) { $s['Plan type'] = 'Actual' } else { $s['Plan type'] = 'Estimated' }
    $s['SQL build']                 = $A.Build
    $s['Statements']                = $st.Count
    $s['Query hash (main)']         = $m.QueryHash
    $s['Query plan hash (main)']    = $m.QueryPlanHash
    $s['CE model']                  = (@($st | ForEach-Object { $_.CEVersion } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
    $s['Optimization level']        = (@($st | ForEach-Object { $_.OptmLevel } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
    $s['Early abort reason']        = (@($st | ForEach-Object { $_.EarlyAbort } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
    $s['Est. subtree cost (total)'] = Format-Num $A.TotalCost 3
    $s['Est. rows (main stmt)']     = Format-Num $m.EstRows 0
    $s['DOP (max)']                 = Format-Num $A.MaxDOP 0
    $s['Non-parallel reason']       = (@($st | ForEach-Object { $_.NonParallelReason } | Where-Object { $_ } | Select-Object -Unique) -join ', ')
    $s['Compile time ms']           = Format-Num (Get-Sum $st 'CompileTimeMs') 0
    $s['Compile CPU ms']            = Format-Num (Get-Sum $st 'CompileCpuMs') 0
    $s['Compile memory KB']         = Format-Num (Get-Sum $st 'CompileMemoryKB') 0
    $s['Cached plan size KB']       = Format-Num (Get-Sum $st 'CachedPlanSizeKB') 0
    $s['Memory requested KB']       = Format-Num (Get-Sum $st 'RequestedMemKB') 0
    $s['Memory granted KB']         = Format-Num $A.GrantedMemKB 0
    $s['Memory max used KB']        = Format-Num $A.MaxUsedMemKB 0
    $s['Grant wait ms']             = Format-Num (Get-Sum $st 'GrantWaitMs') 0
    $s['Elapsed ms (actual)']       = Format-Num $A.ElapsedMs 0
    $s['CPU ms (actual)']           = Format-Num $A.CpuMs 0
    $s['Logical reads (actual)']    = Format-Num $A.LogicalReads 0
    $s['Physical reads (actual)']   = Format-Num (Get-Sum $ops 'ActPhysicalReads') 0
    $s['Operators']                 = $ops.Count
    $s['Seeks']                     = & $count { $_.PhysicalOp -match 'Seek' }
    $s['Scans']                     = & $count { $_.PhysicalOp -match 'Scan' -and $_.PhysicalOp -ne 'Constant Scan' }
    $s['Key/RID lookups']           = & $count { $_.PhysicalOp -in 'Key Lookup', 'RID Lookup' }
    $s['Nested loops joins']        = & $count { $_.PhysicalOp -eq 'Nested Loops' }
    $s['Hash joins']                = & $count { $_.PhysicalOp -eq 'Hash Match' -and $_.LogicalOp -notmatch 'Aggregate|Flow Distinct' }
    $s['Merge joins']               = & $count { $_.PhysicalOp -eq 'Merge Join' }
    $s['Adaptive joins']            = & $count { $_.PhysicalOp -eq 'Adaptive Join' }
    $s['Hash aggregates']           = & $count { $_.PhysicalOp -eq 'Hash Match' -and $_.LogicalOp -match 'Aggregate|Flow Distinct' }
    $s['Stream aggregates']         = & $count { $_.PhysicalOp -eq 'Stream Aggregate' }
    $s['Sorts']                     = & $count { $_.PhysicalOp -match 'Sort' }
    $s['Spools']                    = & $count { $_.PhysicalOp -match 'Spool' }
    $s['Parallelism (exchange) ops']= & $count { $_.PhysicalOp -eq 'Parallelism' }
    $s['Batch mode ops']            = & $count { $_.ExecMode -eq 'Batch' }
    $s['Warnings']                  = $A.Warnings.Count
    $s['Missing index suggestions'] = $A.MissingIndexes.Count

    $sk = $ops | Where-Object { $null -ne $_.Skew } | Sort-Object Skew -Descending | Select-Object -First 1
    if ($sk) { $s['Max est/actual skew'] = "x$($sk.Skew) (Stmt $($sk.Stmt) Node $($sk.NodeId) $($sk.PhysicalOp))" } else { $s['Max est/actual skew'] = '' }

    $topOp = $ops | Sort-Object OwnCost -Descending | Select-Object -First 1
    if ($topOp) { $s['Top operator (est. cost)'] = "Node $($topOp.NodeId) $($topOp.PhysicalOp) $($topOp.Object) ($($topOp.CostPct)%)" } else { $s['Top operator (est. cost)'] = '' }

    $slow = $ops | Where-Object { $null -ne $_.ActElapsedMs } | Sort-Object ActElapsedMs -Descending | Select-Object -First 1
    if ($slow) { $s['Slowest operator (actual)'] = "Node $($slow.NodeId) $($slow.PhysicalOp) $($slow.Object) ($(Format-Num $slow.ActElapsedMs 0) ms)" } else { $s['Slowest operator (actual)'] = '' }

    $tw = $A.Waits | Sort-Object WaitMs -Descending | Select-Object -First 1
    if ($tw) { $s['Top wait (actual)'] = "$($tw.WaitType) $(Format-Num $tw.WaitMs 0) ms" } else { $s['Top wait (actual)'] = '' }
    $s['Trace flags'] = $A.TraceFlags
    return $s
}

#endregion

#region ---------- Findings ----------

function Get-PlanFindings {
    param($A)
    $f = New-Object System.Collections.ArrayList
    $add = { param($sev, $where, $msg) [void]$f.Add([pscustomobject]@{ Severity = $sev; Where = $where; Message = $msg }) }

    foreach ($s in $A.Statements) {
        $w = "Stmt $($s.Index)"
        if ($s.EarlyAbort -eq 'TimeOut') { & $add 'Medium' $w 'Optimizer timed out (StatementOptmEarlyAbortReason=TimeOut); the chosen plan may be far from optimal. Simplify the query or reduce join count.' }
        if ($s.EarlyAbort -eq 'MemoryLimitExceeded') { & $add 'High' $w 'Optimization stopped due to memory limit; plan quality is likely poor.' }
        if ($s.NonParallelReason -and $s.Cost -ge 50) { & $add 'Medium' $w "Expensive statement (cost $(Format-Num $s.Cost 1)) is forced serial: $($s.NonParallelReason)." }
        if ($s.GrantWaitMs -gt 0) { & $add 'High' $w "Query waited $(Format-Num $s.GrantWaitMs 0) ms for a memory grant (RESOURCE_SEMAPHORE pressure)." }
        if ($s.GrantedMemKB -ge 25600 -and $null -ne $s.MaxUsedMemKB -and $s.MaxUsedMemKB -lt 0.1 * $s.GrantedMemKB) {
            & $add 'Medium' $w "Excessive memory grant: granted $(Format-Num $s.GrantedMemKB 0) KB, used $(Format-Num $s.MaxUsedMemKB 0) KB (over-estimated rows feeding sorts/hashes?)."
        }
        if ($s.ElapsedMs -gt 1000 -and $s.CpuMs -ge 0 -and $s.ElapsedMs -gt 2 * $s.CpuMs) {
            & $add 'Info' $w "Elapsed ($(Format-Num $s.ElapsedMs 0) ms) is much higher than CPU ($(Format-Num $s.CpuMs 0) ms): time was spent waiting - check waits."
        }
    }

    foreach ($wr in $A.Warnings) {
        $sev = 'Info'
        if ($wr.Text -match 'Spill') { $sev = 'High' }
        elseif ($wr.Text -match 'PlanAffectingConvert' -and $wr.Text -match 'Seek Plan') { $sev = 'High' }
        elseif ($wr.Text -match 'NoJoinPredicate') { $sev = 'High' }
        elseif ($wr.Text -match 'PlanAffectingConvert|ColumnsWithNoStatistics|MemoryGrantWarning|UnmatchedIndexes') { $sev = 'Medium' }
        & $add $sev $wr.Where ('Warning: ' + $wr.Text)
    }

    $rowGoalOps = 0
    foreach ($o in $A.Operators) {
        $w = "Stmt $($o.Stmt) Node $($o.NodeId)"
        if ($o.PhysicalOp -in 'Key Lookup', 'RID Lookup') {
            $n = $o.EstExecs; if ($null -ne $o.ActExecs) { $n = $o.ActExecs }
            if ($n -ge 1000) { & $add 'Medium' $w "$($o.PhysicalOp) on $($o.Object) runs ~$(Format-Num $n 0) times; a covering index (INCLUDE the looked-up columns) would remove it." }
            elseif ($o.CostPct -ge 20) { & $add 'Info' $w "$($o.PhysicalOp) on $($o.Object) is $($o.CostPct)% of statement cost." }
        }
        if ($o.PhysicalOp -match 'Scan' -and $o.PhysicalOp -ne 'Constant Scan' -and $o.CostPct -ge 20 -and ([double]$o.TableCardinality -ge 100000 -or [double]$o.EstRowsRead -ge 100000)) {
            $msg = "$($o.PhysicalOp) on $($o.Object) is $($o.CostPct)% of cost over ~$(Format-Num ([math]::Max([double]$o.TableCardinality, [double]$o.EstRowsRead)) 0) rows"
            if ($o.Predicate) { $msg += "; residual predicate: $(Limit-Text $o.Predicate 120) (seekable index?)" }
            & $add 'Medium' $w ($msg + '.')
        }
        if ($null -ne $o.Skew -and $o.Skew -ge 10 -and [math]::Max([double]$o.EstRows, [double]$o.ActRowsPerExec) -ge 1000) {
            $sev = 'Medium'; if ($o.CostPct -ge 10 -or $o.Skew -ge 100) { $sev = 'High' }
            & $add $sev $w "Cardinality misestimate x$($o.Skew) ($($o.SkewDir)) at $($o.PhysicalOp) $($o.Object): est $(Format-Num $o.EstRows 0) vs actual $(Format-Num $o.ActRowsPerExec 0) rows/exec."
        }
        if ($o.PhysicalOp -eq 'Index Spool' -and $o.LogicalOp -eq 'Eager Spool') { & $add 'High' $w 'Eager Index Spool: the optimizer builds a temporary index at runtime - usually a missing index on the inner side of a nested loop.' }
        elseif ($o.PhysicalOp -match 'Spool' -and $o.CostPct -ge 10) { & $add 'Info' $w "$($o.PhysicalOp) ($($o.LogicalOp)) is $($o.CostPct)% of cost." }
        if ($o.PhysicalOp -match 'Sort' -and $o.CostPct -ge 25) { & $add 'Info' $w "Sort is $($o.CostPct)% of cost; an index providing this order could remove it." }
        if ($o.Udf) { & $add 'Medium' $w "Scalar UDF call(s): $($o.Udf) - often forces serial plans and row-by-row execution." }
        if ($o.Table -like '*@*' -and $o.EstRows -le 1 -and ($null -eq $o.ActRowsPerExec -or $o.ActRowsPerExec -ge 100)) {
            & $add 'Info' $w "Table variable $($o.Table) estimated at 1 row; if it holds many rows consider a temp table or deferred compilation (compat 150+)."
        }
        if ($null -ne $o.RowGoalEst) { $rowGoalOps++ }
    }
    if ($rowGoalOps -gt 0) { & $add 'Info' 'Plan' "$rowGoalOps operator(s) use a row goal (TOP/EXISTS/FAST n/IN) - estimates are deliberately reduced." }

    foreach ($p in $A.Parameters) {
        if ($null -ne $p.Runtime -and $null -ne $p.Compiled -and $p.Runtime -ne $p.Compiled) {
            & $add 'Medium' "Stmt $($p.Stmt)" "Parameter $($p.Name) compiled with $($p.Compiled) but ran with $($p.Runtime) - possible parameter sensitivity."
        }
    }
    foreach ($mi in $A.MissingIndexes) {
        $sev = 'Info'; if ($mi.Impact -ge 30) { $sev = 'Medium' }
        & $add $sev "Stmt $($mi.Stmt)" "Missing index on $($mi.Table) (impact $(Format-Num $mi.Impact 1)%): $($mi.Ddl)"
    }
    $stale = @($A.Stats | Where-Object { $_.ModificationCount -gt 0 } | Sort-Object ModificationCount -Descending | Select-Object -First 5)
    foreach ($sx in $stale) {
        & $add 'Info' "Stmt $($sx.Stmt)" "Stats $($sx.Table).$($sx.Statistics): $(Format-Num $sx.ModificationCount 0) modifications since last update ($($sx.LastUpdate)), sampled $(Format-Num $sx.SamplingPercent 1)%."
    }
    foreach ($wt in @($A.Waits | Sort-Object WaitMs -Descending | Select-Object -First 3)) {
        & $add 'Info' "Stmt $($wt.Stmt)" "Wait $($wt.WaitType): $(Format-Num $wt.WaitMs 0) ms ($(Format-Num $wt.WaitCount 0) waits)."
    }
    return @($f)
}

function Get-CrossFindings {
    param($Items)
    $f = New-Object System.Collections.ArrayList
    if ($Items.Count -lt 2) { return @() }
    $tags = @(); for ($i = 0; $i -lt $Items.Count; $i++) { $tags += (Get-PlanTag $i) }

    $qh  = @($Items | ForEach-Object { $_.Analysis.Main.QueryHash } | Where-Object { $_ } | Select-Object -Unique)
    $qph = @($Items | ForEach-Object { $_.Analysis.Main.QueryPlanHash } | Where-Object { $_ } | Select-Object -Unique)
    if ($qh.Count -eq 1 -and $qph.Count -gt 1) {
        [void]$f.Add('Same query hash, different plan hashes: the same query was compiled into different plans. Typical causes: parameter sensitivity, statistics changes, CE/compatibility level, SET options, schema/index changes.')
    } elseif ($qh.Count -gt 1) {
        [void]$f.Add('Different query hashes: these plans are for structurally different queries (or a rewritten query). Differences may be by design.')
    } elseif ($qph.Count -eq 1) {
        [void]$f.Add('Identical query plan hash: plan shape is the same. Differences are runtime/environmental (data volume, waits, blocking, memory, I/O).')
    }

    $describe = {
        param([string]$Name, $Getter, [double]$Ratio)
        $vals = @(); for ($i = 0; $i -lt $Items.Count; $i++) { $vals += , @($tags[$i], (& $Getter $Items[$i])) }
        $nums = @($vals | Where-Object { $null -ne $_[1] -and [double]$_[1] -gt 0 })
        if ($nums.Count -ge 2) {
            $mn = $nums | Sort-Object { [double]$_[1] } | Select-Object -First 1
            $mx = $nums | Sort-Object { [double]$_[1] } -Descending | Select-Object -First 1
            $r = [double]$mx[1] / [double]$mn[1]
            if ($r -ge $Ratio) { [void]$f.Add("$Name differs x$([math]::Round($r,1)): lowest $($mn[0]) = $(Format-Num $mn[1] 2), highest $($mx[0]) = $(Format-Num $mx[1] 2).") }
        }
    }
    & $describe 'Estimated cost' { param($p) $p.Analysis.TotalCost } 2
    & $describe 'Actual elapsed ms' { param($p) $p.Analysis.ElapsedMs } 1.5
    & $describe 'Actual CPU ms' { param($p) $p.Analysis.CpuMs } 1.5
    & $describe 'Actual logical reads' { param($p) $p.Analysis.LogicalReads } 2
    & $describe 'Memory granted KB' { param($p) $p.Analysis.GrantedMemKB } 2
    & $describe 'Server avg duration ms' { param($p) if ($p.Runtime -and $p.Runtime.Contains('Server: avg duration ms')) { ConvertTo-Num $p.Runtime['Server: avg duration ms'] } } 1.5

    foreach ($key in 'CE model', 'DOP (max)', 'SQL build', 'Trace flags', 'Non-parallel reason', 'Early abort reason', 'Plan type') {
        $vals = @(); for ($i = 0; $i -lt $Items.Count; $i++) { $vals += "$($tags[$i])=$($Items[$i].Analysis.Summary[$key])" }
        $u = @($Items | ForEach-Object { [string]$_.Analysis.Summary[$key] } | Select-Object -Unique)
        if ($u.Count -gt 1) { [void]$f.Add("$key differs: " + ($vals -join '; ')) }
    }

    # join / operator strategy
    $joinKeys = 'Nested loops joins', 'Hash joins', 'Merge joins', 'Adaptive joins', 'Key/RID lookups', 'Scans', 'Seeks', 'Sorts', 'Spools', 'Batch mode ops'
    $diffs = @()
    foreach ($k in $joinKeys) {
        $u = @($Items | ForEach-Object { [string]$_.Analysis.Summary[$k] } | Select-Object -Unique)
        if ($u.Count -gt 1) {
            $v = @(); for ($i = 0; $i -lt $Items.Count; $i++) { $v += "$($tags[$i])=$($Items[$i].Analysis.Summary[$k])" }
            $diffs += "$k ($($v -join ', '))"
        }
    }
    if ($diffs.Count) { [void]$f.Add('Operator strategy differs: ' + ($diffs -join '; ')) }

    # access paths per table
    $tables = @{}
    for ($i = 0; $i -lt $Items.Count; $i++) {
        foreach ($o in $Items[$i].Analysis.Operators) {
            if (-not $o.Table -or $o.PhysicalOp -notmatch 'Scan|Seek|Lookup') { continue }
            if (-not $tables.ContainsKey($o.Table)) { $tables[$o.Table] = @{} }
            if (-not $tables[$o.Table].ContainsKey($i)) { $tables[$o.Table][$i] = New-Object System.Collections.Generic.List[string] }
            $tables[$o.Table][$i].Add("$($o.Index) $($o.PhysicalOp)".Trim())
        }
    }
    foreach ($t in ($tables.Keys | Sort-Object)) {
        $per = @()
        for ($i = 0; $i -lt $Items.Count; $i++) {
            if ($tables[$t].ContainsKey($i)) { $per += ((@($tables[$t][$i] | Sort-Object -Unique)) -join ' + ') } else { $per += '(not accessed)' }
        }
        if (@($per | Select-Object -Unique).Count -gt 1) {
            $txt = @(); for ($i = 0; $i -lt $Items.Count; $i++) { $txt += "$($tags[$i]): $($per[$i])" }
            [void]$f.Add("Access path to $t differs -> " + ($txt -join ' | '))
        }
    }

    # compiled parameter values
    $pn = @($Items | ForEach-Object { $_.Analysis.Parameters } | ForEach-Object { $_.Name } | Select-Object -Unique)
    foreach ($n in $pn) {
        $v = @(); for ($i = 0; $i -lt $Items.Count; $i++) {
            $p = $Items[$i].Analysis.Parameters | Where-Object { $_.Name -eq $n } | Select-Object -First 1
            if ($p) { $v += "$($tags[$i])=$($p.Compiled)" } else { $v += "$($tags[$i])=n/a" }
        }
        $vals = @($v | ForEach-Object { $_.Substring($_.IndexOf('=') + 1) } | Select-Object -Unique)
        if ($vals.Count -gt 1) { [void]$f.Add("Parameter $n compiled with different values: " + ($v -join ', ') + ' (sniffing candidate).') }
    }

    foreach ($pair in @(@('Spill', 'Spills'), @('PlanAffectingConvert', 'Implicit conversion warnings'), @('NoJoinPredicate', 'No-join-predicate warnings'))) {
        $has = @(); for ($i = 0; $i -lt $Items.Count; $i++) { if (@($Items[$i].Analysis.Warnings | Where-Object { $_.Text -match $pair[0] }).Count) { $has += $tags[$i] } }
        if ($has.Count -gt 0 -and $has.Count -lt $Items.Count) { [void]$f.Add("$($pair[1]) only in: " + ($has -join ', ')) }
    }
    $hasMi = @(); for ($i = 0; $i -lt $Items.Count; $i++) { if ($Items[$i].Analysis.MissingIndexes.Count) { $hasMi += $tags[$i] } }
    if ($hasMi.Count -gt 0 -and $hasMi.Count -lt $Items.Count) { [void]$f.Add('Missing index suggestions only in: ' + ($hasMi -join ', ')) }

    return @($f)
}

#endregion

#region ---------- AI prompt / exports ----------

function New-AiPrompt {
    param($Items, [string]$Notes, [bool]$IncludeXml, [int]$TopOps)
    $sb = New-Object System.Text.StringBuilder
    function Add-Line { param([string]$t = '') [void]$sb.AppendLine($t) }
    $n = $Items.Count

    Add-Line '# SQL Server Execution Plan Comparison - Analysis Request'
    Add-Line ''
    Add-Line '## Your role and task'
    Add-Line "You are a senior SQL Server performance engineer. Below are $n execution plan(s), labelled P1..P$n, pre-processed by a script into structured facts (metrics, operator trees, warnings, parameters, statistics, waits). Analyze them and:"
    Add-Line '1. Explain the material differences between the plans (shape, access paths, join/aggregate strategy, memory grants, parallelism, cardinality estimates). Ignore cosmetic differences.'
    Add-Line '2. Determine the most likely root cause(s) of the differences (e.g. parameter sensitivity, stale/missing statistics, CE model or compatibility level, SET options, schema/index changes, row goals, data growth). Cite evidence as plan label + Stmt/Node + metric.'
    Add-Line '3. Decide which plan is better for this workload and why. If only estimated plans are available, state that costs are optimizer estimates, not time, and what the judgment depends on.'
    Add-Line '4. List the problems in each plan ranked by likely impact (spills, misestimates, lookups, implicit conversions, large scans, excessive or insufficient grants, serial plans, UDFs, spools).'
    Add-Line '5. Recommend fixes ranked by impact and risk: indexes (give T-SQL), query rewrites, statistics maintenance, Query Store plan forcing / hints, configuration. Note trade-offs (write overhead, plan forcing risks, etc.).'
    Add-Line '6. List the additional information that would confirm the diagnosis (e.g. actual plan, row counts, index definitions, Query Store history).'
    Add-Line ''
    Add-Line 'Rules: rely on the data given; mark assumptions explicitly; "cost %" values are estimated operator costs relative to their statement; the "Script findings" are heuristics - verify them rather than repeating them.'
    Add-Line ''
    Add-Line 'Answer format: 1) Executive summary (3-5 bullets) 2) Key differences table 3) Root cause analysis 4) Issues per plan 5) Recommendations with T-SQL 6) Open questions.'
    Add-Line ''
    Add-Line '## User context'
    if ([string]::IsNullOrWhiteSpace($Notes)) { Add-Line '(none provided)' } else { Add-Line $Notes.Trim() }
    Add-Line ''

    Add-Line '## Plan inventory'
    Add-Line '| Plan | Label | Type | Source | Statements | Est. cost |'
    Add-Line '|---|---|---|---|---|---|'
    for ($i = 0; $i -lt $n; $i++) {
        $p = $Items[$i]; $a = $p.Analysis
        Add-Line ("| {0} | {1} | {2} | {3} | {4} | {5} |" -f (Get-PlanTag $i), (ConvertTo-MdCell $p.Label), $a.Summary['Plan type'], (ConvertTo-MdCell $p.Source), $a.Statements.Count, (Format-Num $a.TotalCost 3))
    }
    Add-Line ''

    Add-Line '## Side-by-side metrics (* = differs between plans)'
    $hdr = '| Metric |'; $sep = '|---|'
    for ($i = 0; $i -lt $n; $i++) { $hdr += " $(Get-PlanTag $i) |"; $sep += '---|' }
    Add-Line $hdr; Add-Line $sep
    foreach ($row in (Get-ComparisonRows $Items)) {
        if ($row.Metric -eq 'Source') { continue }
        $nonEmpty = @($row.Values | Where-Object { "$_" -ne '' -and "$_" -ne '0' })
        if ($nonEmpty.Count -eq 0) { continue }
        $mark = ''; if ($row.Differs) { $mark = ' *' }
        Add-Line ('| ' + (ConvertTo-MdCell $row.Metric) + $mark + ' | ' + ((@($row.Values | ForEach-Object { ConvertTo-MdCell $_ })) -join ' | ') + ' |')
    }
    Add-Line ''

    if ($n -ge 2) {
        Add-Line '## Cross-plan differences detected by the script'
        foreach ($c in $script:CrossFindings) { Add-Line "- $c" }
        if (-not $script:CrossFindings.Count) { Add-Line '- (none detected)' }
        Add-Line ''
    }

    Add-Line '## Plan details'
    for ($i = 0; $i -lt $n; $i++) {
        $p = $Items[$i]; $a = $p.Analysis; $tag = Get-PlanTag $i
        Add-Line ''
        Add-Line "### $tag - $($p.Label)"
        Add-Line "- Type: $($a.Summary['Plan type']); build $($a.Build); source: $($p.Source)"
        if ($p.Runtime -and $p.Runtime.Count) { Add-Line ('- Server runtime stats: ' + ((@($p.Runtime.Keys | ForEach-Object { "$($_ -replace '^Server: ','')=$($p.Runtime[$_])" })) -join ', ')) }

        $stmtsToShow = @($a.Statements | Sort-Object Cost -Descending | Select-Object -First 10 | Sort-Object Index)
        if ($a.Statements.Count -gt $stmtsToShow.Count) { Add-Line "- Showing the 10 most expensive of $($a.Statements.Count) statements." }

        foreach ($s in $stmtsToShow) {
            $share = ''
            if ($a.TotalCost -gt 0 -and $null -ne $s.Cost) { $share = ", $([math]::Round(100 * $s.Cost / $a.TotalCost, 1))% of plan" }
            Add-Line ''
            Add-Line "#### $tag Stmt $($s.Index) ($($s.Type), est. cost $(Format-Num $s.Cost 3)$share)"
            $txt = [string]$s.Text; if ($txt.Length -gt 4000) { $txt = $txt.Substring(0, 4000) + "`n-- (truncated)" }
            Add-Line '```sql'; Add-Line $txt.Trim(); Add-Line '```'
            $bits = @()
            if ($s.DOP) { $bits += "DOP=$($s.DOP)" }
            if ($s.NonParallelReason) { $bits += "NonParallelPlanReason=$($s.NonParallelReason)" }
            if ($s.CEVersion) { $bits += "CE=$($s.CEVersion)" }
            if ($s.OptmLevel) { $bits += "OptmLevel=$($s.OptmLevel)" }
            if ($s.EarlyAbort) { $bits += "EarlyAbort=$($s.EarlyAbort)" }
            if ($s.GrantedMemKB -gt 0 -or $s.RequestedMemKB -gt 0) { $bits += "Grant: requested=$(Format-Num $s.RequestedMemKB 0)KB granted=$(Format-Num $s.GrantedMemKB 0)KB used=$(Format-Num $s.MaxUsedMemKB 0)KB" }
            if ($null -ne $s.ElapsedMs) { $bits += "Elapsed=$(Format-Num $s.ElapsedMs 0)ms CPU=$(Format-Num $s.CpuMs 0)ms" }
            if ($s.CompileTimeMs) { $bits += "Compile=$(Format-Num $s.CompileTimeMs 0)ms" }
            if ($bits.Count) { Add-Line ('- ' + ($bits -join '; ')) }

            $ps = @($a.Parameters | Where-Object { $_.Stmt -eq $s.Index })
            if ($ps.Count) {
                Add-Line '- Parameters:'
                Add-Line '  | Name | Type | Compiled value | Runtime value |'
                Add-Line '  |---|---|---|---|'
                foreach ($pp in $ps) { Add-Line ("  | {0} | {1} | {2} | {3} |" -f (ConvertTo-MdCell $pp.Name), (ConvertTo-MdCell $pp.DataType), (ConvertTo-MdCell $pp.Compiled), (ConvertTo-MdCell $pp.Runtime)) }
            }
            $ws = @($a.Warnings | Where-Object { $_.Stmt -eq $s.Index })
            if ($ws.Count) { Add-Line '- Warnings:'; foreach ($w in $ws) { Add-Line "  - $($w.Where): $($w.Text)" } }
            $ms = @($a.MissingIndexes | Where-Object { $_.Stmt -eq $s.Index })
            if ($ms.Count) { Add-Line '- Missing index suggestions (optimizer, unvalidated):'; foreach ($m in $ms) { Add-Line "  - impact $(Format-Num $m.Impact 1)%: $($m.Ddl)" } }

            Add-Line '- Operator tree (est=estimated rows per execution, act=actual rows total, cost=% of statement):'
            Add-Line '```text'; Add-Line $s.Tree; Add-Line '```'

            $top = @($a.Operators | Where-Object { $_.Stmt -eq $s.Index } | Sort-Object OwnCost -Descending | Select-Object -First $TopOps)
            if ($top.Count) {
                Add-Line "- Top $($top.Count) operators by estimated cost:"
                Add-Line '  | Node | Operator | Object | Est rows | Act rows/exec | Execs | Cost % | Skew | Time ms | Warnings |'
                Add-Line '  |---|---|---|---|---|---|---|---|---|---|'
                foreach ($o in $top) {
                    $execs = $o.ActExecs; if ($null -eq $execs) { $execs = $o.EstExecs }
                    Add-Line ("  | {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} |" -f $o.NodeId, (ConvertTo-MdCell ("$($o.PhysicalOp)" + $(if ($o.LogicalOp -ne $o.PhysicalOp) { " ($($o.LogicalOp))" } else { '' }))), (ConvertTo-MdCell $o.Object), (Format-Num $o.EstRows 0), (Format-Num $o.ActRowsPerExec 0), (Format-Num $execs 0), $o.CostPct, $o.Skew, (Format-Num $o.ActElapsedMs 0), (ConvertTo-MdCell (@($o.Warnings) -join '; ')))
                }
            }
            $ss = @($a.Stats | Where-Object { $_.Stmt -eq $s.Index } | Select-Object -First 30)
            if ($ss.Count) {
                Add-Line '- Statistics used by the optimizer:'
                Add-Line '  | Table | Statistics | Last update | Modifications | Sampling % |'
                Add-Line '  |---|---|---|---|---|'
                foreach ($x in $ss) { Add-Line ("  | {0} | {1} | {2} | {3} | {4} |" -f (ConvertTo-MdCell $x.Table), (ConvertTo-MdCell $x.Statistics), $x.LastUpdate, (Format-Num $x.ModificationCount 0), (Format-Num $x.SamplingPercent 1)) }
            }
            $wt = @($a.Waits | Where-Object { $_.Stmt -eq $s.Index } | Sort-Object WaitMs -Descending | Select-Object -First 10)
            if ($wt.Count) { Add-Line ('- Waits: ' + ((@($wt | ForEach-Object { "$($_.WaitType)=$(Format-Num $_.WaitMs 0)ms" })) -join ', ')) }
        }

        if ($p.Findings.Count) {
            Add-Line ''
            Add-Line "#### $tag script findings (heuristics)"
            foreach ($fd in ($p.Findings | Sort-Object @{ Expression = { @{ High = 0; Medium = 1; Info = 2 }[$_.Severity] } })) { Add-Line "- [$($fd.Severity)] $($fd.Where): $($fd.Message)" }
        }
    }

    if ($IncludeXml) {
        Add-Line ''
        Add-Line '## Raw plan XML'
        for ($i = 0; $i -lt $n; $i++) {
            $x = $Items[$i].Xml -replace '>\s+<', '><'
            if ($x.Length -gt 300000) { $x = $x.Substring(0, 300000) + '<!-- truncated -->' }
            Add-Line ''; Add-Line "### $(Get-PlanTag $i) XML"; Add-Line '```xml'; Add-Line $x; Add-Line '```'
        }
    }
    return $sb.ToString()
}

function Get-ComparisonRows {
    param($Items)
    $rows = New-Object System.Collections.ArrayList
    if ($Items.Count -eq 0) { return @() }
    $keys = New-Object System.Collections.Generic.List[string]
    $keys.Add('Source')
    foreach ($p in $Items) { foreach ($k in $p.Analysis.Summary.Keys) { if (-not $keys.Contains($k)) { $keys.Add($k) } } }
    foreach ($p in $Items) { if ($p.Runtime) { foreach ($k in $p.Runtime.Keys) { if (-not $keys.Contains($k)) { $keys.Add($k) } } } }
    foreach ($k in $keys) {
        $vals = @()
        foreach ($p in $Items) {
            if ($k -eq 'Source') { $vals += [string]$p.Source }
            elseif ($p.Analysis.Summary.Contains($k)) { $vals += [string]$p.Analysis.Summary[$k] }
            elseif ($p.Runtime -and $p.Runtime.Contains($k)) { $vals += [string]$p.Runtime[$k] }
            else { $vals += '' }
        }
        $differs = $Items.Count -gt 1 -and @($vals | Select-Object -Unique).Count -gt 1
        [void]$rows.Add([pscustomobject]@{ Metric = $k; Values = $vals; Differs = $differs })
    }
    return @($rows)
}

function Get-ExportObject {
    $plans = @()
    for ($i = 0; $i -lt $script:Plans.Count; $i++) {
        $p = $script:Plans[$i]; $a = $p.Analysis
        $plans += [pscustomobject]@{
            Tag = (Get-PlanTag $i); Label = $p.Label; Source = $p.Source; Runtime = $p.Runtime
            Summary = $a.Summary; Statements = $a.Statements; Operators = $a.Operators; Warnings = $a.Warnings
            MissingIndexes = $a.MissingIndexes; Parameters = $a.Parameters; Statistics = $a.Stats; Waits = $a.Waits
            Findings = $p.Findings
        }
    }
    return [pscustomobject]@{ Generated = (Get-Date).ToString('s'); CrossPlanFindings = $script:CrossFindings; Plans = $plans }
}

#endregion

#region ---------- Plan list management ----------

function Get-UniqueLabel {
    param([string]$Label, $Exclude)
    if ([string]::IsNullOrWhiteSpace($Label)) { $Label = 'Plan' }
    $existing = @($script:Plans | Where-Object { -not [object]::ReferenceEquals($_, $Exclude) } | ForEach-Object { $_.Label })
    $l = $Label; $k = 2
    while ($existing -contains $l) { $l = "$Label ($k)"; $k++ }
    return $l
}

function Add-Plan {
    param([string]$Xml, [string]$Label, [string]$Source, $Runtime)
    $a = Get-PlanAnalysis -PlanXml $Xml
    $item = [pscustomobject]@{
        Label = (Get-UniqueLabel $Label $null); Source = $Source; Xml = $Xml; Analysis = $a
        Runtime = $Runtime; Findings = @(Get-PlanFindings $a)
    }
    [void]$script:Plans.Add($item)
}

function Add-PlanFiles {
    param([string[]]$Files)
    $errors = @()
    foreach ($f in $Files) {
        try {
            $xml = [IO.File]::ReadAllText($f)
            Add-Plan -Xml $xml -Label ([IO.Path]::GetFileNameWithoutExtension($f)) -Source $f -Runtime $null
        } catch { $errors += "$([IO.Path]::GetFileName($f)): $($_.Exception.Message)" }
    }
    Update-AllViews
    if ($errors.Count) { [void][Windows.Forms.MessageBox]::Show(($errors -join "`r`n"), 'Some files could not be loaded', 'OK', 'Warning') }
}

#endregion

#region ---------- Server capture dialog ----------

function Show-ServerDialog {
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text = 'Capture plans from SQL Server'
    $dlg.Size = New-Object Drawing.Size(1150, 680)
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = $form.Font

    $grid = New-Object Windows.Forms.DataGridView
    $grid.Dock = 'Fill'; $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = 'DisplayedCells'; $grid.RowHeadersVisible = $false
    $dlg.Controls.Add($grid)

    $top = New-Object Windows.Forms.Panel
    $top.Dock = 'Top'; $top.Height = 100
    $dlg.Controls.Add($top)

    $bottom = New-Object Windows.Forms.Panel
    $bottom.Dock = 'Bottom'; $bottom.Height = 42
    $dlg.Controls.Add($bottom)

    function New-Lbl($t, $x, $y) { $l = New-Object Windows.Forms.Label; $l.Text = $t; $l.AutoSize = $true; $l.Location = New-Object Drawing.Point($x, ($y + 3)); $top.Controls.Add($l); $l }
    function New-Txt($x, $y, $w) { $t = New-Object Windows.Forms.TextBox; $t.Location = New-Object Drawing.Point($x, $y); $t.Width = $w; $top.Controls.Add($t); $t }

    [void](New-Lbl 'Server' 8 8);   $txtServer = New-Txt 60 8 200; $txtServer.Text = 'localhost'
    [void](New-Lbl 'Database' 272 8); $txtDb = New-Txt 335 8 150; $txtDb.Text = 'master'
    [void](New-Lbl 'Auth' 497 8)
    $cboAuth = New-Object Windows.Forms.ComboBox; $cboAuth.DropDownStyle = 'DropDownList'; $cboAuth.Location = New-Object Drawing.Point(535, 8); $cboAuth.Width = 130
    [void]$cboAuth.Items.AddRange(@('Windows', 'SQL login')); $cboAuth.SelectedIndex = 0; $top.Controls.Add($cboAuth)
    [void](New-Lbl 'User' 677 8); $txtUser = New-Txt 712 8 120
    [void](New-Lbl 'Password' 842 8); $txtPwd = New-Txt 905 8 120; $txtPwd.UseSystemPasswordChar = $true
    $txtUser.Enabled = $false; $txtPwd.Enabled = $false
    $cboAuth.Add_SelectedIndexChanged({ $on = $cboAuth.SelectedIndex -eq 1; $txtUser.Enabled = $on; $txtPwd.Enabled = $on })

    [void](New-Lbl 'Source' 8 40)
    $cboMode = New-Object Windows.Forms.ComboBox; $cboMode.DropDownStyle = 'DropDownList'; $cboMode.Location = New-Object Drawing.Point(60, 40); $cboMode.Width = 330
    [void]$cboMode.Items.AddRange(@(
        'query_hash: Query Store + plan cache (all plan hashes)',
        'query_hash: Query Store only (plan history)',
        'query_hash: plan cache only (current plans)',
        'query_plan_hash: Query Store + plan cache',
        'Query Store: query_id list (all plans)',
        'Query Store: plan_id list',
        'Query Store: query text contains',
        'Plan cache: query text contains'))
    $cboMode.SelectedIndex = 0; $top.Controls.Add($cboMode)
    [void](New-Lbl 'Search' 400 40); $txtSearch = New-Txt 450 40 375
    [void](New-Lbl 'Top' 835 40)
    $numTop = New-Object Windows.Forms.NumericUpDown; $numTop.Location = New-Object Drawing.Point(865, 40); $numTop.Width = 60; $numTop.Minimum = 1; $numTop.Maximum = 500; $numTop.Value = 50; $top.Controls.Add($numTop)
    $btnFetch = New-Object Windows.Forms.Button; $btnFetch.Text = 'Fetch'; $btnFetch.Location = New-Object Drawing.Point(940, 38); $btnFetch.Width = 85; $top.Controls.Add($btnFetch)

    $hint = New-Lbl 'Hashes/IDs: comma-separated (0x5A6A..., 0x1B2C...). Query Store runs in the selected database and keeps plan history; the plan cache is server-wide but holds only current plans. Select 2+ rows, then "Add selected".' 8 72
    $hint.ForeColor = [Drawing.Color]::DimGray

    $btnAdd = New-Object Windows.Forms.Button; $btnAdd.Text = 'Add selected plans'; $btnAdd.Width = 150; $btnAdd.Location = New-Object Drawing.Point(8, 8); $bottom.Controls.Add($btnAdd)
    $btnClose = New-Object Windows.Forms.Button; $btnClose.Text = 'Close'; $btnClose.Width = 90; $btnClose.Location = New-Object Drawing.Point(166, 8); $bottom.Controls.Add($btnClose)
    $lblInfo = New-Object Windows.Forms.Label; $lblInfo.AutoSize = $true; $lblInfo.Location = New-Object Drawing.Point(270, 14); $bottom.Controls.Add($lblInfo)
    $btnClose.Add_Click({ $dlg.Close() })

    $qsSelect = @'
SELECT TOP (@top)
    CONCAT('QS q', p.query_id, ' p', p.plan_id) AS label,
    'Query Store' AS plan_source,
    CONVERT(varchar(18), q.query_hash, 1) AS query_hash,
    CONVERT(varchar(18), p.query_plan_hash, 1) AS query_plan_hash,
    p.query_id, p.plan_id, p.is_forced_plan,
    CAST(rs.execs AS bigint) AS execs, rs.avg_duration_ms, rs.avg_cpu_ms, rs.avg_logical_reads, rs.max_duration_ms,
    CONVERT(datetime2(0), p.initial_compile_start_time) AS creation_time,
    CONVERT(datetime2(0), p.last_execution_time) AS last_execution_time,
    DB_NAME() AS database_name,
    LEFT(qt.query_sql_text, 200) AS query_text,
    p.query_plan AS plan_xml
FROM sys.query_store_plan AS p
JOIN sys.query_store_query AS q ON q.query_id = p.query_id
JOIN sys.query_store_query_text AS qt ON qt.query_text_id = q.query_text_id
OUTER APPLY (
    SELECT SUM(r.count_executions) AS execs,
           CAST(SUM(r.avg_duration * r.count_executions) / NULLIF(SUM(r.count_executions), 0) / 1000.0 AS decimal(18,2)) AS avg_duration_ms,
           CAST(SUM(r.avg_cpu_time * r.count_executions) / NULLIF(SUM(r.count_executions), 0) / 1000.0 AS decimal(18,2)) AS avg_cpu_ms,
           CAST(SUM(r.avg_logical_io_reads * r.count_executions) / NULLIF(SUM(r.count_executions), 0) AS decimal(18,0)) AS avg_logical_reads,
           CAST(MAX(r.max_duration) / 1000.0 AS decimal(18,2)) AS max_duration_ms
    FROM sys.query_store_runtime_stats AS r
    WHERE r.plan_id = p.plan_id) AS rs
WHERE p.query_plan IS NOT NULL AND {FILTER}
ORDER BY p.query_id, p.plan_id;
'@
    $cacheSelect = @'
SELECT TOP (@top)
    CONCAT('Cache ', CONVERT(varchar(18), qs.query_plan_hash, 1)) AS label,
    'Plan cache' AS plan_source,
    CONVERT(varchar(18), qs.query_hash, 1) AS query_hash,
    CONVERT(varchar(18), qs.query_plan_hash, 1) AS query_plan_hash,
    CAST(qs.execution_count AS bigint) AS execs,
    CAST(qs.total_elapsed_time / 1000.0 / qs.execution_count AS decimal(18,2)) AS avg_duration_ms,
    CAST(qs.total_worker_time / 1000.0 / qs.execution_count AS decimal(18,2)) AS avg_cpu_ms,
    CAST(qs.total_logical_reads / qs.execution_count AS decimal(18,0)) AS avg_logical_reads,
    CAST(qs.max_elapsed_time / 1000.0 AS decimal(18,2)) AS max_duration_ms,
    CONVERT(datetime2(0), qs.creation_time) AS creation_time,
    CONVERT(datetime2(0), qs.last_execution_time) AS last_execution_time,
    DB_NAME(CONVERT(int, pa.value)) AS database_name,
    LEFT(SUBSTRING(st.text, qs.statement_start_offset / 2 + 1,
        (CASE qs.statement_end_offset WHEN -1 THEN DATALENGTH(st.text) ELSE qs.statement_end_offset END - qs.statement_start_offset) / 2 + 1), 200) AS query_text,
    tqp.query_plan AS plan_xml
FROM sys.dm_exec_query_stats AS qs
CROSS APPLY sys.dm_exec_sql_text(qs.sql_handle) AS st
CROSS APPLY sys.dm_exec_text_query_plan(qs.plan_handle, qs.statement_start_offset, qs.statement_end_offset) AS tqp
OUTER APPLY (SELECT value FROM sys.dm_exec_plan_attributes(qs.plan_handle) WHERE attribute = N'dbid') AS pa
WHERE tqp.query_plan IS NOT NULL AND {FILTER}
ORDER BY qs.total_worker_time DESC;
'@

    $btnFetch.Add_Click({
        try {
            $dlg.Cursor = 'WaitCursor'
            $search = $txtSearch.Text.Trim()
            $mode = $cboMode.SelectedIndex

            # Validated literals only, so they are safe to inline (binary literals / integers)
            $hashList = {
                $h = @($search -split '[,\s;]+' | Where-Object { $_ })
                if ($h.Count -eq 0) { throw 'Enter one or more hashes, comma-separated (e.g. 0x5A6A4FC86EECD6DD, 0x1B2C...).' }
                $bad = @($h | Where-Object { $_ -notmatch '^0x[0-9A-Fa-f]{16}$' })
                if ($bad.Count) { throw "Not a valid 8-byte hash: $($bad -join ', ')  (expected 0x + 16 hex digits)" }
                ($h | Select-Object -Unique) -join ', '
            }
            $idList = {
                $ids = @($search -split '[,\s;]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [long]$_ })
                if ($ids.Count -eq 0) { throw 'Enter one or more numeric IDs, comma-separated.' }
                $ids -join ','
            }

            # Each job: which select, the filter, whether @search is used, whether failure is fatal
            $jobs = @()
            switch ($mode) {
                0 { $l = & $hashList
                    $jobs += @{ Name = 'Query Store'; Sql = $qsSelect;    Filter = "q.query_hash IN ($l)";  Param = $false; Optional = $true }
                    $jobs += @{ Name = 'Plan cache';  Sql = $cacheSelect; Filter = "qs.query_hash IN ($l)"; Param = $false; Optional = $true } }
                1 { $l = & $hashList; $jobs += @{ Name = 'Query Store'; Sql = $qsSelect;    Filter = "q.query_hash IN ($l)";  Param = $false; Optional = $false } }
                2 { $l = & $hashList; $jobs += @{ Name = 'Plan cache';  Sql = $cacheSelect; Filter = "qs.query_hash IN ($l)"; Param = $false; Optional = $false } }
                3 { $l = & $hashList
                    $jobs += @{ Name = 'Query Store'; Sql = $qsSelect;    Filter = "p.query_plan_hash IN ($l)";  Param = $false; Optional = $true }
                    $jobs += @{ Name = 'Plan cache';  Sql = $cacheSelect; Filter = "qs.query_plan_hash IN ($l)"; Param = $false; Optional = $true } }
                4 { $l = & $idList; $jobs += @{ Name = 'Query Store'; Sql = $qsSelect; Filter = "p.query_id IN ($l)"; Param = $false; Optional = $false } }
                5 { $l = & $idList; $jobs += @{ Name = 'Query Store'; Sql = $qsSelect; Filter = "p.plan_id IN ($l)";  Param = $false; Optional = $false } }
                6 { if (-not $search) { throw 'Enter text to search for.' }
                    $jobs += @{ Name = 'Query Store'; Sql = $qsSelect; Filter = 'qt.query_sql_text LIKE @search'; Param = $true; Optional = $false } }
                7 { if (-not $search) { throw 'Enter text to search for.' }
                    $jobs += @{ Name = 'Plan cache'; Sql = $cacheSelect; Filter = "st.text LIKE @search AND st.text NOT LIKE N'%dm_exec_query_stats%'"; Param = $true; Optional = $false } }
            }

            $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
            # Builder is an IDictionary: PowerShell turns $csb.Prop = x into a bogus key, so use the indexer with real keywords
            $csb.set_Item('Data Source', $txtServer.Text.Trim())
            $csb.set_Item('Initial Catalog', $txtDb.Text.Trim())
            $csb.set_Item('Application Name', 'ExecutionPlanComparer')
            $csb.set_Item('Connect Timeout', 15)
            $csb.set_Item('TrustServerCertificate', $true)
            if ($cboAuth.SelectedIndex -eq 0) { $csb.set_Item('Integrated Security', $true) }
            else { $csb.set_Item('User ID', $txtUser.Text); $csb.set_Item('Password', $txtPwd.Text) }

            $dt = $null; $notes = @(); $okJobs = 0
            $conn = New-Object System.Data.SqlClient.SqlConnection($csb.get_ConnectionString())
            try {
                $conn.Open()
                foreach ($j in $jobs) {
                    try {
                        $cmd = $conn.CreateCommand()
                        $cmd.CommandText = $j.Sql.Replace('{FILTER}', $j.Filter)
                        $cmd.CommandTimeout = 120
                        [void]$cmd.Parameters.AddWithValue('@top', [int]$numTop.Value)
                        if ($j.Param) { [void]$cmd.Parameters.Add('@search', [Data.SqlDbType]::NVarChar, 4000); $cmd.Parameters['@search'].Value = "%$search%" }
                        $t = New-Object System.Data.DataTable
                        [void](New-Object System.Data.SqlClient.SqlDataAdapter($cmd)).Fill($t)
                        $notes += "$($j.Name): $($t.Rows.Count)"
                        $okJobs++
                        if ($null -eq $dt) { $dt = $t } else { $dt.Merge($t, $false, [Data.MissingSchemaAction]::Add) }
                    } catch {
                        if (-not $j.Optional) { throw }
                        $notes += "$($j.Name): skipped ($($_.Exception.Message))"
                    }
                }
            } finally { $conn.Dispose() }
            if ($okJobs -eq 0) { throw ('No source could be queried. ' + ($notes -join ' | ')) }

            $dv = $dt.DefaultView
            $dv.Sort = 'query_hash, query_plan_hash, plan_source'
            $grid.DataSource = $dv
            if ($grid.Columns.Contains('plan_xml')) { $grid.Columns['plan_xml'].Visible = $false }

            $distinct = @($dt.Rows | ForEach-Object { $_['query_plan_hash'] } | Where-Object { $_ -isnot [DBNull] } | Select-Object -Unique).Count
            $info = "$($dt.Rows.Count) row(s), $distinct distinct query_plan_hash.  [$($notes -join ' | ')]"
            if ($mode -in 0, 1, 2 -and $distinct -le 1 -and $dt.Rows.Count -gt 0) {
                $info += '  Only one plan shape found: the cache keeps just current plans; check Query Store is on for this database.'
            }
            $lblInfo.Text = $info
            $script:ServerDesc = "$($txtServer.Text.Trim())/$($txtDb.Text.Trim()) - $($cboMode.Text)"
        } catch {
            [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Fetch failed', 'OK', 'Error')
        } finally { $dlg.Cursor = 'Default' }
    })

    $btnAdd.Add_Click({
        $added = 0; $errors = @()
        $runtimeCols = [ordered]@{
            'plan_source' = 'Server: plan source'; 'query_id' = 'Server: query_id'; 'plan_id' = 'Server: plan_id'
            'execs' = 'Server: executions'; 'avg_duration_ms' = 'Server: avg duration ms'; 'avg_cpu_ms' = 'Server: avg CPU ms'
            'avg_logical_reads' = 'Server: avg logical reads'; 'max_duration_ms' = 'Server: max duration ms'
            'is_forced_plan' = 'Server: forced plan'; 'creation_time' = 'Server: plan created'; 'last_execution_time' = 'Server: last execution'
            'database_name' = 'Server: database'
        }
        $rows = @($grid.SelectedRows | Sort-Object Index)
        foreach ($row in $rows) {
            $drv = $row.DataBoundItem
            try {
                if ($drv['plan_xml'] -is [DBNull]) { throw 'plan XML is NULL' }
                $rt = [ordered]@{}
                foreach ($c in $runtimeCols.Keys) {
                    if ($drv.Row.Table.Columns.Contains($c) -and -not ($drv[$c] -is [DBNull])) {
                        $v = $drv[$c]
                        if ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd HH:mm:ss') }
                        elseif ($c -in 'query_id', 'plan_id') { $v = [string]$v }
                        elseif ($v -is [long] -or $v -is [int]) { $v = Format-Num $v 0 }
                        elseif ($v -is [decimal] -or $v -is [double]) { $v = Format-Num $v 2 }
                        $rt[$runtimeCols[$c]] = [string]$v
                    }
                }
                Add-Plan -Xml ([string]$drv['plan_xml']) -Label ([string]$drv['label']) -Source $script:ServerDesc -Runtime $rt
                $added++
            } catch { $errors += "$($drv['label']): $($_.Exception.Message)" }
        }
        Update-AllViews
        $lblInfo.Text = "$added plan(s) added."
        if ($errors.Count) { [void][Windows.Forms.MessageBox]::Show(($errors -join "`r`n"), 'Some plans could not be added', 'OK', 'Warning') }
    })

    [void]$dlg.ShowDialog($form)
}

#endregion

#region ---------- Main form ----------

$form = New-Object Windows.Forms.Form
$form.Text = 'Execution Plan Comparer'
$form.Size = New-Object Drawing.Size(1450, 920)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Segoe UI', 9)
$mono = New-Object Drawing.Font('Consolas', 9.5)

function New-MonoBox {
    $t = New-Object Windows.Forms.TextBox
    $t.Multiline = $true; $t.ReadOnly = $true; $t.ScrollBars = 'Both'; $t.WordWrap = $false
    $t.Dock = 'Fill'; $t.Font = $mono; $t.MaxLength = 0; $t.BackColor = [Drawing.Color]::White
    return $t
}
function New-Grid {
    $g = New-Object Windows.Forms.DataGridView
    $g.Dock = 'Fill'; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $false; $g.SelectionMode = 'FullRowSelect'; $g.AutoSizeColumnsMode = 'DisplayedCells'
    $g.BackgroundColor = [Drawing.Color]::White
    return $g
}

# --- split: list on top, tabs below (Fill control added first so docking works)
$split = New-Object Windows.Forms.SplitContainer
$split.Dock = 'Fill'; $split.Orientation = 'Horizontal'
$form.Controls.Add($split)

$toolbar = New-Object Windows.Forms.FlowLayoutPanel
$toolbar.Dock = 'Top'; $toolbar.AutoSize = $true; $toolbar.Padding = New-Object Windows.Forms.Padding(4)
$form.Controls.Add($toolbar)

$status = New-Object Windows.Forms.StatusStrip
$statusLbl = New-Object Windows.Forms.ToolStripStatusLabel
$statusLbl.Text = 'Add two or more plans (files, drag & drop, clipboard, or server).'
[void]$status.Items.Add($statusLbl)
$form.Controls.Add($status)

$lv = New-Object Windows.Forms.ListView
$lv.Dock = 'Fill'; $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.MultiSelect = $true; $lv.HideSelection = $false; $lv.AllowDrop = $true
foreach ($c in @(@('Plan', 45), @('Label', 220), @('Type', 75), @('Stmts', 50), @('Est. cost', 90), @('Elapsed ms', 85), @('High/Med findings', 115), @('Source', 600))) {
    [void]$lv.Columns.Add($c[0], $c[1])
}
$split.Panel1.Controls.Add($lv)

$tabs = New-Object Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$split.Panel2.Controls.Add($tabs)

function New-Tab($title) { $tp = New-Object Windows.Forms.TabPage; $tp.Text = $title; [void]$tabs.TabPages.Add($tp); return $tp }

$tabCmp = New-Tab 'Comparison'
$gridCmp = New-Grid; $tabCmp.Controls.Add($gridCmp)
$gridCmp.Add_DataBindingComplete({
    if ($gridCmp.Columns.Contains('_diff')) {
        $gridCmp.Columns['_diff'].Visible = $false
        foreach ($r in $gridCmp.Rows) { if ($r.Cells['_diff'].Value -eq $true) { $r.DefaultCellStyle.BackColor = [Drawing.Color]::LightYellow } }
    }
    if ($gridCmp.Columns.Count) { $gridCmp.Columns[0].DefaultCellStyle.Font = New-Object Drawing.Font($form.Font, [Drawing.FontStyle]::Bold) }
})

$tabFind = New-Tab 'Findings'
$txtFind = New-MonoBox; $tabFind.Controls.Add($txtFind)

$tabOps = New-Tab 'Operators'
$gridOps = New-Grid; $tabOps.Controls.Add($gridOps)

$tabTree = New-Tab 'Plan trees'
$txtTree = New-MonoBox; $tabTree.Controls.Add($txtTree)

$tabAi = New-Tab 'AI prompt'
$txtPrompt = New-MonoBox; $tabAi.Controls.Add($txtPrompt)
$aiTop = New-Object Windows.Forms.Panel; $aiTop.Dock = 'Top'; $aiTop.Height = 112; $tabAi.Controls.Add($aiTop)
$lblNotes = New-Object Windows.Forms.Label; $lblNotes.Text = 'Context for the AI (what changed, symptoms, goal, compat level, data volumes...):'; $lblNotes.AutoSize = $true; $lblNotes.Location = New-Object Drawing.Point(6, 6); $aiTop.Controls.Add($lblNotes)
$txtNotes = New-Object Windows.Forms.TextBox; $txtNotes.Multiline = $true; $txtNotes.ScrollBars = 'Vertical'; $txtNotes.Location = New-Object Drawing.Point(6, 26); $txtNotes.Size = New-Object Drawing.Size(1380, 50); $txtNotes.Anchor = 'Top, Left, Right'; $aiTop.Controls.Add($txtNotes)
$chkXml = New-Object Windows.Forms.CheckBox; $chkXml.Text = 'Include raw plan XML (large)'; $chkXml.AutoSize = $true; $chkXml.Location = New-Object Drawing.Point(6, 84); $aiTop.Controls.Add($chkXml)
$lblTop = New-Object Windows.Forms.Label; $lblTop.Text = 'Top operators per statement:'; $lblTop.AutoSize = $true; $lblTop.Location = New-Object Drawing.Point(220, 86); $aiTop.Controls.Add($lblTop)
$numOps = New-Object Windows.Forms.NumericUpDown; $numOps.Minimum = 3; $numOps.Maximum = 100; $numOps.Value = 15; $numOps.Width = 55; $numOps.Location = New-Object Drawing.Point(385, 83); $aiTop.Controls.Add($numOps)
$btnRegen = New-Object Windows.Forms.Button; $btnRegen.Text = 'Regenerate'; $btnRegen.Location = New-Object Drawing.Point(455, 81); $btnRegen.Width = 95; $aiTop.Controls.Add($btnRegen)
$lblSize = New-Object Windows.Forms.Label; $lblSize.AutoSize = $true; $lblSize.Location = New-Object Drawing.Point(565, 86); $aiTop.Controls.Add($lblSize)

function Update-Prompt {
    if ($script:Plans.Count -eq 0) { $txtPrompt.Text = ''; $lblSize.Text = ''; return }
    $p = New-AiPrompt -Items $script:Plans -Notes $txtNotes.Text -IncludeXml $chkXml.Checked -TopOps ([int]$numOps.Value)
    $txtPrompt.Text = $p -replace "(?<!`r)`n", "`r`n"
    $lblSize.Text = "$(Format-Num $p.Length 0) chars, ~$(Format-Num ($p.Length / 4) 0) tokens"
}

function Update-AllViews {
    $form.Cursor = 'WaitCursor'
    try {
        $items = @($script:Plans)
        $script:CrossFindings = @(Get-CrossFindings $items)

        # list
        $lv.BeginUpdate(); $lv.Items.Clear()
        for ($i = 0; $i -lt $items.Count; $i++) {
            $p = $items[$i]; $a = $p.Analysis
            $hm = @($p.Findings | Where-Object { $_.Severity -in 'High', 'Medium' }).Count
            $li = New-Object Windows.Forms.ListViewItem((Get-PlanTag $i))
            foreach ($v in @($p.Label, $a.Summary['Plan type'], $a.Statements.Count, (Format-Num $a.TotalCost 3), (Format-Num $a.ElapsedMs 0), $hm, $p.Source)) { [void]$li.SubItems.Add([string]$v) }
            [void]$lv.Items.Add($li)
        }
        $lv.EndUpdate()

        # comparison grid
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add('Metric')
        for ($i = 0; $i -lt $items.Count; $i++) { [void]$dt.Columns.Add("$(Get-PlanTag $i) - $($items[$i].Label)") }
        [void]$dt.Columns.Add('_diff', [bool])
        foreach ($r in (Get-ComparisonRows $items)) {
            $row = $dt.NewRow(); $row[0] = $r.Metric
            for ($i = 0; $i -lt $items.Count; $i++) { $row[$i + 1] = $r.Values[$i] }
            $row['_diff'] = [bool]$r.Differs
            $dt.Rows.Add($row)
        }
        $gridCmp.DataSource = $null; $gridCmp.DataSource = $dt

        # findings
        $sb = New-Object System.Text.StringBuilder
        if ($items.Count -ge 2) {
            [void]$sb.AppendLine('=== CROSS-PLAN DIFFERENCES ===')
            foreach ($c in $script:CrossFindings) { [void]$sb.AppendLine("* $c") }
            [void]$sb.AppendLine('')
        } elseif ($items.Count -eq 1) { [void]$sb.AppendLine('(Add at least one more plan to get cross-plan comparison.)'); [void]$sb.AppendLine('') }
        for ($i = 0; $i -lt $items.Count; $i++) {
            $p = $items[$i]
            [void]$sb.AppendLine("=== $(Get-PlanTag $i) - $($p.Label) ===")
            $sorted = $p.Findings | Sort-Object @{ Expression = { @{ High = 0; Medium = 1; Info = 2 }[$_.Severity] } }
            foreach ($fd in $sorted) { [void]$sb.AppendLine(('[{0,-6}] {1,-28} {2}' -f $fd.Severity, $fd.Where, $fd.Message)) }
            if (-not $p.Findings.Count) { [void]$sb.AppendLine('(no findings)') }
            [void]$sb.AppendLine('')
        }
        $txtFind.Text = $sb.ToString()

        # operators grid
        $od = New-Object System.Data.DataTable
        $cols = @(@('Plan', [string]), @('Stmt', [int]), @('Node', [int]), @('Operator', [string]), @('Logical', [string]), @('Object', [string]),
                  @('Est rows', [double]), @('Est execs', [double]), @('Act rows/exec', [double]), @('Act execs', [double]), @('Cost %', [double]),
                  @('Skew x', [double]), @('Act ms', [double]), @('Logical reads', [double]), @('Mode', [string]), @('Seek predicate', [string]),
                  @('Predicate', [string]), @('Warnings', [string]))
        foreach ($c in $cols) { [void]$od.Columns.Add($c[0], $c[1]) }
        for ($i = 0; $i -lt $items.Count; $i++) {
            foreach ($o in $items[$i].Analysis.Operators) {
                $r = $od.NewRow()
                $r['Plan'] = Get-PlanTag $i; $r['Stmt'] = $o.Stmt; $r['Node'] = $o.NodeId; $r['Operator'] = $o.PhysicalOp; $r['Logical'] = $o.LogicalOp; $r['Object'] = $o.Object
                $r['Est rows'] = ConvertTo-DbValue $o.EstRows; $r['Est execs'] = ConvertTo-DbValue $o.EstExecs; $r['Act rows/exec'] = ConvertTo-DbValue $o.ActRowsPerExec
                $r['Act execs'] = ConvertTo-DbValue $o.ActExecs; $r['Cost %'] = ConvertTo-DbValue $o.CostPct; $r['Skew x'] = ConvertTo-DbValue $o.Skew
                $r['Act ms'] = ConvertTo-DbValue $o.ActElapsedMs; $r['Logical reads'] = ConvertTo-DbValue $o.ActLogicalReads; $r['Mode'] = $o.ExecMode
                $r['Seek predicate'] = Limit-Text $o.SeekPredicate 300; $r['Predicate'] = Limit-Text $o.Predicate 300; $r['Warnings'] = (@($o.Warnings) -join '; ')
                $od.Rows.Add($r)
            }
        }
        $gridOps.DataSource = $null; $gridOps.DataSource = $od

        # trees
        $tb = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $items.Count; $i++) {
            $p = $items[$i]
            [void]$tb.AppendLine("##### $(Get-PlanTag $i) - $($p.Label) ($($p.Analysis.Summary['Plan type']))")
            foreach ($s in $p.Analysis.Statements) {
                [void]$tb.AppendLine("--- Stmt $($s.Index) [$($s.Type)] cost=$(Format-Num $s.Cost 3): $(Limit-Text $s.Text 150)")
                [void]$tb.AppendLine($s.Tree); [void]$tb.AppendLine('')
            }
        }
        $txtTree.Text = $tb.ToString() -replace "(?<!`r)`n", "`r`n"

        Update-Prompt
        $statusLbl.Text = "$($items.Count) plan(s) loaded. " + $(if ($items.Count -lt 2) { 'Add at least 2 to compare.' } else { 'Yellow rows in Comparison differ between plans.' })
    } catch {
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message + "`r`n" + $_.ScriptStackTrace, 'Refresh failed', 'OK', 'Error')
    } finally { $form.Cursor = 'Default' }
}

function New-ToolButton {
    param([string]$Text, [scriptblock]$OnClick)
    $b = New-Object Windows.Forms.Button
    $b.Text = $Text; $b.AutoSize = $true; $b.Height = 28
    $b.Add_Click($OnClick)
    $toolbar.Controls.Add($b)
    return $b
}

[void](New-ToolButton 'Add plan files...' {
    $ofd = New-Object Windows.Forms.OpenFileDialog
    $ofd.Filter = 'Execution plans (*.sqlplan;*.xml)|*.sqlplan;*.xml|All files (*.*)|*.*'
    $ofd.Multiselect = $true; $ofd.Title = 'Select two or more execution plans'
    if ($ofd.ShowDialog($form) -eq 'OK') { Add-PlanFiles $ofd.FileNames }
})

[void](New-ToolButton 'Paste plan XML' {
    try {
        $t = [Windows.Forms.Clipboard]::GetText()
        if ([string]::IsNullOrWhiteSpace($t)) { throw 'Clipboard does not contain text.' }
        $start = $t.IndexOf('<'); if ($start -gt 0) { $t = $t.Substring($start) }
        Add-Plan -Xml $t -Label "Pasted $($script:Plans.Count + 1)" -Source 'Clipboard' -Runtime $null
        Update-AllViews
    } catch { [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Paste failed', 'OK', 'Warning') }
})

[void](New-ToolButton 'From SQL Server...' { Show-ServerDialog })

[void](New-ToolButton 'Rename' {
    if ($lv.SelectedIndices.Count -ne 1) { return }
    $p = $script:Plans[$lv.SelectedIndices[0]]
    $new = [Microsoft.VisualBasic.Interaction]::InputBox('New label (e.g. "Before index", "Slow - sniffed"):', 'Rename plan', $p.Label)
    if (-not [string]::IsNullOrWhiteSpace($new)) { $p.Label = Get-UniqueLabel $new.Trim() $p; Update-AllViews }
})

[void](New-ToolButton 'Remove' {
    $idx = @($lv.SelectedIndices | ForEach-Object { $_ } | Sort-Object -Descending)
    foreach ($i in $idx) { $script:Plans.RemoveAt($i) }
    if ($idx.Count) { Update-AllViews }
})

[void](New-ToolButton 'Move up' {
    if ($lv.SelectedIndices.Count -ne 1) { return }
    $i = $lv.SelectedIndices[0]; if ($i -le 0) { return }
    $p = $script:Plans[$i]; $script:Plans.RemoveAt($i); $script:Plans.Insert($i - 1, $p)
    Update-AllViews; $lv.Items[$i - 1].Selected = $true
})

[void](New-ToolButton 'Clear' { $script:Plans.Clear(); Update-AllViews })

$sep = New-Object Windows.Forms.Label; $sep.Width = 20; $toolbar.Controls.Add($sep)

[void](New-ToolButton 'Copy AI prompt' {
    if ($script:Plans.Count -eq 0) { return }
    Update-Prompt
    [Windows.Forms.Clipboard]::SetText($txtPrompt.Text)
    $statusLbl.Text = "AI prompt copied to clipboard ($($lblSize.Text))."
})

[void](New-ToolButton 'Save AI prompt (.md)...' {
    if ($script:Plans.Count -eq 0) { return }
    Update-Prompt
    $sfd = New-Object Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Markdown (*.md)|*.md|Text (*.txt)|*.txt'; $sfd.FileName = "PlanComparison_$(Get-Date -Format yyyyMMdd_HHmmss).md"
    if ($sfd.ShowDialog($form) -eq 'OK') { [IO.File]::WriteAllText($sfd.FileName, $txtPrompt.Text, (New-Object Text.UTF8Encoding($false))); $statusLbl.Text = "Saved $($sfd.FileName)" }
})

[void](New-ToolButton 'Export JSON...' {
    if ($script:Plans.Count -eq 0) { return }
    $sfd = New-Object Windows.Forms.SaveFileDialog
    $sfd.Filter = 'JSON (*.json)|*.json'; $sfd.FileName = "PlanComparison_$(Get-Date -Format yyyyMMdd_HHmmss).json"
    if ($sfd.ShowDialog($form) -eq 'OK') {
        $json = Get-ExportObject | ConvertTo-Json -Depth 8
        [IO.File]::WriteAllText($sfd.FileName, $json, (New-Object Text.UTF8Encoding($false)))
        $statusLbl.Text = "Saved $($sfd.FileName)"
    }
})

$btnRegen.Add_Click({ Update-Prompt })
$chkXml.Add_CheckedChanged({ Update-Prompt })

$lv.Add_DoubleClick({
    if ($lv.SelectedIndices.Count -ne 1) { return }
    $p = $script:Plans[$lv.SelectedIndices[0]]
    $new = [Microsoft.VisualBasic.Interaction]::InputBox('New label:', 'Rename plan', $p.Label)
    if (-not [string]::IsNullOrWhiteSpace($new)) { $p.Label = Get-UniqueLabel $new.Trim() $p; Update-AllViews }
})
$lv.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [Windows.Forms.DragDropEffects]::Copy } })
$lv.Add_DragDrop({ Add-PlanFiles ([string[]]$_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) })

$form.Add_Shown({ $split.SplitterDistance = 170; $form.Activate() })
[void]$form.ShowDialog()
$form.Dispose()

#endregion

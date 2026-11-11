@echo off
rem ==========================================================================
rem  PBIP_Context.cmd
rem  Drop this file into a Power BI Project (.pbip) folder and run it.
rem  It writes PBIP_Context.txt: a dense, logic-only summary of the project
rem  (tables, columns, DAX, M queries, relationships, and which fields each
rem  visual uses) meant to be SCREENSHOTTED and sent to an AI assistant
rem  together with screenshots of the report itself.
rem
rem  Left out on purpose: everything a report screenshot already shows
rem  (layout, colours, fonts, formatting) and all tooling noise (lineage
rem  tags, annotations, schemas, cultures, auto date tables).
rem
rem  Every item is one line with no blank lines, and short items share a
rem  line, so open the output with Word Wrap ON to fill each screenshot.
rem ==========================================================================
setlocal

rem Short items are packed onto one line up to this many characters.
rem Set it to roughly the number of characters that fit across your window.
set "WIDTH=200"

rem Embedded "Enter Data" columns with at most this many distinct values are
rem listed in full with counts; others are summarised.
set "MAXD=15"

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $s=[IO.File]::ReadAllText('%~f0'); $i=$s.IndexOf('#PS'+'START'); & ([scriptblock]::Create($s.Substring($i))) -Root '%~dp0' -Self '%~nx0' -MaxDistinct %MAXD% -Width %WIDTH% }"

echo.
if /i not "%~1"=="/nopause" pause
endlocal
exit /b 0

#PSSTART
param([string]$Root, [string]$Self, [int]$MaxDistinct = 15, [int]$Width = 200)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions

$Root     = $Root.TrimEnd('\')
$outName  = 'PBIP_Context.txt'
$outPath  = Join-Path $Root $outName
$skipDirs = '.pbi','.git','node_modules'

$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue

$out = New-Object System.Collections.ArrayList
$script:pack = ''
$script:nTables = 0
$script:nMeasures = 0
$script:nVisuals = 0

$blobRx = [regex]'Binary\.FromText\(\s*"([A-Za-z0-9+/=]{20,})"'
$declRx = [regex]'^([A-Za-z_]\w*)\s+(''(?:[^'']|'''')*''|[^\s=]+)\s*(?:=\s*(.*))?$'
$propRx = [regex]'^([A-Za-z_]\w*)\s*:\s*(.*)$'
$exprRx = [regex]'^([A-Za-z_]\w*)\s*=\s*(.*)$'
$relRx  = [regex]'^(''(?:[^'']|'''')*''|[^.]+)\.(.+)$'
$autoRx = 'LocalDateTable_|DateTableTemplate_'
$aggNames = 'Sum','Avg','DistinctCount','Min','Max','Count','Median','StdDev','Var'
$cmpOps   = '=','>','>=','<','<='

# ---------------------------------------------------------------- output ----

function Flush-Pack {
    if ($script:pack.Length -gt 0) { [void]$out.Add($script:pack); $script:pack = '' }
}

# A line of its own
function Add-Line([string]$s) {
    Flush-Pack
    [void]$out.Add($s)
}

# A short item that may share a line with its neighbours
function Add-Packed([string]$s) {
    if ($script:pack.Length -eq 0) { $script:pack = $s }
    elseif (($script:pack.Length + 5 + $s.Length) -le $Width) { $script:pack = $script:pack + '  |  ' + $s }
    else { [void]$out.Add($script:pack); $script:pack = $s }
}

# --------------------------------------------------------------- helpers ----

function Test-InSkippedDir([string]$rel) {
    $segs = $rel.Split('\')
    for ($i = 0; $i -lt $segs.Length - 1; $i++) {
        if ($skipDirs -contains $segs[$i].ToLower()) { return $true }
    }
    return $false
}

# Null-safe walk into deserialised JSON: G $obj 'a' 0 'b'
function G($o) {
    foreach ($k in $args) {
        if ($null -eq $o) { return }
        if ($o -is [System.Collections.IDictionary]) {
            if ($o.ContainsKey($k)) { $o = $o[$k] } else { return }
        }
        elseif (($o -is [System.Collections.IList]) -and ($k -is [int])) {
            if ($k -lt $o.Count) { $o = $o[$k] } else { return }
        }
        else { return }
    }
    if ($null -eq $o) { return }
    return $o
}

function Read-Json([string]$path) {
    return $ser.DeserializeObject([IO.File]::ReadAllText($path))
}

function Try-Json($f) {
    if ($null -eq $f) { return }
    try { return (Read-Json $f.Path) }
    catch { Add-Line "(could not parse $($f.Rel): $($_.Exception.Message))"; return }
}

function From-JsonText($s) {
    if ([string]::IsNullOrEmpty([string]$s)) { return }
    try { return $ser.DeserializeObject([string]$s) } catch { return }
}

# BIM expressions may be a string or an array of lines
function JText($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [array]) { return ($v -join "`n") }
    return [string]$v
}

function Cut([string]$s, [int]$n) {
    if ($s.Length -gt $n) { return $s.Substring(0, $n) + '...' }
    return $s
}

function Unq([string]$n) {
    if ($n.Length -ge 2 -and $n.StartsWith("'") -and $n.EndsWith("'")) {
        return $n.Substring(1, $n.Length - 2).Replace("''", "'")
    }
    return $n
}

# Quote a table name the way DAX does when it needs it
function QT([string]$n) {
    if ($n -match '^\w+$') { return $n }
    return "'" + $n + "'"
}

function Get-Owner([string]$rel, [string]$suffixRx) {
    $m = [regex]::Match($rel, '^(.*?\.(?:' + $suffixRx + '))\\', 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

# Collapse a DAX or M expression onto one line. String literals are left
# alone and line comments become /* */ so nothing after them is swallowed.
function Compress-Expr([string]$t, [bool]$dax) {
    $sb = New-Object System.Text.StringBuilder
    $n = $t.Length
    $i = 0
    while ($i -lt $n) {
        $c = $t[$i]
        $c2 = [char]0
        if (($i + 1) -lt $n) { $c2 = $t[$i + 1] }

        if (($c -eq '"') -or ($dax -and ($c -eq "'"))) {
            $j = $i + 1
            while ($j -lt $n) {
                if ($t[$j] -eq $c) {
                    if ((($j + 1) -lt $n) -and ($t[$j + 1] -eq $c)) { $j += 2; continue }
                    break
                }
                $j++
            }
            if ($j -ge $n) { $j = $n - 1 }
            [void]$sb.Append($t.Substring($i, $j - $i + 1))
            $i = $j + 1
            continue
        }

        if (($c -eq '/') -and ($c2 -eq '*')) {
            $j = $t.IndexOf('*/', $i + 2)
            if ($j -lt 0) { $j = $n } else { $j = $j + 2 }
            [void]$sb.Append(($t.Substring($i, $j - $i) -replace '\s+', ' '))
            $i = $j
            continue
        }

        if ((($c -eq '/') -and ($c2 -eq '/')) -or ($dax -and ($c -eq '-') -and ($c2 -eq '-'))) {
            $j = $t.IndexOf("`n", $i)
            if ($j -lt 0) { $j = $n }
            $cm = $t.Substring($i + 2, $j - $i - 2).Trim().Replace('*/', '* /')
            if ($cm.Length -gt 0) {
                if (($sb.Length -gt 0) -and ($sb[$sb.Length - 1] -ne ' ')) { [void]$sb.Append(' ') }
                [void]$sb.Append('/* ' + $cm + ' */ ')
            }
            $i = $j
            continue
        }

        if ([char]::IsWhiteSpace($c)) {
            if (($sb.Length -gt 0) -and ($sb[$sb.Length - 1] -ne ' ')) { [void]$sb.Append(' ') }
            $i++
            continue
        }

        [void]$sb.Append($c)
        $i++
    }
    return $sb.ToString().Trim()
}

# Decode an Enter Data blob (base64 + deflate + JSON rows) and describe it
# per column on one line instead of dumping every row.
function Get-DataSummary([string]$b64, [string]$after) {
    $bytes = [Convert]::FromBase64String($b64)
    $ms = New-Object IO.MemoryStream (,$bytes)
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $sr = New-Object IO.StreamReader($ds, [Text.Encoding]::UTF8)
    $json = $sr.ReadToEnd()
    $rows = @($ser.DeserializeObject($json))

    $names = @()
    if ($after -match '(?s)type\s+table\s*\[(.*?)\]') {
        foreach ($p in $Matches[1].Split(',')) {
            $names += (($p -split '=')[0] -replace '[#"\\]', '').Trim()
        }
    }

    $ncol = 0
    foreach ($r in $rows) { $n = @($r).Count; if ($n -gt $ncol) { $ncol = $n } }
    if ($rows.Count -eq 0 -or $ncol -eq 0) { return '0 rows' }

    $parts = New-Object System.Collections.ArrayList
    for ($c = 0; $c -lt $ncol; $c++) {
        $name = if ($c -lt $names.Count -and $names[$c]) { $names[$c] } else { "Column$($c + 1)" }
        $counts = @{}
        $order  = New-Object System.Collections.ArrayList
        foreach ($r in $rows) {
            $ra = @($r)
            $v = if ($c -lt $ra.Count -and $null -ne $ra[$c]) { [string]$ra[$c] } else { '' }
            if ($counts.ContainsKey($v)) { $counts[$v]++ } else { $counts[$v] = 1; [void]$order.Add($v) }
        }
        $d = $counts.Count
        if ($d -le $MaxDistinct) {
            $vals = foreach ($k in ($order | Sort-Object)) {
                $label = if ($k -eq '') { '(blank)' } else { $k }
                "$label ($($counts[$k]))"
            }
            [void]$parts.Add("${name}: " + ($vals -join ', '))
        } else {
            $nums = New-Object System.Collections.ArrayList
            $allNum = $true
            foreach ($k in $order) {
                if ($k -eq '') { continue }
                $x = 0.0
                if ([double]::TryParse($k, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$x)) { [void]$nums.Add($x) } else { $allNum = $false; break }
            }
            $info = "$d distinct"
            if ($allNum -and $nums.Count -gt 0) {
                $mm = $nums | Measure-Object -Minimum -Maximum
                $info += ", $($mm.Minimum)..$($mm.Maximum)"
            }
            $sample = ($order | Where-Object { $_ -ne '' } | Select-Object -First 3) -join ', '
            [void]$parts.Add("${name}: $info, e.g. $sample")
        }
    }
    return "$($rows.Count) rows (value (count) per column); " + ($parts -join '; ')
}

# One-line M query, with embedded data blobs replaced by a summary
function Format-M([string]$expr) {
    $sums = @()
    $ms = $blobRx.Matches($expr)
    for ($k = $ms.Count - 1; $k -ge 0; $k--) {
        $m = $ms[$k]
        $g = $m.Groups[1]
        $start = $m.Index + $m.Length
        $after = $expr.Substring($start, [Math]::Min(3000, $expr.Length - $start))
        try { $s = Get-DataSummary $g.Value $after }
        catch { $s = "could not decode ($($_.Exception.Message))" }
        $sums = @($s) + $sums
        $expr = $expr.Substring(0, $g.Index) + '...' + $expr.Substring($g.Index + $g.Length)
    }
    $r = Compress-Expr $expr $false
    foreach ($s in $sums) { $r = $r + ' <embedded data: ' + $s + '>' }
    return $r
}

# ------------------------------------------------------- semantic model ----

function New-Node($kind, $name, $expr) {
    return @{ kind = [string]$kind; name = [string]$name; expr = $expr; props = @{}; kids = (New-Object System.Collections.ArrayList); desc = ''; depth = 0 }
}

# Parse a TMDL file into a tree of nodes, appending top-level objects to $kids
function Read-Tmdl([string]$text, $kids) {
    $root = New-Node 'root' '' $null
    $root.depth = -1
    $root.kids = $kids
    $stack = New-Object System.Collections.ArrayList
    [void]$stack.Add($root)
    $lines = $text -split "`r?`n"
    $desc = ''
    $cont = $null
    $i = 0
    while ($i -lt $lines.Length) {
        $raw = $lines[$i]
        $i++
        $t = $raw.Trim()
        if ($t.Length -eq 0) { continue }
        $ws = $raw.Substring(0, $raw.Length - $raw.TrimStart().Length)
        $depth = $ws.Replace('    ', "`t").Replace(' ', '').Length

        # Body of a multi-line expression: indented two levels past its owner
        if (($null -ne $cont) -and ($depth -ge ($cont.depth + 2))) {
            if ($cont.key) { $cont.node.props[$cont.key] = [string]$cont.node.props[$cont.key] + "`n" + $t }
            else { $cont.node.expr = [string]$cont.node.expr + "`n" + $t }
            continue
        }
        $cont = $null

        if ($t.StartsWith('///')) { $desc = ($desc + ' ' + $t.Substring(3).Trim()).Trim(); continue }

        while (($stack.Count -gt 1) -and ($stack[$stack.Count - 1].depth -ge $depth)) { $stack.RemoveAt($stack.Count - 1) }
        $top = $stack[$stack.Count - 1]

        $m = $propRx.Match($t)
        if ($m.Success) { $top.props[$m.Groups[1].Value] = $m.Groups[2].Value; continue }

        $key = $null
        $node = $null
        $rest = $null
        $m = $exprRx.Match($t)
        if ($m.Success) {
            $key = $m.Groups[1].Value
            $node = $top
            $rest = $m.Groups[2].Value
        }
        else {
            $m = $declRx.Match($t)
            if ($m.Success) {
                $node = New-Node ($m.Groups[1].Value) (Unq ($m.Groups[2].Value)) $null
                if ($m.Groups[3].Success) { $rest = $m.Groups[3].Value }
            }
            elseif ($t -match '^[A-Za-z_]\w*$') { $top.props[$t] = 'true'; continue }
            else {
                $sp = $t.IndexOf(' ')
                if ($sp -lt 0) { continue }
                $node = New-Node ($t.Substring(0, $sp)) ($t.Substring($sp + 1)) $null
            }
            $node.depth = $depth
            $node.desc = $desc
            [void]$top.kids.Add($node)
            [void]$stack.Add($node)
        }
        $desc = ''

        if ($null -ne $rest) {
            $rest = $rest.Trim()
            if ($rest.StartsWith('```')) {
                $body = $rest.Substring(3)
                $e = $body.IndexOf('```')
                if ($e -ge 0) { $body = $body.Substring(0, $e) }
                else {
                    while ($i -lt $lines.Length) {
                        $l = $lines[$i]
                        $i++
                        $e = $l.IndexOf('```')
                        if ($e -ge 0) { $body = $body + "`n" + $l.Substring(0, $e); break }
                        $body = $body + "`n" + $l
                    }
                }
                $rest = $body
            }
            else { $cont = @{ node = $node; key = $key; depth = $depth } }
            if ($key) { $node.props[$key] = $rest } else { $node.expr = $rest }
        }
    }
}

# Convert a model.bim (TMSL JSON) into the same node tree
function Read-Bim($j, $kids) {
    $model = G $j 'model'
    foreach ($t in @(G $model 'tables')) {
        $tn = New-Node 'table' (G $t 'name') $null
        if (G $t 'isHidden') { $tn.props['isHidden'] = 'true' }
        $tn.desc = [string](JText (G $t 'description'))
        if ($null -ne (G $t 'calculationGroup')) {
            $tn.props['calculationGroup'] = 'true'
            foreach ($ci in @(G $t 'calculationGroup' 'calculationItems')) {
                [void]$tn.kids.Add((New-Node 'calculationItem' (G $ci 'name') (JText (G $ci 'expression'))))
            }
        }
        foreach ($c in @(G $t 'columns')) {
            $ex = $null
            if ((G $c 'type') -eq 'calculated') { $ex = [string](JText (G $c 'expression')) }
            $cn = New-Node 'column' (G $c 'name') $ex
            if (G $c 'isHidden') { $cn.props['isHidden'] = 'true' }
            if (G $c 'dataType') { $cn.props['dataType'] = [string](G $c 'dataType') }
            if (G $c 'sortByColumn') { $cn.props['sortByColumn'] = [string](G $c 'sortByColumn') }
            $cn.desc = [string](JText (G $c 'description'))
            [void]$tn.kids.Add($cn)
        }
        foreach ($ms in @(G $t 'measures')) {
            $mn = New-Node 'measure' (G $ms 'name') ([string](JText (G $ms 'expression')))
            if (G $ms 'isHidden') { $mn.props['isHidden'] = 'true' }
            $mn.desc = [string](JText (G $ms 'description'))
            [void]$tn.kids.Add($mn)
        }
        foreach ($h in @(G $t 'hierarchies')) {
            $hn = New-Node 'hierarchy' (G $h 'name') $null
            foreach ($lv in @(G $h 'levels')) { [void]$hn.kids.Add((New-Node 'level' (G $lv 'name') $null)) }
            [void]$tn.kids.Add($hn)
        }
        foreach ($p in @(G $t 'partitions')) {
            $pn = New-Node 'partition' (G $p 'name') ([string](G $p 'source' 'type'))
            if (G $p 'mode') { $pn.props['mode'] = [string](G $p 'mode') }
            $src = JText (G $p 'source' 'expression')
            if ($null -eq $src) { $src = JText (G $p 'source' 'query') }
            if ($null -ne $src) { $pn.props['source'] = $src }
            if (G $p 'source' 'entityName') { $pn.props['entityName'] = [string](G $p 'source' 'entityName') }
            [void]$tn.kids.Add($pn)
        }
        [void]$kids.Add($tn)
    }
    foreach ($r in @(G $model 'relationships')) {
        $rn = New-Node 'relationship' (G $r 'name') $null
        foreach ($k in 'fromTable','fromColumn','toTable','toColumn','crossFilteringBehavior','isActive','fromCardinality','toCardinality') {
            $v = G $r $k
            if ($null -ne $v) { $rn.props[$k] = [string]$v }
        }
        [void]$kids.Add($rn)
    }
    foreach ($e in @(G $model 'expressions')) {
        [void]$kids.Add((New-Node 'expression' (G $e 'name') ([string](JText (G $e 'expression')))))
    }
    foreach ($ro in @(G $model 'roles')) {
        $ron = New-Node 'role' (G $ro 'name') $null
        foreach ($tp in @(G $ro 'tablePermissions')) {
            [void]$ron.kids.Add((New-Node 'tablePermission' (G $tp 'name') ([string](JText (G $tp 'filterExpression')))))
        }
        [void]$kids.Add($ron)
    }
}

function Test-Hidden($n) {
    $v = $n.props['isHidden']
    return ($v -and ($v -ne 'false'))
}

# " /* hidden; description */" or nothing
function Get-Note($n) {
    $a = @()
    if (Test-Hidden $n) { $a += 'hidden' }
    if ($n.desc) { $a += ([string]$n.desc).Replace('*/', '* /') }
    if ($a.Count -gt 0) { return ' /* ' + ($a -join '; ') + ' */' }
    return ''
}

function Get-RelEnd($r, [string]$side) {
    $tbl = $r.props[$side + 'Table']
    $col = [string]$r.props[$side + 'Column']
    if ($tbl) { return (QT ([string]$tbl)) + '[' + $col + ']' }
    $m = $relRx.Match($col)
    if ($m.Success) { return (QT (Unq ($m.Groups[1].Value))) + '[' + (Unq ($m.Groups[2].Value)) + ']' }
    return $col
}

function Write-Model([string]$title, $kids) {
    if ($title -eq '') { $title = '(project root)' }
    Add-Line "=== SEMANTIC MODEL: $title ==="
    $auto = 0

    foreach ($t in $kids) {
        if ($t.kind -ne 'table') { continue }
        if ($t.name -match $autoRx) { $auto++; continue }
        $script:nTables++

        $hdr = 'TABLE ' + $t.name
        $tags = @()
        if (Test-Hidden $t) { $tags += 'hidden' }
        if ($t.props.ContainsKey('calculationGroup')) { $tags += 'calculation group' }
        if ($tags.Count -gt 0) { $hdr += ' (' + ($tags -join ', ') + ')' }
        if ($t.desc) { $hdr += ' - ' + $t.desc }
        Add-Line $hdr

        $cols = @()
        foreach ($c in $t.kids) {
            if (($c.kind -ne 'column') -or ($null -ne $c.expr)) { continue }
            $s = $c.name
            if ($c.props['dataType']) { $s += ' ' + $c.props['dataType'] }
            $x = @()
            if (Test-Hidden $c) { $x += 'hidden' }
            if ($c.props['sortByColumn']) { $x += 'sort by ' + $c.props['sortByColumn'] }
            if ($c.desc) { $x += $c.desc }
            if ($x.Count -gt 0) { $s += ' (' + ($x -join '; ') + ')' }
            $cols += $s
        }
        if ($cols.Count -gt 0) { Add-Line ('columns: ' + ($cols -join ', ')) }

        foreach ($c in $t.kids) {
            if (($c.kind -ne 'column') -or ($null -eq $c.expr)) { continue }
            Add-Packed ((QT $t.name) + '[' + $c.name + ']' + (Get-Note $c) + ' = ' + (Compress-Expr ([string]$c.expr) $true))
        }
        foreach ($c in $t.kids) {
            if ($c.kind -ne 'measure') { continue }
            $script:nMeasures++
            Add-Packed ('[' + $c.name + ']' + (Get-Note $c) + ' = ' + (Compress-Expr ([string]$c.expr) $true))
        }
        foreach ($c in $t.kids) {
            if ($c.kind -ne 'calculationItem') { continue }
            Add-Packed ('calc item "' + $c.name + '" = ' + (Compress-Expr ([string]$c.expr) $true))
        }
        foreach ($c in $t.kids) {
            if ($c.kind -ne 'hierarchy') { continue }
            $lv = @(foreach ($l in $c.kids) { if ($l.kind -eq 'level') { $l.name } })
            Add-Packed ('hierarchy ' + $c.name + ': ' + ($lv -join ' > '))
        }
        Flush-Pack

        foreach ($p in $t.kids) {
            if ($p.kind -ne 'partition') { continue }
            $kind = ([string]$p.expr).Trim()
            $src  = [string]$p.props['source']
            $mode = [string]$p.props['mode']
            $mt = ''
            if ($mode -and ($mode -ne 'import')) { $mt = " ($mode)" }
            if ($kind -eq 'm') { Add-Line ("source M${mt}: " + (Format-M $src)) }
            elseif ($kind -eq 'calculated') { Add-Line ('calculated table = ' + (Compress-Expr $src $true)) }
            elseif ($kind -eq 'entity') { Add-Line ("source entity${mt}: " + [string]$p.props['entityName']) }
            elseif ($src) { Add-Line ("source ${kind}${mt}: " + (Compress-Expr $src $false)) }
        }
    }

    $rels = @()
    foreach ($r in $kids) {
        if ($r.kind -ne 'relationship') { continue }
        $f  = Get-RelEnd $r 'from'
        $to = Get-RelEnd $r 'to'
        if (($f + $to) -match $autoRx) { continue }
        $fc = if ($r.props['fromCardinality'] -eq 'one') { '1' } else { '*' }
        $tc = if ($r.props['toCardinality'] -eq 'many') { '*' } else { '1' }
        $s = "$f ${fc}:$tc $to"
        $x = @()
        if ($r.props['crossFilteringBehavior'] -eq 'bothDirections') { $x += 'both directions' }
        if ($r.props['isActive'] -eq 'false') { $x += 'inactive' }
        if ($x.Count -gt 0) { $s += ' (' + ($x -join ', ') + ')' }
        $rels += $s
    }
    if ($rels.Count -gt 0) {
        Add-Line 'RELATIONSHIPS:'
        foreach ($s in $rels) { Add-Packed $s }
        Flush-Pack
    }

    foreach ($e in $kids) {
        if ($e.kind -eq 'expression') { Add-Line ('shared M ' + $e.name + ' = ' + (Format-M ([string]$e.expr))) }
        elseif ($e.kind -eq 'function') { Add-Line ('DAX function ' + $e.name + ' = ' + (Compress-Expr ([string]$e.expr) $true)) }
    }
    foreach ($ro in $kids) {
        if ($ro.kind -ne 'role') { continue }
        $tp = @(foreach ($p in $ro.kids) { if ($p.kind -eq 'tablePermission') { (QT $p.name) + ': ' + (Compress-Expr ([string]$p.expr) $true) } })
        Add-Line ('ROLE ' + $ro.name + ' (row-level security): ' + ($tp -join '; '))
    }
    if ($auto -gt 0) { Add-Line "($auto auto date/time tables omitted)" }
}

# --------------------------------------------------------------- report ----

function Get-Entity($node, $alias) {
    $ent = G $node 'Expression' 'SourceRef' 'Entity'
    if ($null -eq $ent) {
        $src = G $node 'Expression' 'SourceRef' 'Source'
        if (($null -ne $src) -and $alias -and $alias.ContainsKey([string]$src)) { $ent = $alias[[string]$src] } else { $ent = $src }
    }
    return (QT ([string]$ent))
}

# Field reference / literal from a report query expression
function Format-Field($e, $alias) {
    if ($null -eq $e) { return '?' }
    if (-not ($e -is [System.Collections.IDictionary])) { return [string]$e }
    foreach ($k in 'Column','Measure') {
        $n = G $e $k
        if ($null -ne $n) { return (Get-Entity $n $alias) + '[' + [string](G $n 'Property') + ']' }
    }
    $a = G $e 'Aggregation'
    if ($null -ne $a) {
        $fi = [int](G $a 'Function')
        $fn = "Agg$fi"
        if (($fi -ge 0) -and ($fi -lt $aggNames.Count)) { $fn = $aggNames[$fi] }
        return $fn + '(' + (Format-Field (G $a 'Expression') $alias) + ')'
    }
    $h = G $e 'HierarchyLevel'
    if ($null -ne $h) {
        $hh = G $h 'Expression' 'Hierarchy'
        $pv = G $hh 'Expression' 'PropertyVariationSource'
        if ($null -ne $pv) { return (Get-Entity $pv $alias) + '[' + [string](G $pv 'Property') + '.' + [string](G $h 'Level') + ']' }
        return (Get-Entity $hh $alias) + '[' + [string](G $hh 'Hierarchy') + '.' + [string](G $h 'Level') + ']'
    }
    $lit = G $e 'Literal' 'Value'
    if ($null -ne $lit) { return [string]$lit }
    return (Cut ($ser.Serialize($e)) 200)
}

# Filter condition from a report filter's Where clause
function Format-Cond($c, $alias) {
    if (-not ($c -is [System.Collections.IDictionary])) { return [string]$c }
    $n = G $c 'In'
    if ($null -ne $n) {
        $ex = @(foreach ($x in @(G $n 'Expressions')) { Format-Field $x $alias })
        $vals = @(foreach ($row in @(G $n 'Values')) { (@(foreach ($v in @($row)) { Format-Field $v $alias }) -join '+') })
        if ($vals.Count -gt 20) { $vals = @($vals[0..19]) + "... ($($vals.Count) values)" }
        return ($ex -join '+') + ' in (' + ($vals -join ', ') + ')'
    }
    $n = G $c 'Not'
    if ($null -ne $n) { return 'not (' + (Format-Cond (G $n 'Expression') $alias) + ')' }
    foreach ($k in 'And','Or') {
        $n = G $c $k
        if ($null -ne $n) { return '(' + (Format-Cond (G $n 'Left') $alias) + ' ' + $k.ToLower() + ' ' + (Format-Cond (G $n 'Right') $alias) + ')' }
    }
    $n = G $c 'Comparison'
    if ($null -ne $n) {
        $ki = [int](G $n 'ComparisonKind')
        $op = '?'
        if (($ki -ge 0) -and ($ki -lt $cmpOps.Count)) { $op = $cmpOps[$ki] }
        return (Format-Cond (G $n 'Left') $alias) + " $op " + (Format-Cond (G $n 'Right') $alias)
    }
    $n = G $c 'Between'
    if ($null -ne $n) { return (Format-Cond (G $n 'Expression') $alias) + ' between ' + (Format-Cond (G $n 'LowerBound') $alias) + ' and ' + (Format-Cond (G $n 'UpperBound') $alias) }
    foreach ($k in 'Contains','StartsWith','EndsWith') {
        $n = G $c $k
        if ($null -ne $n) { return (Format-Cond (G $n 'Left') $alias) + ' ' + $k.ToLower() + ' ' + (Format-Cond (G $n 'Right') $alias) }
    }
    return (Format-Field $c $alias)
}

# Filters with a condition; with $keepEmpty also filter-pane fields left on "all"
function Format-Filters($filters, [bool]$keepEmpty) {
    $res = @()
    foreach ($f in @($filters)) {
        if ($null -eq $f) { continue }
        $alias = @{}
        foreach ($fr in @(G $f 'filter' 'From')) { $alias[[string](G $fr 'Name')] = [string](G $fr 'Entity') }
        $conds = @(foreach ($w in @(G $f 'filter' 'Where')) { Cut (Format-Cond (G $w 'Condition') $alias) 300 })
        $type = [string](G $f 'type')
        if ($conds.Count -gt 0) {
            $s = ($conds -join ' and ')
            if ($type -and ($type -ne 'Categorical') -and ($type -ne 'Advanced')) { $s += " [$type]" }
            $res += $s
        }
        elseif ($keepEmpty) {
            $fld = G $f 'field'
            if ($null -eq $fld) { $fld = G $f 'expression' }
            if ($null -ne $fld) { $res += (Format-Field $fld $null) + ' (all)' }
        }
    }
    return $res
}

# Collect every model field referenced anywhere inside a formatting object
function Get-Refs($o, $acc) {
    if ($o -is [System.Collections.IDictionary]) {
        if ($o.ContainsKey('Measure') -or $o.ContainsKey('Column') -or $o.ContainsKey('Aggregation')) {
            $s = Format-Field $o $null
            if (-not $acc.Contains($s)) { [void]$acc.Add($s) }
            return
        }
        foreach ($k in @($o.Keys)) { Get-Refs $o[$k] $acc }
    }
    elseif (($o -is [System.Collections.IList]) -and (-not ($o -is [string]))) {
        foreach ($x in $o) { Get-Refs $x $acc }
    }
}

function Format-VisualLine([string]$type, $title, $parts, [bool]$hidden) {
    $s = $type
    if ($title) { $s += ' "' + ([string]$title).Trim("'") + '"' }
    if ($hidden) { $s += ' (hidden)' }
    return $s + ': ' + ($parts -join '; ')
}

# PBIR visual.json -> one line, or nothing when the visual binds no data
function Format-Visual($v) {
    $vis = G $v 'visual'
    if ($null -eq $vis) { return }
    $parts = @()
    $qs = G $vis 'query' 'queryState'
    if ($qs -is [System.Collections.IDictionary]) {
        foreach ($role in @($qs.Keys)) {
            $fields = @(foreach ($p in @(G $qs $role 'projections')) {
                $n = Format-Field (G $p 'field') $null
                $dn = G $p 'displayName'
                if ($dn) { $n += ' as "' + [string]$dn + '"' }
                $n
            })
            if ($fields.Count -gt 0) { $parts += "${role}: " + ($fields -join ', ') }
        }
    }
    $sorts = @(foreach ($so in @(G $vis 'query' 'sortDefinition' 'sort')) {
        $d = [string](G $so 'direction')
        $d = if ($d -match '^Desc') { 'desc' } else { 'asc' }
        (Format-Field (G $so 'field') $null) + ' ' + $d
    })
    if ($sorts.Count -gt 0) { $parts += 'sort: ' + ($sorts -join ', ') }
    $vf = @(Format-Filters (G $v 'filterConfig' 'filters') $false)
    if ($vf.Count -gt 0) { $parts += 'filter: ' + ($vf -join ' and ') }
    $acc = New-Object System.Collections.ArrayList
    Get-Refs (G $vis 'objects') $acc
    Get-Refs (G $vis 'visualContainerObjects') $acc
    if ($acc.Count -gt 0) { $parts += 'formatting driven by: ' + ($acc -join ', ') }
    if ($parts.Count -eq 0) { return }
    $title = G $vis 'visualContainerObjects' 'title' 0 'properties' 'text' 'expr' 'Literal' 'Value'
    return (Format-VisualLine ([string](G $vis 'visualType')) $title $parts ([bool](G $v 'isHidden')))
}

function Write-ExtMeasures($entities) {
    $any = $false
    foreach ($e in @($entities)) {
        foreach ($m in @(G $e 'measures')) {
            if (-not $any) { Add-Line 'report-level measures:'; $any = $true }
            $script:nMeasures++
            Add-Packed ((QT ([string](G $e 'name'))) + '[' + [string](G $m 'name') + '] = ' + (Compress-Expr ([string](JText (G $m 'expression'))) $true))
        }
    }
    Flush-Pack
}

function Write-Deco($deco) {
    if ($deco.Count -eq 0) { return }
    $d = @(foreach ($k in ($deco.Keys | Sort-Object)) { "$($deco[$k]) $k" })
    Add-Line ('(omitted, no data bound: ' + ($d -join ', ') + ')')
}

# Enhanced report format (PBIR): one folder per page, one per visual
function Write-PbirReport($byRel) {
    $rep = Try-Json $byRel['definition\report.json']
    $ext = Try-Json $byRel['definition\reportextensions.json']
    Write-ExtMeasures (G $ext 'entities')
    $rf = @(Format-Filters (G $rep 'filterConfig' 'filters') $true)
    if ($rf.Count -gt 0) { Add-Line ('report filters: ' + ($rf -join '; ')) }

    $found = @($byRel.Keys | ForEach-Object { if ($_ -match '^definition\\pages\\([^\\]+)\\page\.json$') { $Matches[1] } } | Sort-Object)
    $po = Try-Json $byRel['definition\pages\pages.json']
    $order = @(foreach ($x in @(G $po 'pageOrder')) { ([string]$x).ToLower() })
    $ids = @($order | Where-Object { $found -contains $_ }) + @($found | Where-Object { $order -notcontains $_ })

    foreach ($id in $ids) {
        $pg = Try-Json $byRel["definition\pages\$id\page.json"]
        $hdr = 'PAGE "' + [string](G $pg 'displayName') + '"'
        $tags = @()
        if ((G $pg 'visibility') -eq 'HiddenInViewMode') { $tags += 'hidden' }
        $bt = [string](G $pg 'pageBinding' 'type')
        if ($bt -and ($bt -ne 'Default')) { $tags += $bt.ToLower() }
        if ($tags.Count -gt 0) { $hdr += ' (' + ($tags -join ', ') + ')' }
        $pf = @(Format-Filters (G $pg 'filterConfig' 'filters') $true)
        if ($pf.Count -gt 0) { $hdr += '  |  page filters: ' + ($pf -join '; ') }
        Add-Line $hdr

        $deco = @{}
        $pre = "definition\pages\$id\visuals\"
        foreach ($k in @($byRel.Keys | Where-Object { $_.StartsWith($pre) -and $_.EndsWith('\visual.json') } | Sort-Object)) {
            $v = Try-Json $byRel[$k]
            $s = Format-Visual $v
            if ($s) { $script:nVisuals++; Add-Packed $s }
            else {
                $vt = [string](G $v 'visual' 'visualType')
                if (-not $vt) { $vt = 'group' }
                $deco[$vt] = 1 + [int]$deco[$vt]
            }
        }
        Flush-Pack
        Write-Deco $deco
    }
}

# Legacy single-file report.json
function Write-LegacyReport($r) {
    $cfg = From-JsonText (G $r 'config')
    foreach ($me in @(G $cfg 'modelExtensions')) { Write-ExtMeasures (G $me 'entities') }
    $rf = @(Format-Filters (From-JsonText (G $r 'filters')) $true)
    if ($rf.Count -gt 0) { Add-Line ('report filters: ' + ($rf -join '; ')) }

    foreach ($sec in @(G $r 'sections')) {
        $hdr = 'PAGE "' + [string](G $sec 'displayName') + '"'
        $pf = @(Format-Filters (From-JsonText (G $sec 'filters')) $true)
        if ($pf.Count -gt 0) { $hdr += '  |  page filters: ' + ($pf -join '; ') }
        Add-Line $hdr

        $deco = @{}
        foreach ($vc in @(G $sec 'visualContainers')) {
            $c = From-JsonText (G $vc 'config')
            $sv = G $c 'singleVisual'
            $type = [string](G $sv 'visualType')
            $parts = @()
            $proj = G $sv 'projections'
            if ($proj -is [System.Collections.IDictionary]) {
                foreach ($role in @($proj.Keys)) {
                    $q = @(foreach ($p in @(G $proj $role)) { G $p 'queryRef' })
                    if ($q.Count -gt 0) { $parts += "${role}: " + ($q -join ', ') }
                }
            }
            $vf = @(Format-Filters (From-JsonText (G $vc 'filters')) $false)
            if ($vf.Count -gt 0) { $parts += 'filter: ' + ($vf -join ' and ') }
            $acc = New-Object System.Collections.ArrayList
            Get-Refs (G $sv 'objects') $acc
            Get-Refs (G $sv 'vcObjects') $acc
            if ($acc.Count -gt 0) { $parts += 'formatting driven by: ' + ($acc -join ', ') }
            if ($parts.Count -gt 0) {
                $script:nVisuals++
                $title = G $sv 'vcObjects' 'title' 0 'properties' 'text' 'expr' 'Literal' 'Value'
                Add-Packed (Format-VisualLine $type $title $parts $false)
            }
            else {
                if (-not $type) { $type = 'group' }
                $deco[$type] = 1 + [int]$deco[$type]
            }
        }
        Flush-Pack
        Write-Deco $deco
    }
}

function Write-Report([string]$own, $files) {
    Add-Line "=== REPORT: $own ==="
    $pre = $own + '\'
    $byRel = @{}
    foreach ($f in $files) { $byRel[$f.Rel.Substring($pre.Length).ToLower()] = $f }
    $isPbir = @($byRel.Keys | Where-Object { $_.StartsWith('definition\') }).Count -gt 0
    if ($isPbir) { Write-PbirReport $byRel }
    elseif ($byRel.ContainsKey('report.json')) { Write-LegacyReport (Try-Json $byRel['report.json']) }
    else { Add-Line '(no report definition found)' }
}

# ----------------------------------------------------------------- main ----

Write-Host "Scanning $Root ..."
$all = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force | ForEach-Object {
    [pscustomobject]@{ Path = $_.FullName; Name = $_.Name; Rel = $_.FullName.Substring($Root.Length + 1) }
} | Where-Object { ($_.Name -ne $outName) -and ($_.Name -ne $Self) -and (-not (Test-InSkippedDir $_.Rel)) } | Sort-Object Rel)

Add-Line ('POWER BI PROJECT, LOGIC ONLY (layout and formatting omitted: see the report screenshots). How to read: "[Name] = ..." is a DAX measure; "Table[Name] = ..." a calculated column; "source M" a Power Query; "  |  " separates items sharing a line; relationships read from-side cardinality:to-side (* many, 1 one); visuals read type "title": role: fields.')

$models = New-Object System.Collections.Specialized.OrderedDictionary
foreach ($f in $all) {
    $isT = $f.Name.ToLower().EndsWith('.tmdl')
    $isB = ($f.Name -eq 'model.bim')
    if (-not ($isT -or $isB)) { continue }
    if ($f.Rel -match '\\cultures\\') { continue }
    $own = Get-Owner $f.Rel 'SemanticModel|Dataset'
    if (-not $models.Contains($own)) { $models.Add($own, (New-Object System.Collections.ArrayList)) }
    $kids = $models[$own]
    try {
        if ($isT) { Read-Tmdl ([IO.File]::ReadAllText($f.Path)) $kids }
        else { Read-Bim (Read-Json $f.Path) $kids }
    }
    catch { Add-Line "(could not parse $($f.Rel): $($_.Exception.Message))" }
}
foreach ($own in @($models.Keys)) { Write-Model ([string]$own) $models[$own] }

$reports = New-Object System.Collections.Specialized.OrderedDictionary
foreach ($f in $all) {
    if (-not $f.Name.ToLower().EndsWith('.json')) { continue }
    $own = Get-Owner $f.Rel 'Report'
    if ($own -eq '') { continue }
    if (-not $reports.Contains($own)) { $reports.Add($own, (New-Object System.Collections.ArrayList)) }
    [void]$reports[$own].Add($f)
}
foreach ($own in @($reports.Keys)) { Write-Report ([string]$own) $reports[$own] }

if (($models.Count -eq 0) -and ($reports.Count -eq 0)) {
    Add-Line '(no .SemanticModel or .Report folder found next to this script)'
}
Flush-Pack

$text = ($out -join "`r`n") + "`r`n"
[IO.File]::WriteAllText($outPath, $text, (New-Object Text.UTF8Encoding($false)))
Write-Host "Done. $($script:nTables) tables, $($script:nMeasures) measures, $($script:nVisuals) data visuals."
Write-Host "Output: $outPath ($($out.Count) lines, $($text.Length) characters)"
Write-Host "Tip: open it with Word Wrap on, a small font and a maximised window, then screenshot."

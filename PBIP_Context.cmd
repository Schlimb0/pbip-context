@echo off
rem ==========================================================================
rem  PBIP_Context.cmd
rem  Drop this file into a Power BI Project (.pbip) folder and run it.
rem  It writes PBIP_Context.txt: a file list plus the full text of every
rem  readable project file (PBIR report, TMDL/BIM semantic model, etc.) so an
rem  AI assistant can read the whole project as one document.
rem
rem  Embedded data ("Enter Data" tables) is NOT dumped row by row. Instead:
rem    - columns with few distinct values (<= MAXD) are listed in full with
rem      counts (e.g. Age: 18 (40), 19 (35), 20 (25))
rem    - columns with many distinct values (names, heights...) are summarised
rem      (distinct count, numeric min/max, a few sample values)
rem ==========================================================================
setlocal

rem Max distinct values for a column to be listed in full. Edit to taste.
set "MAXD=15"

powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $s=[IO.File]::ReadAllText('%~f0'); $i=$s.IndexOf('#PS'+'START'); & ([scriptblock]::Create($s.Substring($i))) -Root '%~dp0' -Self '%~nx0' -MaxDistinct %MAXD% }"

echo.
if /i not "%~1"=="/nopause" pause
endlocal
exit /b 0

#PSSTART
param([string]$Root, [string]$Self, [int]$MaxDistinct = 15)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web.Extensions

$Root     = $Root.TrimEnd('\')
$outName  = 'PBIP_Context.txt'
$outPath  = Join-Path $Root $outName
$exts     = '.pbip','.pbir','.pbism','.pbidataset','.pbivis','.json','.tmdl','.bim','.platform','.dax','.m','.pq','.md','.txt','.xml','.yaml','.yml'
$skipDirs = '.pbi','.git','node_modules'

function Test-InSkippedDir([string]$rel) {
    $segs = $rel.Split('\')
    for ($i = 0; $i -lt $segs.Length - 1; $i++) {
        if ($skipDirs -contains $segs[$i].ToLower()) { return $true }
    }
    return $false
}

# Decode an Enter Data blob (base64 + deflate + JSON rows) and describe it
function Get-DataSummary([string]$b64, [string]$after) {
    $bytes = [Convert]::FromBase64String($b64)
    $ms = New-Object IO.MemoryStream (,$bytes)
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $sr = New-Object IO.StreamReader($ds, [Text.Encoding]::UTF8)
    $json = $sr.ReadToEnd()
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = [int]::MaxValue
    $rows = @($ser.DeserializeObject($json))

    $names = @()
    if ($after -match '(?s)type\s+table\s*\[(.*?)\]') {
        foreach ($p in $Matches[1].Split(',')) {
            $names += (($p -split '=')[0] -replace '[#"\\]', '').Trim()
        }
    }

    $ncol = 0
    foreach ($r in $rows) { $n = @($r).Count; if ($n -gt $ncol) { $ncol = $n } }
    if ($rows.Count -eq 0 -or $ncol -eq 0) { return "EMBEDDED TABLE DATA: 0 rows." }

    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add("EMBEDDED TABLE DATA SUMMARY: $($rows.Count) rows, $ncol columns. Columns with <= $MaxDistinct distinct values are listed in full with counts; others are summarised (full rows omitted).")

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
            $parts = foreach ($k in ($order | Sort-Object)) {
                $label = if ($k -eq '') { '(blank)' } else { $k }
                "$label ($($counts[$k]))"
            }
            [void]$lines.Add("  - ${name}: $d distinct values: " + ($parts -join ', '))
        } else {
            $nums = New-Object System.Collections.ArrayList
            $allNum = $true
            foreach ($k in $order) {
                if ($k -eq '') { continue }
                $x = 0.0
                if ([double]::TryParse($k, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$x)) { [void]$nums.Add($x) } else { $allNum = $false; break }
            }
            $info = "$d distinct values (too many to list)"
            if ($allNum -and $nums.Count -gt 0) {
                $m = $nums | Measure-Object -Minimum -Maximum
                $info += ", numeric, min $($m.Minimum), max $($m.Maximum)"
            }
            $sample = ($order | Where-Object { $_ -ne '' } | Select-Object -First 3) -join ', '
            [void]$lines.Add("  - ${name}: $info. Sample: $sample")
        }
    }
    return ($lines -join "`r`n")
}

Write-Host "Scanning $Root ..."
$all = Get-ChildItem -LiteralPath $Root -Recurse -File -Force | ForEach-Object {
    [pscustomobject]@{ File = $_; Rel = $_.FullName.Substring($Root.Length + 1) }
} | Where-Object { $_.File.Name -ne $outName -and $_.File.Name -ne $Self }

$include = @($all | Where-Object { -not (Test-InSkippedDir $_.Rel) -and ($exts -contains $_.File.Extension.ToLower()) })
$skipped = $all.Count - $include.Count

$sb = New-Object System.Text.StringBuilder
$nl = "`r`n"
[void]$sb.Append("################################################################$nl")
[void]$sb.Append("# POWER BI PROJECT CONTEXT FILE$nl")
[void]$sb.Append("# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')$nl")
[void]$sb.Append("# Source folder: $Root$nl#$nl")
[void]$sb.Append("# Section 1 lists the included files. Section 2 contains the full$nl")
[void]$sb.Append("# text of each, starting with a ""FILE:"" header. Embedded data tables$nl")
[void]$sb.Append("# are summarised per column instead of dumped row by row.$nl")
[void]$sb.Append("# Skipped: .pbi cache, .git, binary files, and this script.$nl")
[void]$sb.Append("################################################################$nl$nl")
[void]$sb.Append("======== SECTION 1: FILE LIST ========$nl")
foreach ($i in $include) { [void]$sb.Append($i.Rel + $nl) }
[void]$sb.Append("$nl======== SECTION 2: FILE CONTENTS ========$nl")

$pat = '(Binary\.FromText\(\s*\\?")([A-Za-z0-9+/=]{20,})(\\?")'
$blobs = 0
foreach ($i in $include) {
    $text = [IO.File]::ReadAllText($i.File.FullName)
    $script:summaries = New-Object System.Collections.ArrayList
    if ($text -match 'Binary\.FromText') {
        $script:curText = $text
        $text = [regex]::Replace($text, $pat, [System.Text.RegularExpressions.MatchEvaluator]{
            param($m)
            $start = $m.Index + $m.Length
            $after = $script:curText.Substring($start, [Math]::Min(3000, $script:curText.Length - $start))
            try { $s = Get-DataSummary $m.Groups[2].Value $after }
            catch { $s = "EMBEDDED TABLE DATA: could not decode ($($_.Exception.Message))" }
            [void]$script:summaries.Add($s)
            $m.Groups[1].Value + '<embedded data omitted - see summary>' + $m.Groups[3].Value
        })
    }
    [void]$sb.Append("$nl================================================================$nl")
    [void]$sb.Append("FILE: $($i.Rel)$nl")
    [void]$sb.Append("================================================================$nl")
    [void]$sb.Append($text)
    foreach ($s in $script:summaries) { [void]$sb.Append("$nl$nl[$s$nl]$nl"); $blobs++ }
    [void]$sb.Append($nl)
}
[void]$sb.Append("$nl======== END OF PROJECT CONTEXT ($($include.Count) files included) ========$nl")

[IO.File]::WriteAllText($outPath, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
Write-Host "Done. Included $($include.Count) files, skipped $skipped, summarised $blobs embedded data table(s)."
Write-Host "Output: $outPath"

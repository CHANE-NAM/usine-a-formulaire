# Tools/snapshot_helpers.ps1
# PowerShell 5.1 compatible - ASCII only (no backticks in strings)

function Get-RelativePath {
  param([Parameter(Mandatory)][string]$Root,[Parameter(Mandatory)][string]$Path)
  $rootFull = (Resolve-Path -LiteralPath $Root).Path.TrimEnd('\','/')
  $pathFull = (Resolve-Path -LiteralPath $Path).Path
  if ($pathFull.StartsWith($rootFull,[System.StringComparison]::OrdinalIgnoreCase)) {
    return $pathFull.Substring($rootFull.Length).TrimStart('\','/')
  }
  return $pathFull
}

function Split-CsvHeader {
  param([string]$Line)
  if ($null -eq $Line) { return @() }
  $fields = New-Object System.Collections.Generic.List[string]
  $cur = New-Object System.Text.StringBuilder
  $inQuotes = $false
  for ($i=0; $i -lt $Line.Length; $i++) {
    $ch = $Line[$i]
    if ($ch -eq '"') {
      if ($inQuotes -and ($i+1 -lt $Line.Length) -and $Line[$i+1] -eq '"') {
        [void]$cur.Append('"'); $i++
      } else {
        $inQuotes = -not $inQuotes
      }
    } elseif ($ch -eq ',' -and -not $inQuotes) {
      $fields.Add($cur.ToString().Trim()); $cur.Clear() | Out-Null
    } else {
      [void]$cur.Append($ch)
    }
  }
  $fields.Add($cur.ToString().Trim())
  return $fields | ForEach-Object {
    if ($_ -match '^".*"$') { $_.Substring(1, $_.Length-2).Replace('""','"') } else { $_ }
  }
}

function Get-CsvInfo {
  param([Parameter(Mandatory)][string]$File,[Parameter(Mandatory)][string]$SnapshotDir)
  $rel = Get-RelativePath -Root $SnapshotDir -Path $File
  $bn  = [System.IO.Path]::GetFileNameWithoutExtension($File)

  $header = $null
  try { $header = Get-Content -LiteralPath $File -TotalCount 1 -ErrorAction Stop } catch { $header = $null }
  $cols = Split-CsvHeader -Line $header
  $lines = (Get-Content -LiteralPath $File -ErrorAction SilentlyContinue | Measure-Object -Line).Lines
  if ($lines -gt 0) { $lines-- } else { $lines = 0 }
  $hash = (Get-FileHash -LiteralPath $File -Algorithm SHA256).Hash

  [pscustomobject]@{
    file     = $rel
    name     = $bn
    rows     = $lines
    cols     = $cols.Count
    headers  = $cols
    sha256   = $hash
    bytes    = (Get-Item -LiteralPath $File).Length
  }
}

function Extract-FunctionsFromConcat {
  param([Parameter(Mandatory)][string]$File)
  $txt = Get-Content -LiteralPath $File -Raw -ErrorAction SilentlyContinue
  if ($null -eq $txt) { return @() }
  $names = New-Object System.Collections.Generic.List[string]
  $rx = @(
    '^[\t ]*(?:export\s+)?function\s+([A-Za-z_]\w*)\s*\(',
    '^[\t ]*(?:const|let|var)\s+([A-Za-z_]\w*)\s*=\s*\(',
    '^[\t ]*class\s+([A-Za-z_]\w*)\b',
    '^[\t ]*(?:const|let|var)\s+([A-Za-z_]\w*)\s*=\s*async\s*\(',
    '^[\t ]*([A-Za-z_]\w*)\s*:\s*function\s*\('
  ) -join '|'
  $matches = [regex]::Matches($txt, $rx, 'Multiline,IgnoreCase')
  foreach ($m in $matches) {
    foreach ($g in $m.Groups) {
      if ($g.Index -gt 0 -and $g.Value -and $g.Value -ne $m.Value) {
        if (-not $names.Contains($g.Value)) { [void]$names.Add($g.Value) }
      }
    }
  }
  return $names | Select-Object -Unique
}

function Write-Manifest {
  param(
    [Parameter(Mandatory)][string]$SnapshotDir,
    [string]$RepoRoot = $null
  )
  $snapName = Split-Path -Path $SnapshotDir -Leaf
  $nowIso = (Get-Date).ToString('s') + 'Z'

  $csvFiles = Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Filter *.csv -ErrorAction SilentlyContinue
  $csvInfos = @()
  foreach ($f in $csvFiles) { $csvInfos += Get-CsvInfo -File $f.FullName -SnapshotDir $SnapshotDir }

  $bySheet = @{}
  foreach ($c in $csvInfos) {
    $firstSeg = ($c.file -split '[\\/]')[0]
    if (-not $firstSeg) { $firstSeg = '.' }
    if (-not $bySheet.ContainsKey($firstSeg)) { $bySheet[$firstSeg] = New-Object System.Collections.Generic.List[object] }
    $bySheet[$firstSeg].Add($c)
  }

  $sheets = @()
  foreach ($k in $bySheet.Keys | Sort-Object) {
    $idHint = $null
    if ($k -match '([A-Za-z0-9_-]{6,})$') { $idHint = $Matches[1] }
    $tabs = $bySheet[$k] | Sort-Object file
    $totalRows = ($tabs | Measure-Object rows -Sum).Sum
    $sheets += [pscustomobject]@{
      folder     = $k
      id_hint    = $idHint
      tabs       = $tabs
      total_rows = $totalRows
      tab_count  = $tabs.Count
    }
  }

  $concatFiles = Get-ChildItem -LiteralPath $SnapshotDir -File -Filter 'scripts_*.txt' -ErrorAction SilentlyContinue
  $projects = @()
  foreach ($cf in $concatFiles) {
    $rel = Get-RelativePath -Root $SnapshotDir -Path $cf.FullName
    $lines = (Get-Content -LiteralPath $cf.FullName | Measure-Object -Line).Lines
    $hash  = (Get-FileHash -LiteralPath $cf.FullName -Algorithm SHA256).Hash
    $funcs = (Extract-FunctionsFromConcat -File $cf.FullName) | Select-Object -First 50
    $projects += [pscustomobject]@{
      file      = $rel
      lines     = $lines
      sha256    = $hash
      functions = $funcs
    }
  }

  $manifest = [pscustomobject]@{
    snapshot       = $snapName
    generated_at   = $nowIso
    repo_root      = $RepoRoot
    csv_total      = $csvInfos.Count
    projects_total = $projects.Count
    sheets         = $sheets
    gas_projects   = $projects
  }

  $outPath = Join-Path $SnapshotDir 'manifest.json'
  $json = $manifest | ConvertTo-Json -Depth 20
  $json | Set-Content -LiteralPath $outPath -Encoding UTF8
  Write-Host ("[MANIFEST] {0}" -f $outPath)
  return $manifest
}

function Write-BriefMd {
  param(
    [Parameter(Mandatory)][string]$SnapshotDir,
    [Parameter(Mandatory)][object]$Manifest
  )
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine(('# Snapshot: {0}' -f $Manifest.snapshot))
  [void]$sb.AppendLine('')
  [void]$sb.AppendLine(('- Genere le: {0}' -f $Manifest.generated_at))
  [void]$sb.AppendLine(('- CSV: {0} fichiers' -f $Manifest.csv_total))
  [void]$sb.AppendLine(('- Projets GAS concatenes: {0}' -f $Manifest.projects_total))
  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('## Google Sheets')

  foreach ($s in $Manifest.sheets) {
    $extra = if ($s.id_hint) { (' - id_hint: {0}' -f $s.id_hint) } else { '' }
    [void]$sb.AppendLine(('- {0} (onglets: {1}, lignes: {2}){3}' -f $s.folder, $s.tab_count, $s.total_rows, $extra))
    foreach ($t in $s.tabs | Select-Object -First 8) {
      $heads = ($t.headers | Select-Object -First 12) -join ', '
      [void]$sb.AppendLine(('  - {0}: {1} lignes, {2} colonnes - {3}' -f $t.name, $t.rows, $t.cols, $t.file))
      if ($heads) { [void]$sb.AppendLine(('    - colonnes: {0}' -f $heads)) }
    }
  }

  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('## Projets GAS (fonctions detectees)')
  foreach ($p in $Manifest.gas_projects) {
    [void]$sb.AppendLine(('- {0} - {1} lignes' -f $p.file, $p.lines))
    $funcs = ($p.functions | Select-Object -First 20) -join ', '
    if ($funcs) { [void]$sb.AppendLine(('  - fonctions: {0}' -f $funcs)) }
  }

  $out = Join-Path $SnapshotDir 'brief.md'
  $sb.ToString() | Set-Content -LiteralPath $out -Encoding UTF8
  Write-Host ("[BRIEF] {0}" -f $out)
  return $out
}

function Write-DiffMd {
  param(
    [Parameter(Mandatory)][string]$PrevManifestPath,
    [Parameter(Mandatory)][string]$CurrManifestPath,
    [Parameter(Mandatory)][string]$OutPath
  )
  $prev = Get-Content -LiteralPath $PrevManifestPath -Raw | ConvertFrom-Json
  $curr = Get-Content -LiteralPath $CurrManifestPath -Raw | ConvertFrom-Json

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('# Diff')
  [void]$sb.AppendLine(('- Ancien: {0}' -f $prev.snapshot))
  [void]$sb.AppendLine(('- Nouveau: {0}' -f $curr.snapshot))
  [void]$sb.AppendLine('')

  $pTabs = @{}
  foreach ($s in $prev.sheets) { foreach ($t in $s.tabs) { $pTabs[$t.file] = $t } }
  $cTabs = @{}
  foreach ($s in $curr.sheets) { foreach ($t in $s.tabs) { $cTabs[$t.file] = $t } }

  $added  = @($cTabs.Keys | Where-Object { -not $pTabs.ContainsKey($_) })
  $removed= @($pTabs.Keys | Where-Object { -not $cTabs.ContainsKey($_) })
  $common = @($cTabs.Keys | Where-Object { $pTabs.ContainsKey($_) })

  [void]$sb.AppendLine('## Onglets ajoutes')
  if ($added.Count -eq 0) { [void]$sb.AppendLine('- (aucun)') }
  foreach ($k in $added) { [void]$sb.AppendLine(('- {0}' -f $k)) }

  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('## Onglets supprimes')
  if ($removed.Count -eq 0) { [void]$sb.AppendLine('- (aucun)') }
  foreach ($k in $removed) { [void]$sb.AppendLine(('- {0}' -f $k)) }

  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('## Modifications')
  $changes = 0
  foreach ($k in $common) {
    $a = $pTabs[$k]; $b = $cTabs[$k]
    if ($a.sha256 -ne $b.sha256 -or $a.rows -ne $b.rows -or $a.cols -ne $b.cols) {
      $delta = $b.rows - $a.rows
      $sign = if ($delta -gt 0) {'+'} elseif ($delta -lt 0) {'-'} else {'0'}
      [void]$sb.AppendLine(('- {0} : lignes {1} -> {2} ({3}{4}), colonnes {5} -> {6}' -f $k,$a.rows,$b.rows,$sign,$delta,$a.cols,$b.cols))
      $changes++
    }
  }
  if ($changes -eq 0) { [void]$sb.AppendLine('- (aucun changement detecte sur les CSV communs)') }

  [void]$sb.AppendLine('')
  [void]$sb.AppendLine('## Scripts GAS concat')
  $pIdx = @{}; foreach ($p in $prev.gas_projects){ $pIdx[$p.file]=$p }
  $cIdx = @{}; foreach ($p in $curr.gas_projects){ $cIdx[$p.file]=$p }

  $pOnly = @($pIdx.Keys | Where-Object { -not $cIdx.ContainsKey($_) })
  $cOnly = @($cIdx.Keys | Where-Object { -not $pIdx.ContainsKey($_) })
  $both  = @($cIdx.Keys | Where-Object { $pIdx.ContainsKey($_) })

  if ($cOnly.Count -gt 0) { [void]$sb.AppendLine(('- Ajoutes: {0}' -f ($cOnly -join ', '))) }
  if ($pOnly.Count -gt 0) { [void]$sb.AppendLine(('- Supprimes: {0}' -f ($pOnly -join ', '))) }
  foreach ($k in $both) {
    if ($pIdx[$k].sha256 -ne $cIdx[$k].sha256 -or $pIdx[$k].lines -ne $cIdx[$k].lines) {
      [void]$sb.AppendLine(('- Modifie: {0} ({1} -> {2} lignes)' -f $k,$pIdx[$k].lines,$cIdx[$k].lines))
    }
  }

  $sb.ToString() | Set-Content -LiteralPath $OutPath -Encoding UTF8
  Write-Host ("[DIFF] {0}" -f $OutPath)
  return $OutPath
}

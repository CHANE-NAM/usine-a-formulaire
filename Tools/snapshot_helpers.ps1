# Tools\snapshot_helpers.ps1
# Helpers robustes : Write-Manifest / Write-BriefMd / Write-DiffMd
# Garde-fou anti double-chargement
if ($script:__SNAP_HELPERS_LOADED) { return }
$script:__SNAP_HELPERS_LOADED = $true

function _Get-RelativePath {
  param(
    [Parameter(Mandatory=$true)][string]$BasePath,
    [Parameter(Mandatory=$true)][string]$FullPath
  )
  try {
    $base = (Resolve-Path -LiteralPath $BasePath).Path
    $full = (Resolve-Path -LiteralPath $FullPath).Path
    $baseUri = [Uri]((Join-Path $base [IO.Path]::DirectorySeparatorChar))
    $relUri  = $baseUri.MakeRelativeUri([Uri]$full)
    return [Uri]::UnescapeDataString($relUri.ToString()).Replace('/', [IO.Path]::DirectorySeparatorChar)
  } catch { return $FullPath }
}

function _Get-FileSha1 {
  param([Parameter(Mandatory=$true)][string]$Path)
  try {
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $fs = [System.IO.File]::OpenRead($Path)
    try {
      ($sha1.ComputeHash($fs) | ForEach-Object { $_.ToString('x2') }) -join ''
    } finally {
      $fs.Dispose(); $sha1.Dispose()
    }
  } catch { '' }
}

function _Classify-File {
  param([Parameter(Mandatory=$true)][System.IO.FileInfo]$File)
  $n = $File.Name; $e = $File.Extension.ToLowerInvariant()
  if ($n -like 'scripts_*.txt') { 'concat' }
  elseif ($e -eq '.csv') { 'csv' }
  elseif ($e -eq '.zip') { 'zip' }
  elseif ($e -eq '.json') { 'json' }
  elseif ($e -eq '.md') { 'markdown' }
  elseif ($e -eq '.txt') { 'text' }
  else { 'other' }
}

function _List-SnapshotFiles {
  param(
    [Parameter(Mandatory=$true)][string]$SnapshotDir,
    [Parameter(Mandatory=$true)][string]$RepoRoot
  )
  $items = Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Force -ErrorAction SilentlyContinue
  foreach ($f in $items) {
    [PSCustomObject]@{
      FullPath      = $f.FullName
      RelToSnapshot = (_Get-RelativePath -BasePath $SnapshotDir -FullPath $f.FullName)
      RelToRepo     = (_Get-RelativePath -BasePath $RepoRoot   -FullPath $f.FullName)
      Name          = $f.Name
      Extension     = $f.Extension
      Type          = (_Classify-File -File $f)
      Length        = $f.Length
      LastWriteUtc  = $f.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
      Sha1          = (_Get-FileSha1 -Path $f.FullName)
    }
  }
}

function Write-Manifest {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$SnapshotDir,
    [Parameter(Mandatory=$true)][string]$RepoRoot
  )
  if (-not (Test-Path -LiteralPath $SnapshotDir)) { throw "SnapshotDir introuvable: $SnapshotDir" }
  if (-not (Test-Path -LiteralPath $RepoRoot))   { throw "RepoRoot introuvable: $RepoRoot" }

  $files = @(_List-SnapshotFiles -SnapshotDir $SnapshotDir -RepoRoot $RepoRoot)
  $byType = @{}
  foreach ($g in ($files | Group-Object Type)) { $byType[$g.Name] = $g.Count }

  $manifest = [PSCustomObject]@{
    summary = [PSCustomObject]@{
      snapshotDir = $SnapshotDir
      repoRoot    = $RepoRoot
      generatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssZ')
      counts      = [PSCustomObject]@{
        total  = $files.Count
        byType = $byType
      }
      totalSize   = ($files | Measure-Object Length -Sum).Sum
    }
    files = $files
  }

  $out = Join-Path $SnapshotDir 'manifest.json'
  $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $out -Encoding UTF8
  Write-Host "[MANIFEST] $out"
  return $out
}

function Write-BriefMd {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$SnapshotDir,
    [Parameter(Mandatory=$true)]$Manifest
  )
  if ($Manifest -is [string]) {
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "Manifest introuvable: $Manifest" }
    $m = Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json
  } else { $m = $Manifest }

  $sum   = $m.summary
  $files = $m.files

  # -- byType est un PSCustomObject => passer par .PSObject.Properties
  $countsProps = @()
  if ($sum.counts.byType) {
    $countsProps = $sum.counts.byType.PSObject.Properties |
                   Sort-Object Name |
                   ForEach-Object { @{ Name = $_.Name; Value = $_.Value } }
  }

  $countsRows = @()
  if ($countsProps.Count) {
    $countsRows = $countsProps | ForEach-Object { '| ' + $_.Name + ' | ' + $_.Value + ' |' }
  }

  $topRows = ($files | Sort-Object Length -Descending | Select-Object -First 10 | ForEach-Object {
    '| ' + $_.RelToSnapshot + ' | ' + ([math]::Round($_.Length/1KB,1)) + ' KB |'
  })

  $md = New-Object System.Collections.Generic.List[string]
  $md.Add('# Snapshot brief')
  $md.Add('')
  $md.Add('Informations')
  $md.Add('- Snapshot : ' + $sum.snapshotDir)
  $md.Add('- Genere   : ' + $sum.generatedAt)
  $md.Add('- Fichiers : ' + $sum.counts.total + '  -  Taille totale : ' + ([math]::Round($sum.totalSize/1MB,2)) + ' MB')
  $md.Add('')
  $md.Add('## Repartition par type')
  $md.Add('')
  $md.Add('| Type | Nb |')
  $md.Add('|------|----|')
  if ($countsRows.Count) { $countsRows | ForEach-Object { $md.Add($_) } } else { $md.Add('| (aucun) | 0 |') }
  $md.Add('')
  $md.Add('## Top 10 par taille')
  $md.Add('')
  $md.Add('| Chemin | Taille |')
  $md.Add('|--------|--------|')
  if ($topRows.Count) { $topRows | ForEach-Object { $md.Add($_) } } else { $md.Add('| (aucun) | 0 |') }

  $out = Join-Path $SnapshotDir 'brief.md'
  ($md -join "`n") | Set-Content -LiteralPath $out -Encoding UTF8
  Write-Host "[BRIEF] $out"
  return $out
}


function Write-DiffMd {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$PrevManifestPath,
    [Parameter(Mandatory=$true)][string]$CurrManifestPath,
    [Parameter(Mandatory=$true)][string]$OutPath
  )
  if (-not (Test-Path -LiteralPath $CurrManifestPath)) { throw "CurrManifestPath introuvable: $CurrManifestPath" }
  if (-not (Test-Path -LiteralPath $PrevManifestPath)) {
    Write-Host '[DIFF] Manifest précédent absent -> diff non généré.' -ForegroundColor Yellow
    return $null
  }

  $prev = Get-Content -LiteralPath $PrevManifestPath -Raw | ConvertFrom-Json
  $curr = Get-Content -LiteralPath $CurrManifestPath -Raw | ConvertFrom-Json

  $p = @{}; foreach ($f in $prev.files) { $p[$f.RelToSnapshot] = $f }
  $c = @{}; foreach ($f in $curr.files) { $c[$f.RelToSnapshot] = $f }

  $hs = New-Object System.Collections.Generic.HashSet[string]
  foreach ($k in $p.Keys) { $hs.Add($k) | Out-Null }
  foreach ($k in $c.Keys) { $hs.Add($k) | Out-Null }

  $added=@(); $removed=@(); $changed=@()
  foreach ($k in $hs) {
    $pv = $p[$k]; $cv = $c[$k]
    if ($null -eq $pv -and $null -ne $cv) { $added   += $cv; continue }
    if ($null -ne $pv -and $null -eq $cv) { $removed += $pv; continue }
    if ($null -ne $pv -and $null -ne $cv) {
      if ($pv.Sha1 -ne $cv.Sha1 -or [int64]$pv.Length -ne [int64]$cv.Length) {
        $changed += [PSCustomObject]@{
          Path=$k; FromKB=[math]::Round([double]$pv.Length/1KB,1); ToKB=[math]::Round([double]$cv.Length/1KB,1)
        }
      }
    }
  }

  $md = New-Object System.Collections.Generic.List[string]
  $md.Add('# Diff snapshot')
  $md.Add('')
  $md.Add('Ancien manifest : ' + $PrevManifestPath)
  $md.Add('Nouveau manifest : ' + $CurrManifestPath)
  $md.Add('')

  $md.Add('## Ajouts (' + $added.Count + ')')
  if ($added.Count) {
    foreach ($a in ($added | Sort-Object RelToSnapshot)) {
      $md.Add('* + ' + $a.RelToSnapshot + ' (' + ([math]::Round($a.Length/1KB,1)) + ' KB)')
    }
  } else { $md.Add('* (aucun)') }
  $md.Add('')

  $md.Add('## Suppressions (' + $removed.Count + ')')
  if ($removed.Count) {
    foreach ($r in ($removed | Sort-Object RelToSnapshot)) {
      $md.Add('* - ' + $r.RelToSnapshot + ' (' + ([math]::Round($r.Length/1KB,1)) + ' KB)')
    }
  } else { $md.Add('* (aucune)') }
  $md.Add('')

  $md.Add('## Modifications (' + $changed.Count + ')')
  if ($changed.Count) {
    foreach ($ch in ($changed | Sort-Object Path)) {
      $md.Add('* ' + $ch.Path + ' : ' + $ch.FromKB + ' KB -> ' + $ch.ToKB + ' KB')
    }
  } else { $md.Add('* (aucune)') }

  ($md -join "`n") | Set-Content -LiteralPath $OutPath -Encoding UTF8
  Write-Host "[DIFF] $OutPath"
  return $OutPath
}

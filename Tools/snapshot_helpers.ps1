# Tools\snapshot_helpers.ps1
# Helpers pour manifest/brief/diff — version "safe"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Set-StrictMode -Version Latest

function _Utf8NoBom([string]$Path, [string]$Content) {
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function _TryGitInfo([string]$RepoRoot) {
  $info = [ordered]@{ present = $false; branch = $null; commit = $null; describe = $null }
  try {
    $git = Get-Command git -ErrorAction Stop | Select-Object -First 1
    if ($git) {
      Push-Location $RepoRoot
      try {
        $branch  = (git rev-parse --abbrev-ref HEAD 2>$null)
        $commit  = (git rev-parse --short=8 HEAD 2>$null)
        $describe= (git describe --always --dirty 2>$null)
        if ($commit) {
          $info.present = $true
          $info.branch  = $branch
          $info.commit  = $commit
          $info.describe= $describe
        }
      } finally { Pop-Location }
    }
  } catch { }
  return $info
}

function _RelPath([string]$Base, [string]$Path) {
  try {
    $b = (Resolve-Path -LiteralPath $Base).Path
    $p = (Resolve-Path -LiteralPath $Path).Path
    $uriB = New-Object System.Uri("$b\")
    $uriP = New-Object System.Uri($p)
    return [System.Uri]::UnescapeDataString($uriB.MakeRelativeUri($uriP).ToString()) -replace '/', '\'
  } catch { return $Path }
}

function _HashSha1([string]$Path) {
  try {
    # Get-FileHash est dispo sur PowerShell 5+, SHA1 reste suffisant pour “diff”
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA1 -ErrorAction Stop).Hash.ToLowerInvariant()
  } catch { return $null }
}

function _ListSnapshotFiles([string]$SnapshotDir, [string]$RepoRoot) {
  $items = @()
  Get-ChildItem -LiteralPath $SnapshotDir -Recurse -File -Force | ForEach-Object {
    $rel = _RelPath $SnapshotDir $_.FullName
    $items += [ordered]@{
      relPath        = $rel
      size           = $_.Length
      lastWriteTime  = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
      sha1           = _HashSha1 $_.FullName
      kind           = (
        if ($_.Name -like "scripts_*.txt") { "concat" }
        elseif ($_.Extension -eq ".csv")   { "csv" }
        elseif ($_.Extension -eq ".zip")   { "zip" }
        elseif ($_.Extension -eq ".json")  { "json" }
        elseif ($_.Extension -eq ".md")    { "markdown" }
        else { "other" }
      )
    }
  }
  return ,$items
}

function Write-Manifest {
  <#
    .SYNOPSIS
      Génère manifest.json dans un dossier snapshot.
    .PARAMETER SnapshotDir
      Dossier du snapshot (SNAPSHOT_YYYYMMDD_HHMMSS).
    .PARAMETER RepoRoot
      Racine du repo (utilisée pour infos git + contexte).
    .OUTPUTS
      Chemin du manifest.json écrit.
  #>
  param(
    [Parameter(Mandatory=$true)] [string]$SnapshotDir,
    [Parameter(Mandatory=$true)] [string]$RepoRoot
  )
  if (-not (Test-Path -LiteralPath $SnapshotDir)) { throw "SnapshotDir introuvable: $SnapshotDir" }
  if (-not (Test-Path -LiteralPath $RepoRoot))    { throw "RepoRoot introuvable: $RepoRoot" }

  $createdAt = Get-Date
  $name      = Split-Path -Leaf $SnapshotDir

  $gitInfo = _TryGitInfo $RepoRoot
  $files   = _ListSnapshotFiles $SnapshotDir $RepoRoot

  # Détection de sous-dossiers CSV issus de l'export (optionnel)
  $csvRoots = Get-ChildItem -LiteralPath $SnapshotDir -Directory -ErrorAction SilentlyContinue |
              Where-Object { Test-Path (Join-Path $_.FullName '*.csv') } |
              ForEach-Object { $_.Name }

  $manifest = [ordered]@{
    schemaVersion = 1
    snapshot      = [ordered]@{
      name       = $name
      dir        = $SnapshotDir
      createdAt  = $createdAt.ToString("yyyy-MM-dd HH:mm:ss")
    }
    repo          = [ordered]@{
      root    = $RepoRoot
      git     = $gitInfo
    }
    content       = [ordered]@{
      csvRoots  = $csvRoots
      files     = $files
      counters  = [ordered]@{
        totalFiles = $files.Count
        totalSize  = ($files | Measure-Object -Property size -Sum).Sum
        nbCsv      = ($files | ? { $_.kind -eq 'csv' }).Count
        nbConcat   = ($files | ? { $_.kind -eq 'concat' }).Count
        nbZip      = ($files | ? { $_.kind -eq 'zip' }).Count
      }
    }
  }

  $json = $manifest | ConvertTo-Json -Depth 8
  $out  = Join-Path $SnapshotDir 'manifest.json'
  _Utf8NoBom $out $json
  Write-Host "[MANIFEST] $out"
  return $out
}

function _FormatSize([long]$bytes) {
  if ($bytes -ge 1GB) { "{0:N2} GB" -f ($bytes/1GB) }
  elseif ($bytes -ge 1MB) { "{0:N2} MB" -f ($bytes/1MB) }
  elseif ($bytes -ge 1KB) { "{0:N2} KB" -f ($bytes/1KB) }
  else { "$bytes B" }
}

function Write-BriefMd {
  <#
    .SYNOPSIS
      Génère un brief.md lisible à partir du manifest.
    .PARAMETER SnapshotDir
      Dossier snapshot.
    .PARAMETER Manifest
      Objet manifest (déjà ConvertFrom-Json) OU chemin vers manifest.json.
  #>
  param(
    [Parameter(Mandatory=$true)] [string]$SnapshotDir,
    [Parameter(Mandatory=$true)] $Manifest
  )

  if (-not (Test-Path -LiteralPath $SnapshotDir)) { throw "SnapshotDir introuvable: $SnapshotDir" }

  $man = $Manifest
  if ($Manifest -is [string]) {
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "Manifest introuvable: $Manifest" }
    $man = Get-Content -LiteralPath $Manifest -Raw | ConvertFrom-Json
  }

  # Lecture sûre des champs (manifest “souple”)
  $snapName   = $man.snapshot.name
  $snapDir    = $man.snapshot.dir
  $createdAt  = $man.snapshot.createdAt
  $repoRoot   = $man.repo.root
  $gitInfo    = $man.repo.git
  $files      = $man.content.files
  $counters   = $man.content.counters

  $totalSize  = if ($counters.totalSize) { [long]$counters.totalSize } else { ($files | Measure-Object -Property size -Sum).Sum }
  $nbFiles    = if ($counters.totalFiles) { [int]$counters.totalFiles } else { ($files).Count }

  $top10 = $files | Sort-Object size -Descending | Select-Object -First 10

  $md = @()
  $md += "# Snapshot: $snapName"
  $md += ""
  $md += "- **Créé le :** $createdAt"
  $md += "- **Dossier :** `$snapDir`"
  $md += "- **Repo root :** `$repoRoot`"
  if ($gitInfo.present -eq $true) {
    $md += "- **Git :** branch=`$($gitInfo.branch)`, commit=`$($gitInfo.commit)`, describe=`$($gitInfo.describe)`"
  } else {
    $md += "- **Git :** (non détecté)"
  }
  $md += ""
  $md += "## Contenu"
  $md += "- **Fichiers :** $nbFiles"
  $md += "- **Taille totale :** $(_FormatSize $totalSize)"
  $md += "- **CSV :** $($counters.nbCsv)  •  **Concat :** $($counters.nbConcat)  •  **ZIP :** $($counters.nbZip)"
  $md += ""
  $md += "## Top 10 fichiers par taille"
  $md += ""
  $md += "| Fichier | Taille | Type |"
  $md += "|---|---:|:--|"
  foreach ($f in $top10) {
    $md += "| `$($f.relPath)` | $(_FormatSize([long]$f.size)) | $($f.kind) |"
  }

  $out = Join-Path $SnapshotDir 'brief.md'
  _Utf8NoBom $out ($md -join [Environment]::NewLine)
  Write-Host "[BRIEF] $out"
  return $out
}

function _IndexByRelPath($files) {
  $map = @{}
  foreach ($f in $files) {
    if ($null -ne $f.relPath) {
      $map[$f.relPath.ToString()] = $f
    }
  }
  return $map
}

function _FilesFromManifestFlexible($manifestObj) {
  # Essaye de récupérer une liste homogène d’objets {relPath,size,sha1,kind}
  if ($manifestObj -and $manifestObj.content -and $manifestObj.content.files) {
    return @($manifestObj.content.files)
  }
  # Fallback ultra souple : recréer depuis le disque si on connaît snapshot.dir
  if ($manifestObj -and $manifestObj.snapshot -and $manifestObj.snapshot.dir -and (Test-Path $manifestObj.snapshot.dir)) {
    return _ListSnapshotFiles $manifestObj.snapshot.dir $manifestObj.repo.root
  }
  return @()
}

function Write-DiffMd {
  <#
    .SYNOPSIS
      Compare deux manifests et génère un diff.md (ajouts/suppressions/modifs).
    .PARAMETER PrevManifestPath
      Chemin vers manifest.json précédent.
    .PARAMETER CurrManifestPath
      Chemin vers manifest.json courant.
    .PARAMETER OutPath
      Chemin du diff.md à écrire.
  #>
  param(
    [Parameter(Mandatory=$true)] [string]$PrevManifestPath,
    [Parameter(Mandatory=$true)] [string]$CurrManifestPath,
    [Parameter(Mandatory=$true)] [string]$OutPath
  )

  if (-not (Test-Path -LiteralPath $CurrManifestPath)) { throw "CurrManifestPath introuvable: $CurrManifestPath" }
  if (-not (Test-Path -LiteralPath $PrevManifestPath)) {
    Write-Host "[DIFF] Manifest précédent absent -> diff non généré." -ForegroundColor Yellow
    return $null
  }

  $prev = Get-Content -LiteralPath $PrevManifestPath -Raw | ConvertFrom-Json
  $curr = Get-Content -LiteralPath $CurrManifestPath -Raw | ConvertFrom-Json

  $prevFiles = _FilesFromManifestFlexible $prev
  $currFiles = _FilesFromManifestFlexible $curr

  $iPrev = _IndexByRelPath $prevFiles
  $iCurr = _IndexByRelPath $currFiles

  $added = New-Object System.Collections.Generic.List[object]
  $removed = New-Object System.Collections.Generic.List[object]
  $changed = New-Object System.Collections.Generic.List[object]

  foreach ($k in $iCurr.Keys) {
    if (-not $iPrev.ContainsKey($k)) {
      $added.Add($iCurr[$k])
    } else {
      $p = $iPrev[$k]
      $c = $iCurr[$k]
      $pHash = "$($p.sha1)"
      $cHash = "$($c.sha1)"
      $pSize = [long]$p.size
      $cSize = [long]$c.size
      if (($pHash -ne $cHash -and $pHash) -or ($pSize -ne $cSize)) {
        $changed.Add([ordered]@{
          relPath = $k
          prev = [ordered]@{ size = $pSize; sha1 = $pHash }
          curr = [ordered]@{ size = $cSize; sha1 = $cHash }
        })
      }
    }
  }
  foreach ($k in $iPrev.Keys) {
    if (-not $iCurr.ContainsKey($k)) {
      $removed.Add($iPrev[$k])
    }
  }

  $md = @()
  $md += "# Diff entre snapshots"
  $md += ""
  $md += "- **Ancien :** `$PrevManifestPath`"
  $md += "- **Nouveau :** `$CurrManifestPath`"
  $md += ""
  $md += "## Résumé"
  $md += ""
  $md += "- Ajoutés : **$($added.Count)**"
  $md += "- Supprimés : **$($removed.Count)**"
  $md += "- Modifiés : **$($changed.Count)**"
  $md += ""

  if ($added.Count -gt 0) {
    $md += "## Ajouts"
    $md += "| Fichier | Taille |"
    $md += "|---|---:|"
    foreach ($f in $added | Sort-Object size -Descending) {
      $md += "| `$($f.relPath)` | $(_FormatSize([long]$f.size)) |"
    }
    $md += ""
  }

  if ($removed.Count -gt 0) {
    $md += "## Suppressions"
    $md += "| Fichier | Taille (ancienne) |"
    $md += "|---|---:|"
    foreach ($f in $removed | Sort-Object size -Descending) {
      $md += "| `$($f.relPath)` | $(_FormatSize([long]$f.size)) |"
    }
    $md += ""
  }

  if ($changed.Count -gt 0) {
    $md += "## Modifications"
    $md += "| Fichier | Taille (anc.) | Taille (nouv.) | Hash (anc.) | Hash (nouv.) |"
    $md += "|---|---:|---:|---|---|"
    foreach ($c in $changed | Sort-Object { [math]::Abs([long]$_.curr.size - [long]$_.prev.size) } -Descending) {
      $md += "| `$($c.relPath)` | $(_FormatSize([long]$c.prev.size)) | $(_FormatSize([long]$c.curr.size)) | $($c.prev.sha1) | $($c.curr.sha1) |"
    }
    $md += ""
  }

  _Utf8NoBom $OutPath ($md -join [Environment]::NewLine)
  Write-Host "[DIFF] $OutPath"
  return $OutPath
}

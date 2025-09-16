# Tools\snapshot_rk.ps1
# === Snapshot complet : GAS + CSV + ZIP (+ manifest/brief/diff, + hook docs optionnel) ===
[CmdletBinding()]
param(
  [switch]$NoDocs,     # ne pas appeler gen_docs.ps1
  [switch]$NoCsv,      # sauter l’export CSV
  [bool]$Spy = $true   # activer/désactiver le diagnostic "espion"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function Write-Section($text) {
  Write-Host ""
  Write-Host $text -ForegroundColor Cyan
}

# ------------------------------------------------------------------------------------
# 0) DIAGNOSTIC "ESPION" (peut être désactivé) + base chemins + transcript de log
# ------------------------------------------------------------------------------------
$EnableSpy = [bool]$Spy

# Base fiable pour les chemins (ok en .ps1 et en lancement manuel)
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

# Timestamp global (sert aussi au nom de snapshot et au fichier de log)
$ts = Get-Date -Format "yyyyMMdd_HHmmss"

# Transcript (log)
try {
  $LogsDir = Join-Path $ScriptRoot "logs"
  New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null
  $TranscriptPath = Join-Path $LogsDir ("snapshot_{0}.log" -f $ts)
  Start-Transcript -Path $TranscriptPath -Append | Out-Null
} catch {
  Write-Warning ("[LOG] Start-Transcript a échoué : {0}" -f $_.Exception.Message)
}

if ($EnableSpy) {
  try {
    $thisPath = if ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { Join-Path $ScriptRoot "snapshot_rk.ps1" }
    Write-Host ("[SPY] Analyse du fichier: {0}" -f $thisPath)
    $balCurly = 0; $balParen = 0; $lineNum = 0
    Get-Content -LiteralPath $thisPath | ForEach-Object {
      $lineNum++
      $opensCurly  = ([regex]::Matches($_, '\{')).Count
      $closesCurly = ([regex]::Matches($_, '\}')).Count
      $opensParen  = ([regex]::Matches($_, '\(')).Count
      $closesParen = ([regex]::Matches($_, '\)')).Count
      $balCurly += ($opensCurly - $closesCurly)
      $balParen += ($opensParen - $closesParen)
      if ($_ -match '`\s*$') { Write-Warning ("[SPY] Backtick fin de ligne -> {0}" -f $lineNum) }
      if ($_ -match '\xA0')  { Write-Warning ("[SPY] NBSP (0xA0) détecté -> {0}" -f $lineNum) }
      if ($_ -match '\x200B'){ Write-Warning ("[SPY] Zero-width space détecté -> {0}" -f $lineNum) }
    }
    Write-Host ("[SPY] Balance finale: {{}}={0}  ()={1}  (attendu: 0 / 0)" -f $balCurly, $balParen)
  } catch {
    Write-Warning ("[SPY] Échec diagnostic: {0}" -f $_.Exception.Message)
  }
}

# ------------------------------------------------------------------------------------
# 1) Import des helpers (manifest/brief/diff) — robuste et sans récursion
# ------------------------------------------------------------------------------------
$HelpersCandidates = @(
  (Join-Path $ScriptRoot 'snapshot_helpers.ps1'),
  (Join-Path $ScriptRoot 'Tools\snapshot_helpers.ps1'),
  (Join-Path ((Resolve-Path (Join-Path $ScriptRoot '..')).Path) 'Tools\snapshot_helpers.ps1')
) | Select-Object -Unique

$HelpersPath = $null
foreach ($cand in $HelpersCandidates) {
  if (Test-Path -LiteralPath $cand) { $HelpersPath = $cand; break }
}

$HelpersLoaded = $false
if ($HelpersPath) {
  try {
    . $HelpersPath
    $HelpersLoaded = $true
    Write-Host ("[META] Helpers chargés: {0}" -f $HelpersPath)
  } catch {
    Write-Warning ("[META] Échec chargement helpers: {0}" -f $_.Exception.Message)
  }
} else {
  Write-Host "[META] Helpers absents (Tools\snapshot_helpers.ps1 non trouvé) — manifest/brief/diff seront sautés."
}

# ------------------------------------------------------------------------------------
# 2) Dossiers
# ------------------------------------------------------------------------------------
$Repo      = (Resolve-Path (Join-Path $ScriptRoot "..")).Path
$ExportDir = Join-Path $Repo "export-onglets-csv"
New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null

# ------------------------------------------------------------------------------------
# 3) Snapshot : nom et répertoire
# ------------------------------------------------------------------------------------
$SNAPSHOT_NAME = "SNAPSHOT_$ts"
$SnapDir       = Join-Path $ExportDir $SNAPSHOT_NAME
New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null

Write-Host ("=== SNAPSHOT {0} ===" -f $SNAPSHOT_NAME)
Write-Host ("Repo     : {0}" -f $Repo)
Write-Host ("Snapshot : {0}" -f $SnapDir)

# ------------------------------------------------------------------------------------
# 4) CLASP pull (synchronisation des projets GAS locaux)
# ------------------------------------------------------------------------------------
Write-Section "[1/4] CLASP pull (via backup_gas.ps1) ..."
try {
  & (Join-Path $ScriptRoot "backup_gas.ps1")
} catch {
  Write-Warning ("[CLASP] Échec backup_gas.ps1 : {0}" -f $_.Exception.Message)
}

# ------------------------------------------------------------------------------------
# 5) Concat des scripts GAS par projet -> scripts__*.txt dans le snapshot
# ------------------------------------------------------------------------------------
Write-Section "[2/4] Concat des scripts par projet ..."

$bdd1 = Join-Path $Repo "03_BaseDeDonnées"
$bdd2 = Join-Path $Repo "03_BaseDeDonnees"
if     (Test-Path -LiteralPath $bdd1) { $bddDir = $bdd1 }
elseif (Test-Path -LiteralPath $bdd2) { $bddDir = $bdd2 }
else { $bddDir = $null }

$Projets = @()
$Projets += ,@("[MOTEUR]V2 Usine à Tests",        (Join-Path $Repo "01_Moteur"))
$Projets += ,@("[CONFIG]V2 Usine à Tests",        (Join-Path $Repo "02_configuration"))
if ($bddDir) { $Projets += ,@("[BDD]V2 Tests & Profils", $bddDir) } else { Write-Warning "Dossier BDD introuvable (03_BaseDeDonnées / 03_BaseDeDonnees)." }
$Projets += ,@("[TEMPLATE]V2 Kit de Traitement",  (Join-Path $Repo "04_Templates"))
$Projets += ,@("[BIBLIOTHEQUE]TEMPLATE", (Join-Path $Repo "05_Bibliotheque"))

foreach ($p in $Projets) {
  $pname = $p[0]
  $pdir  = $p[1]

  if (-not (Test-Path -LiteralPath $pdir)) { Write-Warning ("Dossier introuvable: {0}" -f $pdir); continue }

  $safeName = ($pname -replace '[^\w\-]+','_')
  $outTxt   = Join-Path $SnapDir ("scripts_" + $safeName + ".txt")

  $files = Get-ChildItem -LiteralPath $pdir -Recurse -File -ErrorAction SilentlyContinue |
           Where-Object { ($_.Extension -in ".gs",".js",".ts") -or ($_.Name -eq "appsscript.json") }

  if (-not $files) { Write-Warning ("Aucun fichier GAS trouvé dans {0}" -f $pdir); continue }

  ("=== Projet: {0} ({1}) ==={2}" -f $pname, $pdir, [Environment]::NewLine) |
    Out-File -FilePath $outTxt -Encoding UTF8

  foreach ($f in $files) {
    ("{0}--- FILE: {1} ---{0}" -f [Environment]::NewLine, $f.FullName) | Out-File -FilePath $outTxt -Encoding UTF8 -Append
    Get-Content -LiteralPath $f.FullName -Raw | Out-File -FilePath $outTxt -Encoding UTF8 -Append
  }
  Write-Host ("[OK] Concat: {0}" -f $outTxt)
}

# ------------------------------------------------------------------------------------
# 6) Export CSV des 4 classeurs (par IDs) via export-onglets-csv\index.js
# ------------------------------------------------------------------------------------
if (-not $NoCsv) {
  Write-Section "[3/4] Export des onglets -> CSV ..."

  # IDs des classeurs (modifiable ici)
  $Ids = @(
    "1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8", # [BDD]V2 Tests & Profils
    "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ", # [CONFIG] Usine à Tests
    "1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU", # [TEMPLATE]V2 Kit de Traitement
    "1hrcdsMRwx4FuHTvvtJoq2AVh8XTzwp5MErJ3UQ0OA5E"  # [MOTEUR] Usine à Tests
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

  # Chemins (surchargables via variables d'environnement RK_CREDS / RK_TOKEN)
  $credsPath = if ($env:RK_CREDS) { $env:RK_CREDS } else { "C:\secrets\rk_oauth\credentials.json" }
  $tokenPath = if ($env:RK_TOKEN) { $env:RK_TOKEN } else { "C:\secrets\rk_oauth\token.json" }

  # Préconditions minimales
  $IndexJs = Join-Path $ExportDir 'index.js'
  $nodeCmd = Get-Command node -ErrorAction SilentlyContinue

  $canRun = $true
  if (-not (Test-Path -LiteralPath $IndexJs)) { Write-Warning "[CSV] export-onglets-csv\index.js introuvable — étape CSV sautée."; $canRun = $false }
  if (-not $nodeCmd)                          { Write-Warning "[CSV] Node.js (commande 'node') introuvable — étape CSV sautée.";   $canRun = $false }
  if (-not (Test-Path -LiteralPath $credsPath)){ Write-Warning ("[CSV] credentials.json introuvable : {0} — étape CSV sautée." -f $credsPath); $canRun = $false }
  # NOTE: on NE bloque PAS si token.json est absent -> l'exporteur déclenchera le flow OAuth
  if (-not (Test-Path -LiteralPath $tokenPath)) {
    Write-Warning ("[CSV] token.json absent : {0} — un nouveau consentement OAuth sera demandé." -f $tokenPath)
  }
  if (-not $Ids -or $Ids.Count -eq 0)         { Write-Warning "[CSV] Aucun ID de classeur fourni — étape CSV sautée.";            $canRun = $false }

  if ($canRun) {
    # Construction des arguments Node
    $nodeArgs = @(
      "--out",   $SnapDir,
      "--creds", $credsPath,
      "--token", $tokenPath
    ) + ($Ids | ForEach-Object { @("--id", $_) })

    # Log de la commande exacte
    $cmdPreview = "$($nodeCmd.Source) `"$IndexJs`" " + ($nodeArgs | ForEach-Object { if ($_ -match '\s') { '"{0}"' -f $_ } else { $_ } }) -join ' '
    Write-Host "NODE CMD: $cmdPreview" -ForegroundColor Cyan

    # Exécution dans export-onglets-csv pour résoudre .\index.js correctement
    $csvBefore = (Get-ChildItem -LiteralPath $SnapDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count
    Push-Location -LiteralPath $ExportDir
    try {
      & $nodeCmd.Source ".\index.js" @nodeArgs
      if ($LASTEXITCODE -ne 0) { throw "Échec export CSV (node exit code = $LASTEXITCODE)." }
      $csvAfter = (Get-ChildItem -LiteralPath $SnapDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count
      $delta = $csvAfter - $csvBefore
      Write-Host ("[CSV] Export terminé : {0} fichier(s) CSV ajouté(s) dans {1}" -f [math]::Max($delta,0), $SnapDir)
    } catch {
      Write-Warning ("[CSV] Échec export CSV : {0}" -f $_.Exception.Message)
    } finally {
      Pop-Location
    }
  }
} else {
  Write-Host "[CSV] Étape export CSV ignorée (NoCsv)."
}

# ------------------------------------------------------------------------------------
# 7) Manifest / Brief / Diff (si helpers chargés)
# ------------------------------------------------------------------------------------
if ($HelpersLoaded -and (Get-Command Write-Manifest -ErrorAction SilentlyContinue)) {
  try {
    $manifest = Write-Manifest -SnapshotDir $SnapDir -RepoRoot $Repo
    $briefMd  = Write-BriefMd  -SnapshotDir $SnapDir -Manifest $manifest

    $prev = Get-ChildItem -LiteralPath $ExportDir -Directory |
            Where-Object { $_.FullName -ne $SnapDir -and (Test-Path (Join-Path $_.FullName 'manifest.json')) } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($prev) {
      $prevManifest = Join-Path $prev.FullName 'manifest.json'
      $diffArgs = @{
        PrevManifestPath = $prevManifest
        CurrManifestPath = (Join-Path $SnapDir 'manifest.json')
        OutPath          = (Join-Path $SnapDir 'diff.md')
      }
      Write-DiffMd @diffArgs | Out-Null
      Write-Host ("[DIFF] {0}" -f $diffArgs.OutPath)
    } else {
      Write-Host "[DIFF] Aucun manifest précédent existant — génération de diff sautée."
    }
  } catch {
    Write-Warning ("[META] Échec génération manifest/brief/diff : {0}" -f $_.Exception.Message)
  }
} else {
  Write-Host "[META] Helpers indisponibles — étape manifest/brief/diff ignorée."
}

# ------------------------------------------------------------------------------------
# 8) HOOK optionnel : génération de documents “AI / État / Utilisateurs”
# ------------------------------------------------------------------------------------
if (-not $NoDocs) {
  try {
    $GenDocs = Join-Path $ScriptRoot "gen_docs.ps1"
    if (Test-Path -LiteralPath $GenDocs) {
      Write-Section "[opt] Génération des documents (hook gen_docs.ps1) ..."
      & $GenDocs -RepoRoot $Repo -SnapshotDir $SnapDir -ExportDir $ExportDir -Timestamp $ts
    } else {
      Write-Host "[DOCS] gen_docs.ps1 absent — étape ignorée."
    }
  } catch {
    Write-Warning ("[DOCS] gen_docs.ps1 a échoué : {0}" -f $_.Exception.Message)
  }
} else {
  Write-Host "[DOCS] Génération de documents ignorée (NoDocs)."
}

# ------------------------------------------------------------------------------------
# 9) ZIP du snapshot
# ------------------------------------------------------------------------------------
Write-Section "[4/4] ZIP du snapshot ..."
$zipPath = Join-Path $ExportDir ($SNAPSHOT_NAME + ".zip")
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path $SnapDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host ("[ZIP] Archive: {0}" -f $zipPath)

# ------------------------------------------------------------------------------------
# 10) RÉTENTION DES SNAPSHOTS (garder les N derniers)
# ------------------------------------------------------------------------------------
$RetentionCount = 12
try {
  $allSnaps = Get-ChildItem -LiteralPath $ExportDir -Directory |
              Where-Object { $_.Name -like 'SNAPSHOT_*' } |
              Sort-Object LastWriteTime -Descending

  if ($allSnaps.Count -gt $RetentionCount) {
    $toDelete = $allSnaps | Select-Object -Skip $RetentionCount
    foreach ($old in $toDelete) {
      try {
        $oldZip = Join-Path $ExportDir ($old.Name + '.zip')
        if (Test-Path -LiteralPath $oldZip) { Remove-Item -LiteralPath $oldZip -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $old.FullName -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ("[RETENTION] Supprimé: {0}" -f $old.Name)
      } catch {
        Write-Warning ("[RETENTION] Échec suppression {0} : {1}" -f $old.Name, $_.Exception.Message)
      }
    }
  }
} catch {
  Write-Warning ("[RETENTION] Échec du traitement de rétention : {0}" -f $_.Exception.Message)
}

Write-Host ""
Write-Host ("[DONE] Snapshot: {0}" -f $SnapDir)

try { Stop-Transcript | Out-Null } catch {}

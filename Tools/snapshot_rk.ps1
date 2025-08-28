# Tools\snapshot_rk.ps1
# === Snapshot complet : GAS + CSV + ZIP ===
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

# -----------------------------
# 0) DIAGNOSTIC "ESPION" (peut être désactivé)
# -----------------------------
$EnableSpy = $true
if ($EnableSpy) {
  try {
    $thisPath = $MyInvocation.MyCommand.Path
    Write-Host "[SPY] Analyse du fichier: $thisPath"
    $balCurly = 0; $balParen = 0; $lineNum = 0
    Get-Content -LiteralPath $thisPath | ForEach-Object {
      $lineNum++
      $opensCurly  = ([regex]::Matches($_, '\{')).Count
      $closesCurly = ([regex]::Matches($_, '\}')).Count
      $opensParen  = ([regex]::Matches($_, '\(')).Count
      $closesParen = ([regex]::Matches($_, '\)')).Count
      $balCurly += ($opensCurly - $closesCurly)
      $balParen += ($opensParen - $closesParen)
      if ($_ -match '`\s*$') { Write-Warning "[SPY] Backtick fin de ligne -> $lineNum" }
      if ($_ -match '\xA0')  { Write-Warning "[SPY] NBSP (0xA0) detecte -> $lineNum" }
      if ($_ -match '\x200B'){ Write-Warning "[SPY] Zero-width space detecte -> $lineNum" }
    }
    Write-Host ("[SPY] Balance finale: {{}}={0}  ()={1}  (attendu: 0 / 0)" -f $balCurly, $balParen)
  } catch {
    Write-Warning "[SPY] Echec diagnostic: $($_.Exception.Message)"
  }
}

# -----------------------------
# 1) Dossiers
# -----------------------------
$Repo      = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ExportDir = Join-Path $Repo "export-onglets-csv"
$LogsDir   = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $LogsDir,$ExportDir | Out-Null

# -----------------------------
# 2) Timestamp + snapshot
# -----------------------------
$ts            = Get-Date -Format "yyyyMMdd_HHmmss"
$SNAPSHOT_NAME = "SNAPSHOT_$ts"
$SnapDir       = Join-Path $ExportDir $SNAPSHOT_NAME
New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null

Write-Host "=== SNAPSHOT $SNAPSHOT_NAME ==="
Write-Host ("Repo     : {0}" -f $Repo)
Write-Host ("Snapshot : {0}" -f $SnapDir)

# -----------------------------
# 3) CLASP pull
# -----------------------------
Write-Host ""
Write-Host "[1/4] CLASP pull (via backup_gas.ps1) ..."
& (Join-Path $PSScriptRoot "backup_gas.ps1")

# -----------------------------
# 4) Concat des scripts
# -----------------------------
Write-Host ""
Write-Host "[2/4] Concat des scripts par projet ..."

# BDD : gère le nom avec/sans accents (affectation classique)
$bdd1 = Join-Path $Repo "03_BaseDeDonnées"
$bdd2 = Join-Path $Repo "03_BaseDeDonnees"
if (Test-Path $bdd1) {
  $bddDir = $bdd1
} elseif (Test-Path $bdd2) {
  $bddDir = $bdd2
} else {
  $bddDir = $null
}

# Liste des projets : paires [0]=name ; [1]=dir (évite @{...})
$Projets = @()
$Projets += ,@("[MOTEUR]V2 Usine à Tests",        (Join-Path $Repo "01_Moteur"))
$Projets += ,@("[CONFIG]V2 Usine à Tests",        (Join-Path $Repo "02_configuration"))
if ($bddDir) {
  $Projets += ,@("[BDD]V2 Tests & Profils",      $bddDir)
} else {
  Write-Warning "Dossier BDD introuvable (03_BaseDeDonnées / 03_BaseDeDonnees)."
}
$Projets += ,@("[TEMPLATE]V2 Kit de Traitement",  (Join-Path $Repo "04_Templates"))

foreach ($p in $Projets) {
  $pname = $p[0]
  $pdir  = $p[1]

  if (-not (Test-Path $pdir)) {
    Write-Warning ("Dossier introuvable: {0}" -f $pdir)
    continue
  }

  # Nom de fichier "safe" (ASCII: lettres/chiffres/underscore/tiret)
  $safeName = ($pname -replace '[^\w\-]+','_')
  $outTxt   = Join-Path $SnapDir ("scripts_" + $safeName + ".txt")

  # Filtrage ROBUSTE (pas de -Include)
  $files = Get-ChildItem -Path $pdir -Recurse -File -ErrorAction SilentlyContinue |
           Where-Object { ($_.Extension -in ".gs",".js",".ts") -or ($_.Name -eq "appsscript.json") }

  if (-not $files) {
    Write-Warning ("Aucun fichier GAS trouve dans {0}" -f $pdir)
    continue
  }

  # En-tete de projet (sans backticks)
  ("=== Projet: {0} ({1}) ==={2}" -f $pname, $pdir, [Environment]::NewLine) |
    Out-File -FilePath $outTxt -Encoding UTF8

  foreach ($f in $files) {
    ("{0}--- FILE: {1} ---{0}" -f [Environment]::NewLine, $f.FullName) |
      Out-File -FilePath $outTxt -Encoding UTF8 -Append
    Get-Content -LiteralPath $f.FullName -Raw |
      Out-File -FilePath $outTxt -Encoding UTF8 -Append
  }

  Write-Host ("[OK] Concat: {0}" -f $outTxt)
}

# -----------------------------
# 5) Export CSV des 4 classeurs (par IDs)
# -----------------------------
Write-Host ""
Write-Host "[3/4] Export des onglets -> CSV ..."
$Ids = @(
  "1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8", # [BDD]V2 Tests & Profils
  "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ", # [CONFIG] Usine à Tests
  "1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU", # [TEMPLATE]V2 Kit de Traitement
  "1hrcdsMRwx4FuHTvvtJoq2AVh8XTzwp5MErJ3UQ0OA5E"  # [MOTEUR] Usine à Tests
)

$nodeArgs = @("--out", $SnapDir) + ($Ids | ForEach-Object { @("--id", $_) })
Push-Location $ExportDir
try {
  node ".\index.js" @nodeArgs
} finally {
  Pop-Location
}

# -----------------------------
# 6) ZIP
# -----------------------------
Write-Host ""
Write-Host "[4/4] ZIP du snapshot ..."
$zipPath = Join-Path $ExportDir ($SNAPSHOT_NAME + ".zip")
if (Test-Path $zipPath) {
  Remove-Item $zipPath -Force
}
Compress-Archive -Path $SnapDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host ("[ZIP] Archive: {0}" -f $zipPath)

Write-Host ""
Write-Host ("[DONE] Snapshot: {0}" -f $SnapDir)

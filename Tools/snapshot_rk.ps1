# Tools\snapshot_rk.ps1
# === Snapshot complet R&K : GAS + CSV ===
# Exécution depuis la racine du repo :
#   powershell -ExecutionPolicy Bypass -File ".\Tools\snapshot_rk.ps1"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

# -- Répertoires —
$Repo = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ExportDir = Join-Path $Repo "export-onglets-csv"
$LogsDir = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

# -- Timestamp & chemins snapshot —
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$SNAPSHOT_NAME = "SNAPSHOT_$ts"
$SnapDir = Join-Path $ExportDir $SNAPSHOT_NAME
New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null

Write-Host "=== SNAPSHOT $SNAPSHOT_NAME ==="
Write-Host "Repo       : $Repo"
Write-Host "ExportDir  : $ExportDir"
Write-Host "Snapshot   : $SnapDir"

# -- 1) Rafraîchir les Apps Script (CLASP) —
Write-Host "`n[1/4] CLASP pull (via backup_gas.ps1) ..."
& "$PSScriptRoot\backup_gas.ps1"

# -- 2) Concaténer les scripts de chaque projet —
Write-Host "`n[2/4] Concat des scripts par projet ..."

# Dossiers des 4 projets (adaptés à ton arborescence)
$Projets = @(
  @{ name = "[MOTEUR]V2 Usine à Tests";     dir = Join-Path $Repo "01_Moteur" },
  @{ name = "[CONFIG]V2 Usine à Tests";     dir = Join-Path $Repo "02_configuration" },
  @{ name = "[BDD]V2 Tests & Profils";      dir = Join-Path $Repo "03_BaseDeDonnées" },
  @{ name = "[TEMPLATE]V2 Kit de Traitement"; dir = Join-Path $Repo "04_Templates" }
)

foreach ($p in $Projets) {
  $pname = $p.name
  $pdir  = $p.dir
  if (-not (Test-Path $pdir)) { Write-Warning "Dossier introuvable: $pdir"; continue }

  $outTxt = Join-Path $SnapDir ("scripts_" + ($pname -replace '[^\w\-]+','_') + ".txt")
  $files = Get-ChildItem -Path $pdir -Recurse -File -Include *.gs,*.js,*.ts,appsscript.json -ErrorAction SilentlyContinue

  if (-not $files) { Write-Warning "Aucun fichier GAS trouvé dans $pdir"; continue }

  # Assemble tout dans un unique .txt, avec séparateurs
  "=== Projet: $pname ($pdir) ===`r`n" | Out-File -FilePath $outTxt -Encoding UTF8
  foreach ($f in $files) {
    "`r`n--- FILE: $($f.FullName) ---`r`n" | Out-File -FilePath $outTxt -Encoding UTF8 -Append
    Get-Content $f.FullName -Raw | Out-File -FilePath $outTxt -Encoding UTF8 -Append
  }
  Write-Host "Concat: $outTxt"
}   # <-- fermeture du foreach ($p in $Projets)

# -- 3) Export CSV de tous les onglets des 4 Google Sheets —
Write-Host "`n[3/4] Export des onglets -> CSV ..."

# Chemins des .gsheet (les raccourcis locaux que tu m'as donnés)
$GSheets = @(
  "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\01_Moteur\[MOTEUR]V2 Usine à Tests.gsheet",
  "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\02_configuration\[CONFIG]V2 Usine à Tests.gsheet",
  "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\03_BaseDeDonnées\[BDD]V2 Tests & Profils.gsheet",
  "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\04_Templates\[TEMPLATE]V2 Kit de Traitement.gsheet"
)

# Appelle le script Node avec :
#   --out  : dossier du snapshot pour que les CSV tombent dedans
#   --gsheet ... : les 4 chemins .gsheet (le script Node va extraire les IDs)
$nodeArgs = @("--out", $SnapDir) + ($GSheets | ForEach-Object { @("--gsheet", $_) })
Push-Location $ExportDir
try {
  node ".\index.js" @nodeArgs
} finally {
  Pop-Location
}

# -- 4) ZIP final du snapshot —
Write-Host "`n[4/4] ZIP du snapshot ..."
$zipPath = Join-Path $ExportDir ($SNAPSHOT_NAME + ".zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $SnapDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host "ZIP créé : $zipPath"
Write-Host "`nTerminé. Snapshot: $SnapDir"

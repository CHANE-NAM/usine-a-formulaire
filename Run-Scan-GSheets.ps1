# --- Les paramètres DOIVENT être en première position ---
param(
  [string]$Root = "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\05_DOSSIER - Cible Génération",
  [string]$Name = "Flauger Stéphane",
  [string]$Out  = "G:\Mon Drive\scan_gsheets_hits.csv"
)

# Run-Scan-GSheets.ps1 — Lance le scan Google Sheets (via .gsheet) avec venv auto + dépendances
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

# --- Localisation projet (dossier de CE script) ---
$ProjectDir = $PSScriptRoot
$VenvDir    = Join-Path $ProjectDir ".venv"
$Activate   = Join-Path $VenvDir "Scripts\Activate.ps1"
$PyScript   = Join-Path $ProjectDir "scan_gsheets_in_folder.py"
$Creds      = Join-Path $ProjectDir "credentials.json"

function Assert-File([string]$path, [string]$hint) {
  if (-not (Test-Path -LiteralPath $path)) { throw "Manquant: $path`n$hint" }
}
function Assert-Dir([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { throw "Dossier introuvable: $path" }
}

# --- Pré-checks ---
Assert-File $PyScript "Créez le fichier Python dans $ProjectDir (scan_gsheets_in_folder.py)."
Assert-File $Creds   "Placez votre credentials.json (OAuth Desktop) dans $ProjectDir."
Assert-Dir  $Root

# --- Python dispo ? ---
try { & python --version | Out-Null } catch { throw "Python n'est pas disponible dans le PATH. Installe Python 3.10+." }

# --- venv ---
if (-not (Test-Path -LiteralPath $VenvDir)) {
  Write-Host "[SETUP] Création de l'environnement virtuel…" -ForegroundColor Cyan
  & python -m venv $VenvDir
}

# --- Activation (bypass si politique restrictive) ---
$prevPolicy = Get-ExecutionPolicy
if ($prevPolicy -ne 'Bypass') { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force }
. $Activate

# --- Dépendances ---
$needInstall = $false
$checkCode = @'
import importlib, sys
mods = ["googleapiclient.discovery","google_auth_oauthlib","google.oauth2.credentials"]
try:
    for m in mods:
        importlib.import_module(m)
    sys.exit(0)
except Exception:
    sys.exit(1)
'@
python -c $checkCode
if ($LASTEXITCODE -ne 0) { $needInstall = $true }

if ($needInstall) {
  Write-Host "[SETUP] Installation des dépendances…" -ForegroundColor Cyan
  python -m pip install --upgrade pip
  python -m pip install google-auth google-auth-oauthlib google-api-python-client
}

# --- Exécution ---
Write-Host "`n[RUN] Scan en cours…" -ForegroundColor Green
$cmd = @(
  $PyScript,
  '-root', $Root,
  '-name', $Name,
  '-out',  $Out
)
python @cmd
$exit = $LASTEXITCODE
if ($exit -ne 0) { throw "Le script Python a retourné le code $exit." }

Write-Host "`n[OK] Terminé. Résultats: $Out" -ForegroundColor Green
if (Test-Path -LiteralPath $Out) {
  Write-Host "Ouverture du CSV…" -ForegroundColor DarkGreen
  Invoke-Item -LiteralPath $Out
}

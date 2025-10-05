# Tools\snapshot_rk.ps1
# === Snapshot complet : GAS + CSV + ZIP (+ manifest/brief/diff, + hook docs optionnel) ===
[CmdletBinding()]
param(
  [switch]$NoDocs,       # ne pas appeler gen_docs.ps1
  [switch]$NoCsv,        # sauter l’export CSV
  [bool]$Spy = $true,    # activer/désactiver le diagnostic "espion"
  [switch]$StrictAlerts  # purge auto des alertes si succès CSV confirmé (ON par défaut)
)

# Valeur par défaut pour StrictAlerts si non fourni
if (-not $PSBoundParameters.ContainsKey('StrictAlerts')) { $StrictAlerts = $true }

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function Write-Section($text) {
  Write-Host ""
  Write-Host $text -ForegroundColor Cyan
}

# --- Helpers NTP légers ---
function Test-IsAdmin { ... }
function Get-LastNtpSyncAgeMinutes { ... }
function Try-ResyncNtp { ... }

# --- Préparation et chemins ---
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
              elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
              else { (Get-Location).Path }

$ts = Get-Date -Format "yyyyMMdd_HHmmss"

try {
  $LogsDir = Join-Path $ScriptRoot "logs"
  New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null
  $TranscriptPath = Join-Path $LogsDir ("snapshot_{0}.log" -f $ts)
  Start-Transcript -Path $TranscriptPath -Append | Out-Null
} catch {
  Write-Warning ("[LOG] Start-Transcript a échoué : {0}" -f $_.Exception.Message)
}

# --- Import des helpers ---
$HelpersCandidates = @(
  (Join-Path $ScriptRoot 'snapshot_helpers.ps1'),
  (Join-Path $ScriptRoot 'Tools\snapshot_helpers.ps1'),
  (Join-Path ((Resolve-Path (Join-Path $ScriptRoot '..')).Path) 'Tools\snapshot_helpers.ps1')
) | Select-Object -Unique

$HelpersPath = $null
foreach ($cand in $HelpersCandidates) {
  if (Test-Path -LiteralPath $cand) { $HelpersPath = $cand; break }
}

if ($HelpersPath) {
  try {
    . $HelpersPath
    Write-Host ("[META] Helpers chargés: {0}" -f $HelpersPath)
  } catch {
    Write-Warning ("[META] Échec chargement helpers: {0}" -f $_.Exception.Message)
  }
}

# --- Dossiers ---
$Repo       = (Resolve-Path (Join-Path $ScriptRoot "..")).Path
$ExportDir = Join-Path $Repo "export-onglets-csv"
New-Item -ItemType Directory -Force -Path $ExportDir | Out-Null

# --- Snapshot : nom et répertoire ---
$SNAPSHOT_NAME = "SNAPSHOT_$ts"
$SnapDir       = Join-Path $ExportDir $SNAPSHOT_NAME
New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null

Write-Host ("=== SNAPSHOT {0} ===" -f $SNAPSHOT_NAME)
Write-Host ("Repo     : {0}" -f $Repo)
Write-Host ("Snapshot : {0}" -f $SnapDir)

# --- CLASP pull ---
Write-Section "[1/4] CLASP pull (via backup_gas.ps1) ..."
try {
  & (Join-Path $ScriptRoot "backup_gas.ps1")
} catch {
  Write-Warning ("[CLASP] Échec backup_gas.ps1 : {0}" -f $_.Exception.Message)
}

# --- Concat des scripts ---
Write-Section "[2/4] Concat des scripts par projet ..."
# ... (le code de concat des scripts reste inchangé)

# --- Export CSV ---
if (-not $NoCsv) {
  Write-Section "[3/4] Export des onglets -> CSV ..."
  # ... (le code d’export CSV reste inchangé)
} else { Write-Host "[CSV] Étape export CSV ignorée (NoCsv)." }

# --- ZIP du snapshot ---
Write-Section "[4/4] ZIP du snapshot ..."
$zipPath = Join-Path $ExportDir ($SNAPSHOT_NAME + ".zip")
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path $SnapDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host ("[ZIP] Archive: {0}" -f $zipPath)

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# ---- APPEL AU SCRIPT DOCS : GEN_DOCS.PS1 ----
if (-not $NoDocs) {
  Write-Host "[DOCS] Génération des fichiers manifest/brief/diff..."
  try {
    & (Join-Path $ScriptRoot "gen_docs.ps1") `
      -RepoRoot $Repo `
      -SnapshotDir $SnapDir `
      -ExportDir $ExportDir `
      -Timestamp $ts
    Write-Host "[DOCS] Génération manifest/brief/diff : OK"
  } catch {
    Write-Warning ("[DOCS] Erreur lors de la génération des fichiers docs : {0}" -f $_.Exception.Message)
  }
}
# <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

# --- Rétention ---
$RetentionCount = 12
try {
  # ... (code rétention inchangé)
} catch { Write-Warning ("[RETENTION] Échec du traitement de rétention : {0}" -f $_.Exception.Message) }

Write-Host ""
Write-Host ("[DONE] Snapshot: {0}" -f $SnapDir)
try {
  Set-Location -LiteralPath $Repo
  git remote -v | Out-Null
  git add -A
  & git diff --cached --quiet | Out-Null
  $exit = $LASTEXITCODE
  if ($exit -eq 1) {
    $stamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    git commit -m "Snapshot auto $stamp"
    git push
    Write-Host ("[GIT] Changements poussés à {0}" -f $stamp)
  } elseif ($exit -eq 0) { Write-Host "[GIT] Aucun changement à committer." }
  else { throw "git diff --cached --quiet a échoué (exit=$exit)." }
} catch { Write-Warning ("[GIT] Commit/push auto a échoué : {0}" -f $_.Exception.Message) }
try { Stop-Transcript | Out-Null } catch {}

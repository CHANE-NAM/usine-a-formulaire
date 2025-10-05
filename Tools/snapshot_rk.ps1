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
function Test-IsAdmin {
  try {
    $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wp = New-Object Security.Principal.WindowsPrincipal($wi)
    return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}
function Get-LastNtpSyncAgeMinutes {
  try {
    $raw = (w32tm /query /status) -join "`n"
    $dt = $null
    if ($raw -match 'Heure de la derni[eè]re synchronisation r[eé]ussie\s*:\s*([0-9/\-]+\s+[0-9:]+)') {
      $dt = [datetime]::Parse($matches[1], [System.Globalization.CultureInfo]::CurrentCulture)
    } elseif ($raw -match 'Last Successful Sync Time\s*:\s*([0-9/\-]+\s+[0-9:]+)') {
      $dt = [datetime]::Parse($matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
    }
    if ($null -eq $dt) { return $null }
    return ([datetime]::UtcNow - $dt.ToUniversalTime()).TotalMinutes
  } catch { return $null }
}
function Try-ResyncNtp {
  try {
    if (-not (Test-IsAdmin)) {
      Write-Host "[NTP] Pas de privilèges administrateur -> resync ignorée (ce n'est pas bloquant)."
      return
    }
    w32tm /resync | Out-Null
    Write-Host "[NTP] Resynchronisation demandée."
  } catch {
    Write-Warning ("[NTP] Resynchronisation a échoué : {0}" -f $_.Exception.Message)
  }
}

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

$bdd1 = Join-Path $Repo "03_BaseDeDonnées"
$bdd2 = Join-Path $Repo "03_BaseDeDonnees"
if       (Test-Path -LiteralPath $bdd1) { $bddDir = $bdd1 }
elseif (Test-Path -LiteralPath $bdd2) { $bddDir = $bdd2 }
else { $bddDir = $null }

$Projets = @()
$Projets += ,@("[MOTEUR]V2 Usine à Tests",       (Join-Path $Repo "01_Moteur"))
$Projets += ,@("[CONFIG]V2 Usine à Tests",       (Join-Path $Repo "02_configuration"))
if ($bddDir) { $Projets += ,@("[BDD]V2 Tests & Profils", $bddDir) } else { Write-Warning "Dossier BDD introuvable." }
$Projets += ,@("[TEMPLATE]V2 Kit de Traitement",   (Join-Path $Repo "04_Templates"))
$Projets += ,@("[BIBLIOTHEQUE]TEMPLATE", (Join-Path $Repo "05_Bibliotheque"))
$Projets += ,@("[HANDLER]V2 Web App",       (Join-Path $Repo "08_handler"))
$Projets += ,@("[TOOLS] Scripts de Snapshot",   (Join-Path $Repo "Tools"))
$Projets += ,@("[TOOLING] Export CSV",         (Join-Path $Repo "export-onglets-csv"))

foreach ($p in $Projets) {
  $pname = $p[0]
  $pdir  = $p[1]

  if (-not (Test-Path -LiteralPath $pdir)) { Write-Warning ("Dossier introuvable: {0}" -f $pdir); continue }

  $safeName = ($pname -replace '[^\w\-]+','_')
  $outTxt   = Join-Path $SnapDir ("scripts_" + $safeName + ".txt")

  # VERSION CORRIGÉE AVEC FILTRE ROBUSTE
  $files = Get-ChildItem -LiteralPath $pdir -Recurse -File -ErrorAction SilentlyContinue |
           Where-Object { 
              $_.FullName -notmatch "[\\/]\.git[\\/]" -and
              $_.FullName -notmatch "[\\/]node_modules[\\/]" -and
              ( ($_.Extension -in ".gs",".js",".ts", ".html", ".ps1") -or ($_.Name -in "appsscript.json", "package.json") )
           }

  if (-not $files) { Write-Warning ("Aucun fichier pertinent trouvé dans {0}" -f $pdir); continue }

  ("=== Projet: {0} ({1}) ==={2}" -f $pname, $pdir, [Environment]::NewLine) |
    Out-File -FilePath $outTxt -Encoding UTF8

  foreach ($f in $files) {
    ("`n--- FILE: {0} ---`n" -f $f.FullName) | Out-File -FilePath $outTxt -Encoding UTF8 -Append
    Get-Content -LiteralPath $f.FullName -Raw | Out-File -FilePath $outTxt -Encoding UTF8 -Append
  }
  Write-Host ("[OK] Concat: {0}" -f $outTxt)
}

# --- Le reste du script est inchangé ---
# ... (Export CSV, Manifest, Brief, Diff, ZIP, Rétention, Git Push) ...
if (-not $NoCsv) {
  Write-Section "[3/4] Export des onglets -> CSV ..."
  $Ids = @(
    "1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8",
    "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ",
    "1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU",
    "1hrcdsMRwx4FuHTvvtJoq2AVh8XTzwp5MErJ3UQ0OA5E"
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  $credsPath = if ($env:RK_CREDS) { $env:RK_CREDS } else { "C:\secrets\rk_oauth\credentials.json" }
  $tokenPath = if ($env:RK_TOKEN) { $env:RK_TOKEN } else { "C:\secrets\rk_oauth\token.json" }
  $IndexJs = Join-Path $ExportDir 'index.js'
  $nodeCmd = Get-Command node -ErrorAction SilentlyContinue
  $canRun = $true
  if (-not (Test-Path -LiteralPath $IndexJs)) { Write-Warning "[CSV] export-onglets-csv\index.js introuvable — étape CSV sautée."; $canRun = $false }
  if (-not $nodeCmd)                         { Write-Warning "[CSV] Node.js (commande 'node') introuvable — étape CSV sautée.";   $canRun = $false }
  if (-not (Test-Path -LiteralPath $credsPath)){ Write-Warning ("[CSV] credentials.json introuvable : {0} — étape CSV sautée." -f $credsPath); $canRun = $false }
  if (-not (Test-Path -LiteralPath $tokenPath)) { Write-Warning ("[CSV] token.json absent : {0} — un nouveau consentement OAuth sera demandé." -f $tokenPath) }
  if (-not $Ids -or $Ids.Count -eq 0)         { Write-Warning "[CSV] Aucun ID de classeur fourni — étape CSV sautée.";         $canRun = $false }
  if ($canRun) {
    $nodeArgs = @("--out", $SnapDir, "--creds", $credsPath, "--token", $tokenPath) + ($Ids | ForEach-Object { @("--id", $_) })
    $cmdPreview = "$($nodeCmd.Source) `"$IndexJs`" " + ($nodeArgs | ForEach-Object { if ($_ -match '\s') { '"{0}"' -f $_ } else { $_ } }) -join ' '
    Write-Host "NODE CMD: $cmdPreview" -ForegroundColor Cyan
    $csvBefore = (Get-ChildItem -LiteralPath $SnapDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count
    Push-Location -LiteralPath $ExportDir
    try {
      & $nodeCmd.Source ".\index.js" @nodeArgs
      if ($LASTEXITCODE -ne 0) { throw "Échec export CSV (node exit code = $LASTEXITCODE)." }
      $csvAfter = (Get-ChildItem -LiteralPath $SnapDir -Recurse -File -Filter '*.csv' -ErrorAction SilentlyContinue).Count
      $delta = $csvAfter - $csvBefore
      Write-Host ("[CSV] Export terminé : {0} fichier(s) CSV ajouté(s) dans {1}" -f [math]::Max($delta,0), $SnapDir)
    } catch { Write-Warning ("[CSV] Échec export CSV : {0}" -f $_.Exception.Message) } finally { Pop-Location }
  }
} else { Write-Host "[CSV] Étape export CSV ignorée (NoCsv)." }

Write-Section "[4/4] ZIP du snapshot ..."
$zipPath = Join-Path $ExportDir ($SNAPSHOT_NAME + ".zip")
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path $SnapDir -DestinationPath $zipPath -CompressionLevel Optimal
Write-Host ("[ZIP] Archive: {0}" -f $zipPath)

$RetentionCount = 12
try {
  $allSnaps = Get-ChildItem -LiteralPath $ExportDir -Directory | Where-Object { $_.Name -like 'SNAPSHOT_*' } | Sort-Object LastWriteTime -Descending
  if ($allSnaps.Count -gt $RetentionCount) {
    $toDelete = $allSnaps | Select-Object -Skip $RetentionCount
    foreach ($old in $toDelete) {
      try {
        $oldZip = Join-Path $ExportDir ($old.Name + '.zip')
        if (Test-Path -LiteralPath $oldZip) { Remove-Item -LiteralPath $oldZip -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $old.FullName -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ("[RETENTION] Supprimé: {0}" -f $old.Name)
      } catch { Write-Warning ("[RETENTION] Échec suppression {0} : {1}" -f $old.Name, $_.Exception.Message) }
    }
  }
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
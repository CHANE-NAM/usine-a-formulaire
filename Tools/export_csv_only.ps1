# Tools/export_csv_only.ps1
param(
  [string[]]$Ids = @(
    "1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8",
    "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ",
    "1XwyTt9hcFLd-_IrCYuKY4_E6Dw9aUrls-AGQp65dzDU",
    "1hrcdsMRwx4FuHTvvtJoq2AVh8XTzwp5MErJ3UQ0OA5E"
  ),
  [int]$MaxRetry = 3
)

$ErrorActionPreference = "Stop"
$Repo      = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ExportDir = Join-Path $Repo "export-onglets-csv"

# Trouve le dernier snapshot; sinon en crée un nouveau
$snap = Get-ChildItem -Path $ExportDir -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $snap) {
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  $snapName = "SNAPSHOT_$ts"
  $SnapDir = Join-Path $ExportDir $snapName
  New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null
} else {
  $SnapDir = $snap.FullName
}
Write-Host ("Export -> {0}" -f $SnapDir)

$nodeArgsBase = @("--out", $SnapDir)

foreach ($id in $Ids) {
  $ok = $false
  for ($i=1; $i -le $MaxRetry -and -not $ok; $i++) {
    Write-Host ("[{0}] Tentative {1}/{2}..." -f $id, $i, $MaxRetry)
    Push-Location $ExportDir
    try {
      node ".\index.js" @($nodeArgsBase + @("--id", $id))
      $ok = $true
      Write-Host ("[{0}] OK" -f $id)
    } catch {
      Write-Warning ("[{0}] Echec tentative {1}: {2}" -f $id, $i, $_.Exception.Message)
      Start-Sleep -Seconds ([Math]::Min(30, 2 * $i))
    } finally {
      Pop-Location
    }
  }
  if (-not $ok) { Write-Warning ("[{0}] ABANDON apres {1} essais." -f $id, $MaxRetry) }
}

Write-Host "Termine."



param(
  [string]$SnapshotDir
)

$ErrorActionPreference = "Stop"
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$Tools = Join-Path $ScriptRoot "Tools"
$Export = Join-Path $ScriptRoot "export-onglets-csv"

if (-not (Test-Path -LiteralPath (Join-Path $Tools "gen_docs.ps1"))) {
  throw "Tools\gen_docs.ps1 introuvable."
}

if (-not $SnapshotDir) {
  if (-not (Test-Path -LiteralPath $Export)) { throw "Dossier export-onglets-csv introuvable." }
  $latest = Get-ChildItem -LiteralPath $Export -Directory |
            Where-Object { $_.Name -like 'SNAPSHOT_*' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $latest) { throw "Aucun snapshot trouvé dans $Export." }
  $SnapshotDir = $latest.FullName
}

Write-Host ("[INFO] Génération de doc sur : {0}" -f $SnapshotDir)
& (Join-Path $Tools "gen_docs.ps1") -RepoRoot $ScriptRoot -SnapshotDir $SnapshotDir -ExportDir $Export

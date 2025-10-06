# gen_synthese.ps1 — Génère la synthèse markdown à partir des fichiers snapshot
param(
  [string]$SnapshotDir = ""
)

# Si aucun dossier snapshot fourni, prend le dernier trouvé
if (-not $SnapshotDir -or -not (Test-Path -LiteralPath $SnapshotDir)) {
  $ExportDir = Join-Path $PSScriptRoot "..\export-onglets-csv"
  $lastSnap = Get-ChildItem -LiteralPath $ExportDir -Directory | Where-Object { $_.Name -like 'SNAPSHOT_*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $lastSnap) { Write-Error "Aucun snapshot trouvé !"; exit 1 }
  $SnapshotDir = $lastSnap.FullName
}

# Charge le modèle de prompt (modèle markdown)
$templatePath = Join-Path $PSScriptRoot "modele_synthese.md"
if (-not (Test-Path $templatePath)) { Write-Error "Modèle modele_synthese.md introuvable !"; exit 1 }
$template = Get-Content $templatePath -Raw

# Charge les fichiers du snapshot
$briefPath = Join-Path $SnapshotDir "brief.md"
$manifestPath = Join-Path $SnapshotDir "manifest.json"
$diffPath = Join-Path $SnapshotDir "diff.md"

# Insert sections
function Get-SectionOrEmpty($path) {
  if (Test-Path $path) { return Get-Content $path -Raw }
  else { return "_Section absente_" }
}
# Variables : à enrichir selon ton workflow (par exemple en extrayant depuis le brief, manifest ou un CSV spécifique)
$variables = "_À compléter selon le projet_"

# Compose la synthèse
$out = $template `
    -replace "{{date}}", (Get-Date -Format "yyyy-MM-dd HH:mm") `
    -replace "{{section_brief}}", (Get-SectionOrEmpty $briefPath) `
    -replace "{{section_manifest}}", (Get-SectionOrEmpty $manifestPath) `
    -replace "{{section_diff}}", (Get-SectionOrEmpty $diffPath) `
    -replace "{{section_variables}}", $variables

# Chemin du dossier Syntheses à la racine du projet
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$SyntheseDir = Join-Path $ProjectRoot "Syntheses"
if (-not (Test-Path $SyntheseDir)) { New-Item -ItemType Directory -Path $SyntheseDir | Out-Null }

# Nom de la synthèse (date + heure)
$syntheseName = "synthese_session_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".md"
$outPath = Join-Path $SyntheseDir $syntheseName

# Écrit le fichier final
Set-Content -Path $outPath -Value $out -Encoding UTF8

Write-Host "✅ Synthèse générée ici : $outPath"
# Optionnel : Copie dans le presse-papiers
# $out | Set-Clipboard

exit 0

# Build-IA-Kit.ps1 — prépare un kit minimal pour l'IA
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Racine projet
$Repo = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
$Exp  = Join-Path $Repo 'export-onglets-csv'
if (-not (Test-Path $Exp)) { throw "Dossier introuvable : $Exp" }

# Dernier snapshot
$Snap = Get-ChildItem $Exp -Directory | ? Name -like 'SNAPSHOT_*' | Sort LastWriteTime -Descending | Select -First 1
if (-not $Snap) { throw "Aucun snapshot trouvé dans $Exp" }

# Fichiers sources
$Brief = Join-Path $Snap.FullName 'docs\00_SESSION_BRIEF.md'
$Etat  = Join-Path $Snap.FullName 'docs\etat\etat_projet.md'
$Diff  = Join-Path $Snap.FullName 'diff.md'
$AiDir = Join-Path $Snap.FullName 'docs\ai'

# Dossier de sortie
$Kit = Join-Path $Snap.FullName 'IA_kit'
New-Item -ItemType Directory -Force -Path $Kit | Out-Null

# Copie BRIEF, état, diff (s’ils existent)
$toCopy = @()
if (Test-Path $Brief) { $toCopy += $Brief }
if (Test-Path $Etat)  { $toCopy += $Etat }
if (Test-Path $Diff)  { $toCopy += $Diff }
foreach ($f in $toCopy) { Copy-Item -LiteralPath $f -Destination $Kit -Force }

# Sélectionne les docs/ai pertinents en lisant diff.md
$WantedAi = New-Object System.Collections.Generic.HashSet[string]
if ((Test-Path $Diff) -and (Test-Path $AiDir)) {
  $d = Get-Content -LiteralPath $Diff -Raw -Encoding UTF8
  $lines = $d -split "`r?`n"
  foreach ($line in $lines) {
    # détecte des chemins qui pointent aux dossiers projet
    if ($line -match '01_Moteur'         ) { $null = $WantedAi.Add('[MOTEUR]V2_Usine_à_Tests.md') }
    if ($line -match '02_configuration'  ) { $null = $WantedAi.Add('[CONFIG]V2_Usine_à_Tests.md') }
    if ($line -match '03_BaseDeDonn(é|e)es') { $null = $WantedAi.Add('[BDD]V2_Tests_&_Profils.md') }
    if ($line -match '04_Templates'      ) { $null = $WantedAi.Add('[TEMPLATE]V2_Kit_de_Traitement.md') }
    # en plus : si le diff touche les fichiers "scripts_<nom>.txt"
    if ($line -match 'scripts_(.+?)\.txt') {
      $base = $Matches[1] + '.md'
      $null = $WantedAi.Add($base)
    }
  }
}

# Copie jusqu’à 3 docs AI maximum
$copiedAi = @()
if (Test-Path $AiDir) {
  $aiFiles = Get-ChildItem -LiteralPath $AiDir -File -Filter '*.md'
  if ($WantedAi.Count -eq 0) {
    # si on n'a rien détecté, prends le(s) 1–2 fichier(s) AI le(s) plus récent(s)
    $aiPick = $aiFiles | Sort LastWriteTime -Descending | Select -First 2
  } else {
    $aiPick = foreach ($name in $WantedAi) { $aiFiles | Where-Object { $_.Name -eq $name } }
    $aiPick = $aiPick | Where-Object { $_ } | Select-Object -First 3
  }
  foreach ($f in $aiPick) {
    Copy-Item -LiteralPath $f.FullName -Destination $Kit -Force
    $copiedAi += $f.Name
  }
}

Write-Host "[IA-KIT] Dossier prêt : $Kit"
Write-Host "[IA-KIT] Inclus :"
foreach ($b in (Get-ChildItem $Kit -File | Select -ExpandProperty Name)) { Write-Host " - $b" }

# Ouvre le dossier et met le BRIEF en presse-papiers si présent
if (Test-Path $Brief) {
  $txt = Get-Content -Raw -LiteralPath $Brief -Encoding UTF8
  $txt | Set-Clipboard
  Write-Host "[IA-KIT] BRIEF copié dans le presse-papiers."
}
Invoke-Item $Kit

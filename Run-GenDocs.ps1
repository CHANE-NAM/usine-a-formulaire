# Run-GenDocs.ps1 — lance Tools\gen_docs.ps1 sur le dernier snapshot, en -Command (params nommés)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Racine fiable
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }

$Repo   = [string]$ScriptRoot
$ExpDir = Join-Path $Repo 'export-onglets-csv'
if (-not (Test-Path -LiteralPath $ExpDir)) { throw "Dossier manquant: $ExpDir" }

$LatestSnap = (Get-ChildItem -LiteralPath $ExpDir -Directory |
               Where-Object { $_.Name -like 'SNAPSHOT_*' } |
               Sort-Object LastWriteTime -Descending |
               Select-Object -First 1).FullName
if (-not $LatestSnap) { throw "Aucun snapshot trouvé dans $ExpDir" }

$Gen = Join-Path $Repo 'Tools\gen_docs.ps1'
if (-not (Test-Path -LiteralPath $Gen)) { throw "Tools\gen_docs.ps1 introuvable: $Gen" }

Write-Host "[WRAP] Repo=$Repo"
Write-Host "[WRAP] Snap=$LatestSnap"
Write-Host "[WRAP] Exp =$ExpDir"
Write-Host "[WRAP] Script=$Gen"

# Construit une commande -Command **nommée** et correctement échappée (espaces gérés)
$cmd = "& `"$Gen`" -RepoRoot `"$Repo`" -SnapshotDir `"$LatestSnap`" -ExportDir `"$ExpDir`""

# Lance d'abord PowerShell 7 si présent, sinon Windows PowerShell
$pwsh7 = Join-Path $Env:ProgramFiles 'PowerShell\7\pwsh.exe'
if (Test-Path -LiteralPath $pwsh7) {
  Write-Host "[WRAP] pwsh.exe -Command ... (nommé)"
  & $pwsh7 -NoLogo -ExecutionPolicy Bypass -Command $cmd
} else {
  Write-Host "[WRAP] powershell.exe -Command ... (nommé)"
  powershell -NoLogo -ExecutionPolicy Bypass -Command $cmd
}

Write-Host "[WRAP] Terminé."

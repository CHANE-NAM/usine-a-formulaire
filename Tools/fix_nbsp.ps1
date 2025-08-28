# Tools/fix_nbsp.ps1
$path = Join-Path $PSScriptRoot 'snapshot_rk.ps1'
if (-not (Test-Path $path)) { Write-Error "Introuvable: $path"; exit 1 }

# Lire, remplacer NBSP (0xA0) -> ' ' et ZWSP (0x200B) -> ''
$content = Get-Content -LiteralPath $path -Raw
$content = $content -replace "\xA0"," " -replace "\x200B",""

# Réécrire avec encodage UTF-8 BOM (évite les artefacts en console)
$utf8bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($path, $content, $utf8bom)

Write-Host "Nettoyage OK: $path"

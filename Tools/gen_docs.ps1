param(
  [Parameter(Mandatory=$true)] [string]$RepoRoot,
  [Parameter(Mandatory=$true)] [string]$SnapshotDir,
  [Parameter(Mandatory=$true)] [string]$ExportDir,
  [Parameter(Mandatory=$false)] [string]$Timestamp
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function New-Dir([string]$p) { New-Item -ItemType Directory -Force -Path $p | Out-Null }

function Write-FileUtf8([string]$Path,[string]$Content) {
  $enc = New-Object System.Text.UTF8Encoding($false) # sans BOM
  [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Guess-CodeFence([string]$filePath) {
  $ext = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
  switch ($ext) {
    ".js"   { "javascript" }
    ".gs"   { "javascript" }
    ".ts"   { "typescript" }
    ".html" { "html" }
    ".json" { "json" }
    default { "" }
  }
}

# Dossiers de sortie dans le snapshot
$DocsRoot = Join-Path $SnapshotDir "docs"
$DirAI    = Join-Path $DocsRoot "ai"
$DirEtat  = Join-Path $DocsRoot "etat"
$DirUsers = Join-Path $DocsRoot "users"
New-Dir $DocsRoot; New-Dir $DirAI; New-Dir $DirEtat; New-Dir $DirUsers

# -----------------------------------------------------------------------------
# 1) AI-FRIENDLY : reconstruire des markdowns lisibles par IA à partir des scripts__*.txt
# -----------------------------------------------------------------------------
$concatFiles = Get-ChildItem -LiteralPath $SnapshotDir -File -Filter "scripts__*.txt" -ErrorAction SilentlyContinue
foreach ($cf in $concatFiles) {
  $projectName = [System.IO.Path]::GetFileNameWithoutExtension($cf.Name) -replace '^scripts__',''
  $outPath = Join-Path $DirAI ($projectName + ".md")

  $lines = Get-Content -LiteralPath $cf.FullName -Raw -Encoding UTF8
  # Split sur les délimiteurs; on garde les noms de fichiers capturés
  $parts = [System.Text.RegularExpressions.Regex]::Split($lines, "(?m)^\s*---\s*FILE:\s*(.+?)\s*---\s*$")

  $md = @"
# $projectName

> Généré automatiquement depuis **$($cf.Name)** — snapshot: **$([System.IO.Path]::GetFileName($SnapshotDir))**.

"@

  # Ensuite par paires (filePath, content)
  for ($i = 1; $i -lt $parts.Count; $i += 2) {
    $filePath = $parts[$i].Trim()
    $content  = if ($i + 1 -lt $parts.Count) { $parts[$i+1] } else { "" }
    $lang = Guess-CodeFence $filePath
@"
## $filePath

```$lang
$content

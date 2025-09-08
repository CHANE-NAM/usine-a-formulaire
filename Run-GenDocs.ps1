param(
  [string]$SnapshotDir
)

$ErrorActionPreference = "Stop"

# Racine repo fiable
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path }
else { (Get-Location).Path }

$Tools   = Join-Path $ScriptRoot "Tools"
$Export  = Join-Path $ScriptRoot "export-onglets-csv"
$Gen     = Join-Path $Tools "gen_docs.ps1"

if (-not (Test-Path -LiteralPath $Gen))    { throw "Tools\gen_docs.ps1 introuvable: $Gen" }
if (-not (Test-Path -LiteralPath $Export)) { throw "Dossier export-onglets-csv introuvable: $Export" }

# Si aucun snapshot passé, prendre le plus récent
if (-not $SnapshotDir) {
  $latest = Get-ChildItem -LiteralPath $Export -Directory |
            Where-Object { $_.Name -like 'SNAPSHOT_*' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $latest) { throw "Aucun snapshot trouvé dans $Export." }
  $SnapshotDir = $latest.FullName
}

Write-Host ("[INFO] Génération de doc sur : {0}" -f $SnapshotDir)

# 1) Tentative directe (splatting)
try {
  $parms = @{
    RepoRoot    = $ScriptRoot
    SnapshotDir = $SnapshotDir
    ExportDir   = $Export
  }
  & $Gen @parms
  return
} catch {
  Write-Warning ("[WRAP] Appel direct a échoué ({0}) — fallback powershell.exe -File..." -f $_.Exception.Message)
}

# 2) Fallback universel via powershell.exe -File (compatible 5.1)
function Resolve-PowerShellExe {
  try {
    $cmd = Get-Command powershell -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { return $cmd.Source }
  } catch {}
  $default = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
  return $default
}
function Q([string]$s){ '"' + $s + '"' }

$psExe = Resolve-PowerShellExe

$argList = @(
  "-NoLogo", "-ExecutionPolicy", "Bypass",
  "-File",       (Q $Gen),
  "-RepoRoot",   (Q $ScriptRoot),
  "-SnapshotDir",(Q $SnapshotDir),
  "-ExportDir",  (Q $Export)
)

$proc = Start-Process -FilePath $psExe -ArgumentList $argList -Wait -PassThru
if ($proc.ExitCode -ne 0) {
  throw "gen_docs.ps1 a échoué (exit $($proc.ExitCode))."
}

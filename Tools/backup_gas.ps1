# backup_gas.ps1 — Sauvegarde Apps Script -> GitHub
# 1) fait un clasp pull dans chaque sous-dossier contenant .clasp.json
# 2) commit/push uniquement s'il y a des changements
# 3) journalise la sortie dans tools\logs

$ErrorActionPreference = "Stop"

# === À ADAPTER si besoin ===
# Chemins dynamiques basés sur l’emplacement du script
$RepoPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path   # remonte d’un dossier depuis Tools/
$LogsDir  = Join-Path $PSScriptRoot "logs"                        # logs dans Tools\logs

# ===========================

# S'assure que les dossiers existent
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

# Ajoute les chemins usuels au PATH pour l'exécution planifiée
# (npm global pour clasp, git.exe)
$env:Path = "$env:AppData\npm;C:\Program Files\Git\cmd;$env:Path"

# Démarre un log horodaté
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $LogsDir "backup_$ts.log"
Start-Transcript -Path $logFile -Force | Out-Null

try {
    Write-Host "=== Backup GAS -> GitHub ($ts) ==="

    # Vérifs outils
    git --version | Out-Host
    clasp --version | Out-Host

    # Va à la racine du dépôt
    Set-Location $RepoPath

    # 1) CLASP PULL dans tous les sous-projets détectés
    $projects = Get-ChildItem -Path $RepoPath -Recurse -Filter ".clasp.json" -File -Force |
                Select-Object -ExpandProperty DirectoryName -Unique
    if (-not $projects) {
        Write-Host "Aucun .clasp.json trouvé. (Ignorable si tout est local)"
    } else {
        foreach ($dir in $projects) {
            Write-Host "`n--- clasp pull : $dir ---"
            Push-Location $dir
            try {
                clasp pull | Out-Host
            } catch {
                Write-Warning "clasp pull a échoué dans $dir : $($_.Exception.Message)"
            } finally {
                Pop-Location
            }
        }
    }

    # 2) Commit/push uniquement s'il y a des modifs
    Set-Location $RepoPath
    Write-Host "RepoPath = $RepoPath"

    $changes = git status --porcelain
    if ($changes) {
        Write-Host "`nDes changements détectés -> commit & push"
        git add -A | Out-Host
        $msg = "Backup auto $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        git commit -m $msg | Out-Host
        git push | Out-Host
    } else {
        Write-Host "`nAucun changement à sauvegarder."
    }

    # Nettoyage des logs > 14 jours (optionnel)
    Get-ChildItem $LogsDir -Filter "backup_*.log" |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-14) } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    Write-Host "`n=== Terminé avec succès ==="
}
catch {
    Write-Error "Échec du backup : $($_.Exception.Message)"
}
finally {
    Stop-Transcript | Out-Null
}

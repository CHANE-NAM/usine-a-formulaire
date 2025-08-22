# =================================================================================================
# 🔧 Sauvegarde Apps Script -> GitHub (backup_gas.ps1)
# -------------------------------------------------------------------------------------------------
# ▶️ Exécution manuelle (depuis la racine du dépôt) :
#    powershell -ExecutionPolicy Bypass -File ".\Tools\backup_gas.ps1"
#
# ⏱️ Automatisation (tâche quotidienne à 22:00 via Invite de commandes/PowerShell) :
#    schtasks /Create /SC DAILY /ST 22:00 /TN "Backup GAS vers GitHub" /TR ^
#     "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\Tools\backup_gas.ps1""" /F
#  - Pour lancer la tâche tout de suite :  schtasks /Run /TN "Backup GAS vers GitHub"
#  - Pour la supprimer :                   schtasks /Delete /TN "Backup GAS vers GitHub" /F
#
# (Alternative GUI : Planificateur de tâches → Créer une tâche de base → Quotidien → Action: powershell.exe
#  Arguments : -NoProfile -ExecutionPolicy Bypass -File "CHEMIN\Tools\backup_gas.ps1")
# =================================================================================================

# --- Confort d'affichage (évite les caractères bizarres dans la console) ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ErrorActionPreference = "Stop"

# === Chemins dynamiques basés sur l’emplacement du script ===================
# (Tools\backup_gas.ps1 -> remonte d’un dossier vers la racine du repo)
$RepoPath = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$LogsDir  = Join-Path $PSScriptRoot "logs"

# === Prépare PATH pour l'exécution planifiée (Git + Node/npm pour clasp) ====
# (Fix: construire chaque chemin séparément pour éviter l'erreur Join-Path)
$npmPath      = if ($env:APPDATA)      { Join-Path -Path $env:APPDATA      -ChildPath 'npm' }                   else { $null }
$nodeGlobal   =                         'C:\Program Files\nodejs'
$gitCmdSystem =                         'C:\Program Files\Git\cmd'
$gitCmdUser   = if ($env:LOCALAPPDATA) { Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Programs\Git\cmd' }      else { $null }

$pathsToTry = @($npmPath, $nodeGlobal, $gitCmdSystem, $gitCmdUser) | Where-Object { $_ -and (Test-Path $_) }
if ($pathsToTry.Count -gt 0) { $env:Path = ($pathsToTry -join ';') + ';' + $env:Path }

# Optionnel : autoriser les chemins longs Git sous Windows
try { git config --global core.longpaths true | Out-Null } catch {}

# === Logs ===================================================================
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $LogsDir "backup_$ts.log"
Start-Transcript -Path $logFile -Force | Out-Null

try {
    Write-Host "=== Backup GAS -> GitHub ($ts) ==="

    # Vérif des outils
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git introuvable dans le PATH. Installe Git ou ajuste PATH."
    }
    $hasClasp = $true
    if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
        $hasClasp = $false
        Write-Warning "clasp introuvable. Étape Apps Script ignorée. (Installer: npm i -g @google/clasp ; puis clasp login)"
    }

    # Aller à la racine du dépôt
    Set-Location $RepoPath
    if (-not (Test-Path (Join-Path $RepoPath ".git"))) {
        # Au cas où Tools/ serait déplacé, tente de retrouver la racine via Git
        $gitRoot = (& git rev-parse --show-toplevel) 2>$null
        if ($LASTEXITCODE -eq 0 -and $gitRoot) {
            $RepoPath = $gitRoot.Trim()
            Set-Location $RepoPath
        } else {
            throw "Le dossier '$RepoPath' n'est pas un dépôt Git (.git introuvable)."
        }
    }
    Write-Host "RepoPath = $RepoPath"

    # 1) CLASP PULL dans tous les sous-projets détectés
    if ($hasClasp) {
        $projects = Get-ChildItem -Path $RepoPath -Recurse -Force -File -Filter ".clasp.json" |
                    Select-Object -ExpandProperty DirectoryName -Unique
        if (-not $projects) {
            Write-Host "Aucun .clasp.json trouvé. (Ignorable si tout est local)"
        } else {
            foreach ($dir in $projects) {
                Write-Host "`n--- clasp pull : $dir ---"
                Push-Location $dir
                try {
                    # (Option) Lancer une fonction Apps Script avant/à la place du pull :
                    # clasp run exportAllSheetsToCsv | Out-Host
                    clasp pull | Out-Host
                } catch {
                    Write-Warning "clasp pull a échoué dans $dir : $($_.Exception.Message)"
                } finally {
                    Pop-Location
                }
            }
        }
    }

    # 2) Commit/push uniquement s'il y a des modifs
    Set-Location $RepoPath
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
    Get-ChildItem $LogsDir -Filter "backup_*.log" -ErrorAction SilentlyContinue |
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

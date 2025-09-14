# =================================================================================
# FinaliserKit.ps1 — Rename Google Sheet à partir de la CONFIG (ligne 16)
# - Lit le nom dans l’onglet "Paramètres Généraux", colonne "Nom_Fichier_Complet"
# - Renomme le classeur cible (ID en dur) : d’abord via Sheets (title), fallback Drive (name)
# - Auth : priorise gcloud ADC (scopes cloud-platform, spreadsheets, drive), fallback sur clasp
# - Ajoute X-Goog-User-Project pour éviter les 403 "quota project required"
# =================================================================================

# ---- CONSTANTES (en dur) ----
[string]$ConfigSheetId        = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ"
[int]   $RowIndex             = 16
[string]$HeaderName           = "Nom_Fichier_Complet"
[string]$SheetTabName         = "Paramètres Généraux"
[string]$TargetSpreadsheetId  = "1t_0aDs6Kv1ZF-Upn4Zsn9BNoFlKrUM2FXo6ZaisrXOE"  # <= EN DUR

Write-Host -ForegroundColor Cyan "Étape: Renommage à partir de la CONFIG (ligne $RowIndex / '$HeaderName')…"
Write-Host -ForegroundColor Yellow "Classeur cible : $TargetSpreadsheetId"

# ---- Auth : gcloud ADC d'abord, sinon clasp ----
function Get-AccessToken {
  try {
    $tok = (& gcloud auth application-default print-access-token 2>$null).Trim()
    if ($tok) { return $tok }
  } catch {}
  $credPath = Join-Path $env:USERPROFILE ".clasprc.json"
  if (Test-Path -LiteralPath $credPath) {
    $cred = Get-Content -LiteralPath $credPath -Raw | ConvertFrom-Json
    if ($cred.token.access_token) { return $cred.token.access_token }
  }
  throw "Aucun access_token. Lance : gcloud auth application-default login --scopes=""https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/drive"""
}

function Get-QuotaProject {
  try {
    $adcJson = (& gcloud auth application-default print-json 2>$null)
    if ($adcJson) {
      $adc = $adcJson | ConvertFrom-Json
      if ($adc.quota_project_id) { return $adc.quota_project_id }
    }
  } catch {}
  try {
    $p = (& gcloud config get-value project 2>$null).Trim()
    if ($p -and $p -ne "(unset)") { return $p }
  } catch {}
  return $null
}

function New-AuthHeaders {
  param([string]$accessToken)
  $h = @{ Authorization = "Bearer $accessToken" }
  $qp = Get-QuotaProject
  if ($qp) { $h["X-Goog-User-Project"] = $qp }
  return $h
}

# ---- Helpers HTTP ----
function UrlEncode([string]$s) { [System.Uri]::EscapeDataString($s) }
function GET($url)            { Invoke-RestMethod -Method Get    -Uri $url -Headers $script:authHeader }
function POST_JSON($url,$obj) { Invoke-RestMethod -Method Post   -Uri $url -Headers $script:authHeader -ContentType "application/json" -Body ($obj | ConvertTo-Json -Depth 8) }
function PATCH_JSON($url,$obj){ Invoke-RestMethod -Method Patch  -Uri $url -Headers $script:authHeader -ContentType "application/json" -Body ($obj | ConvertTo-Json -Depth 8) }

function Show-HttpError($err) {
  Write-Host -ForegroundColor Red $err.Exception.Message
  if ($err.Exception.Response -and $err.Exception.Response.GetResponseStream()) {
    try {
      $sr = New-Object System.IO.StreamReader($err.Exception.Response.GetResponseStream())
      $body = $sr.ReadToEnd()
      if ($body) { Write-Host -ForegroundColor Red $body }
    } catch {}
  }
}

# ---- Construire l’auth header (avec quota project) ----
$accessToken     = Get-AccessToken
$script:authHeader = New-AuthHeaders -accessToken $accessToken

try {
  # --- DIAG : afficher l’email si possible ---
  try {
    $me = GET "https://www.googleapis.com/oauth2/v2/userinfo"
    if ($me.email) { Write-Host -ForegroundColor Yellow "Compte Google utilisé : $($me.email)" }
  } catch {}

  # --- Drive : confirmer l’objet et son type (avec quota project) ---
  $driveMeta = $null
  try {
    $driveMeta = GET ("https://www.googleapis.com/drive/v3/files/{0}?fields=id,name,mimeType,owners(emailAddress,displayName)" -f $TargetSpreadsheetId)
    $owner = $driveMeta.owners[0].emailAddress
    Write-Host -ForegroundColor Yellow ("Drive OK → Nom actuel: ""{0}"" | Type: {1} | Propriétaire: {2}" -f $driveMeta.name, $driveMeta.mimeType, $owner)
  } catch {
    Write-Host -ForegroundColor Red "❌ Drive metadata KO (ID invalide, API non activée ou pas d’accès). Détails :"
    Show-HttpError $_
    throw "Impossible de lire le fichier Drive cible."
  }

  # --- 1) Lire le NOM cible dans la CONFIG ---
  $rngHeaders = ("'{0}'!A1:Z1" -f $SheetTabName)
  $urlHeaders = "https://sheets.googleapis.com/v4/spreadsheets/$ConfigSheetId/values/" + (UrlEncode $rngHeaders)
  $headersResp = $null
  try {
    $headersResp = GET $urlHeaders
  } catch {
    Write-Host -ForegroundColor Red "❌ Lecture en-têtes CONFIG échouée."
    Show-HttpError $_
    throw "Impossible de lire les en-têtes de la CONFIG."
  }
  if (-not $headersResp.values -or $headersResp.values.Count -eq 0) { throw "CONFIG: ligne d’en-têtes vide." }
  $headers = $headersResp.values[0] | ForEach-Object { [string]$_ }

  $colIndex = $headers.IndexOf($HeaderName)
  if ($colIndex -lt 0) { throw "CONFIG: colonne '$HeaderName' introuvable." }

  $colA1    = [char]([int][char]'A' + $colIndex)  # A,B,C...
  $rngValue = ("'{0}'!{1}{2}:{1}{2}" -f $SheetTabName, $colA1, $RowIndex)
  $urlValue = "https://sheets.googleapis.com/v4/spreadsheets/$ConfigSheetId/values/" + (UrlEncode $rngValue)
  $valResp  = $null
  try {
    $valResp = GET $urlValue
  } catch {
    Write-Host -ForegroundColor Red "❌ Lecture valeur CONFIG échouée."
    Show-HttpError $_
    throw "Impossible de lire la valeur cible dans la CONFIG."
  }
  $newName  = ($valResp.values | ForEach-Object { $_[0] }) | Select-Object -First 1
  if ([string]::IsNullOrWhiteSpace($newName)) { throw "CONFIG: nom vide à la ligne $RowIndex (colonne '$HeaderName')." }
  Write-Host -ForegroundColor Yellow "Nouveau nom attendu : $newName"

  $renamed = $false
  $sheetsTriedError = $null

  # --- 2) Tenter via Sheets (si c’est bien un Google Spreadsheet) ---
  if ($driveMeta.mimeType -eq "application/vnd.google-apps.spreadsheet") {
    try {
      $batchUrl = "https://sheets.googleapis.com/v4/spreadsheets/$TargetSpreadsheetId:batchUpdate"
      $body = @{
        requests = @(
          @{
            updateSpreadsheetProperties = @{
              properties = @{ title = $newName }
              fields     = "title"
            }
          }
        )
      }
      POST_JSON $batchUrl $body | Out-Null
      Write-Host -ForegroundColor Green "✅ Renommage via Sheets OK → $newName"
      $renamed = $true
    } catch {
      Write-Host -ForegroundColor Yellow "⚠️ Échec via Sheets (on tentera Drive). Détails :"
      Show-HttpError $_
      $sheetsTriedError = $_
    }
  } else {
    Write-Host -ForegroundColor Yellow "Type non-Spreadsheet (mimeType=$($driveMeta.mimeType)) → on passera par Drive."
  }

  # --- 3) Fallback via Drive (name) ---
  if (-not $renamed) {
    try {
      $driveUrl = "https://www.googleapis.com/drive/v3/files/$TargetSpreadsheetId"
      PATCH_JSON $driveUrl @{ name = $newName } | Out-Null
      Write-Host -ForegroundColor Green "✅ Renommage via Drive OK → $newName"
      $renamed = $true
    } catch {
      Write-Host -ForegroundColor Red "❌ Renommage via Drive KO. Détails :"
      Show-HttpError $_
      if ($sheetsTriedError) { Write-Host -ForegroundColor Yellow "Note: Échec Sheets affiché plus haut également." }
      throw "Impossible de renommer le fichier (Sheets et Drive ont échoué)."
    }
  }
}
catch {
  Write-Host -ForegroundColor Red $_
  exit 1
}

// ============================================================================
// FICHIER : Temp_LinkSeed.gs  (TEMPORAIRE / DEBUG)
// VERSION : 1.0
// RÔLE    : Lier un fichier Réponses DEV à un Type_Test et stocker le mapping.
// ============================================================================

function dev_linkSeedSheet(typeTest, langue) {
  const name = `[DEV] ${typeTest} – Réponses`;

  // 1) Retrouver (ou créer) le fichier de réponses
  let ssid = '';
  const it = DriveApp.getFilesByName(name);
  if (it.hasNext()) {
    ssid = it.next().getId();
    Logger.log(`Trouvé : ${name} → ${ssid}`);
  } else if (typeof dev_seedResponseSheet === 'function') {
    // Si tu as déjà la fonction de "seed" complète, on l’utilise.
    ssid = dev_seedResponseSheet(typeTest, langue);
    Logger.log(`Créé via dev_seedResponseSheet : ${name} → ${ssid}`);
  } else {
    // Fallback minimal : crée un Google Sheet vide avec un onglet standard.
    const ss = SpreadsheetApp.create(name);
    ss.getSheets()[0].setName('Feuille 1');
    ssid = ss.getId();
    Logger.log(`Créé (minimal) : ${name} → ${ssid}`);
  }

  // 2) Enregistrer le mapping dans les Script Properties
  const sp = PropertiesService.getScriptProperties();
  let map = {};
  try { map = JSON.parse(sp.getProperty('RESPONSES_SSID_BY_TEST') || '{}'); } catch(_) {}
  map[typeTest] = ssid;
  sp.setProperty('RESPONSES_SSID_BY_TEST', JSON.stringify(map));

  // 3) S’assurer que le nom d’onglet cible est défini
  if (!sp.getProperty('RESPONSES_SHEET_NAME')) {
    sp.setProperty('RESPONSES_SHEET_NAME', 'Feuille 1');
  }

  Logger.log(
    `Mappé : ${typeTest} → ${ssid} (onglet="${sp.getProperty('RESPONSES_SHEET_NAME')}")`
  );
  return ssid;
}

// Raccourcis pratiques depuis la barre "Exécuter"
function run_linkSeed_Adap_FR(){ dev_linkSeedSheet('r&K_Adaptabilite','FR'); }
function run_linkSeed_Resi_FR(){ dev_linkSeedSheet('r&K_Resilience','FR'); }
function run_linkSeed_Crea_FR(){ dev_linkSeedSheet('r&K_Creativite','FR'); }

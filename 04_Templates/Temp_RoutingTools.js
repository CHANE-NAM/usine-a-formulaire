// ============================================================================
// FICHIER : Temp_RoutingTools.gs  (TEMPORAIRE / DEBUG)
// VERSION : 1.0
// RÔLE    : Inspecter et ajuster le routage des feuilles de réponses (Dry-run).
// ============================================================================

function Temp_debugRouting() {
  const sp = PropertiesService.getScriptProperties();
  const force = sp.getProperty('RESPONSES_SSID_FORCE') || '';
  const mapJson = sp.getProperty('RESPONSES_SSID_BY_TEST') || '{}';
  let map = {};
  try { map = JSON.parse(mapJson); } catch(_) {}
  const sheetName = sp.getProperty('RESPONSES_SHEET_NAME') || 'Réponses au formulaire 1';

  let type = '';
  try { type = (getTestConfiguration().Type_Test || '').trim(); } catch(_) {}

  Logger.log('Type_Test lu = %s', type || '(vide)');
  Logger.log('RESPONSES_SSID_FORCE = %s', force || '(vide)');
  Logger.log('RESPONSES_SHEET_NAME = %s', sheetName);
  Logger.log('Mappings (RESPONSES_SSID_BY_TEST) = %s', JSON.stringify(map));

  let effective = '';
  let reason = '';
  if (force) { effective = force; reason = 'FORCE'; }
  else if (type && map[type]) { effective = map[type]; reason = 'MAP['+type+']'; }
  else { reason = 'AUCUN (config/type ou mapping manquant)'; }

  Logger.log('→ SSID EFFECTIVE = %s (raison: %s)', effective || '(indéterminée)', reason);

  if (effective) {
    try {
      const ss = SpreadsheetApp.openById(effective);
      Logger.log('→ Nom du fichier ciblé : %s', ss.getName());
      const sh = ss.getSheetByName(sheetName) || ss.getSheets()[0];
      Logger.log('→ Onglet ciblé : %s', sh.getName());
    } catch(e) {
      Logger.log('⚠️ Impossible d’ouvrir la SSID effective: ' + e.message);
    }
  }
}

function Temp_clearResponsesRouting() {
  PropertiesService.getScriptProperties().deleteProperty('RESPONSES_SSID_FORCE');
  Logger.log('RESPONSES_SSID_FORCE supprimée.');
}

function Temp_showMappings() {
  const mapJson = PropertiesService.getScriptProperties().getProperty('RESPONSES_SSID_BY_TEST') || '{}';
  Logger.log('Mappings actuels = ' + mapJson);
}

function Temp_pointToTest(typeTest) {
  const sp = PropertiesService.getScriptProperties();
  let map = {};
  try { map = JSON.parse(sp.getProperty('RESPONSES_SSID_BY_TEST') || '{}'); } catch(_) {}
  const ssid = map[typeTest];
  if (!ssid) { Logger.log('⚠️ Aucun mapping pour ' + typeTest); return; }
  sp.setProperty('RESPONSES_SSID_FORCE', ssid);
  Logger.log('FORCE défini → %s', ssid);
}

function Temp_setResponsesSheetName(name) {
  PropertiesService.getScriptProperties().setProperty('RESPONSES_SHEET_NAME', String(name||''));
  Logger.log('RESPONSES_SHEET_NAME = ' + name);
}

// ============================================================================
// WRAPPERS SANS ARGUMENTS (TEMPORAIRE / DEBUG) — v1.1
// Permettent d'exécuter depuis le menu "Exécuter" sans passer d'arguments.
// ============================================================================
function run_pointToTest_Adap() { Temp_pointToTest('r&K_Adaptabilite'); }
function run_pointToTest_Resi() { Temp_pointToTest('r&K_Resilience'); }
function run_pointToTest_Crea() { Temp_pointToTest('r&K_Creativite'); }

function run_setSheet_Feuilles1() { Temp_setResponsesSheetName('Feuille 1'); }
function run_setSheet_Reponses()  { Temp_setResponsesSheetName('Réponses au formulaire 1'); }

function run_debugRouting()  { Temp_debugRouting(); }
function run_showMappings()  { Temp_showMappings(); }
function run_clearForce()    { Temp_clearResponsesRouting(); }

// Pointer automatiquement selon le Type_Test lu dans CONFIG (si présent)
function run_pointToConfigType() {
  try {
    const t = (getTestConfiguration().Type_Test || '').trim();
    if (!t) throw new Error('Type_Test vide dans CONFIG');
    Temp_pointToTest(t);
  } catch (e) {
    Logger.log('⚠️ ' + e.message);
  }
}

// --- TEMP : aligne l’ancienne clé sur la FORCE courante ---
function Temp_applyForceToLegacy() {
  const sp = PropertiesService.getScriptProperties();
  const force = sp.getProperty('RESPONSES_SSID_FORCE') || '';
  if (!force) { Logger.log('⚠️ Aucune FORCE définie'); return; }
  sp.setProperty('RESPONSES_SSID', force);
  Logger.log('RESPONSES_SSID (legacy) = ' + force);
}

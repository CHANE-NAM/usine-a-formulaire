/**********************************************
 * Injecteur de scénarios — r&K_Environnement
 * VERSION : 1.5
 * - Retire onOpen (aucun menu autonome)
 * - Corrige createResponse(...) pour LIST/MC (valeur string)
 * - Plusieurs scénarios (stable_rapide, instable_lent, k_fort, r_fort, alterne, median, stress)
 * - Alternance fiable par index d’échelle
 * - Fallback sûr sur sélecteurs requis (auto-choix 1)
 *
 * Envoie des réponses au Google Form ciblé
 * pour tester le pipeline (calculs, emails, PDF).
 *
 * Ouvre le CONFIG central par ID :
 *  1) via getSystemIds().ID_CONFIG si dispo
 *  2) sinon via la constante globale ID_FEUILLE_CONFIGURATION
 * Supporte les questions échelle même si le titre
 * ne contient pas l’ID "ENV001" (fallback par ordre).
 **********************************************/

// ========= Réglages rapides =========
const INJECT_DEFAULT = {
  rowIndex: 9,                     // ← n° de ligne dans l’onglet "Paramètres Généraux" (CONFIG)
  nbSubmissions: 1,                // nombre de soumissions par clic
  langue: 'FR',                    // valeur à injecter si un champ langue existe
  emailTest: 'dev.scenario+rk@example.com'
};

// ========= API publique (fonctions appelables) =========
function injectScenarioStableLent()      { _injectScenario({ type: 'stable_lent',      ...INJECT_DEFAULT }); }
function injectScenarioTurbulentRapide() { _injectScenario({ type: 'turbulent_rapide', ...INJECT_DEFAULT }); }
function injectScenarioMixte()           { _injectScenario({ type: 'mixte',            ...INJECT_DEFAULT }); }
function injectScenarioStableRapide()    { _injectScenario({ type: 'stable_rapide',    ...INJECT_DEFAULT }); }
function injectScenarioInstableLent()    { _injectScenario({ type: 'instable_lent',    ...INJECT_DEFAULT }); }
function injectScenarioKFort()           { _injectScenario({ type: 'k_fort',           ...INJECT_DEFAULT }); }
function injectScenarioRFort()           { _injectScenario({ type: 'r_fort',           ...INJECT_DEFAULT }); }
function injectScenarioAlterne()         { _injectScenario({ type: 'alterne',          ...INJECT_DEFAULT }); }
function injectScenarioMedian()          { _injectScenario({ type: 'median',           ...INJECT_DEFAULT }); }
function injectScenarioStressTest()      { _injectScenario({ type: 'stress',           ...INJECT_DEFAULT, nbSubmissions: 3 }); }

// ================== COEUR ==================
function _injectScenario(opts) {
  const scenario  = (opts && opts.type) || 'stable_lent';
  const rowIndex  = (opts && opts.rowIndex) || INJECT_DEFAULT.rowIndex;
  const nb        = (opts && opts.nbSubmissions) || 1;
  const langue    = (opts && opts.langue) || 'FR';
  const emailTest = (opts && opts.emailTest) || INJECT_DEFAULT.emailTest;

  // 1) CONFIG (ouvre le bon fichier)
  const cfg = _getCfgRow(rowIndex); // lit la ligne "Paramètres Généraux"
  if (!cfg.ID_Formulaire_Cible) {
    throw new Error("CONFIG ligne " + cfg._rowIndex + " : ID_Formulaire_Cible manquant.");
  }

  const form  = FormApp.openById(String(cfg.ID_Formulaire_Cible).trim());
  const items = form.getItems();

  // 2) BDD : profil de chaque ENVxxx (Stabilité / Vitesse)
  const profilSequence = _getProfilSequenceFromBDD(); // ex: ['ENV_STABILITE','ENV_STABILITE','ENV_VITESSE',...]

  // 3) FABRICATION/ENVOI
  for (let k = 0; k < nb; k++) {
    const resp = form.createResponse();
    let scaleIndex = 0; // réinitialisé à chaque soumission

    items.forEach(it => {
      const t = it.getType();

      // a) Échelles (SCALE)
      if (t === FormApp.ItemType.SCALE) {
        const scale = it.asScaleItem();
        const title = scale.getTitle() || '';

        // 1er essai : extraire ENVxxx du titre
        let profil = null;
        const idMatch = title.match(/(ENV\d{3,})/i);
        if (idMatch) {
          profil = _profilFromId(idMatch[1]);
        }
        // Fallback : prendre le profil à l'index courant
        if (!profil) {
          profil = profilSequence[scaleIndex] || 'ENV_STABILITE';
        }

        const min = scale.getLowerBound();
        const max = scale.getUpperBound();
        const val = _valueForScenario(profil, min, max, scenario, scaleIndex);
        resp.withItemResponse(scale.createResponse(val));

        scaleIndex++;

      // b) Champs texte (email, nom/entreprise, langue libre)
      } else if (t === FormApp.ItemType.TEXT) {
        const ti = it.asTextItem();
        const title = (ti.getTitle() || '').toLowerCase();

        if (title.match(/mail|e-?mail/)) {
          resp.withItemResponse(ti.createResponse(emailTest));
        } else if (title.match(/nom|name|entreprise|company/)) {
          const label =
              scenario === 'stable_lent'       ? 'Entreprise ALPHA (Stable & Lent)' :
              scenario === 'turbulent_rapide'  ? 'Entreprise BETA (Turbulent & Rapide)' :
              scenario === 'stable_rapide'     ? 'Entreprise DELTA (Stable & Rapide)' :
              scenario === 'instable_lent'     ? 'Entreprise EPSILON (Instable & Lent)' :
              scenario === 'k_fort'            ? 'Entreprise KAPPA (Très K)' :
              scenario === 'r_fort'            ? 'Entreprise RHO (Très r)' :
              scenario === 'alterne'           ? 'Entreprise SIGMA (Alterné)' :
              scenario === 'median'            ? 'Entreprise OMEGA (Médian)' :
                                                  'Entreprise GAMMA (Mixte)';
          resp.withItemResponse(ti.createResponse(label));
        } else if (title.match(/langue|language/)) {
          resp.withItemResponse(ti.createResponse(langue === 'FR' ? 'Français' : langue));
        }

      // c) Listes / Choix multiples (langue, consentement, etc.)
      } else if (t === FormApp.ItemType.MULTIPLE_CHOICE || t === FormApp.ItemType.LIST) {
        const sel   = (t === FormApp.ItemType.MULTIPLE_CHOICE) ? it.asMultipleChoiceItem() : it.asListItem();
        const title = (sel.getTitle() || '').toLowerCase();
        const choices = sel.getChoices() || [];

        // Gestion d’un sélecteur de langue (FR/EN/ES/DE…)
        if (title.match(/langue|language/)) {
          const target = (langue || 'FR').toString();
          const wanted = target.toUpperCase() === 'FR'
            ? /(fran|français|\bFR\b)/i
            : new RegExp(target, 'i');

          const hit = choices.find(c => wanted.test(String(c.getValue && c.getValue())));
          const value = hit ? hit.getValue() : (choices[0] ? choices[0].getValue() : 'Français');

          // IMPORTANT : passer une *valeur string*, pas l’objet Choice
          resp.withItemResponse(sel.createResponse(value));

        } else {
          // Fallback : si l’item est requis et non mappé, choisir la 1ère option
          if (sel.isRequired && choices.length > 0) {
            resp.withItemResponse(sel.createResponse(choices[0].getValue()));
          }
        }
      }
      // autres types ignorés
    });

    const submitted = resp.submit(); // déclenche onFormSubmit côté Sheet de réponses
    Logger.log('[OK] Scenario %s → ResponseId=%s | EditUrl=%s',
      scenario,
      submitted.getId && submitted.getId(),
      submitted.getEditResponseUrl && submitted.getEditResponseUrl());
  }

  Logger.log('Injection terminée : %s envoi(s) pour la ligne CONFIG %s.', nb, cfg._rowIndex);
}

// =============== Accès CONFIG & BDD ===============
function _getCfgRow(rowIndex) {
  const { ss, sheet } = _openConfig_();
  const lastCol  = sheet.getLastColumn();
  const headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values   = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  if (!values || values.every(v => v === '' || v == null)) {
    throw new Error('Ligne CONFIG vide ou invalide: ' + rowIndex);
  }
  const cfg = {};
  headers.forEach((h, i) => { if (h) cfg[String(h)] = values[i]; });
  cfg._rowIndex = rowIndex;
  return cfg;
}

function _openConfig_() {
  // 1) via getSystemIds().ID_CONFIG si dispo
  try {
    if (typeof getSystemIds === 'function') {
      const ids = getSystemIds();
      if (ids && ids.ID_CONFIG) {
        const ss = SpreadsheetApp.openById(ids.ID_CONFIG);
        const sheet = ss.getSheetByName('Paramètres Généraux');
        if (!sheet) throw new Error("Onglet 'Paramètres Généraux' introuvable.");
        return { ss, sheet };
      }
    }
  } catch (e) {
    // on retente via la constante
  }
  // 2) via constante globale ID_FEUILLE_CONFIGURATION
  if (typeof ID_FEUILLE_CONFIGURATION === 'string' && ID_FEUILLE_CONFIGURATION.trim() !== '') {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
    const sheet = ss.getSheetByName('Paramètres Généraux');
    if (!sheet) throw new Error("Onglet 'Paramètres Généraux' introuvable.");
    return { ss, sheet };
  }
  throw new Error('Impossible d’ouvrir la feuille CONFIG (ni ID_CONFIG ni ID_FEUILLE_CONFIGURATION).');
}

function _getProfilSequenceFromBDD() {
  const ids = (typeof getSystemIds === 'function') ? getSystemIds() : null;
  if (!ids || !ids.ID_BDD) throw new Error('ID_BDD introuvable dans sys_ID_Fichiers.');
  const bdd = SpreadsheetApp.openById(ids.ID_BDD);
  const qSheet = bdd.getSheetByName('Questions_r&K_Environnement_FR');
  if (!qSheet) throw new Error("BDD: onglet Questions_r&K_Environnement_FR introuvable");

  const qData = qSheet.getDataRange().getValues();
  const headers = qData.shift();
  const typeCol   = headers.indexOf('TypeQuestion');
  const paramsCol = headers.indexOf('Paramètres (JSON)');

  const seq = [];
  qData.forEach(r => {
    const type = r[typeCol];
    const pj   = r[paramsCol];
    if (String(type).toUpperCase() === 'ECHELLE_NOTE' && pj) {
      try {
        const p = JSON.parse(pj);
        if (p && p.profil) seq.push(String(p.profil));
      } catch (_) {}
    }
  });
  return seq;
}

function _profilFromId(envId) {
  try {
    const ids = (typeof getSystemIds === 'function') ? getSystemIds() : null;
    if (!ids || !ids.ID_BDD) return null;
    const bdd = SpreadsheetApp.openById(ids.ID_BDD);
    const qSheet = bdd.getSheetByName('Questions_r&K_Environnement_FR');
    if (!qSheet) return null;

    const qData = qSheet.getDataRange().getValues();
    const headers = qData.shift();
    const idCol     = headers.indexOf('ID');
    const paramsCol = headers.indexOf('Paramètres (JSON)');
    const row = qData.find(r => String(r[idCol]).toUpperCase() === String(envId).toUpperCase());
    if (!row) return null;
    const p = JSON.parse(row[paramsCol] || '{}');
    return p.profil || null;
  } catch(_) {
    return null;
  }
}

// =============== Génération de valeurs ===============
function _valueForScenario(profil, min, max, scenario, idx /* index d’échelle */) {
  const clamp = (v) => Math.max(min, Math.min(max, v));
  const span  = max - min;
  const rnd   = () => Math.random();

  const hi  = () => clamp(Math.round(max - span * 0.05 * rnd()));                  // très haut
  const lo  = () => clamp(Math.round(min + span * 0.05 * rnd()));                  // très bas
  const mid = () => clamp(Math.round(min + span * (0.45 + 0.10 * (rnd() - 0.5)))); // milieu ±

  const isStab    = String(profil || '').toUpperCase().indexOf('STABILITE') >= 0;
  const isVitesse = String(profil || '').toUpperCase().indexOf('VITESSE')   >= 0;

  switch (scenario) {
    case 'stable_lent':
      return isStab ? hi() : isVitesse ? lo() : mid();
    case 'turbulent_rapide':
      return isStab ? lo() : isVitesse ? hi() : mid();
    case 'stable_rapide':
      return isStab ? hi() : isVitesse ? hi() : mid();
    case 'instable_lent':
      return isStab ? lo() : isVitesse ? lo() : mid();
    case 'k_fort':
      return isStab ? hi() : mid();
    case 'r_fort':
      return isVitesse ? hi() : mid();
    case 'alterne':
      return (idx % 2 === 0) ? hi() : lo();
    case 'median':
      return mid();
    case 'mixte':
    default:
      if (isStab)    return clamp(Math.round(min + span * (0.55 + 0.15 * (rnd() - 0.5))));
      if (isVitesse) return clamp(Math.round(min + span * (0.50 + 0.20 * (rnd() - 0.5))));
      return mid();
  }
}



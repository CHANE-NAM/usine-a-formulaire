// =================================================================================
// FICHIER : Utils V2.gs (Projet MOTEUR)
// RÔLE    : Fonctions utilitaires pour le Moteur / Usine
// VERSION : 3.1 - getConfigurationFromRow tolérant (résolution sheet/ID/nom) + ECHELLE_NOTE robuste
// =================================================================================

// ⚙️ ID de la feuille de configuration centrale (CONFIG)
const ID_FEUILLE_CONFIGURATION = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

// ------------------------------------
// Helpers génériques
// ------------------------------------
function _normHeader(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // enlève les accents
    .trim();
}
function _normKey(s) {
  return _normHeader(s).toLowerCase();
}
function _normLabel(s){ return _normHeader(s).toLowerCase(); }

// Résout une "feuille de config" depuis : Sheet | Spreadsheet | nom d’onglet | ID spreadsheet | null
function _resolveConfigSheet(sheetLike) {
  // 1) Si on a déjà une Sheet
  if (sheetLike && typeof sheetLike.getLastColumn === 'function' && typeof sheetLike.getRange === 'function') {
    return sheetLike; // c'est bien une Sheet
  }

  // Liste des noms possibles de l’onglet "Paramètres Généraux"
  const CANDIDATE_NAMES = ['Paramètres Généraux','Parametres Generaux','Parameters','Parametres'];

  // 2) Si on a un Spreadsheet
  if (sheetLike && typeof sheetLike.getSheetByName === 'function') {
    for (const name of CANDIDATE_NAMES) {
      const sh = sheetLike.getSheetByName(name);
      if (sh) return sh;
    }
    throw new Error("Onglet 'Paramètres Généraux' introuvable dans le Spreadsheet transmis.");
  }

  // 3) Si on a une chaîne : essayer d’abord comme ID de spreadsheet, sinon comme nom d’onglet dans la CONFIG
  if (typeof sheetLike === 'string' && sheetLike.trim() !== '') {
    const val = sheetLike.trim();

    // a) Essai comme ID de Spreadsheet
    try {
      const ss = SpreadsheetApp.openById(val);
      for (const name of CANDIDATE_NAMES) {
        const sh = ss.getSheetByName(name);
        if (sh) return sh;
      }
      throw new Error("Onglet 'Paramètres Généraux' introuvable dans le spreadsheet " + val);
    } catch (e) {
      // b) Essai comme nom d’onglet dans la CONFIG centrale
      const cfgSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
      const sh = cfgSS.getSheetByName(val);
      if (sh) return sh;
      // c) fallback : on continuera vers CONFIG + noms candidats
    }
  }

  // 4) Fallback : CONFIG centrale + noms candidats
  const cfgSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
  for (const name of CANDIDATE_NAMES) {
    const sh = cfgSS.getSheetByName(name);
    if (sh) return sh;
  }
  throw new Error("Impossible de localiser l'onglet 'Paramètres Généraux' dans la CONFIG centrale.");
}

// ------------------------------------
// IDs système (CONFIG → onglet sys_ID_Fichiers)
// ------------------------------------
function getSystemIds() {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIGURATION);
  const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
  if (!idSheet) throw new Error("L'onglet 'sys_ID_Fichiers' est introuvable dans CONFIG.");
  const data = idSheet.getDataRange().getValues();
  const ids = {};
  data.slice(1).forEach(row => {
    if (row[0] && row[1]) ids[row[0]] = row[1];
  });
  return ids;
}

// ------------------------------------
// Lecture d'une ligne de CONFIG (format horizontal)
//  - Accepte rowIndex numérique OU chaîne ("9", "Ligne 9", etc.)
//  - Si invalide, choisit automatiquement une ligne "En construction"
//    (priorité à la dernière SANS ID_Formulaire_Cible), sinon la dernière.
//  - Retourne un objet { en-tête -> valeur } + cfg._rowIndex pour le logging.
// ------------------------------------
function getConfigurationFromRow(sheetLike, rowIndex) {
  const sheet = _resolveConfigSheet(sheetLike);

  // 1) Normaliser la valeur reçue (nombre, string, etc.)
  let idx = rowIndex;
  if (idx !== 0 && !idx) idx = ''; // null/undefined -> ''

  // Essayer d’extraire un entier (ex: "9", " 9 ", "Ligne 9" -> 9)
  if (typeof idx !== 'number') {
    const m = String(idx).match(/\d+/);
    idx = m ? Number(m[0]) : NaN;
  }

  // 2) Lecture des en-têtes
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const mapIdx = {};
  headers.forEach((h,i) => { mapIdx[_normLabel(h)] = i; });

  const iStatut = mapIdx[_normLabel('Statut')];
  const iIdForm = mapIdx[_normLabel('ID_Formulaire_Cible')];

  // 3) Si l’index n’est pas exploitable, sélectionner automatiquement une ligne "En construction"
  if (!idx || isNaN(idx) || idx < 2) {
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    if (rowCount === 0) throw new Error('getConfigurationFromRow: la feuille CONFIG est vide.');

    const data = sheet.getRange(2, 1, rowCount, lastCol).getValues();

    // a) collecter les lignes "En construction"
    const lignesEC = [];
    data.forEach((row, k) => {
      const statut = (iStatut != null) ? String(row[iStatut] || '').trim().toLowerCase() : '';
      if (statut === 'en construction') {
        const hasIdForm = (iIdForm != null) && String(row[iIdForm] || '').trim() !== '';
        lignesEC.push({ rowIndex: k + 2, hasIdForm, row });
      }
    });

    if (lignesEC.length === 0) {
      throw new Error('getConfigurationFromRow: aucune ligne "En construction" trouvée et rowIndex fourni invalide.');
    }

    // b) Priorité : la DERNIÈRE sans ID_Formulaire_Cible, sinon la DERNIÈRE tout court
    let pick = lignesEC.filter(r => !r.hasIdForm).pop();
    if (!pick) pick = lignesEC.pop();

    idx = pick.rowIndex;
  }

  // 4) Index final validé
  if (!idx || isNaN(idx) || idx < 2) {
    throw new Error('getConfigurationFromRow: rowIndex invalide (' + rowIndex + ')');
  }

  // 5) Renvoi de la config de la ligne choisie
  const values = sheet.getRange(idx, 1, 1, lastCol).getValues()[0];
  const cfg = {};
  headers.forEach((h, i) => { if (h) cfg[h] = values[i]; });
  cfg._rowIndex = idx;
  return cfg;
}

// ------------------------------------
// Langues
// ------------------------------------
function getLangueFullName(code) {
  const map = { FR: 'Français', EN: 'English', ES: 'Español', DE: 'Deutsch' };
  return map[String(code || '').toUpperCase()] || code;
}

// ------------------------------------
// Options pour QCU/QRM
// ------------------------------------
function buildChoices(optionsString, params) {
  if (params && Array.isArray(params.options) && params.options.length > 0) {
    return params.options.map(o => (o && o.libelle) ? String(o.libelle) : '').filter(Boolean);
  }
  if (!optionsString) return [];
  return String(optionsString).split(';').map(s => s.trim()).filter(Boolean);
}

// ------------------------------------
// Création d'items dans le Google Form
//  - gère QRM, QCU, ECHELLE, ECHELLE_NOTE (robuste), EMAIL, TEXTE_COURT
//  - remplace [LIEN_FICHIER:Nom] dans la description si présent
// ------------------------------------
function creerItemFormulaire(form, type, titre, optionsString, description, paramsJSONString) {
  // 1) Résolution [LIEN_FICHIER:...]
  let finalDescription = description;
  const placeholderRegex = /\[LIEN_FICHIER:(.*?)\]/;
  const match = description ? description.match(placeholderRegex) : null;

  if (match && match[1]) {
    const nomFichier = match[1].trim();
    try {
      const systemIds = getSystemIds();
      const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
      const listeFichiersSheet = bdd.getSheetByName('Liste_Fichiers_Drive');

      if (listeFichiersSheet) {
        const data = listeFichiersSheet.getDataRange().getValues();
        const fileRow = data.find(row => String(row[0] || '').trim() === nomFichier);
        if (fileRow && fileRow[1]) {
          const fileId = String(fileRow[1]).trim();
          const fileUrl = `https://drive.google.com/file/d/${fileId}/view`;
          finalDescription = description.replace(placeholderRegex, fileUrl);
        } else {
          finalDescription = description.replace(placeholderRegex, `[ERREUR: Fichier '${nomFichier}' introuvable dans la BDD]`);
        }
      } else {
        finalDescription = description.replace(placeholderRegex, `[ERREUR: Onglet 'Liste_Fichiers_Drive' introuvable]`);
      }
    } catch (e) {
      Logger.log("Erreur lien fichier : " + e.message);
      finalDescription = description.replace(placeholderRegex, `[ERREUR SCRIPT: ${e.message}]`);
    }
  }

  // 2) Parsing JSON souple
  let params = null;
  if (paramsJSONString && String(paramsJSONString).trim() !== '') {
    try { params = JSON.parse(paramsJSONString); } catch (e) { params = null; }
  }

  // 3) Construction des choix si nécessaire
  const choices = buildChoices(optionsString, params);

  // 4) Création de l'item selon le type
  let item = null;
  const formItemType = type ? String(type).toUpperCase() : '';

  if (formItemType.startsWith('QRM')) {
    if (choices.length > 0) {
      item = form.addCheckboxItem().setTitle(titre).setChoiceValues(choices).setRequired(true);
    } else {
      item = form.addParagraphTextItem().setTitle("[Erreur QRM: Options manquantes] " + titre);
    }

  } else if (formItemType.startsWith('QCU')) {
    if (choices.length > 0) {
      item = form.addMultipleChoiceItem().setTitle(titre).setChoiceValues(choices).setRequired(true);
    } else {
      item = form.addParagraphTextItem().setTitle("[Erreur QCU: Options manquantes] " + titre);
    }

  } else if (formItemType === 'ECHELLE_NOTE') { // ✅ v3.1 robuste : min/max OU echelle_min/echelle_max + labels tolérants
    if (params) {
      const eMin = (params.echelle_min ?? params.min);
      const eMax = (params.echelle_max ?? params.max);
      if (eMin != null && eMax != null) {
        const scaleItem = form.addScaleItem()
          .setTitle(titre)
          .setBounds(Number(eMin), Number(eMax))
          .setRequired(true);

        // Labels : label_min / libelle_min / labelMin  (idem pour _max)
        const lmin = (params.label_min ?? params.libelle_min ?? params.labelMin);
        const lmax = (params.label_max ?? params.libelle_max ?? params.labelMax);
        if (lmin && lmax) scaleItem.setLabels(String(lmin), String(lmax));

        item = scaleItem;
      } else {
        item = form.addParagraphTextItem()
          .setTitle("[Erreur ECHELLE_NOTE: Paramètres JSON incomplets (min/max)] " + titre);
      }
    } else {
      item = form.addParagraphTextItem()
        .setTitle("[Erreur ECHELLE_NOTE: Paramètres JSON absents] " + titre);
    }

  } else if (formItemType === 'ECHELLE') { // compat historique
    const parts = optionsString ? optionsString.split(';').map(s => s.trim()) : [];
    const min = parts[0] ? Number(parts[0]) : 1;
    const max = parts[parts.length - 1] ? Number(parts[parts.length - 1]) : 5;
    const scaleItem = form.addScaleItem().setTitle(titre).setBounds(min, max).setRequired(true);
    item = scaleItem;

  } else if (formItemType === 'EMAIL') {
    const textItem = form.addTextItem().setTitle(titre).setRequired(true);
    const emailValidation = FormApp.createTextValidation().requireTextIsEmail().build();
    item = textItem.setValidation(emailValidation);

  } else if (formItemType === 'TEXTE_COURT') {
    item = form.addTextItem().setTitle(titre).setRequired(true);

  } else {
    item = form.addParagraphTextItem().setTitle("[Type Inconnu: " + type + "] " + titre);
  }

  // 5) HelpText (évite d’écraser les labels d’échelle)
  if (finalDescription && item && typeof item.setHelpText === 'function') {
    if (formItemType !== 'ECHELLE' && formItemType !== 'ECHELLE_NOTE') {
      item.setHelpText(finalDescription);
    }
  }
}

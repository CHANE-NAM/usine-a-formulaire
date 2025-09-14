/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Utilities.gs
 * == VERSION : 12.1 - Mise à jour de la logique de configuration pour l'architecture V4.
 * == RÔLE    : Boîte à outils, lecture de configuration, et utilitaires de données.
 * =================================================================================
 */

/**
 * Récupère la configuration du test en utilisant le classeur du kit fourni comme contexte.
 * @param {Spreadsheet} kitSpreadsheet Le classeur du kit de traitement.
 * @returns {Object} L'objet de configuration.
 */
function getTestConfiguration(kitSpreadsheet) {
  if (!kitSpreadsheet) {
    throw new Error("getTestConfiguration a été appelée sans classeur de kit valide.");
  }

  // Stratégie 1 (préférée) : Lire la configuration depuis le fichier CONFIG global.
  const ids = getSystemIds();
  if (ids && ids.ID_CONFIG) {
    const cfgFromGlobal = _tryReadKeyValueOrHorizontalConfig(ids.ID_CONFIG, ['Paramètres Généraux'], kitSpreadsheet);
    if (cfgFromGlobal && String(cfgFromGlobal.Type_Test || '').trim() !== '') {
      return cfgFromGlobal;
    }
  }

  // Stratégie 2 (fallback) : Lire la configuration depuis l'onglet '[CONFIG]' du kit lui-même.
  const cfgFromKit = _tryReadKeyValueOrHorizontalConfig(null, ['[CONFIG]', 'CONFIG'], kitSpreadsheet);
  if (cfgFromKit && String(cfgFromKit.Type_Test || '').trim() !== '') {
    return cfgFromKit;
  }

  throw new Error("Impossible de trouver la configuration pour ce test (ID: " + kitSpreadsheet.getId() + "). Vérifiez le fichier [CONFIG] global ou l'onglet [CONFIG] local.");
}


/**
 * Tente de lire une configuration.
 * @param {string|null} fileId L'ID du fichier à ouvrir (si null, utilise kitSpreadsheet).
 * @param {Array<string>} possibleSheetNames Les noms d'onglets possibles.
 * @param {Spreadsheet} kitSpreadsheet Le classeur du kit actif (toujours requis pour le contexte).
 * @returns {Object|null} L'objet de configuration ou null.
 */
function _tryReadKeyValueOrHorizontalConfig(fileId, possibleSheetNames, kitSpreadsheet) {
  try {
    const ss = fileId ? SpreadsheetApp.openById(fileId) : kitSpreadsheet;
    if (!ss) return null;

    let sh = null;
    for (const name of possibleSheetNames) {
      sh = ss.getSheetByName(name);
      if (sh) break;
    }
    if (!sh) return null;

    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return null;

    const headersRow = data[0].map(h => String(h || '').trim());

    // Gestion du format Clé/Valeur vertical
    const header0 = headersRow[0].toLowerCase();
    if ((headersRow.length <= 3) && (header0.includes('clé') || header0.includes('cle') || header0.includes('key'))) {
      const cfg = {};
      for (let i = 1; i < data.length; i++) {
        const k = String(data[i][0] || '').trim();
        if (k) cfg[k] = data[i][1];
      }
      return cfg;
    }

    // --- MODIFICATION V4 ---
    // La logique ci-dessous est mise à jour pour chercher par ID de Formulaire d'abord.
    // -------------------------
    const idx = {};
    headersRow.forEach((h, i) => { if (h) idx[h] = i; });

    const kitId = kitSpreadsheet.getId();
    let targetRow = null;

    // NOUVELLE STRATÉGIE V4 : On cherche d'abord par l'ID du formulaire lié à la feuille de réponses.
    const formUrl = kitSpreadsheet.getFormUrl();
    if (formUrl && idx['ID_Formulaire_Cible'] != null) {
      const formIdMatch = formUrl.match(/[-\w]{25,}/); // Regex pour extraire l'ID de l'URL
      if (formIdMatch) {
        const formId = formIdMatch[0];
        targetRow = data.slice(1).find(r => String(r[idx['ID_Formulaire_Cible']] || '') === formId);
      }
    }

    // ANCIENNE STRATÉGIE (pour la compatibilité) : Si non trouvé, on cherche par l'ID du Sheet.
    if (!targetRow && idx['ID_Sheet_Cible'] != null) {
      targetRow = data.slice(1).find(r => String(r[idx['ID_Sheet_Cible']] || '') === kitId);
    }

    if (!targetRow) {
      throw new Error(`Configuration introuvable pour le kit ID "${kitId}" dans l'onglet "${sh.getName()}". Vérifiez les colonnes 'ID_Formulaire_Cible' ou 'ID_Sheet_Cible'.`);
    }

    const cfg = {};
    headersRow.forEach((h, i) => { if (h) cfg[h] = targetRow[i]; });
    return cfg;
    // --- FIN DE LA MODIFICATION V4 ---

  } catch (e) {
    Logger.log('_tryReadKeyValueOrHorizontalConfig KO pour fileId ' + fileId + ' / kit ' + (kitSpreadsheet ? kitSpreadsheet.getName() : 'N/A') + ' : ' + e.message);
    throw e;
  }
}


/**
 * Lit les IDs système depuis une feuille pilote centrale.
 */
function getSystemIds() {
  const ID_FEUILLE_PILOTE = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";
  try {
    const configSS = SpreadsheetApp.openById(ID_FEUILLE_PILOTE);
    const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
    if (!idSheet) throw new Error("L'onglet 'sys_ID_Fichiers' est introuvable.");
    const data = idSheet.getDataRange().getValues();
    const ids = {};
    data.slice(1).forEach(row => {
      if (row[0] && row[1]) { ids[row[0]] = row[1]; }
    });
    return ids;
  } catch (e) {
    Logger.log("Impossible de charger les ID système : " + e.toString());
    throw new Error("Impossible de charger les ID système. Erreur: " + e.message);
  }
}


/**
 * Détecte la langue de la réponse initiale de l'utilisateur.
 */
function getOriginalLanguage(reponses) {
  const langueRepondantBrute = reponses['Langue___Language'] || reponses['Langue / Language'] || 'Français';
  const mapLangue = { 'Français': 'FR', 'English': 'EN', 'Español': 'ES', 'Deutsch': 'DE' };
  return mapLangue[langueRepondantBrute] || 'FR';
}


/**
 * Récupère le contenu d'un gabarit d'email depuis la BDD.
 */
function getGabaritEmail(idGabarit, langueCode) {
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const gabaritsSheet = bdd.getSheetByName("Gabarits_Emails");
  if (!gabaritsSheet) throw new Error("L'onglet 'Gabarits_Emails' est introuvable.");
  const data = gabaritsSheet.getDataRange().getValues();
  const headers = data.shift();
  const idCol = headers.indexOf('ID_Gabarit');
  const langCol = headers.indexOf('Langue');
  const gabaritRow = data.find(row => row[idCol] === idGabarit && row[langCol].toUpperCase() === langueCode.toUpperCase());
  if (!gabaritRow) throw new Error(`Aucun gabarit trouvé pour l'ID '${idGabarit}' et la langue '${langueCode}'.`);

  const gabarit = {};
  headers.forEach((header, index) => {
    if (header) { gabarit[header] = gabaritRow[index]; }
  });
  return gabarit;
}


/**
 * Formate le texte de détail des scores pour l'email.
 */
function formatScoresDetails(resultats, niveauDetails, typeTest, langueCode) {
  if (niveauDetails === 'Simple' || !resultats.scoresData || Object.keys(resultats.scoresData).length === 0) {
    return "";
  }
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const formatSheet = bdd.getSheetByName("sys_Formatage_Scores");
    if (!formatSheet) return "Erreur: Onglet 'sys_Formatage_Scores' introuvable.\n";
    const formatData = formatSheet.getDataRange().getValues();
    const formatHeaders = formatData.shift();
    const typeTestCol = formatHeaders.indexOf('Type_Test');
    const regle = formatData.find(row => row[typeTestCol] === typeTest);
    if (!regle) return `Aucune règle d'affichage trouvée pour le test '${typeTest}'.\n`;

    const regleMap = {};
    formatHeaders.forEach((h, i) => regleMap[h] = regle[i]);
    const T = loadTraductions(langueCode);
    let scoresText = (regleMap.Texte_Intro || "Voici le détail de vos scores :") + "\n";

    if (regleMap.Mode_Affichage === 'Simple') {
      let scoresArray = Object.entries(resultats.scoresData).map(([code, score]) => ({
        code_profil: code,
        nom_profil: resultats.mapCodeToName[code] || code,
        score: score
      }));
      if (regleMap.Tri_Scores === 'Décroissant') {
        scoresArray.sort((a, b) => b.score - a.score);
      } else if (regleMap.Tri_Scores === 'Croissant') {
        scoresArray.sort((a, b) => a.score - b.score);
      }
      scoresArray.forEach(item => {
        let ligne = regleMap.Format_Ligne.replace(/{{nom_profil}}/g, item.nom_profil)
          .replace(/{{score}}/g, item.score)
          .replace(/{{suffixe_points}}/g, T.SUFFIXE_POINTS || 'points');
        scoresText += ligne + "\n";
      });
    } else if (regleMap.Mode_Affichage === 'Dichotomie') {
      const axes = [
        { nom: (T.AXE_EI || "Extraversion (E) vs Introversion (I)"), p1: 'E', p2: 'I' },
        { nom: (T.AXE_SN || "Sensation (S) vs Intuition (N)"), p1: 'S', p2: 'N' },
        { nom: (T.AXE_TF || "Pensée (T) vs Sentiment (F)"), p1: 'T', p2: 'F' },
        { nom: (T.AXE_JP || "Jugement (J) vs Perception (P)"), p1: 'J', p2: 'P' }
      ];
      axes.forEach(axe => {
        let ligne = regleMap.Format_Ligne.replace(/{{axe_nom}}/g, axe.nom)
          .replace(/{{score1}}/g, resultats.scoresData[axe.p1] || 0)
          .replace(/{{score2}}/g, resultats.scoresData[axe.p2] || 0);
        scoresText += ligne + "\n";
      });
    }
    return scoresText;
  } catch (e) {
    Logger.log(`ERREUR CRITIQUE DANS formatScoresDetails : ${e.toString()}`);
    return "Impossible d'afficher le détail des scores en raison d'une erreur.\n";
  }
}


/**
 * Charge les chaînes de caractères traduites pour une langue donnée.
 */
function loadTraductions(langueCode) {
  if (!langueCode) {
    throw new Error("Le code de langue fourni à loadTraductions est indéfini.");
  }
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const traductionsSheet = bdd.getSheetByName("traductions");
  if (!traductionsSheet) throw new Error("L'onglet 'traductions' est introuvable.");
  const data = traductionsSheet.getDataRange().getValues();
  const headers = data.shift();
  const langColIndex = headers.findIndex(h => h && String(h).trim().toLowerCase() === langueCode.toLowerCase());
  if (langColIndex === -1) throw new Error(`La colonne de langue '${langueCode}' est introuvable dans l'onglet "traductions".`);

  const traductions = {};
  const keyColIndex = 0;
  data.forEach(row => {
    if (row[keyColIndex]) { traductions[row[keyColIndex]] = row[langColIndex]; }
  });
  return traductions;
}


/**
 * Trouve les pièces jointes à inclure dans l'email.
 */
function findAttachments(config, profilCode, niveauPJ, langueCode) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const pjSheet = bdd.getSheetByName("sys_PiecesJointes");
    if (!pjSheet) { return []; }
    const data = pjSheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = {
      type: headers.indexOf('Type_Test'),
      profil: headers.indexOf('Profil_Code'),
      niveau: headers.indexOf('Email_Niveau'),
      langue: headers.indexOf('Langue'),
      id: headers.indexOf('ID_Fichier_Drive')
    };
    if (Object.values(idx).some(i => i === -1)) {
      Logger.log("Avertissement : une ou plusieurs colonnes sont manquantes dans 'sys_PiecesJointes'.");
      return [];
    }

    const niveauNumRequis = parseInt(String(niveauPJ).replace(/[^0-9]/g, ''), 10) || 1;
    const idsFichiersTrouves = new Set();
    data.forEach(row => {
      const typeMatch = (row[idx.type] || '').toString().toUpperCase() === (config.Type_Test || '').toUpperCase();
      const profilMatch = (row[idx.profil] === profilCode || row[idx.profil] === 'TOUS');
      const langueMatch = (row[idx.langue] === langueCode || row[idx.langue] === 'TOUS');
      const niveauMatch = (row[idx.niveau] > 0 && row[idx.niveau] <= niveauNumRequis);

      if (typeMatch && profilMatch && niveauMatch && langueMatch && row[idx.id]) {
        idsFichiersTrouves.add(row[idx.id]);
      }
    });

    const fichiers = [];
    idsFichiersTrouves.forEach(id => {
      try {
        fichiers.push(DriveApp.getFileById(id).getBlob());
      } catch (e) {
        Logger.log(`Impossible d'accéder au fichier Drive avec l'ID : ${id}`);
      }
    });
    return fichiers;
  } catch (e) {
    Logger.log(`Erreur critique dans findAttachments : ${e.toString()}`);
    return [];
  }
}
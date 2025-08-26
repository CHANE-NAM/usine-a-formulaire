// =================================================================================
// == FICHIER : Utilities.gs
// == VERSION : 9.1 (Multi-sources + lecture format horizontal OU clé→valeur)
// == RÔLE  : Boîte à outils du Kit de Traitement.
// =================================================================================

const ID_FEUILLE_PILOTE = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

/**
 * Récupère la configuration du test en cours avec une stratégie multi-sources :
 * 1) CONFIG global (ID_CONFIG) → onglet 'Paramètres Généraux'
 * 2) Ancien template V2 (ID_TEMPLATE_TRAITEMENT_V2) → 'Paramètres Généraux'
 * 3) Legacy "pilote" (ID_FEUILLE_PILOTE) mappant la sheet active via 'ID_Sheet_Cible'
 */
function getTestConfiguration() {
  const ids = getSystemIds();

  // 1) Préféré : CONFIG global
  const cfgFromConfig = _tryReadKeyValueOrHorizontalConfig(ids.ID_CONFIG, [
    'Paramètres Généraux','Parametres Generaux','Parameters','Parametres'
  ]);
  if (cfgFromConfig && String(cfgFromConfig.Type_Test || '').trim() !== '') {
    return cfgFromConfig;
  }

  // 2) Fallback : ancien template V2
  const cfgFromTemplateV2 = _tryReadKeyValueOrHorizontalConfig(ids.ID_TEMPLATE_TRAITEMENT_V2, [
    'Paramètres Généraux','Parametres Generaux','Parameters','Parametres'
  ]);
  if (cfgFromTemplateV2 && String(cfgFromTemplateV2.Type_Test || '').trim() !== '') {
    return cfgFromTemplateV2;
  }

  // 3) Dernier recours : logique legacy "pilote" (comportement historique)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const idSheetActuelle = ss.getId();
    const piloteSheet = SpreadsheetApp.openById(ID_FEUILLE_PILOTE);
    const paramsSheet = piloteSheet.getSheetByName("Paramètres Généraux");
    if (!paramsSheet) { throw new Error("L'onglet 'Paramètres Généraux' est introuvable dans la feuille pilote."); }
    const data = paramsSheet.getDataRange().getValues();
    const headers = data.shift();
    const idSheetColIndex = headers.indexOf('ID_Sheet_Cible');
    if (idSheetColIndex === -1) { throw new Error("La colonne 'ID_Sheet_Cible' est introuvable dans la feuille pilote."); }
    const configRow = data.find(row => row[idSheetColIndex] === idSheetActuelle);
    if (!configRow) { throw new Error("Impossible de trouver la configuration pour ce test (ID: " + idSheetActuelle + ")."); }
    const configuration = {};
    headers.forEach((header, index) => {
      if (header) { configuration[header] = configRow[index]; }
    });
    if (!configuration.Type_Test || String(configuration.Type_Test).trim() === '') {
      throw new Error("Type_Test manquant dans la configuration legacy (feuille pilote).");
    }
    return configuration;
  } catch (e) {
    throw new Error("Impossible de trouver la configuration pour ce test (CONFIG, TEMPLATE_V2 et legacy ont échoué). Détail: " + e.message);
  }
}

/**
 * Lit une config depuis un fichier (ID Drive) et un des noms d’onglet possibles.
 * Supporte 2 formats :
 *  - Clé→Valeur (2 colonnes, ex. 'Type_Test' | 'r&K_Environnement')
 *  - Tableau horizontal (1 ligne de données sous les entêtes, ex. 'Type_Test' est une colonne)
 * Retourne null si le fichier/onglet n’existe pas ou si aucune ligne valide n’est trouvée.
 */
function _tryReadKeyValueOrHorizontalConfig(fileId, possibleSheetNames) {
  try {
    if (!fileId) return null;
    const ss = SpreadsheetApp.openById(fileId);

    // Trouver un onglet parmi ceux proposés
    let sh = null;
    for (const name of possibleSheetNames) {
      sh = ss.getSheetByName(name);
      if (sh) break;
    }
    if (!sh) return null;

    const data = sh.getDataRange().getValues();
    if (!data || data.length < 2) return null;

    const headersRow = data[0].map(h => String(h || '').trim());
    const nbCols = headersRow.length;

    // Heuristique : format Clé→Valeur si 2 colonnes OU si la 1ère colonne s’appelle "Clé"/"Key"
    const header0 = headersRow[0].toLowerCase();
    const isKeyValue = (nbCols <= 3) && (header0.includes('clé') || header0.includes('cle') || header0.includes('key'));

    if (isKeyValue) {
      // Lecture Clé→Valeur
      const cfg = {};
      for (let i = 1; i < data.length; i++) {
        const k = String(data[i][0] || '').trim();
        const v = data[i][1];
        if (k) cfg[k] = v;
      }
      return cfg;
    }

    // Sinon : format horizontal (entêtes en ligne 1, données sur une/des lignes)
    const idx = {};
    headersRow.forEach((h, i) => { if (h) idx[h] = i; });

    // Choix de la ligne cible :
    // 1) si 'ID_Sheet_Cible' présent, on essaie de matcher l'ID de la feuille active
    const activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    let target = null;
    if (idx['ID_Sheet_Cible'] != null) {
      target = data.slice(1).find(r => String(r[idx['ID_Sheet_Cible']] || '') === activeId);
    }
    // 2) sinon, première ligne où 'Type_Test' est non vide
    if (!target && idx['Type_Test'] != null) {
      target = data.slice(1).find(r => String(r[idx['Type_Test']] || '').trim() !== '');
    }
    // 3) sinon, première ligne non vide
    if (!target) {
      target = data.slice(1).find(r => r.some(c => String(c || '').trim() !== ''));
    }
    if (!target) return null;

    const cfg = {};
    headersRow.forEach((h, i) => { if (h) cfg[h] = target[i]; });
    return cfg;

  } catch (e) {
    Logger.log('_tryReadKeyValueOrHorizontalConfig KO for ' + fileId + ' : ' + e.message);
    return null;
  }
}

/**
 * Lit l'onglet 'sys_ID_Fichiers' de la feuille de configuration centrale (feuille pilote).
 * Renvoie un dictionnaire { Clé → ID } (ex.: { ID_BDD: "...", ID_CONFIG: "...", ... }).
 */
function getSystemIds() {
  try {
    const configSS = SpreadsheetApp.openById(ID_FEUILLE_PILOTE);
    const idSheet = configSS.getSheetByName('sys_ID_Fichiers');
    if (!idSheet) { throw new Error("L'onglet 'sys_ID_Fichiers' est introuvable."); }
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
 * Détecte correctement la langue de la réponse initiale de l'utilisateur.
 */
function getOriginalLanguage(reponses) {
  const langueRepondantBrute = reponses['Langue___Language'] || reponses['Langue / Language'] || 'Français';
  const mapLangue = { 'Français': 'FR', 'English': 'EN', 'Español': 'ES', 'Deutsch': 'DE' };
  return mapLangue[langueRepondantBrute] || 'FR';
}

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
        { nom: (T.AXE_SN || "Sensation (S) vs Intuition (N)"),  p1: 'S', p2: 'N' },
        { nom: (T.AXE_TF || "Pensée (T) vs Sentiment (F)"),    p1: 'T', p2: 'F' },
        { nom: (T.AXE_JP || "Jugement (J) vs Perception (P)"),  p1: 'J', p2: 'P' }
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
    Logger.log(`ERREUR CRITIQUE DANS formatScoresDetails (universel): ${e.toString()}`);
    return "Impossible d'afficher le détail des scores en raison d'une erreur.\n";
  }
}

/**
 * Charge les chaînes de caractères traduites pour une langue donnée.
 * @version CORRIGÉE : Utilise .trim() pour ignorer les espaces dans les en-têtes.
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
  // MODIFICATION : Ajout de .trim() pour ignorer les espaces avant/après les noms de colonnes
  const langColIndex = headers.findIndex(h => h && String(h).trim().toLowerCase() === langueCode.toLowerCase());
  if (langColIndex === -1) throw new Error(`La colonne de langue '${langueCode}' est introuvable dans l'onglet "traductions".`);
  const traductions = {};
  const keyColIndex = 0;
  data.forEach(row => {
    if (row[keyColIndex]) { traductions[row[keyColIndex]] = row[langColIndex]; }
  });
  return traductions;
}

function buildAndSendEmails(config, reponse, resultats, langueCode, isDebugMode, destinatairesSurcharge = {}) {
  try {
    const idGabarit = config.ID_Gabarit_Email_Repondant;
    if (!idGabarit) {
      throw new Error("La colonne 'ID_Gabarit_Email_Repondant' n'est pas définie dans la configuration du test.");
    }
    const gabarit = getGabaritEmail(idGabarit, langueCode);
    const T = loadTraductions(langueCode);
    const variables = {
      nom_repondant: reponse.nomRepondant || 'Participant',
      Type_Test: config.Type_Test || '',
      profil_titre: resultats.titreProfil || resultats.profilFinal || '',
      profil_description: resultats.descriptionProfil || 'Aucune description disponible.',
      scores_details: formatScoresDetails(resultats, gabarit.Niveau_Details_Resultats, config.Type_Test, langueCode).replace(/\n/g, '<br>'),
      formateur_nom: config.Formateur_Nom || 'Votre Formateur',
      formateur_consultant: gabarit.formateur_consultant || 'Votre Consultant Certifié'
    };
    let corpsHtml = gabarit.Corps_HTML;
    if (!corpsHtml) {
      throw new Error(`Le gabarit d'e-mail '${idGabarit}' n'a pas de contenu dans la colonne 'Corps_HTML'.`);
    }

    let sujet = gabarit.Sujet;
    for (const [key, value] of Object.entries(variables)) {
      const regex = new RegExp(`\\{${key}\\}`, 'g');
      sujet = sujet.replace(regex, value);
      corpsHtml = corpsHtml.replace(regex, value);
    }

    const piecesJointes = findAttachments(config.Type_Test, resultats.profilFinal, gabarit.Niveau_Pieces_Jointes, langueCode);
    const adressesUniques = new Set();
    const useSurcharge = destinatairesSurcharge && Object.keys(destinatairesSurcharge).length > 0;
    if (useSurcharge) {
      if (destinatairesSurcharge.repondant && reponse.emailRepondant) { adressesUniques.add(reponse.emailRepondant); }
      if (destinatairesSurcharge.formateur && destinatairesSurcharge.formateurEmail) { adressesUniques.add(destinatairesSurcharge.formateurEmail); }
      if (destinatairesSurcharge.patron && destinatairesSurcharge.patronEmail) { adressesUniques.add(destinatairesSurcharge.patronEmail); }
      if (destinatairesSurcharge.test && destinatairesSurcharge.test.trim() !== '') {
        destinatairesSurcharge.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email));
      }
    } else {
      if (config.Repondant_Email_Actif === 'Oui' && reponse.emailRepondant) { adressesUniques.add(reponse.emailRepondant); }
      if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) { adressesUniques.add(config.Patron_Email); }
      if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) { adressesUniques.add(config.Formateur_Email); }
    }
    if (config.Developpeur_Email) { adressesUniques.add(config.Developpeur_Email); }
    adressesUniques.forEach(adresse => {
      try {
        let sujetFinal = sujet;
        if (adresse.toLowerCase() !== (reponse.emailRepondant || "").toLowerCase()) {
          sujetFinal = (T.PREFIXE_COPIE_EMAIL || "Copie : ") + sujet;
        }
        MailApp.sendEmail({
          to: adresse,
          subject: sujetFinal,
          htmlBody: corpsHtml,
          attachments: piecesJointes
        });
      } catch (e) {
        Logger.log(`Echec de l'envoi à ${adresse}. Erreur: ${e.message}`);
      }
    });
  } catch (err) {
    Logger.log("ERREUR CRITIQUE dans buildAndSendEmails : " + err.toString() + "\n" + err.stack);
  }
}

function findAttachments(typeTest, profilCode, niveauPJ, langueCode) {
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
      const typeSheet = row[idx.type] ? row[idx.type].toString().toUpperCase() : '';
      const typeTestUpper = typeTest ? typeTest.toUpperCase() : '';
      const typeMatch = (typeSheet === typeTestUpper);
      const profilSheet = row[idx.profil] ? row[idx.profil].toString().toUpperCase() : '';
      const profilCodeUpper = profilCode ? profilCode.toUpperCase() : '';
      const profilMatch = (profilSheet === profilCodeUpper || profilSheet === 'TOUS');
      const langueSheet = row[idx.langue] ? row[idx.langue].toString().toUpperCase() : '';
      const langueCodeUpper = langueCode ? langueCode.toUpperCase() : '';
      const langueMatch = (langueSheet === langueCodeUpper || langueSheet === 'TOUS');

      const niveauMatch = (row[idx.niveau] > 0 && row[idx.niveau] <= niveauNumRequis);

      if (typeMatch && profilMatch && niveauMatch && langueMatch && row[idx.id]) {
        idsFichiersTrouves.add(row[idx.id]);
      }
    });
    if (idsFichiersTrouves.size === 0) return [];
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

function mapQuestionsById(bdd, nomFeuille) {
  const sheet = bdd.getSheetByName(nomFeuille);
  if (!sheet) { throw new Error(`Feuille de questions '${nomFeuille}' introuvable.`); }
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idCol = headers.indexOf('ID');
  const paramsCol = headers.indexOf('Paramètres (JSON)');
  const mapById = {};
  data.forEach(row => {
    const qId = row[idCol];
    if (qId) {
      mapById[qId] = { id: qId, params: row[paramsCol] };
    }
  });
  return mapById;
}

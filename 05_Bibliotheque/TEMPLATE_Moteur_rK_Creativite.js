/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Moteur_rK_Creativite.gs
 * == VERSION : 3.1 - Correction d'une erreur de syntaxe dans la gestion d'erreur.
 * == RÔLE    : Moteur de calcul dédié pour le test r&K Créativité.
 * - Restaure la logique de calcul complète et correcte.
 * - Assure une recherche de profil robuste.
 * =================================================================================
 */

// Fonction de normalisation de texte (robuste aux accents, majuscules, espaces)
function _crea_normStr(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
}

/**
 * Calcule les résultats pour le test r&K Créativité.
 * @param {object} reponses - L'objet contenant les réponses de l'utilisateur.
 * @param {string} langueCible - Le code de la langue pour les résultats (ex: 'FR').
 * @param {object} config - L'objet de configuration du test.
 * @param {string} langueOrigine - Le code de la langue du formulaire.
 * @returns {object} Un objet contenant tous les résultats, scores et données de profil.
 */
function calculerResultats_rK_Creativite(reponses, langueCible, config, langueOrigine) {
  Logger.log("--- ✅ EXÉCUTION MOTEUR CRÉATIVITÉ VERSION 3.1 (CANARI) CONFIRMÉE ---");
  try {
    const { questionsMap } = _crea_chargerQuestionsAvecAxe(config.Type_Test, _normLang(langueOrigine));
    let scoresParAxe = {
      "Idéation": { r: 0, K: 0, total: 0 }, "Sélection": { r: 0, K: 0, total: 0 },
      "Innovation": { r: 0, K: 0, total: 0 }, "Gestion des contraintes": { r: 0, K: 0, total: 0 },
      "Mise en œuvre": { r: 0, K: 0, total: 0 }
    };
    let total_r_global = 0, total_K_global = 0;

    for (const enTete in reponses) {
      if (!enTete.includes(':')) continue;
      const idQuestion = enTete.split(':')[0].trim();
      const qConfig = questionsMap[idQuestion];
      if (qConfig && reponses[enTete]) {
        let scoreTemp = { scoresData: {} };
        _aiguillerCalcul(qConfig.parametres.mode, reponses[enTete], qConfig.parametres, scoreTemp);
        const score_r = scoreTemp.scoresData.r || 0, score_K = scoreTemp.scoresData.K || 0;
        total_r_global += score_r;
        total_K_global += score_K;
        const axe = qConfig.axe;
        if (axe && scoresParAxe[axe]) {
          scoresParAxe[axe].r += score_r;
          scoresParAxe[axe].K += score_K;
          scoresParAxe[axe].total += score_r + score_K;
        }
      }
    }

    const grand_total_global = total_r_global + total_K_global;
    const pourcentage_r = (grand_total_global > 0) ? (total_r_global / grand_total_global) * 100 : 0;
    const pourcentage_k = 100 - pourcentage_r;

    const profilFinal = _crea_determinerProfil(pourcentage_r);
    const profilsDataBrutes = _crea_chargerDonneesProfils(config.Type_Test, langueCible);
    
    // Logique robuste pour trouver le profil correspondant
    const profilFinalNormalise = _crea_normStr(profilFinal);
    const profilData = profilsDataBrutes.find(row => {
      const codeProfilNormalise = _crea_normStr(row.Code_Profil || row.Profil);
      return codeProfilNormalise === profilFinalNormalise;
    }) || {};

    Logger.log(`[ESPION Moteur Créativité] Profil calculé: "${profilFinal}". Profil trouvé dans la BDD: ${Object.keys(profilData).length > 0 ? 'Oui' : 'NON'}`);
    
    // Assemblage final des données
    const finalData = {
      ...profilData,
      profilFinal: profilFinal,
      Titre_Profil: profilData.Titre_Profil || profilFinal,
      Pourcentage_r: parseFloat(pourcentage_r.toFixed(1)),
      Pourcentage_K: parseFloat(pourcentage_k.toFixed(1)),
      Score_Ideation: scoresParAxe["Idéation"].total > 0 ? parseFloat(((scoresParAxe["Idéation"].r / scoresParAxe["Idéation"].total) * 10).toFixed(1)) : 0,
      Score_Selection: scoresParAxe["Sélection"].total > 0 ? parseFloat(((scoresParAxe["Sélection"].r / scoresParAxe["Sélection"].total) * 10).toFixed(1)) : 0,
      Score_Innovation: scoresParAxe["Innovation"].total > 0 ? parseFloat(((scoresParAxe["Innovation"].r / scoresParAxe["Innovation"].total) * 10).toFixed(1)) : 0,
      Score_Contraintes: scoresParAxe["Gestion des contraintes"].total > 0 ? parseFloat(((scoresParAxe["Gestion des contraintes"].r / scoresParAxe["Gestion des contraintes"].total) * 10).toFixed(1)) : 0,
      Score_MiseenOeuvre: scoresParAxe["Mise en œuvre"].total > 0 ? parseFloat(((scoresParAxe["Mise en œuvre"].r / scoresParAxe["Mise en œuvre"].total) * 10).toFixed(1)) : 0,
    };
    
    finalData.scoresData = {"Pourcentage Exploratoire (r)": finalData.Pourcentage_r, "Pourcentage Structuré (K)": finalData.Pourcentage_K};
    finalData.mapCodeToName = {"Pourcentage Exploratoire (r)": "Pourcentage Exploratoire (r)", "Pourcentage Structuré (K)": "Pourcentage Structuré (K)"};
    
    return finalData;

  } catch (e) {
    Logger.log(`!!!! ERREUR FATALE dans calculerResultats_rK_Creativite !!!!\nMessage : ${e.message}\nStack Trace : ${e.stack}`);
    throw e;
  }
}

/**
 * Détermine le nom du profil de créativité en fonction du score en pourcentage 'r'.
 */
function _crea_determinerProfil(pourcentage_r) {
  if (pourcentage_r >= 80) return "Créativité très exploratoire";
  if (pourcentage_r >= 60) return "Créativité exploratoire";
  if (pourcentage_r >= 41) return "Créativité équilibrée";
  if (pourcentage_r >= 21) return "Créativité structurée";
  return "Créativité très structurée";
}

/**
 * Charge les questions et leurs métadonnées (ID, Paramètres, Axe) depuis la BDD.
 */
function _crea_chargerQuestionsAvecAxe(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Questions_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) throw new Error(`Feuille de questions introuvable: ${nomFeuille}`);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h || '').trim());
    const idCol = headers.indexOf('ID'), paramsCol = headers.indexOf('Paramètres (JSON)'), axeCol = headers.indexOf('Axe');
    if (idCol === -1 || paramsCol === -1 || axeCol === -1) throw new Error("Colonnes 'ID', 'Paramètres (JSON)' ou 'Axe' manquantes.");
    const questionsMap = {};
    data.forEach(row => {
      const id = row[idCol], paramsJSON = row[paramsCol], axe = row[axeCol];
      if (id && paramsJSON && axe) {
        try { questionsMap[id] = { id: id, parametres: JSON.parse(paramsJSON), axe: axe }; } 
        catch (e) { Logger.log(`Erreur de parsing JSON pour la question ID ${id} dans ${nomFeuille}`); }
      }
    });
    return { questionsMap };
  } catch (e) {
    Logger.log("Erreur critique _crea_chargerQuestionsAvecAxe: " + e.message);
    throw e;
  }
}

/**
 * Charge les données brutes des profils depuis la BDD.
 */
function _crea_chargerDonneesProfils(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Profils_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) throw new Error(`Feuille de profils introuvable: '${nomFeuille}'.`);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h || '').trim());
    return data.map(row => {
      let obj = {};
      headers.forEach((header, index) => { if (header) obj[header] = row[index]; });
      return obj;
    });
  } catch (e) {
    Logger.log("Erreur critique dans _crea_chargerDonneesProfils: " + e.message);
    throw e;
  }
}
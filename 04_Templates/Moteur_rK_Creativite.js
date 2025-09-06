/**
 * =================================================================================
 * == FICHIER : Moteur_rK_Creativite.gs
 * == VERSION : 2.2 - Nettoyage automatique des en-têtes de profils
 * == RÔLE    : Moteur de calcul dédié pour le test r&K Créativité.
 * =================================================================================
 */

function calculerResultats_rK_Creativite(reponses, langueCible, config, langueOrigine) {
  try {
    const { questionsMap } = _chargerQuestionsAvecAxe(config.Type_Test, _normLang(langueOrigine));
    let resultatsBruts = { scoresData: {} };
    let scoresParAxe = {
      "Idéation": { r: 0, K: 0, total: 0 }, "Sélection": { r: 0, K: 0, total: 0 },
      "Innovation": { r: 0, K: 0, total: 0 }, "Gestion des contraintes": { r: 0, K: 0, total: 0 },
      "Mise en œuvre": { r: 0, K: 0, total: 0 }
    };

    for (const enTete in reponses) {
      if (!enTete.includes(':')) continue;
      const idQuestion = enTete.split(':')[0].trim();
      const qConfig = questionsMap[idQuestion];
      
      if (qConfig && reponses[enTete]) {
        let scoreTemp = { scoresData: {} };
        _aiguillerCalcul(qConfig.parametres.mode, reponses[enTete], qConfig.parametres, scoreTemp);
        const score_r = scoreTemp.scoresData.r || 0;
        const score_K = scoreTemp.scoresData.K || 0;
        
        resultatsBruts.scoresData.r = (resultatsBruts.scoresData.r || 0) + score_r;
        resultatsBruts.scoresData.K = (resultatsBruts.scoresData.K || 0) + score_K;

        const axe = qConfig.axe;
        if (axe && scoresParAxe[axe]) {
          scoresParAxe[axe].r += score_r;
          scoresParAxe[axe].K += score_K;
          scoresParAxe[axe].total += score_r + score_K;
        }
      }
    }

    const total_r_global = resultatsBruts.scoresData.r || 0;
    const total_K_global = resultatsBruts.scoresData.K || 0;
    const grand_total_global = total_r_global + total_K_global;
    const pourcentage_r = (grand_total_global > 0) ? (total_r_global / grand_total_global) * 100 : 0;

    const profilFinal = _determinerProfilCreativite(pourcentage_r);
    const profilsDataBrutes = _chargerDonneesProfilsBrutes_V2(config.Type_Test, langueCible); // Appel de la nouvelle fonction robuste
    const profilData = profilsDataBrutes.find(row => row.Code_Profil === profilFinal) || {};

    const finalData = {
      ...profilData,
      profilFinal: profilFinal,
      "Pourcentage Exploratoire (r)": parseFloat(pourcentage_r.toFixed(1)),
      "Pourcentage Structuré (K)": parseFloat((100 - pourcentage_r).toFixed(1)),
      Score_Ideation: scoresParAxe["Idéation"].total > 0 ? parseFloat(((scoresParAxe["Idéation"].r / scoresParAxe["Idéation"].total) * 10).toFixed(1)) : 0,
      Score_Selection: scoresParAxe["Sélection"].total > 0 ? parseFloat(((scoresParAxe["Sélection"].r / scoresParAxe["Sélection"].total) * 10).toFixed(1)) : 0,
      Score_Innovation: scoresParAxe["Innovation"].total > 0 ? parseFloat(((scoresParAxe["Innovation"].r / scoresParAxe["Innovation"].total) * 10).toFixed(1)) : 0,
      Score_Contraintes: scoresParAxe["Gestion des contraintes"].total > 0 ? parseFloat(((scoresParAxe["Gestion des contraintes"].r / scoresParAxe["Gestion des contraintes"].total) * 10).toFixed(1)) : 0,
      Score_MiseenOeuvre: scoresParAxe["Mise en œuvre"].total > 0 ? parseFloat(((scoresParAxe["Mise en œuvre"].r / scoresParAxe["Mise en œuvre"].total) * 10).toFixed(1)) : 0,
    };
    
    finalData.scoresData = {
        "Pourcentage Exploratoire (r)": finalData["Pourcentage Exploratoire (r)"],
        "Pourcentage Structuré (K)": finalData["Pourcentage Structuré (K)"]
    };
    finalData.mapCodeToName = {
        "Pourcentage Exploratoire (r)": "Pourcentage Exploratoire (r)",
        "Pourcentage Structuré (K)": "Pourcentage Structuré (K)"
    };

    return finalData;

  } catch (e) {
    Logger.log(`!!!! ERREUR FATALE dans calculerResultats_rK_Creativite !!!!\nMessage : ${e.message}\nStack Trace : ${e.stack}`);
    throw e;
  }
}

function _determinerProfilCreativite(pourcentage_r) {
  if (pourcentage_r >= 80) return "Créativité très exploratoire";
  if (pourcentage_r >= 60) return "Créativité exploratoire";
  if (pourcentage_r >= 41) return "Créativité équilibrée";
  if (pourcentage_r >= 21) return "Créativité structurée";
  return "Créativité très structurée";
}

function _chargerQuestionsAvecAxe(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Questions_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) throw new Error(`Feuille introuvable: ${nomFeuille}`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idCol = headers.indexOf('ID');
    const paramsCol = headers.indexOf('Paramètres (JSON)');
    const axeCol = headers.indexOf('Axe');

    if (idCol === -1 || paramsCol === -1 || axeCol === -1) {
      throw new Error("Colonnes 'ID', 'Paramètres (JSON)' ou 'Axe' manquantes.");
    }

    const questionsMap = {};
    data.forEach(row => {
      const id = row[idCol];
      const paramsJSON = row[paramsCol];
      const axe = row[axeCol];
      if (id && paramsJSON && axe) {
        try {
          questionsMap[id] = {
            id: id,
            parametres: JSON.parse(paramsJSON),
            axe: axe
          };
        } catch (e) { /* ignore parse errors */ }
      }
    });
    return { questionsMap };
  } catch (e) {
    Logger.log("Erreur critique _chargerQuestionsAvecAxe: " + e.message);
    return { questionsMap: {} };
  }
}

/**
 * ## FONCTION UTILITAIRE AMÉLIORÉE ##
 * Charge les données brutes des profils et nettoie les en-têtes.
 */
function _chargerDonneesProfilsBrutes_V2(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Profils_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) {
      Logger.log(`Avertissement: L'onglet de profils '${nomFeuille}' est introuvable.`);
      return [];
    }
    const data = sheet.getDataRange().getValues();
    // La ligne ci-dessous est la correction clé : elle nettoie chaque en-tête.
    const headers = data.shift().map(h => String(h || '').trim());
    
    const jsonData = data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        if (header) obj[header] = row[index];
      });
      return obj;
    });
    return jsonData;
  } catch (e) {
    Logger.log("Erreur critique dans _chargerDonneesProfilsBrutes_V2: " + e.message);
    return [];
  }
}
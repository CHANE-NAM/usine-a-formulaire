/**
 * =================================================================================
 * == FICHIER : Moteur_rK_Creativite.js
 * == VERSION : 3.1 - Définitive
 * == RÔLE    : Moteur de calcul dédié pour le test r&K Créativité.
 * Cette version assure que TOUTES les données du profil (y compris Titre_Profil)
 * sont correctement chargées et fusionnées dans le résultat final.
 * =================================================================================
 */

// ======================= SECTION DE DÉBOGAGE (ESPIONS) =======================
const DEBUG_MODE_CREATIVITE = true; 

function _log_crea(flag, ...args) {
  if (DEBUG_MODE_CREATIVITE && flag) {
    const message = args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : arg).join(' ');
    Logger.log(`[ESPION Créativité V3.1] ${message}`);
  }
}
// =================================================================================


function calculerResultats_rK_Creativite(reponses, langueCible, config, langueOrigine) {
  _log_crea(true, '-> DÉMARRAGE MOTEUR CRÉATIVITÉ');
  try {
    // 1. CHARGEMENT DES QUESTIONS
    const { questionsMap } = _chargerQuestionsAvecAxe(config.Type_Test, _normLang(langueOrigine));
    _log_crea(true, `${Object.keys(questionsMap || {}).length} questions chargées.`);

    // 2. CALCUL DES SCORES BRUTS
    let resultatsBruts = { scoresData: { r: 0, K: 0 } };
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
        
        resultatsBruts.scoresData.r += score_r;
        resultatsBruts.scoresData.K += score_K;

        const axe = qConfig.axe;
        if (axe && scoresParAxe[axe]) {
          scoresParAxe[axe].r += score_r;
          scoresParAxe[axe].K += score_K;
          scoresParAxe[axe].total += score_r + score_K;
        }
      }
    }

    // 3. CALCUL DES POURCENTAGES ET DÉTERMINATION DU PROFIL
    const grand_total_global = resultatsBruts.scoresData.r + resultatsBruts.scoresData.K;
    const pourcentage_r = (grand_total_global > 0) ? (resultatsBruts.scoresData.r / grand_total_global) * 100 : 0;
    const profilFinalCode = _determinerProfilCreativite(pourcentage_r);
    _log_crea(true, `Calcul terminé : %r = ${pourcentage_r.toFixed(1)}% -> Code_Profil = "${profilFinalCode}"`);

    // 4. CHARGEMENT DE TOUTES LES DONNÉES DU PROFIL ASSOCIÉ
    const profilsDataBrutes = _chargerDonneesProfilsBrutes_V2(config.Type_Test, langueCible);
    const profilData = profilsDataBrutes.find(row => row.Code_Profil === profilFinalCode) || {};
    _log_crea(true, `Données du profil chargées. ${Object.keys(profilData).length} colonnes trouvées.`);
    
    // 5. ASSEMBLAGE DE L'OBJET DE RÉSULTATS FINAL
    const finalData = {
      ...profilData, 
      
      profilFinal: profilFinalCode, // On garde le code technique
      // On s'assure que Titre_Profil est bien défini, même s'il est vide dans le sheet
      Titre_Profil: profilData.Titre_Profil || profilFinalCode, 
      
      Pourcentage_r: parseFloat(pourcentage_r.toFixed(1)),
      Pourcentage_K: parseFloat((100 - pourcentage_r).toFixed(1)),
      
      Score_Ideation: scoresParAxe["Idéation"].total > 0 ? parseFloat(((scoresParAxe["Idéation"].r / scoresParAxe["Idéation"].total) * 10).toFixed(1)) : 0,
      Score_Selection: scoresParAxe["Sélection"].total > 0 ? parseFloat(((scoresParAxe["Sélection"].r / scoresParAxe["Sélection"].total) * 10).toFixed(1)) : 0,
      Score_Innovation: scoresParAxe["Innovation"].total > 0 ? parseFloat(((scoresParAxe["Innovation"].r / scoresParAxe["Innovation"].total) * 10).toFixed(1)) : 0,
      Score_Contraintes: scoresParAxe["Gestion des contraintes"].total > 0 ? parseFloat(((scoresParAxe["Gestion des contraintes"].r / scoresParAxe["Gestion des contraintes"].total) * 10).toFixed(1)) : 0,
      Score_MiseenOeuvre: scoresParAxe["Mise en œuvre"].total > 0 ? parseFloat(((scoresParAxe["Mise en œuvre"].r / scoresParAxe["Mise en œuvre"].total) * 10).toFixed(1)) : 0,

      scoresData: {
          "Pourcentage Exploratoire (r)": parseFloat(pourcentage_r.toFixed(1)),
          "Pourcentage Structuré (K)": parseFloat((100 - pourcentage_r).toFixed(1))
      },
      mapCodeToName: {
          "Pourcentage Exploratoire (r)": "Exploratoire (r)",
          "Pourcentage Structuré (K)": "Structuré (K)"
      }
    };
    
    _log_crea(true, '<- FIN MOTEUR. Objet final assemblé et prêt à être envoyé.');
    return finalData;

  } catch (e) {
    Logger.log(`!!!! ERREUR FATALE dans calculerResultats_rK_Creativite !!!!\nMessage : ${e.message}\nStack Trace : ${e.stack}`);
    throw e;
  }
}

// Les fonctions de support restent les mêmes
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
    const headers = data.shift().map(h => String(h || '').trim());
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
          questionsMap[id] = { id: id, parametres: JSON.parse(paramsJSON), axe: axe };
        } catch (e) { /* ignore les erreurs de parsing */ }
      }
    });
    return { questionsMap };
  } catch (e) {
    Logger.log("Erreur critique _chargerQuestionsAvecAxe: " + e.message);
    return { questionsMap: {} };
  }
}

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
/**
 * =================================================================================
 * == FICHIER : Moteur_rK_Resilience.js
 * == VERSION : 1.6 - FINAL : Ajout de Recommandation_Generale dans les résultats.
 * == RÔLE    : Moteur de calcul dédié pour le test r&K Résilience.
 * =================================================================================
 */

function calculerResultats_rK_Resilience(reponses, langueCible, config, langueOrigine) {
  try {
    const langOrigineNorm = _normLang(langueOrigine);
    const questionsMap = _chargerQuestions(config.Type_Test, langOrigineNorm);
    let resultatsBruts = {
      scoresData: {},
      scoresNormalisesParAxe: { Echec: [], Changement: [], Ressources: [], Crise: [], Objectifs: [] }
    };

    // 1. Calcul des scores bruts et répartition
    for (const enTete in reponses) {
      if (!enTete.includes(':')) continue;
      const idQuestion = enTete.split(':')[0].trim();
      const qConfig = questionsMap[idQuestion];
      if (qConfig && reponses[enTete]) {
        const params = qConfig.parametres;
        const reponse = reponses[enTete];
        _aiguillerCalcul(params.mode, reponse, params, resultatsBruts);
        let scoreNormalise = 0;
        if (params.mode.toUpperCase() === 'QCU_CAT') {
          if (params.options && Array.isArray(params.options)) {
            const repNorm = _normStr(reponse);
            const optionChoisie = params.options.find(opt => _normStr(opt.libelle) === repNorm);
            if (optionChoisie && optionChoisie.profil === 'r') scoreNormalise = 10;
            else scoreNormalise = 0;
          }
        } else if (params.mode.toUpperCase() === 'LIKERT_5') {
          const valNum = parseFloat(String(reponse).replace(',', '.'));
          if (!isNaN(valNum)) {
            if (params.profil === 'r') scoreNormalise = ((valNum - 1) / 4) * 10;
            else scoreNormalise = ((5 - valNum) / 4) * 10;
          }
        }
        if (params.axe && resultatsBruts.scoresNormalisesParAxe[params.axe]) {
          resultatsBruts.scoresNormalisesParAxe[params.axe].push(scoreNormalise);
        }
      }
    }

    // 2. Calcul des pourcentages globaux
    const total_r = resultatsBruts.scoresData['r'] || 0;
    const total_K = resultatsBruts.scoresData['K'] || 0;
    const grand_total = total_r + total_K;
    const pourcentage_r = (grand_total > 0) ? (total_r / grand_total) * 100 : 0;
    const pourcentage_K = (grand_total > 0) ? (total_K / grand_total) * 100 : 0;

    // 3. Détermination du niveau de résilience
    const profilsDataBrutes = _chargerDonneesProfilsBrutes(config.Type_Test, langueCible);
    const niveauResilience = _determinerNiveauResilience(pourcentage_r, pourcentage_K, profilsDataBrutes);
    const profilData = profilsDataBrutes.find(row => row.Code_Profil === niveauResilience) || {};
    
    // 4. Calcul des scores moyens par axe
    let axesData = {};
    for (const axe in resultatsBruts.scoresNormalisesParAxe) {
      const scores = resultatsBruts.scoresNormalisesParAxe[axe];
      const scoreMoyen = scores.length > 0 ? (scores.reduce((a, b) => a + b, 0) / scores.length) : 0;
      axesData[axe] = {
        score: scoreMoyen.toFixed(1),
        interpretation: scoreMoyen > 7.5 ? "Point fort" : scoreMoyen >= 4 ? "Équilibré" : "Point de vigilance"
      };
    }
    
    // 5. Assemblage final
    const finalData = {
      Votre_nom_et_prenom: reponses['Votre nom et prenom'] || reponses['Votre_nom_et_prenom'] || 'Non renseigné',
      Date_Rapport: new Date().toLocaleDateString('fr-FR'),
      Score_Global_R: pourcentage_r.toFixed(0),
      Score_Global_K: pourcentage_K.toFixed(0),
      Niveau_Resilience: niveauResilience,
      Titre_Profil: profilData.Titre_Profil || niveauResilience,
      Message_Clef: profilData.Message_Clef || "Message clé non configuré.",
      
      // ====================== CORRECTION AJOUTÉE ICI ======================
      Recommandation_Generale: profilData.Recommandation_Generale || "Recommandation non disponible.",
      // ====================================================================

      Score_Echec: axesData.Echec.score, Interpretation_Echec: axesData.Echec.interpretation, Recommandations_Echec: profilData.Reco_Axe_Echec || "N/A",
      Score_Changement: axesData.Changement.score, Interpretation_Changement: axesData.Changement.interpretation, Recommandations_Changement: profilData.Reco_Axe_Changement || "N/A",
      Score_Ressources: axesData.Ressources.score, Interpretation_Ressources: axesData.Ressources.interpretation, Recommandations_Ressources: profilData.Reco_Axe_Ressources || "N/A",
      Score_Crise: axesData.Crise.score, Interpretation_Crise: axesData.Crise.interpretation, Recommandations_Crise: profilData.Reco_Axe_Crise || "N/A",
      Score_Objectifs: axesData.Objectifs.score, Interpretation_Objectifs: axesData.Objectifs.interpretation, Recommandations_Objectifs: profilData.Reco_Axe_Objectifs || "N/A",
      Reco_Rep_Potentiel: profilData.Reco_Rep_Potentiel || "N/A",
      Reco_Rep_Stress: profilData.Reco_Rep_Stress || "N/A",
      Reco_Rep_Difficulte: profilData.Reco_Rep_Difficulte || "N/A",
      Reco_Manager_Potentiel: profilData.Reco_Manager_Potentiel || "N/A",
      Reco_Manager_Stress: profilData.Reco_Manager_Stress || "N/A",
      Reco_Manager_Difficulte: profilData.Reco_Manager_Difficulte || "N/A",
      Reco_Entourage_Potentiel: profilData.Reco_Entourage_Potentiel || "N/A",
      Reco_Entourage_Stress: profilData.Reco_Entourage_Stress || "N/A",
      Reco_Entourage_Difficulte: profilData.Reco_Entourage_Difficulte || "N/A",
      Action_1: profilData.Action_1 || "Action prioritaire 1 non configurée.",
      Action_2: profilData.Action_2 || "Action prioritaire 2 non configurée.",
      Action_3: profilData.Action_3 || "Action prioritaire 3 non configurée.",
      scoresData: { r: pourcentage_r, K: pourcentage_K },
      profilFinal: niveauResilience,
      mapCodeToName: { r: "Résilience (r)", K: "Stabilité (K)" }
    };
    return finalData;

  } catch (e) {
    Logger.log("!!!! ERREUR FATALE dans calculerResultats_rK_Resilience !!!!\nMessage : " + e.message + "\nStack Trace : " + e.stack);
    throw e;
  }
}

function _determinerNiveauResilience(pourcentage_r, pourcentage_K, profilsData) {
  for (const row of profilsData) {
    const codeProfil = row.Code_Profil;
    const seuilStr = row.Seuil_Score || '';
    const match = seuilStr.match(/([RK])\s*(>=|<=|)\s*(\d{1,2})(-(\d{1,2}))?%/);
    if (match) {
      const profilSeuil = match[1];
      const operateur = match[2];
      const valeur1 = parseInt(match[3], 10);
      const valeur2 = match[5] ? parseInt(match[5], 10) : null;
      const scoreCible = (profilSeuil === 'R') ? pourcentage_r : pourcentage_K;
      if (valeur2 !== null) {
        if (scoreCible >= valeur1 && scoreCible <= valeur2) return codeProfil;
      } else if (operateur === '>=') {
        if (scoreCible >= valeur1) return codeProfil;
      } else if (operateur === '<=') {
        if (scoreCible <= valeur1) return codeProfil;
      }
    }
  }
  return "Indéterminé";
}

function _chargerDonneesProfilsBrutes(typeTest, langue) {
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
    const headers = data.shift();
    const jsonData = data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        if (header) obj[header] = row[index];
      });
      return obj;
    });
    return jsonData;
  } catch (e) {
    Logger.log("Erreur critique dans _chargerDonneesProfilsBrutes: " + e.message);
    return [];
  }
}
/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Moteur_rK.gs
 * == RÔLE    : Moteur de calcul pour les tests de la famille r&K.
 * =================================================================================
 */

/**
 * Moteur de calcul pour le test r&K Résilience.
 */
function calculerResultats_rK_Resilience(reponsesUtilisateur, langueCible, config, langueOrigine) {
  _log(DEBUG_FLOW, '-> Moteur r&K Résilience');
  return _executerCalcul_rK(reponsesUtilisateur, langueCible, config, langueOrigine);
}

/**
 * Moteur de calcul pour le test r&K Environnement.
 */
function calculerResultats_rK_Environnement(reponsesUtilisateur, langueCible, config) {
  _log(DEBUG_FLOW, '-> Moteur r&K Environnement');
  return _executerCalcul_rK(reponsesUtilisateur, langueCible, config, 'FR'); // Langue origine non pertinente ici
}

/**
 * Moteur de calcul pour le test r&K Créativité.
 */
function calculerResultats_rK_Creativite(reponsesUtilisateur, langueCible, config, langueOrigine) {
  _log(DEBUG_FLOW, '-> Moteur r&K Créativité');
  return _executerCalcul_rK(reponsesUtilisateur, langueCible, config, langueOrigine);
}

/**
 * Moteur de calcul générique pour les tests r&K basés sur des axes.
 */
function _executerCalcul_rK(reponsesUtilisateur, langueCible, config, langueOrigine) {
  let resultats = { scoresData: {} };
  const langCibleNorm = _normLang(langueCible);
  const langOrigineNorm = _normLang(langueOrigine);
  const questionsMapOrigine = _chargerQuestions(config.Type_Test, langOrigineNorm);
  if (!questionsMapOrigine) {
    Logger.log(`ERREUR FATALE: Impossible de charger les questions pour ${config.Type_Test}_${langOrigineNorm}.`);
    return resultats;
  }

  // Calcul des scores par axe
  for (const enTeteComplet in reponsesUtilisateur) {
    const idQuestion = enTeteComplet.split(':')[0].trim();
    const questionConfig = questionsMapOrigine[idQuestion];
    if (questionConfig && questionConfig.parametres.axes && reponsesUtilisateur[enTeteComplet]) {
      const valeurNumerique = parseFloat(String(reponsesUtilisateur[enTeteComplet]).replace(',', '.'));
      if (!isNaN(valeurNumerique)) {
        questionConfig.parametres.axes.forEach(axe => {
          const axeTrim = axe.trim();
          resultats.scoresData[axeTrim] = (resultats.scoresData[axeTrim] || 0) + valeurNumerique;
        });
      }
    }
  }

  // Détermination du profil final et chargement des données de profil
  if (Object.keys(resultats.scoresData).length > 0) {
    const profilEtReco = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibleNorm);
    resultats = { ...resultats, ...profilEtReco };
    const profilsMap = _chargerProfils(config.Type_Test, langCibleNorm);
    const infosProfilComplet = profilsMap[resultats.profilFinal];

    if (infosProfilComplet) {
      resultats = { ...resultats, ...infosProfilComplet };
    }
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  return resultats;
}
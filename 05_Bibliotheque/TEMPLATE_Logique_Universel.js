/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Logique_Universel.gs
 * == VERSION : 12.1 - Ajout d'espions de comparaison pour le débogage final.
 * == RÔLE    : Aiguilleur principal vers les moteurs de calcul.
 * =================================================================================
 */

// ======================= SECTION DE DÉBOGAGE (ESPIONS) =======================
const DEBUG_MODE = true;
const DEBUG_DATA_LOADING = true;
const DEBUG_FLOW = false;
const DEBUG_SCORING = false;

/**
 * Fonction utilitaire pour l'affichage conditionnel des logs de débogage.
 */
function _log(flag, ...args) {
  if (DEBUG_MODE && flag) {
    const message = args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : arg).join(' ');
    Logger.log(`[ESPION] ${message}`);
  }
}
// =================================================================================

/**
 * POINT D'ENTRÉE PRINCIPAL
 * Aiguille vers le bon moteur de calcul en fonction de la configuration.
 */
function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  // On nettoie la valeur de Moteur_Calcul pour éviter les erreurs dues aux espaces.
  const moteur = String(config.Moteur_Calcul || '').trim();
  _log(DEBUG_FLOW, `-> calculerResultats : Démarrage pour Moteur_Calcul="${moteur}".`);

  // ESPION DE DÉBOGAGE FINAL pour vérifier la comparaison stricte.
  Logger.log(`[COMPARAISON] Est-ce que "${moteur}" === "r&K_Creativite" ? Réponse : ${moteur === 'r&K_Creativite'}`);
  Logger.log(`[COMPARAISON] Longueur de moteur: ${moteur.length} vs Longueur attendue: 14`);

  // --- Aiguillage vers les moteurs de calcul spécifiques ---
  // On utilise maintenant config.Moteur_Calcul comme aiguilleur.
  switch (moteur) {
    case 'r&K_Resilience':
      return calculerResultats_rK_Resilience(reponsesUtilisateur, langueCible, config, langueOrigine);
    case 'r&K_Environnement':
      return calculerResultats_rK_Environnement(reponsesUtilisateur, langueCible, config);
    case 'r&K_Creativite':
      return calculerResultats_rK_Creativite(reponsesUtilisateur, langueCible, config, langueOrigine);
  }

  // --- Calcul standard pour les autres tests (MBTI, Couleurs, ANCRES, et r&K Adaptabilité) ---
  let resultats = { scoresData: {}, sousTotauxParMode: {} };
  const langCibleNorm = _normLang(langueCible);
  const langOrigineNorm = _normLang(langueOrigine);
  const questionsMapOrigine = _chargerQuestions(config.Type_Test, langOrigineNorm);

  if (!questionsMapOrigine) {
    Logger.log(`ERREUR FATALE: Impossible de charger les questions de la langue d'origine ${langOrigineNorm}.`);
    return resultats;
  }

  _executerCalculStandard(reponsesUtilisateur, questionsMapOrigine, resultats, config.nbQuestions);

  // Cas particulier pour r&K Adaptabilité en attendant son moteur dédié.
  if (config.Type_Test === 'r&K_Adaptabilite') {
    const total_r = resultats.scoresData['r'] || 0;
    const total_K = resultats.scoresData['K'] || 0;
    const grand_total = total_r + total_K;
    let pourcentage_r = (grand_total > 0) ? (total_r / grand_total) * 100 : 0;
    let pourcentage_K = (grand_total > 0) ? (total_K / grand_total) * 100 : 0;
    resultats.scoresData = { r: pourcentage_r, K: pourcentage_K };
  }

  resultats.scoresMaxPossible = _calculerScoresMaxPossibles(config.Type_Test, langOrigineNorm);

  if (Object.keys(resultats.scoresData).length > 0) {
    const profilEtReco = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibleNorm);
    resultats = { ...resultats, ...profilEtReco };
    const profilsMap = _chargerProfils(config.Type_Test, langCibleNorm);
    const infosProfilComplet = profilsMap[resultats.profilFinal];

    _log(DEBUG_DATA_LOADING, 'PROFIL RECHERCHÉ :', `"${resultats.profilFinal}"`);
    _log(DEBUG_DATA_LOADING, 'PROFILS DISPONIBLES :', Object.keys(profilsMap));
    _log(DEBUG_DATA_LOADING, 'DONNÉES TROUVÉES :', infosProfilComplet);

    if (infosProfilComplet) {
      resultats = { ...resultats, ...infosProfilComplet };
    }
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  _log(DEBUG_FLOW, `<- calculerResultats : Terminé. Profil Final: "${resultats.profilFinal}".`);
  return resultats;
}
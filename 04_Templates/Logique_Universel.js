/**
 * =================================================================================
 * == FICHIER : Logique_Universel.gs
 * == VERSION : 18.3 - Ajout d'un système de débogage avec interrupteurs
 * == RÔLE    : Aiguilleur principal et conteneur des logiques de calcul standards.
 * =================================================================================
 */

// ======================= SECTION DE DÉBOGAGE (ESPIONS) =======================
const DEBUG_MODE = true; // INTERRUPTEUR GÉNÉRAL : Mettre à false pour désactiver TOUS les espions.

// --- Interrupteurs spécifiques ---
const DEBUG_DATA_LOADING = true;  // (RECOMMANDÉ) Espionne le chargement des données de profil (notre bug actuel).
const DEBUG_FLOW = false;         // Espionne le flux général (entrée/sortie des fonctions).
const DEBUG_SCORING = false;      // Espionne le calcul des scores question par question.

/**
 * Fonction utilitaire pour l'affichage conditionnel des logs de débogage.
 */
function _log(flag, ...args) {
  if (DEBUG_MODE && flag) {
    const message = args.map(arg => typeof arg === 'object' ? JSON.stringify(arg) : arg).join(' ');
    Logger.log(`[ESPION] ${message}`);
  }
}
// =================================================================================


/**
 * POINT D'ENTRÉE PRINCIPAL (REMANIÉ)
 * Aiguille vers le bon moteur de calcul en fonction du Type_Test.
 */
function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  _log(DEBUG_FLOW, `-> calculerResultats : Démarrage pour Type_Test="${config.Type_Test}".`);

  // --- Aiguillage vers les moteurs de calcul spécifiques et complexes ---
  if (config.Type_Test === 'r&K_Resilience') {
    return calculerResultats_rK_Resilience(reponsesUtilisateur, langueCible, config, langueOrigine);
  }
  
  if (config.Type_Test === 'r&K_Environnement') {
    return calculerResultats_rK_Environnement(reponsesUtilisateur, langueCible, config);
  }

  if (config.Type_Test === 'r&K_Creativite') {
    return calculerResultats_rK_Creativite(reponsesUtilisateur, langueCible, config, langueOrigine);
  }

  // --- Appels prospectifs pour les futurs moteurs de calcul ---
  if (config.Type_Test === 'r&K_Adaptabilite') {
    Logger.log(`Aiguillage vers Moteur_rK_Adaptabilite (à créer). Calcul standard appliqué en attendant.`);
    // return calculerResultats_rK_Adaptabilite(reponsesUtilisateur, langueCible, config, langueOrigine);
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
  
  _executerCalcul(reponsesUtilisateur, questionsMapOrigine, resultats, config.nbQuestions);
  
  // Pour le test r&K Adaptabilité, on calcule un pourcentage en attendant son moteur dédié.
  if (config.Type_Test === 'r&K_Adaptabilite') {
    const total_r = resultats.scoresData['r'] || 0;
    const total_K = resultats.scoresData['K'] || 0;
    const grand_total = total_r + total_K;
    let pourcentage_r = 0, pourcentage_K = 0;
    if (grand_total > 0) {
      pourcentage_r = (total_r / grand_total) * 100;
      pourcentage_K = (total_K / grand_total) * 100;
    }
    resultats.scoresData = { r: pourcentage_r, K: pourcentage_K };
  }
  
  resultats.scoresMaxPossible = _calculerScoresMaxPossibles(config.Type_Test, langOrigineNorm);
  
  if (Object.keys(resultats.scoresData).length > 0) {
    const profilEtReco = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibleNorm);
    resultats = { ...resultats, ...profilEtReco };
    const profilsMap = _chargerProfils(config.Type_Test, langCibleNorm);
    const infosProfilComplet = profilsMap[resultats.profilFinal];

    // --- ESPIONS POUR LE BUG DE CORRESPONDANCE ---
    _log(DEBUG_DATA_LOADING, 'PROFIL RECHERCHÉ :', `"${resultats.profilFinal}"`);
    _log(DEBUG_DATA_LOADING, 'PROFILS DISPONIBLES :', Object.keys(profilsMap));
    _log(DEBUG_DATA_LOADING, 'DONNÉES TROUVÉES :', infosProfilComplet);
    // --- FIN DES ESPIONS ---

    if (infosProfilComplet) {
      resultats = { ...resultats, ...infosProfilComplet };
    }
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  _log(DEBUG_FLOW, `<- calculerResultats : Terminé. Profil Final: "${resultats.profilFinal}".`);
  return resultats;
}


// =================================================================================
// SECTION - MOTEURS DE CALCUL STANDARDS ET COMMUNS
// =================================================================================

function _executerCalcul(reponses, questionsMap, resultats, nbQuestionsLimite) {
  let questionsTraitees = 0;
  const limite = nbQuestionsLimite || Object.keys(questionsMap).length;
  for (const enTeteComplet in reponses) {
    if (questionsTraitees >= limite) break;
    if (!enTeteComplet.includes(':')) continue;

    const idQuestion = enTeteComplet.split(':')[0].trim();
    const questionConfig = questionsMap[idQuestion];
    if (questionConfig && reponses[enTeteComplet]) {
      _aiguillerCalcul(questionConfig.parametres.mode, reponses[enTeteComplet], questionConfig.parametres, resultats);
      questionsTraitees++;
    }
  }
}

function _aiguillerCalcul(mode, reponse, parametres, resultats) {
  var m = String(mode || '').replace(/\s+/g, ' ').trim().toUpperCase();
  _log(DEBUG_SCORING, `Aiguillage : mode="${m}", reponse="${reponse}"`);
  switch (m) {
    case 'QCU_CAT':      _traiterQCU_CAT(reponse, parametres, resultats);    break;
    case 'ECHELLE_NOTE': _traiterECHELLE_NOTE(reponse, parametres, resultats); break;
    case 'LIKERT_5':     _traiterECHELLE_NOTE(reponse, parametres, resultats); break;
    default:
      break;
  }
}

function _traiterQCU_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;
  const repNorm = _normStr(reponseUtilisateur);
  let optionTrouvee = parametres.options.find(opt => _normStr(opt.libelle) === repNorm);
  if (optionTrouvee && optionTrouvee.profil) {
    const profil = optionTrouvee.profil;
    const valeur = (typeof optionTrouvee.valeur === 'number') ? optionTrouvee.valeur : 1;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
    _log(DEBUG_SCORING, `_traiterQCU_CAT : Ajout de ${valeur} au profil ${profil}. Total: ${resultats.scoresData[profil]}`);
  }
}

function _traiterECHELLE_NOTE(reponseUtilisateur, parametres, resultats) {
  let profil = parametres.profil;
  if (!profil && parametres.options && parametres.options[0] && parametres.options[0].profil) {
    profil = parametres.options[0].profil;
  }
  if (!profil) return;
  const valeurNumerique = parseFloat(String(reponseUtilisateur).replace(',', '.'));
  if (!isNaN(valeurNumerique)) {
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeurNumerique;
    _log(DEBUG_SCORING, `_traiterECHELLE_NOTE : Ajout de ${valeurNumerique} au profil ${profil}. Total: ${resultats.scoresData[profil]}`);
  }
}


// =================================================================================
// SECTION - UTILITAIRES DE PROFIL, DONNÉES ET NORMALISATION
// =================================================================================

function _determinerProfilFinal(scoresData, typeTest, langue) {
  if (!scoresData || Object.keys(scoresData).length === 0) return { profilFinal: "" };
  
  if (typeTest === 'r&K_Environnement') {
    return _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue);
  }
  
  if (String(typeTest || '').toUpperCase().startsWith('MBTI')) {
    let profil = "";
    profil += (scoresData.E || 0) > (scoresData.I || 0) ? 'E' : 'I';
    profil += (scoresData.S || 0) > (scoresData.N || 0) ? 'S' : 'N';
    profil += (scoresData.T || 0) > (scoresData.F || 0) ? 'T' : 'F';
    profil += (scoresData.J || 0) > (scoresData.P || 0) ? 'J' : 'P';
    return { profilFinal: profil };
  } else {
    const profilFinal = Object.keys(scoresData).reduce((a, b) => scoresData[a] > scoresData[b] ? a : b);
    return { profilFinal: profilFinal };
  }
}

function _chargerProfils(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Profils_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    // LIGNE CORRIGÉE : On nettoie les en-têtes pour les rendre robustes.
    const headers = data.shift().map(h => String(h || '').trim()); 
    
    const profilsMap = {};
    let codeColIndex = headers.indexOf('Code_Profil');
    if (codeColIndex === -1) {
       // Tentative de fallback pour une ancienne nomenclature
       const fallbackIndex = headers.indexOf('Profil');
       if (fallbackIndex === -1) {
         _log(DEBUG_DATA_LOADING, "ERREUR _chargerProfils : Colonne 'Code_Profil' ou 'Profil' introuvable.");
         return {};
       }
       codeColIndex = fallbackIndex;
    }

    data.forEach(row => {
      const codeProfil = row[codeColIndex];
      if (codeProfil) {
        profilsMap[codeProfil] = {};
        headers.forEach((header, index) => {
          if (header) profilsMap[codeProfil][header] = row[index];
        });
      }
    });
    return profilsMap;
  } catch (e) {
    Logger.log("Erreur critique _chargerProfils: " + e.message);
    return {};
  }
}

function _creerMapCodeVersNom(profilsMap) {
  const map = {};
  for (const code in profilsMap) {
    map[code] = profilsMap[code].Titre_Profil || profilsMap[code].titre || code;
  }
  return map;
}

function _chargerQuestions(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Questions_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) throw new Error(`Feuille introuvable: ${nomFeuille}`);
    const data = sheet.getDataRange().getValues();
    const headersRaw = data.shift();
    const headers = (headersRaw || []).map(h => String(h || '').replace(/^\uFEFF/, '').replace(/^"|"$/g, '').trim());
    const idCol     = headers.indexOf('ID');
    const paramsCol = headers.indexOf('Paramètres (JSON)');
    if (idCol === -1 || paramsCol === -1) throw new Error("Colonnes ID ou 'Paramètres (JSON)' manquantes.");

    const questionsMap = {};
    data.forEach(row => {
      const id = row[idCol];
      const paramsJSON = row[paramsCol];
      if (id && paramsJSON) {
        try {
          const parametres = JSON.parse(paramsJSON);
          if (parametres.mode) { questionsMap[id] = { id: id, parametres: parametres }; }
        } catch (e) { Logger.log(`Erreur parsing JSON pour ID '${id}': ${e.message}`); }
      }
    });
    return questionsMap;
  } catch (e) {
    Logger.log("Erreur critique _chargerQuestions: " + e.message);
    return null;
  }
}

function _calculerScoresMaxPossibles(typeTest, langue) {
  const questionsMap = _chargerQuestions(typeTest, langue);
  if (!questionsMap) {
    Logger.log(`AVERTISSEMENT: Impossible de charger les questions pour ${typeTest}_${langue} pour calculer les scores max.`);
    return {};
  }
  const maxScores = {};
  for (const id in questionsMap) {
    const question = questionsMap[id];
    const params = question.parametres;
    const mode = (params.mode || '').toUpperCase();
    if (mode === 'ECHELLE_NOTE' || mode === 'LIKERT_5') {
      const profil = params.profil;
      const maxValue = params.echelle_max || params.max || 0;
      if (profil && typeof maxValue === 'number') {
        maxScores[profil] = (maxScores[profil] || 0) + maxValue;
      }
    } else if (mode === 'QCU_CAT' || mode === 'QRM_CAT') {
      if (params.options && Array.isArray(params.options)) {
        let maxContributionParProfil = {};
        params.options.forEach(opt => {
          if (opt.profil) {
            const valeur = (typeof opt.valeur === 'number') ? opt.valeur : 1;
            if (mode === 'QCU_CAT') {
               maxContributionParProfil[opt.profil] = Math.max(maxContributionParProfil[opt.profil] || 0, valeur);
            } else { // QRM_CAT
               maxContributionParProfil[opt.profil] = (maxContributionParProfil[opt.profil] || 0) + valeur;
            }
          }
        });
        for (const profil in maxContributionParProfil) {
          maxScores[profil] = (maxScores[profil] || 0) + maxContributionParProfil[profil];
        }
      }
    }
  }
  return maxScores;
}

function _normStr(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[\u2019\u2018]/g, "'").replace(/[\u201C\u201D]/g, '"')
    .replace(/[«»]/g, '').replace(/[\u2013\u2014]/g, '-')
    .replace(/\u00A0/g, ' ').replace(/\s+/g, ' ')
    .trim().toLowerCase();
}

function _normLang(s) {
  const x = _normStr(s);
  if (!x) return 'FR'; // Fallback sur FR si vide
  if (/^fr|fran|french/.test(x)) return 'FR';
  if (/^en|angl|english|uk|us/.test(x)) return 'EN';
  return x.toUpperCase();
}
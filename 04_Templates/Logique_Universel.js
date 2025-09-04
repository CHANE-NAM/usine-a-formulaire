// =================================================================================
// == FICHIER : Logique_Universel.gs
// == VERSION : 16.0 - Généralisation du calcul du score total possible.
// ==           (Précédent: 15.1 - Correction du bug de lecture des réponses pour le rapport PDF r&K Env.)
// =================================================================================

// --- MOTEUR DE RECOMMANDATION STANDARD "r&K" ---
/**
 * Moteur de détermination de profil par seuils (Utilisé uniquement par r&K_Environnement).
 */
function _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue) {
    Logger.log(`[ESPION][r&K] Démarrage du moteur de recommandation pour le test "${typeTest}".`);
    try {
        const totalPoints = Object.values(scoresData).reduce((sum, val) => sum + (Number(val) || 0), 0);
        if (totalPoints === 0) { return { profilFinal: "Indéterminé", Recommandation: "" }; }

        const profilMajoritaireCode = Object.keys(scoresData).reduce((a, b) => (scoresData[a] || 0) > (scoresData[b] || 0) ? a : b);
        const scoreMajoritaire = scoresData[profilMajoritaireCode] || 0;
        Logger.log(`[ESPION][r&K] Profil dominant: "${profilMajoritaireCode}" avec un score de ${scoreMajoritaire.toFixed(1)}%.`);

        const systemIds = getSystemIds();
        const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
        const nomFeuilleProfils = `Profils_${typeTest}_${langue}`;
        const sheetProfils = bdd.getSheetByName(nomFeuilleProfils);
        if (!sheetProfils) {
          Logger.log(`[ESPION][r&K] AVERTISSEMENT: Feuille de profils introuvable: "${nomFeuilleProfils}".`);
          return { profilFinal: profilMajoritaireCode, Recommandation: "" };
        }
        
        const data = sheetProfils.getDataRange().getValues();
        const headers = data.shift().map(h => String(h || '').trim());
        const idx = { profil: headers.indexOf('Code_Profil') > -1 ? headers.indexOf('Code_Profil') : headers.indexOf('Profil'), seuil: headers.indexOf('Seuil_Score'), destinataire: headers.indexOf('Destinataire'), axe: headers.indexOf('Axe'), reco: headers.indexOf('Recommandation') };
        
        for (const row of data) {
            const dest = String(row[idx.destinataire] || '').trim();
            const axe = String(row[idx.axe] || '').trim();
            const seuilStr = String(row[idx.seuil] || '').trim();
            
            if (dest === 'Répondant' && axe === 'Développer potentiel' && _parseSeuilScore_rK(seuilStr, profilMajoritaireCode, scoreMajoritaire)) {
                const profilFinalTrouve = String(row[idx.profil] || profilMajoritaireCode);
                const recommandationTrouvee = String(row[idx.reco] || '');
                return { profilFinal: profilFinalTrouve, Recommandation: recommandationTrouvee };
            }
        }
        return { profilFinal: profilMajoritaireCode, Recommandation: "" };
    } catch (e) {
        Logger.log("ERREUR CRITIQUE dans _determinerProfilFinalParSeuils_rK: " + e.message);
        return { profilFinal: "Erreur de calcul", Recommandation: "" };
    }
}

function _parseSeuilScore_rK(seuilStr, codeProfilMajoritaire, scorePourcentage) {
    if (!seuilStr || !codeProfilMajoritaire) return false;
    const seuil = String(seuilStr).trim();
    const profilSeuilMatch = seuil.toUpperCase().split(' ')[0];
    if (profilSeuilMatch !== codeProfilMajoritaire.toUpperCase()) return false;
    const matchSimple = seuil.match(/(>=|<=)\s*(\d+)/);
    if (matchSimple) {
        const operateur = matchSimple[1]; const valeurSeuil = parseInt(matchSimple[2], 10);
        if (operateur === '>=') return scorePourcentage >= valeurSeuil;
        if (operateur === '<=') return scorePourcentage <= valeurSeuil;
    }
    const matchPlage = seuil.match(/(\d+)-(\d+)/);
    if (matchPlage) {
        const min = parseInt(matchPlage[1], 10); const max = parseInt(matchPlage[2], 10);
        return scorePourcentage >= min && scorePourcentage <= max;
    }
    return false;
}

/**
 * Moteur de calcul DÉDIÉ pour le test r&K_Environnement (v5.1 - Corrigé).
 * Calcule les pourcentages ET les données détaillées par thème pour le rapport PDF.
 */
function _calculerResultats_rK_Environnement_dedie(reponsesUtilisateur, questionsMap) {
    Logger.log('[ESPION][ENV] Moteur de calcul r&K (v5.1 - Corrigé) activé.');
    
    // --- 1. Calcul des points bruts (logique data-driven) ---
    let totaux_points = {};
    for (const enTete in reponsesUtilisateur) {
        if (String(enTete).startsWith('ENV')) {
            const idQuestion = enTete.split(':')[0].trim();
            const questionConfig = questionsMap[idQuestion];
            const reponseNum = parseInt(reponsesUtilisateur[enTete], 10);
            const scoringModel = questionConfig ? questionConfig.parametres.scoring_model : null;
            if (!isNaN(reponseNum) && scoringModel) {
                const echelle = scoringModel.echelle || 11;
                const profilDirect = scoringModel.direct;
                const profilInverse = scoringModel.inverse;
                if (profilDirect) { totaux_points[profilDirect] = (totaux_points[profilDirect] || 0) + reponseNum; }
                if (profilInverse) { totaux_points[profilInverse] = (totaux_points[profilInverse] || 0) + (echelle - reponseNum); }
            }
        }
    }
    const total_K_points = totaux_points['K'] || 0;
    const total_r_points = totaux_points['r'] || 0;
    const grand_total_points = total_K_points + total_r_points;

    // --- 2. Calcul des pourcentages globaux ---
    let pourcentage_K = 0, pourcentage_r = 0;
    if (grand_total_points > 0) {
        pourcentage_K = (total_K_points / grand_total_points) * 100;
        pourcentage_r = (total_r_points / grand_total_points) * 100;
    }

    // --- 3. Calcul détaillé par thème pour le rapport PDF ---
    const THEMES = ["Concurrence & Pression du marché", "Clients & Demande", "Technologies & Innovation", "Réglementation & Cadre juridique", "Ressources humaines & Compétences", "Financement & Accès aux capitaux", "Fournisseurs & Logistique", "Ressources & Infrastructures matérielles", "Image & Réputation sectorielle", "Partenariats & Réseaux", "Territoire & Environnement géographique", "Tendances sociétales & culturelles", "Contexte économique global", "Risques & Sécurité", "Opportunités de croissance & Marchés"];
    const interpK = x => x >= 7 ? "Environnement plutôt stable et prévisible" : x <= 3 ? "Environnement plutôt instable / changeant" : "Stabilité modérée avec quelques variations";
    const interpr = x => x >= 7 ? "Changements rapides / forte dynamique" : x <= 3 ? "Changements lents / faible dynamique" : "Vitesse de changement modérée";
    
    let flat = {};
    let sumK_themes = 0, sumr_themes = 0, countK_themes = 0, countr_themes = 0;

    for (let t = 0; t < 15; t++) {
        let k_vals = [], r_vals = [];
        for (let i = 1; i <= 4; i++) {
            const qNum = t * 4 + i;
            const id = 'ENV' + String(qNum).padStart(3, '0');
            const questionConfig = questionsMap[id];
            
            // --- DEBUT DE LA CORRECTION ---
            // Recherche de la réponse de l'utilisateur de manière robuste, sans dépendre de "titre_court".
            const reponse = reponsesUtilisateur[Object.keys(reponsesUtilisateur).find(k => k.startsWith(id))];
            // --- FIN DE LA CORRECTION ---
            
            if (reponse != null && !isNaN(reponse) && questionConfig) {
                if (questionConfig.parametres.dimension === 'Stabilité') k_vals.push(Number(reponse));
                else if (questionConfig.parametres.dimension === 'Vitesse') r_vals.push(Number(reponse));
            }
        }

        const avgK = k_vals.length > 0 ? k_vals.reduce((a, b) => a + b, 0) / k_vals.length : null;
        const avgr = r_vals.length > 0 ? r_vals.reduce((a, b) => a + b, 0) / r_vals.length : null;

        if(avgK != null) { sumK_themes += avgK; countK_themes++; }
        if(avgr != null) { sumr_themes += avgr; countr_themes++; }

        const n = t + 1;
        flat[`Nom_Theme_${n}`] = THEMES[t];
        flat[`Score_Stabilite_Theme_${n}`] = avgK != null ? avgK.toFixed(1) : "";
        flat[`Interpretation_Stabilite_Theme_${n}`] = avgK != null ? interpK(avgK) : "";
        flat[`Score_Vitesse_Theme_${n}`] = avgr != null ? avgr.toFixed(1) : "";
        flat[`Interpretation_Vitesse_Theme_${n}`] = avgr != null ? interpr(avgr) : "";
        flat[`Recommandations_Theme_${n}`] = ""; // Placeholder pour le futur
    }

    const scoreK_global = countK_themes > 0 ? (sumK_themes / countK_themes) : 0;
    const scorer_global = countr_themes > 0 ? (sumr_themes / countr_themes) : 0;

    flat['Score_Stabilite'] = scoreK_global.toFixed(1);
    flat['Interpretation_Stabilite'] = interpK(scoreK_global);
    flat['Score_Vitesse'] = scorer_global.toFixed(1);
    flat['Interpretation_Vitesse'] = interpr(scorer_global);

    Logger.log(`[ESPION][ENV] Données PDF générées. Score K global: ${scoreK_global.toFixed(1)}, Score r global: ${scorer_global.toFixed(1)}`);

    return { 
        scoresData: { K: pourcentage_K, r: pourcentage_r }, // Pour l'email
        ...flat // Pour le PDF
    };
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
  if (!x) return '';
  if (/^fr|fran|french/.test(x)) return 'FR';
  if (/^en|angl|english|uk|us/.test(x)) return 'EN';
  return x.toUpperCase();
}


// ======================= DÉBUT BLOC AJOUTÉ =======================
/**
 * Calcule les scores maximums possibles pour chaque profil d'un test.
 * @param {string} typeTest - Le type de test (ex: 'ANCRES').
 * @param {string} langue - Le code langue (ex: 'FR').
 * @returns {object} Un objet associant chaque code de profil à son score maximum.
 */
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
        // Pour une question, on identifie la contribution maximale à chaque profil.
        let maxContributionParProfil = {};
        
        params.options.forEach(opt => {
          if (opt.profil) {
            const valeur = (typeof opt.valeur === 'number') ? opt.valeur : 1;
            // Pour un QCU, un profil ne peut recevoir que la plus haute valeur de cette question.
            // Pour un QRM, un profil peut recevoir la somme des valeurs (si plusieurs options cochables pour un même profil).
            if (mode === 'QCU_CAT') {
               maxContributionParProfil[opt.profil] = Math.max(maxContributionParProfil[opt.profil] || 0, valeur);
            } else { // QRM_CAT
               maxContributionParProfil[opt.profil] = (maxContributionParProfil[opt.profil] || 0) + valeur;
            }
          }
        });

        // On ajoute ces contributions maximales aux totaux.
        for (const profil in maxContributionParProfil) {
          maxScores[profil] = (maxScores[profil] || 0) + maxContributionParProfil[profil];
        }
      }
    }
  }
  Logger.log(`Scores max possibles calculés pour ${typeTest}: ${JSON.stringify(maxScores)}`);
  return maxScores;
}
// ======================= FIN BLOC AJOUTÉ =======================


function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  Logger.log(`Démarrage du calcul des résultats pour le Type_Test: "${config.Type_Test}". Langue Origine: ${langueOrigine}, Langue Cible: ${langueCible}`);
  let resultats = { scoresData: {}, sousTotauxParMode: {} };
  
  const langCibleNorm = _normLang(langueCible);
  const langOrigineNorm = _normLang(langueOrigine);
  
  if (config.Type_Test === 'r&K_Environnement') {
    const questionsMap = _chargerQuestions(config.Type_Test, langOrigineNorm);
    resultats = _calculerResultats_rK_Environnement_dedie(reponsesUtilisateur, questionsMap);
  } else {
    const questionsMapOrigine = _chargerQuestions(config.Type_Test, langOrigineNorm);
    if (!questionsMapOrigine) {
      Logger.log(`ERREUR FATALE: Impossible de charger les questions de la langue d'origine ${langOrigineNorm}.`);
      return resultats;
    }
    _executerCalcul(reponsesUtilisateur, questionsMapOrigine, resultats, config.nbQuestions);
  }
  
  const testsPourcentage_rK = ['r&K_Resilience', 'r&K_Adaptabilite', 'r&K_Creativite'];
  if (testsPourcentage_rK.includes(config.Type_Test)) {
      const total_r = resultats.scoresData['r'] || 0;
      const total_K = resultats.scoresData['K'] || 0;
      const grand_total = total_r + total_K;
      let pourcentage_r = 0, pourcentage_K = 0;
      if (grand_total > 0) {
          pourcentage_r = (total_r / grand_total) * 100;
          pourcentage_K = (total_K / grand_total) * 100;
      }
      resultats.scoresData = { r: pourcentage_r, K: pourcentage_K };
      Logger.log(`[ESPION r&K %] Conversion en pourcentage : r=${pourcentage_r.toFixed(1)}%, K=${pourcentage_K.toFixed(1)}%`);
  }
  
  // ======================= DÉBUT BLOC MODIFIÉ =======================
  // On attache les scores maximums possibles à l'objet de résultats.
  if (config.Type_Test !== 'r&K_Environnement') {
      resultats.scoresMaxPossible = _calculerScoresMaxPossibles(config.Type_Test, langOrigineNorm);
  }
  // ======================= FIN BLOC MODIFIÉ =======================
  
  if (Object.keys(resultats.scoresData).length > 0) {
    const profilEtReco = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibleNorm);
    resultats = { ...resultats, ...profilEtReco }; 
    const profilsMap = _chargerProfils(config.Type_Test, langCibleNorm);
    const infosProfilComplet = profilsMap[resultats.profilFinal];
    if (infosProfilComplet) { resultats = { ...resultats, ...infosProfilComplet }; }
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  Logger.log(`Calculs terminés. Profil Final: "${resultats.profilFinal}"`);
  return resultats;
}

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
  }
}

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
    const headers = data.shift();
    const profilsMap = {};
    const codeColIndex = headers.indexOf('Code_Profil') > -1 ? headers.indexOf('Code_Profil') : headers.indexOf('Profil');
    if (codeColIndex === -1) return {};
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
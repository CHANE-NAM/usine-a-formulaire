// =================================================================================
// == FICHIER : Logique_Universel.js
// == VERSION : 11.0 - Ajout du calcul du score maximum possible pour tous les tests.
// ==           (Précédent: 10.6 - La logique MBTI fonctionne avec les versions.)
// =================================================================================

// --- MOTEUR DE RECOMMANDATION STANDARD "r&K" ---
/**
 * Analyse une chaîne de seuil de score et vérifie si un score donné correspond.
 */
function _parseSeuilScore_rK(seuilStr, codeProfilMajoritaire, scorePourcentage) {
    if (!seuilStr || !codeProfilMajoritaire) return false;
    const seuil = String(seuilStr).trim();

    const profilSeuilMatch = seuil.toUpperCase().split(' ')[0];
    if (profilSeuilMatch !== codeProfilMajoritaire.toUpperCase()) {
        return false;
    }

    const matchSimple = seuil.match(/(>=|<=)\s*(\d+)/);
    if (matchSimple) {
        const operateur = matchSimple[1];
        const valeurSeuil = parseInt(matchSimple[2], 10);
        if (operateur === '>=') return scorePourcentage >= valeurSeuil;
        if (operateur === '<=') return scorePourcentage <= valeurSeuil;
    }

    const matchPlage = seuil.match(/(\d+)-(\d+)/);
    if (matchPlage) {
        const min = parseInt(matchPlage[1], 10);
        const max = parseInt(matchPlage[2], 10);
        return scorePourcentage >= min && scorePourcentage <= max;
    }

    return false;
}

/**
 * Moteur de détermination de profil et de recommandation SPÉCIFIQUE aux tests "r&K".
 */
function _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue) {
    Logger.log(`[ESPION][r&K] Démarrage du moteur de recommandation pour le test "${typeTest}".`);
    try {
        const totalPoints = Object.values(scoresData).reduce((sum, val) => sum + (Number(val) || 0), 0);
        if (totalPoints === 0) {
            Logger.log("[ESPION][r&K] Le total des points est à zéro. Impossible de déterminer un profil.");
            return { profilFinal: "Indéterminé", Recommandation: "" };
        }

        const profilMajoritaireCode = Object.keys(scoresData).reduce((a, b) => (scoresData[a] || 0) > (scoresData[b] || 0) ? a : b);
        const scoreMajoritaire = scoresData[profilMajoritaireCode] || 0;
        const pourcentage = (scoreMajoritaire / totalPoints) * 100;
        Logger.log(`[ESPION][r&K] Profil dominant: "${profilMajoritaireCode}" avec un score de ${scoreMajoritaire}/${totalPoints} (${pourcentage.toFixed(1)}%).`);

        const systemIds = getSystemIds();
        const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
        const nomFeuilleProfils = `Profils_${typeTest}_${langue}`;
        const sheetProfils = bdd.getSheetByName(nomFeuilleProfils);
        if (!sheetProfils) {
          Logger.log(`[ESPION][r&K] AVERTISSEMENT: Feuille de profils introuvable: "${nomFeuilleProfils}". Aucune recommandation ne sera chargée.`);
          return { profilFinal: profilMajoritaireCode, Recommandation: "" };
        }
        Logger.log(`[ESPION][r&K] Lecture de la feuille de profils: "${nomFeuilleProfils}".`);
        
        const data = sheetProfils.getDataRange().getValues();
        const headers = data.shift().map(h => String(h || '').trim());
        
        const idx = {
            profil: headers.indexOf('Code_Profil') > -1 ? headers.indexOf('Code_Profil') : headers.indexOf('Profil'),
            seuil: headers.indexOf('Seuil_Score'),
            destinataire: headers.indexOf('Destinataire'),
            axe: headers.indexOf('Axe'),
            reco: headers.indexOf('Recommandation')
        };
        
        if (idx.profil === -1 || idx.seuil === -1 || idx.destinataire === -1 || idx.axe === -1 || idx.reco === -1) {
            Logger.log(`[ESPION][r&K] ERREUR: Colonnes manquantes dans "${nomFeuilleProfils}". Requis: Code_Profil (ou Profil), Seuil_Score, Destinataire, Axe, Recommandation`);
            return { profilFinal: profilMajoritaireCode, Recommandation: "" };
        }

        for (const row of data) {
            const dest = String(row[idx.destinataire] || '').trim();
            const axe = String(row[idx.axe] || '').trim();
            const seuilStr = String(row[idx.seuil] || '').trim();
            
            if (dest === 'Répondant' && axe === 'Développer potentiel' && _parseSeuilScore_rK(seuilStr, profilMajoritaireCode, pourcentage)) {
                const profilFinalTrouve = String(row[idx.profil] || profilMajoritaireCode);
                const recommandationTrouvee = String(row[idx.reco] || '');
                Logger.log(`[ESPION][r&K] SUCCÈS: Ligne de recommandation trouvée. Profil: "${profilFinalTrouve}". Recommandation: "${recommandationTrouvee.substring(0, 50)}..."`);
                return {
                    profilFinal: profilFinalTrouve,
                    Recommandation: recommandationTrouvee
                };
            }
        }
        
        Logger.log(`[ESPION][r&K] AVERTISSEMENT: Aucune recommandation correspondante trouvée pour le profil "${profilMajoritaireCode}" avec un score de ${pourcentage.toFixed(1)}%.`);
        return { profilFinal: profilMajoritaireCode, Recommandation: "" };

    } catch (e) {
        Logger.log("ERREUR CRITIQUE dans _determinerProfilFinalParSeuils_rK: " + e.message);
        return { profilFinal: "Erreur de calcul", Recommandation: "" };
    }
}

/**
 * Moteur de calcul DÉDIÉ pour le test r&K_Environnement.
 */
function _calculerResultats_rK_Environnement_dedie(reponsesUtilisateur) {
    Logger.log('[ESPION][ENV] Moteur de calcul dédié pour r&K_Environnement activé.');
    const scores = { K: 0, r: 0 };
    let countK = 0, countR = 0;
    for (const enTete in reponsesUtilisateur) {
        const valNum = parseFloat(String(reponsesUtilisateur[enTete]).replace(',', '.'));
        if (isNaN(valNum)) continue;

        const m = String(enTete).match(/^ENV(\d{3})/);
        if (m) {
            const itemNum = parseInt(m[1], 10);
            const a = (itemNum - 1) % 4; // 0, 1, 2, 3
            if (a < 2) { // K
                scores.K += valNum;
                countK++;
            } else { // r
                scores.r += valNum;
                countR++;
            }
        }
    }
    if (countK > 0) scores.K /= countK;
    if (countR > 0) scores.r /= countR;
    
    Logger.log(`[ESPION][ENV] Scores bruts (moyennes): K=${scores.K.toFixed(2)}, r=${scores.r.toFixed(2)}`);
    return { scoresData: scores };
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
  // ... (autres langues si besoin)
  return x.toUpperCase();
}

/**
 * Calcule les résultats d'un test, y compris les scores de l'utilisateur et les scores maximums possibles.
 * @param {object} reponsesUtilisateur - Les réponses du formulaire.
 * @param {string} langueCible - Le code de la langue (ex: 'FR').
 * @param {object} config - L'objet de configuration du test.
 * @param {string} langueOrigine - La langue originale de la réponse.
 * @returns {object} Un objet contenant les scores, le profil final et les scores maximums possibles.
 */
function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  Logger.log(`[ESPION] Démarrage du calcul des résultats pour le Type_Test: "${config.Type_Test}".`);
  let resultats = { scoresData: {}, sousTotauxParMode: {} };
  
  const langCibN = _normLang(langueCible);
  
  if (config.Type_Test === 'r&K_Environnement') {
    resultats = _calculerResultats_rK_Environnement_dedie(reponsesUtilisateur);
  } else {
    const langOriN = _normLang(langueOrigine);
    const questionsMapCible = _chargerQuestions(config.Type_Test, langCibN || langueCible);
    if (!questionsMapCible) {
      Logger.log('[ESPION] AVERTISSEMENT: Impossible de charger les questions pour ' + (config && config.Type_Test) + '. Le calcul est interrompu.');
      return resultats;
    }
    
    // ==================== DÉBUT DE LA MODIFICATION V11.0 ====================
    // Calcul des scores maximums possibles
    const scoresMax = {};
    const nbQuestionsUtiliser = config.nbQuestions || Object.keys(questionsMapCible).length;
    let questionsPrisesEnCompte = 0;

    for (const id in questionsMapCible) {
      if (questionsPrisesEnCompte >= nbQuestionsUtiliser) break; // Respecte la limite de questions
      
      const q = questionsMapCible[id];
      const params = q.parametres;
      if (!params || !params.mode) continue;

      let profil = params.profil;
      let maxVal = 0;

      const mode = String(params.mode).toUpperCase();
      if (mode === 'ECHELLE_NOTE' || mode === 'LIKERT_5') {
        maxVal = params.echelle_max || params.max || 0;
        // Pour LIKERT_5, le profil est parfois dans la première option
        if (!profil && params.options && params.options[0] && params.options[0].profil) {
          profil = params.options[0].profil;
        }
      } else if (mode === 'QCU_CAT' || mode === 'QRM_CAT') {
        if (params.options && params.options.length > 0) {
          // Le max est la plus haute valeur possible parmi les options
          maxVal = Math.max(...params.options.map(opt => (typeof opt.valeur === 'number' ? opt.valeur : 0)));
          // Le profil est celui de la première option (supposition)
          profil = params.options[0].profil;
        }
      }
      
      if (profil && maxVal > 0) {
        scoresMax[profil] = (scoresMax[profil] || 0) + maxVal;
      }
      questionsPrisesEnCompte++;
    }
    
    resultats.scoresMaxPossible = scoresMax;
    Logger.log('[ESPION] Scores maximums possibles calculés: ' + JSON.stringify(scoresMax));
    // ===================== FIN DE LA MODIFICATION V11.0 =====================

    _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats, config.nbQuestions);
  }
  
  Logger.log('[ESPION] Scores bruts calculés: ' + JSON.stringify(resultats.scoresData));
  if (Object.keys(resultats.scoresData).length > 0) {
    const profilEtReco = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibN || langueCible);
    resultats = { ...resultats, ...profilEtReco }; 
    
    const profilsMap = _chargerProfils(config.Type_Test, langCibN || langueCible);
    const infosProfilComplet = profilsMap[resultats.profilFinal];
    if (infosProfilComplet) {
      resultats = { ...resultats, ...infosProfilComplet };
    }
    
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  Logger.log("[ESPION] Calculs terminés. Objet de résultats final (partiel): " + 
             `profilFinal="${resultats.profilFinal}", ` +
             `Recommandation="${(resultats.Recommandation || '').substring(0,50)}..."`);
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
    if (questionConfig) {
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
      if (__DBG) Logger.log('Mode de traitement inconnu: "%s" → réponse ignorée', mode);
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
  // === DÉBUT DE LA MODIFICATION V10.5 ===
  // On cherche le profil à impacter.
  // D'abord à la racine des paramètres (pour ECHELLE_NOTE standard)...
  let profil = parametres.profil; 
  
  // ...sinon, on le cherche dans la première option (pour la compatibilité avec LIKERT_5)
  if (!profil && parametres.options && parametres.options[0] && parametres.options[0].profil) {
    profil = parametres.options[0].profil;
  }
  
  // Si on ne trouve toujours pas de profil, on arrête.
  if (!profil) return;
  // === FIN DE LA MODIFICATION V10.5 ===

  const valeurNumerique = parseFloat(String(reponseUtilisateur).replace(',', '.'));
  if (!isNaN(valeurNumerique)) {
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeurNumerique;
  }
}

function _determinerProfilFinal(scoresData, typeTest, langue) {
  if (!scoresData || Object.keys(scoresData).length === 0) return { profilFinal: "" };
  
  if (String(typeTest || '').toLowerCase().startsWith('r&k_')) {
    return _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue);
  }
  
  // ==================== DÉBUT MODIFICATION V10.6 ====================
  if (String(typeTest || '').toUpperCase().startsWith('MBTI')) {
  // ===================== FIN MODIFICATION V10.6 =====================
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
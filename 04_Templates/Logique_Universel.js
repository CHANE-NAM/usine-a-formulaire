// =================================================================================
// == FICHIER : Logique_Universel.js
// == VERSION : 10.2 - Moteur de calcul dédié pour r&K_Environnement pour robustesse.
// ==           Amélioration des espions pour tracer la traduction et la recherche.
// =================================================================================

// --- MOTEUR DE RECOMMANDATION STANDARD "r&K" ---
/**
 * Analyse une chaîne de seuil de score (ex: "R >= 80%", "K 60-79%") et vérifie si un score donné correspond.
 * @param {string} seuilStr - La chaîne de caractères du seuil.
 * @param {string} codeProfilMajoritaire - Le code du profil dominant (ex: 'R', 'K').
 * @param {number} scorePourcentage - Le score en pourcentage de l'utilisateur.
 * @returns {boolean} True si le score correspond au seuil, sinon false.
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
 * Lit l'onglet de profil correspondant et trouve la recommandation multi-critères.
 * @param {Object} scoresData - Les scores bruts (ex: {R: 15, K: 5}).
 * @param {string} typeTest - Le nom du test (ex: 'r&K_Adaptabilite').
 * @param {string} langue - Le code de la langue (ex: 'FR').
 * @returns {Object} Un objet contenant `profilFinal` et `Recommandation`.
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
            profil: headers.indexOf('Profil'),
            seuil: headers.indexOf('Seuil_Score'),
            destinataire: headers.indexOf('Destinataire'),
            axe: headers.indexOf('Axe'),
            reco: headers.indexOf('Recommandation')
        };
        if (Object.values(idx).some(i => i === -1)) {
            Logger.log(`[ESPION][r&K] ERREUR: Colonnes manquantes dans "${nomFeuilleProfils}". Requis: Profil, Seuil_Score, Destinataire, Axe, Recommandation`);
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
 * Calcule les scores K et r sans dépendre de la feuille Questions_...
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


// Compteurs de debug
var __DBG_QCU_CAT_MISS = 0;
var __DBG_QRM_CAT_MISS = 0;
var __DBG_LIKERT_MISS  = 0;

// Normalisation robuste de chaînes
function _normStr(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[\u2019\u2018]/g, "'").replace(/[\u201C\u201D]/g, '"')
    .replace(/[«»]/g, '').replace(/[\u2013\u2014]/g, '-')
    .replace(/\u00A0/g, ' ').replace(/\s+/g, ' ')
    .trim().toLowerCase();
}


// Normalisation de codes langues
function _normLang(s) {
  const x = _normStr(s);
  if (!x) return '';
  if (/^fr|fran|french/.test(x)) return 'FR';
  if (/^en|angl|english|uk|us/.test(x)) return 'EN';
  if (/^es|espag|span/.test(x)) return 'ES';
  if (/^de|allem|german/.test(x)) return 'DE';
  if (/^it|ital/.test(x)) return 'IT';
  if (/^pt|portug/.test(x)) return 'PT';
  const m = x.match(/^[a-z]{2}$/);
  return m ? x.toUpperCase() : x.toUpperCase();
}

function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  Logger.log(`[ESPION] Démarrage du calcul des résultats pour le Type_Test: "${config.Type_Test}".`);
  let resultats = { scoresData: {}, sousTotauxParMode: {} };
  
  const langCibN = _normLang(langueCible);
  
  // --- MODIFICATION V10.2 START: Aiguillage vers moteur dédié pour r&K_Environnement ---
  if (config.Type_Test === 'r&K_Environnement') {
    resultats = _calculerResultats_rK_Environnement_dedie(reponsesUtilisateur);
  } else {
    // Logique standard pour tous les autres tests
    const langOriN = _normLang(langueOrigine);
    const questionsMapCible = _chargerQuestions(config.Type_Test, langCibN || langueCible);
    if (!questionsMapCible) {
      Logger.log('[ESPION] AVERTISSEMENT: Impossible de charger les questions pour ' + (config && config.Type_Test) + '. Le calcul est interrompu.');
      return resultats;
    }
    if (!langOriN || langOriN === langCibN) {
      _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats);
    } else {
      const questionsMapOrigine = _chargerQuestions(config.Type_Test, langOriN);
      if (!questionsMapOrigine) {
        Logger.log('[ESPION] AVERTISSEMENT: Feuille de questions de la langue d\'origine introuvable. Tentative de calcul direct.');
        _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats);
      } else {
        _traduireEtExecuterCalcul(reponsesUtilisateur, questionsMapOrigine, questionsMapCible, resultats);
      }
    }
  }
  // --- MODIFICATION V10.2 END ---
  
  Logger.log('[ESPION] Scores bruts calculés: ' + JSON.stringify(resultats.scoresData));
  
  if (config.Type_Test === 'r&K_Environnement') {
      const scoresTraduits = { K: resultats.scoresData.K || 0, r: resultats.scoresData.r || 0 };
      resultats.scoresData = scoresTraduits;
      Logger.log('[ESPION][Traducteur] Scores normalisés pour le moteur r&K: ' + JSON.stringify(resultats.scoresData));
  }

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

function _traduireEtExecuterCalcul(reponsesUtilisateur, questionsMapOrigine, questionsMapCible, resultats) {
  for (const enTeteComplet in reponsesUtilisateur) {
    if (!enTeteComplet.includes(':')) continue;
    const idQuestion = enTeteComplet.split(':')[0].trim();
    const qc = questionsMapCible[idQuestion];
    if (!qc) continue;

    if (qc.parametres.mode === 'ECHELLE_NOTE') {
      _aiguillerCalcul(qc.parametres.mode, reponsesUtilisateur[enTeteComplet], qc.parametres, resultats);
    } else {
      const qo = questionsMapOrigine[idQuestion];
      if (qo && qo.parametres && qo.parametres.options) {
        const reponsesArray = String(reponsesUtilisateur[enTeteComplet]).split(',').map(r => r.trim());
        reponsesArray.forEach(reponseSimple => {
          const idx = qo.parametres.options.findIndex(opt => _normStr(opt.libelle) === _normStr(reponseSimple));
          if (idx !== -1 && qc.parametres.options && qc.parametres.options[idx]) {
            _aiguillerCalcul(qc.parametres.mode, qc.parametres.options[idx].libelle, qc.parametres, resultats);
          } else {
            const optDirect = (qc.parametres.options || []).find(opt => _normStr(opt.libelle) === _normStr(reponseSimple));
            if (optDirect) _aiguillerCalcul(qc.parametres.mode, optDirect.libelle, qc.parametres, resultats);
          }
        });
      } else {
        _aiguillerCalcul(qc.parametres.mode, reponsesUtilisateur[enTeteComplet], qc.parametres, resultats);
      }
    }
  }
}

function _executerCalcul(reponses, questionsMap, resultats) {
  for (const enTeteComplet in reponses) {
    if (!enTeteComplet.includes(':')) continue;
    const idQuestion = enTeteComplet.split(':')[0].trim();
    const questionConfig = questionsMap[idQuestion];
    if (questionConfig) {
      _aiguillerCalcul(questionConfig.parametres.mode, reponses[enTeteComplet], questionConfig.parametres, resultats);
    }
  }
}

function _aiguillerCalcul(mode, reponse, parametres, resultats) {
  var m = String(mode || '').replace(/\s+/g, ' ').trim().toUpperCase();
  switch (m) {
    case 'QCU_CAT':      _traiterQCU_CAT(reponse, parametres, resultats);    break;
    case 'ECHELLE_NOTE': _traiterECHELLE_NOTE(reponse, parametres, resultats); break;
    default:
      if (__DBG) Logger.log('Mode de traitement inconnu: "%s" → réponse ignorée', mode);
      break;
  }
}

function _traiterQCU_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;
  const repNorm = _normStr(reponseUtilisateur);
  let optionTrouvee = parametres.options.find(opt => _normStr(opt.libelle) === repNorm);
  if (!optionTrouvee) {
    const n = parseInt(String(reponseUtilisateur).trim(), 10);
    if (!isNaN(n) && n >= 1 && n <= parametres.options.length) optionTrouvee = parametres.options[n - 1];
  }
  if (optionTrouvee && optionTrouvee.profil) {
    const profil = optionTrouvee.profil;
    const valeur = (typeof optionTrouvee.valeur === 'number') ? optionTrouvee.valeur : 1;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
  }
}

function _traiterECHELLE_NOTE(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.profil) return;
  const valeurNumerique = parseFloat(String(reponseUtilisateur).replace(',', '.'));
  if (!isNaN(valeurNumerique)) {
    resultats.scoresData[parametres.profil] = (resultats.scoresData[parametres.profil] || 0) + valeurNumerique;
  }
}

function _determinerProfilFinal(scoresData, typeTest, langue) {
  if (!scoresData || Object.keys(scoresData).length === 0) return { profilFinal: "" };
  if (String(typeTest || '').toLowerCase().startsWith('r&k_')) {
    return _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue);
  }
  if (String(typeTest || '').toUpperCase() === 'MBTI') {
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


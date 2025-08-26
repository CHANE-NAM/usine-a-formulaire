// =================================================================================
// == FICHIER : Logique_Universel.gs
// == VERSION : 9.8 (normLang + fallback si feuille origine manquante + match robuste
// ==             + LIKERT_5 + moteur Environnement + logs)
// == RÔLE  : Moteur de calcul universel capable de traiter n'importe quel test.
// == STATUT : PÉRENNE (TEMPLATE + Sheets réponses)
// =================================================================================

// Compteurs de debug (limiter le bruit de logs par exécution)
var __DBG_QCU_CAT_MISS = 0;
var __DBG_QRM_CAT_MISS = 0;
var __DBG_LIKERT_MISS  = 0;

// Normalisation robuste de chaînes (accents, apostrophes typographiques, tirets, NBSP, casse)
function _normStr(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // accents → sans accents
    .replace(/[\u2019\u2018]/g, "'")                  // ’ ‘ → '
    .replace(/[\u201C\u201D]/g, '"')                  // “ ” → "
    .replace(/[«»]/g, '')                             // guillemets français → rien
    .replace(/[\u2013\u2014]/g, '-')                  // – — → -
    .replace(/\u00A0/g, ' ')                           // NBSP → espace
    .replace(/\s+/g, ' ')                              // espaces multiples → simple
    .trim()
    .toLowerCase();
}


// Normalisation de codes langues (FR, EN, ES, DE, IT, PT, …)
function _normLang(s) {
  const x = _normStr(s);
  if (!x) return '';
  // français
  if (/^fr|fran|french/.test(x)) return 'FR';
  // anglais
  if (/^en|angl|english|uk|us/.test(x)) return 'EN';
  // espagnol
  if (/^es|espag|span/.test(x)) return 'ES';
  // allemand
  if (/^de|allem|german/.test(x)) return 'DE';
  // italien
  if (/^it|ital/.test(x)) return 'IT';
  // portugais
  if (/^pt|portug/.test(x)) return 'PT';
  // par défaut, si ressemble à 2 lettres
  const m = x.match(/^[a-z]{2}$/);
  return m ? x.toUpperCase() : x.toUpperCase(); // sinon renvoyer tel quel en MAJ
}

function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  // --- Logs de debug (aperçu des clés et valeurs) ---
  try {
    const keys = Object.keys(reponsesUtilisateur || {});
    Logger.log('DEBUG(keys[0..4]) → ' + JSON.stringify(keys.slice(0, 5)));
    if (keys.length) {
      const peek = {};
      keys.slice(0, 5).forEach(k => peek[k] = reponsesUtilisateur[k]);
      Logger.log('DEBUG(values[0..4]) → ' + JSON.stringify(peek));
    }
  } catch(e){}

  // Auto-détection/normalisation du type
  let typeLC = String((config && config.Type_Test) || '').toLowerCase().trim();
  const autoType   = _detectTestTypeAuto(reponsesUtilisateur);
  const hasEnvKeys = _hasEnvKeys(reponsesUtilisateur);

  if (!typeLC && autoType) {
    config.Type_Test = autoType;
    typeLC = autoType.toLowerCase();
    Logger.log('Type_Test auto-détecté: ' + config.Type_Test);
  }

  // Normalisation langues (ex: 'Français' → 'FR')
  const langCibN = _normLang(langueCible);
  const langOriN = _normLang(langueOrigine);

  // ——— CAS PRIORITAIRE : r&K_Environnement ———
  try {
    if (typeLC.indexOf('environnement') !== -1 || hasEnvKeys) {
      if (!config.Type_Test) config.Type_Test = 'r&K_Environnement';
      Logger.log('Moteur ENV activé (raison: ' +
                 (typeLC.indexOf('environnement') !== -1 ? 'Type_Test' : 'Clés ENV détectées') + ').');

      let res = calculerResultats_rK_Environnement(reponsesUtilisateur, langCibN || langueCible, config);

      // Enrichissement optionnel via Profils_*
      const profilsMap = (typeof _chargerProfils === 'function')
        ? _chargerProfils(config.Type_Test, langCibN || langueCible)
        : {};
      if (res.profilFinal && profilsMap[res.profilFinal]) {
        res = { ...res, ...profilsMap[res.profilFinal] };
      }

      Logger.log("Calculs terminés. Résultats : " + JSON.stringify(res));
      return res;
    }
  } catch (e) {
    Logger.log("Bypass moteur ENV (fallback universel) : " + e.message);
  }

  // ——— FLOWS GENERIQUES (autres tests) ———
  let resultats = { scoresData: {}, sousTotauxParMode: {} };

  const profilsMap = (typeof _chargerProfils === 'function')
    ? _chargerProfils(config.Type_Test, langCibN || langueCible)
    : {};

  const questionsMapCible = _chargerQuestions(config.Type_Test, langCibN || langueCible);
  if (!questionsMapCible) {
    Logger.log('Avertissement: _chargerQuestions a retourné null pour ' + (config && config.Type_Test));
    Logger.log("Calculs terminés. Résultats : " + JSON.stringify(resultats));
    return resultats;
  }

  // 1) Même langue (après normalisation) → exécuter directement
  if (!langOriN || langOriN === langCibN) {
    _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats);
  } else {
    // 2) Langues différentes → tenter la "traduction" classique
    const questionsMapOrigine = _chargerQuestions(config.Type_Test, langOriN);
    if (!questionsMapOrigine) {
      // 2a) Fallback : si la feuille origine n’existe pas, traiter quand même avec la langue cible
      Logger.log('Avertissement: _chargerQuestions(orig) introuvable pour langue=' + langOriN + '. Fallback direct → langue cible.');
      _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats);
    } else {
      // 2b) Traduction classique par position d’options (si les options diffèrent réellement)
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
            const reponseTexte = reponsesUtilisateur[enTeteComplet];
            const reponsesArray = String(reponseTexte).split(',').map(r => r.trim());
            reponsesArray.forEach(reponseSimple => {
              const idx = qo.parametres.options.findIndex(opt => _normStr(opt.libelle) === _normStr(reponseSimple));
              if (idx !== -1 && qc.parametres.options && qc.parametres.options[idx]) {
                const optCible = qc.parametres.options[idx];
                _aiguillerCalcul(qc.parametres.mode, optCible.libelle, qc.parametres, resultats);
              } else {
                // si pas trouvé côté origine, tenter le match direct sur la cible (au cas où les libellés sont identiques)
                const optDirect = (qc.parametres.options || []).find(opt => _normStr(opt.libelle) === _normStr(reponseSimple));
                if (optDirect) {
                  _aiguillerCalcul(qc.parametres.mode, optDirect.libelle, qc.parametres, resultats);
                }
              }
            });
          } else {
            // si pas d’options côté origine, essayer direct sur la cible
            const reponseTexte = reponsesUtilisateur[enTeteComplet];
            _aiguillerCalcul(qc.parametres.mode, reponseTexte, qc.parametres, resultats);
          }
        }
      }
    }
  }

  if (Object.keys(resultats.scoresData).length > 0) {
    resultats.profilFinal = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langCibN || langueCible);
    const toutes = profilsMap[resultats.profilFinal];
    if (toutes) resultats = { ...resultats, ...toutes };
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }

  Logger.log("Calculs terminés. Résultats : " + JSON.stringify(resultats));
  return resultats;
}

/**
 * Auto-détecte le type de test à partir des en-têtes de la feuille de réponses.
 * Couvre : ENV###, ADA###, RES###, CRE### et motif "RKxx:" (→ Adaptabilité).
 */
function _detectTestTypeAuto(reponses) {
  let seenENV=false, seenRK=false, seenADA=false, seenRES=false, seenCRE=false;
  for (const k in reponses) {
    const kk = String(k || '');
    if (/^ENV\s*\d{3}/i.test(kk)) seenENV = true;
    if (/^RK\d{2}\s*:/.test(kk))   seenRK  = true; // "RK01: ..."
    if (/^ADA\s*\d{3}/i.test(kk))  seenADA = true;
    if (/^RES\s*\d{3}/i.test(kk))  seenRES = true;
    if (/^CRE\s*\d{3}/i.test(kk))  seenCRE = true;
  }
  if (seenENV) return 'r&K_Environnement';
  if (seenADA) return 'r&K_Adaptabilite';
  if (seenRES) return 'r&K_Resilience';
  if (seenCRE) return 'r&K_Creativite';
  if (seenRK)  return 'r&K_Adaptabilite';
  return '';
}

/** Présence d’en-têtes ENV### (même “nettoyés”). */
function _hasEnvKeys(reponses) {
  for (const k in reponses) {
    if (/^ENV\s*\d{3}/i.test(k)) return true;
    if (/^ENV\d{3}_/i.test(k))   return true;
  }
  return false;
}

/* ============================================================================
 * Moteur — r&K_Environnement
 * ============================================================================ */
function calculerResultats_rK_Environnement(reponse, langueCible, config) {
  const envVals = {};
  for (const k in reponse) {
    const m = String(k).match(/^ENV\s*(\d{3})/i);
    if (m) {
      const n = parseInt(m[1], 10);
      const raw = reponse[k];
      const v = (typeof raw === 'number') ? raw : Number(String(raw).replace(',', '.'));
      if (!isNaN(v)) envVals[n] = v;
    }
  }

  const THEMES = [
    "Concurrence & Pression du marché","Clients & Demande","Technologies & Innovation",
    "Réglementation & Cadre juridique","Ressources humaines & Compétences","Financement & Accès aux capitaux",
    "Fournisseurs & Logistique","Ressources & Infrastructures matérielles","Image & Réputation sectorielle",
    "Partenariats & Réseaux","Territoire & Environnement géographique","Tendances sociétales & culturelles",
    "Contexte économique global","Risques & Sécurité","Opportunités de croissance & Marchés",
  ];

  const avg = (a,b) => (a+b)/2;
  const round2 = n => +n.toFixed(2);
  const interpK = (x) => x>=7 ? "Environnement plutôt stable et prévisible"
                    : x<=3 ? "Environnement plutôt instable / changeant"
                           : "Stabilité modérée avec quelques variations";
  const interpr = (x) => x>=7 ? "Changements rapides / forte dynamique"
                    : x<=3 ? "Changements lents / faible dynamique"
                           : "Vitesse de changement modérée";

  const themes = [];
  let sumK=0, sumR=0, cntK=0, cntR=0;

  for (let t=0; t<15; t++) {
    const base = t*4;
    const K1 = envVals[base+1], K2 = envVals[base+2];
    const R1 = envVals[base+3], R2 = envVals[base+4];

    const hasK = (K1!=null && K2!=null);
    const hasR = (R1!=null && R2!=null);

    const k = hasK ? avg(K1, K2) : null;
    const r = hasR ? avg(R1, R2) : null;

    if (k!=null) { sumK += k; cntK++; }
    if (r!=null) { sumR += r; cntR++; }

    themes.push({
      name: THEMES[t],
      stabilite: k!=null ? round2(k) : "",
      vitesse:   r!=null ? round2(r) : "",
      interpretStab: k!=null ? interpK(k) : "",
      interpretVit:  r!=null ? interpr(r) : "",
      reco: ""
    });
  }

  const scoreK = cntK ? round2(sumK/cntK) : 0;
  const scoreR = cntR ? round2(sumR/cntR) : 0;

  const hi = 6.5, lo = 3.5;
  let titreProfil = "";
  if (scoreK >= hi && scoreR <= lo) titreProfil = "Stable & Lent";
  else if (scoreK >= hi && scoreR >= hi) titreProfil = "Stable & Rapide";
  else if (scoreK <= lo && scoreR >= hi) titreProfil = "Instable & Rapide";
  else if (scoreK <= lo && scoreR <= lo) titreProfil = "Instable & Lent";
  else if (scoreK >= scoreR)            titreProfil = "Plutôt Stable";
  else                                  titreProfil = "Plutôt Rapide";

  const flat = {
    Score_Stabilite: scoreK,
    Interpretation_Stabilite: interpK(scoreK),
    Score_Vitesse: scoreR,
    Interpretation_Vitesse: interpr(scoreR),
    Titre_Profil: titreProfil,
    profilFinal: titreProfil
  };
  themes.forEach((th, i) => {
    const n = i+1;
    flat[`Nom_Theme_${n}`] = th.name;
    flat[`Score_Stabilite_Theme_${n}`] = th.stabilite;
    flat[`Interpretation_Stabilite_Theme_${n}`] = th.interpretStab;
    flat[`Score_Vitesse_Theme_${n}`] = th.vitesse;
    flat[`Interpretation_Vitesse_Theme_${n}`] = th.interpretVit;
    flat[`Recommandations_Theme_${n}`] = th.reco;
  });

  return {
    scoresData: { K: scoreK, r: scoreR },
    sousTotauxParMode: { K: scoreK, r: scoreR },
    mapCodeToName: { K: "Stabilité (K)", r: "Vitesse (r)" },
    themes,
    ...flat
  };
}

// =================================================================================
// == FLOWS & OUTILS GENERIQUES
// =================================================================================

function _executerCalcul(reponses, questionsMap, resultats) {
  for (const enTeteComplet in reponses) {
    if (!enTeteComplet.includes(':')) continue;
    const idQuestion = enTeteComplet.split(':')[0].trim();
    const questionConfig = questionsMap[idQuestion];
    if (questionConfig) {
      const reponse = reponses[enTeteComplet];
      const mode = questionConfig.parametres.mode;
      const parametres = questionConfig.parametres;
      _aiguillerCalcul(mode, reponse, parametres, resultats);
    }
  }
}

function _aiguillerCalcul(mode, reponse, parametres, resultats) {
  // normalise le mode : trim, casse, espaces multiples
  var m = String(mode || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

  switch (m) {
    case 'QCU_DIRECT':   _traiterQCU_DIRECT(reponse, parametres, resultats); break;
    case 'QCU_CAT':      _traiterQCU_CAT(reponse, parametres, resultats);    break;
    case 'QRM_CAT':      _traiterQRM_CAT(reponse, parametres, resultats);    break;
    case 'ECHELLE_NOTE': _traiterECHELLE_NOTE(reponse, parametres, resultats); break;
    case 'LIKERT_5':     _traiterLIKERT_5(reponse, parametres, resultats);   break;

    // petites variantes fréquentes qu’on tolère aussi
    case 'QCU':
      _traiterQCU_CAT(reponse, parametres, resultats);                        break;
    case 'LIKERT':
    case 'LIKERT5':
      _traiterLIKERT_5(reponse, parametres, resultats);                       break;

    default:
      Logger.log('Mode de traitement inconnu: "%s" → réponse ignorée', mode);
      break;
  }
}


// Likert 5 points (options libellés + valeurs 1..5), match robuste
function _traiterLIKERT_5(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.options) return;
  const repN = _normStr(reponseUtilisateur);
  const opt = parametres.options.find(o => _normStr(o.libelle) === repN);
  if (opt && opt.profil && typeof opt.valeur === 'number') {
    resultats.scoresData[opt.profil] = (resultats.scoresData[opt.profil] || 0) + opt.valeur;
  } else {
    if (__DBG_LIKERT_MISS < 3) {
      Logger.log('LIKERT_5: aucune option ne correspond à "%s". Exemples: %s',
        reponseUtilisateur, JSON.stringify((parametres.options||[]).slice(0,3).map(o=>o.libelle)));
      __DBG_LIKERT_MISS++;
    }
  }
}

function _traiterQCU_DIRECT(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.profil) return;
  resultats.scoresData[parametres.profil] = reponseUtilisateur;
}

// QCU: match texte (normalisé) OU index numérique (1..n)
function _traiterQCU_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;

  const repStr  = String(reponseUtilisateur).trim();
  const repNorm = _normStr(repStr);

  // 1) correspondance par libellé normalisé
  let optionTrouvee = parametres.options.find(opt => _normStr(opt.libelle) === repNorm);

  // 2) fallback: si "3" → 3e option
  if (!optionTrouvee) {
    const n = parseInt(repStr, 10);
    if (!isNaN(n) && n >= 1 && n <= parametres.options.length) {
      optionTrouvee = parametres.options[n - 1];
    }
  }

  if (optionTrouvee && optionTrouvee.profil) {
    const profil = optionTrouvee.profil;
    const valeur = (typeof optionTrouvee.valeur === 'number') ? optionTrouvee.valeur : 1;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
  } else {
    if (typeof __DBG_QCU_CAT_MISS === 'number' && __DBG_QCU_CAT_MISS < 5) {
      Logger.log('QCU_CAT: aucune option ne correspond à "%s". Exemples: %s',
        reponseUtilisateur, JSON.stringify((parametres.options||[]).slice(0,3).map(o=>o.libelle)));
      __DBG_QCU_CAT_MISS++;
    }
  }
}

// QRM: idem, et accepte "1,3" comme multi-index
function _traiterQRM_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;

  const reponsesArray = String(reponseUtilisateur)
    .split(',')
    .map(r => r.trim())
    .filter(Boolean);

  reponsesArray.forEach(reponse => {
    const repNorm = _normStr(reponse);

    let optionTrouvee = parametres.options.find(opt => _normStr(opt.libelle) === repNorm);

    if (!optionTrouvee) {
      const n = parseInt(reponse, 10);
      if (!isNaN(n) && n >= 1 && n <= parametres.options.length) {
        optionTrouvee = parametres.options[n - 1];
      }
    }

    if (optionTrouvee && optionTrouvee.profil && typeof optionTrouvee.valeur === 'number') {
      const profil = optionTrouvee.profil;
      const valeur = optionTrouvee.valeur;
      resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
    } else {
      if (typeof __DBG_QRM_CAT_MISS === 'number' && __DBG_QRM_CAT_MISS < 5) {
        Logger.log('QRM_CAT: aucune option ne correspond à "%s". Exemples: %s',
          reponse, JSON.stringify((parametres.options||[]).slice(0,3).map(o=>o.libelle)));
        __DBG_QRM_CAT_MISS++;
      }
    }
  });
}


function _traiterECHELLE_NOTE(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.profil) return;
  const valeurNumerique = parseFloat(String(reponseUtilisateur).replace(',', '.'));
  if (!isNaN(valeurNumerique)) {
    const profil = parametres.profil;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeurNumerique;
  }
}

function _determinerProfilFinalParSeuils(scoresData, typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuilleProfils = `Profils_${typeTest}_${langue}`;
    const sheetProfils = bdd.getSheetByName(nomFeuilleProfils);
    if (!sheetProfils) return "";
    const dataProfils = sheetProfils.getRange("A2:C" + sheetProfils.getLastRow()).getValues();
    const totalPoints = Object.values(scoresData).reduce((sum, val) => sum + val, 0);
    if (totalPoints === 0) return "";
    const profilMajoritaire = Object.keys(scoresData).reduce((a, b) => scoresData[a] > scoresData[b] ? a : b);
    const scoreMajoritaire = scoresData[profilMajoritaire];
    const pourcentage = (scoreMajoritaire / totalPoints) * 100;
    for (const row of dataProfils) {
      const nomProfil = row[0];
      const conditionSeuil = row[1];
      if (!nomProfil || !conditionSeuil) continue;
      const codeProfilSeuil = conditionSeuil.split(' ')[0];
      if (codeProfilSeuil.toUpperCase() !== String(profilMajoritaire).toUpperCase()) continue;
      if (conditionSeuil.includes('>=')) {
        const seuil = parseFloat(conditionSeuil.replace(/[^0-9.]/g, ''));
        if (pourcentage >= seuil) return nomProfil;
      } else if (conditionSeuil.includes('<=')) {
        const seuil = parseFloat(conditionSeuil.replace(/[^0-9.]/g, ''));
        if (pourcentage <= seuil) return nomProfil;
      } else if (conditionSeuil.includes('-')) {
        const parts = conditionSeuil.match(/(\d+)-(\d+)/);
        if (parts) {
          const min = parseInt(parts[1], 10);
          const max = parseInt(parts[2], 10);
          if (pourcentage >= min && pourcentage <= max) return nomProfil;
        }
      }
    }
    return profilMajoritaire;
  } catch (e) {
    Logger.log("Erreur dans _determinerProfilFinalParSeuils: " + e.message);
    return "";
  }
}

function _determinerProfilFinal(scoresData, typeTest, langue) {
  if (!scoresData || Object.keys(scoresData).length === 0) return "";
  const testsAvecSeuils = ["r&K_Adaptabilite", "r&K_Resilience", "r&K_Creativite"];
  if (testsAvecSeuils.some(t => String(typeTest || '').toUpperCase() === t.toUpperCase())) {
    return _determinerProfilFinalParSeuils(scoresData, typeTest, langue);
  }
  if (String(typeTest || '').toUpperCase() === 'MBTI') {
    let profil = "";
    profil += (scoresData.E || 0) > (scoresData.I || 0) ? 'E' : 'I';
    profil += (scoresData.S || 0) > (scoresData.N || 0) ? 'S' : 'N';
    profil += (scoresData.T || 0) > (scoresData.F || 0) ? 'T' : 'F';
    profil += (scoresData.J || 0) > (scoresData.P || 0) ? 'J' : 'P';
    return profil;
  } else {
    return Object.keys(scoresData).reduce((a, b) => scoresData[a] > scoresData[b] ? a : b);
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
    // Normalisation en-têtes (BOM/quotes/espaces)
    const headersRaw = data.shift();
    const headers = (headersRaw || []).map(h =>
      String(h || '')
        .replace(/^\uFEFF/, '')    // BOM
        .replace(/^"|"$/g, '')     // guillemets englobants
        .trim()
    );

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

// =================================================================================
function enrichirVariablesEnv(resultats, reponses, config, langueCible) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const sheetName = `Questions_${config.Type_Test}_${langueCible}`;
    const sheet = bdd.getSheetByName(sheetName);
    if (!sheet) { Logger.log(`enrichirVariablesEnv: feuille introuvable: ${sheetName}`); return resultats; }

    const data = sheet.getDataRange().getValues();
    // Normalisation en-têtes
    const headersRaw = data.shift();
    const headers = (headersRaw || []).map(h =>
      String(h || '').replace(/^\uFEFF/, '').replace(/^"|"$/g, '').trim()
    );

    const colParams = headers.indexOf('Paramètres (JSON)');
    const colTypeQ  = headers.indexOf('TypeQuestion');
    const colID     = headers.indexOf('ID');

    const THEME_ORDER = [
      'Concurrence & Pression du marché','Clients & Demande','Technologies & Innovation',
      'Réglementation & Cadre juridique','Ressources humaines & Compétences','Financement & Accès aux capitaux',
      'Fournisseurs & Logistique','Ressources & Infrastructures matérielles','Image & Réputation sectorielle',
      'Partenariats & Réseaux','Territoire & Environnement géographique','Tendances sociétales & culturelles',
      'Contexte économique global','Risques & Sécurité','Opportunités de croissance & Marchés'
    ];
    function avg(list){ return (list&&list.length)?(list.reduce((a,b)=>a+Number(b||0),0)/list.length):null;}
    function round1(n){ return (n==null)?'':Number(n.toFixed(1));}
    function interpStab(v){ if(v==null) return ''; return v<=3?'Très r': v<=5?'r': v<=7?'K':'Très K'; }
    function interpVit(v){ if(v==null) return ''; return v<=3?'Très lent': v<=5?'Lent': v<=7?'Rapide':'Très rapide'; }

    const bucket = {}; THEME_ORDER.forEach(n=>bucket[n]={stab:[],vit:[]});

    const envMap = {};
    for (const h in reponses) {
      const m = String(h).match(/^ENV\s*(\d{3})/i);
      if (!m) continue;
      const v = Number(String(reponses[h]).replace(',','.'));
      if (!isNaN(v)) envMap[m[1]] = v;
    }

    data.forEach(row=>{
      const typeQ = (row[colTypeQ]||'').trim();
      if (typeQ !== 'ECHELLE_NOTE') return;
      let p={}; try{ p=JSON.parse(row[colParams]||'{}'); }catch(_){}
      const theme = p.theme||''; const dim = p.dimension||'';
      const id = (row[colID]||'').toString().replace(/[^\d]/g,''); // 'ENV001' -> '001'
      if (!theme || !envMap[id] || !bucket.hasOwnProperty(theme)) return;
      if (dim==='Stabilité') bucket[theme].stab.push(envMap[id]);
      else if (dim==='Vitesse') bucket[theme].vit.push(envMap[id]);
    });

    resultats.themes = [];
    let gS=[], gV=[];
    THEME_ORDER.forEach((tName,i)=>{
      const s = avg(bucket[tName].stab), v = avg(bucket[tName].vit);
      if (s!=null) gS = gS.concat(bucket[tName].stab);
      if (v!=null) gV = gV.concat(bucket[tName].vit);
      const s1 = round1(s), v1 = round1(v), sI = interpStab(s), vI = interpVit(v);
      resultats.themes.push({name:tName, stabilite:s1, vitesse:v1, interpretStab:sI, interpretVit:vI, reco:''});
      const j=i+1;
      resultats['Nom_Theme_'+j]=tName;
      resultats['Score_Stabilite_Theme_'+j]=s1;
      resultats['Interpretation_Stabilite_Theme_'+j]=sI;
      resultats['Score_Vitesse_Theme_'+j]=v1;
      resultats['Interpretation_Vitesse_Theme_'+j]=vI;
      resultats['Recommandations_Theme_'+j]='';
    });

    function maybeSet(o,k,v){ if(o[k]==null || o[k]==='') o[k]=v; }
    const gStab = round1(avg(gS)), gVit = round1(avg(gV));
    maybeSet(resultats,'Score_Stabilite', gStab);
    maybeSet(resultats,'Interpretation_Stabilite', interpStab(gStab||null));
    maybeSet(resultats,'Score_Vitesse', gVit);
    maybeSet(resultats,'Interpretation_Vitesse', interpVit(gVit||null));

  } catch(e) {
    Logger.log('enrichirVariablesEnv() ERREUR: ' + e.message + '\n' + e.stack);
  }
  return resultats;
}

/* ============================================================================
 * CHARGEUR DE PROFILS
 * ============================================================================ */
function _chargerProfils(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuille = `Profils_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuille);
    if (!sheet) throw new Error(`Feuille introuvable: ${nomFeuille}`);

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const profilsMap = {};
    const codeColIndex = headers.indexOf('Code_Profil') > -1 ? headers.indexOf('Code_Profil') : headers.indexOf('Profil');
    if (codeColIndex === -1) throw new Error("Colonne 'Code_Profil' ou 'Profil' introuvable.");

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
    Logger.log("Erreur critique _chargerProfils: " + e.message + "\n" + e.stack);
    return {};
  }
}

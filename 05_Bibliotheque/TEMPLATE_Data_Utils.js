/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Data_Utils.gs
 * == RÔLE    : Fonctions utilitaires pour l'accès aux données (BDD), la
 * ==           détermination des profils et la normalisation de textes.
 * =================================================================================
 */

/**
 * Retourne les identifiants des classeurs principaux du système.
 * Remplacer les valeurs par les vrais IDs de vos fichiers.
 */
function getSystemIds() {
  return {
    ID_BDD: '1m2MGBd0nyiAl3qw032B6Nfj7zQL27bRSBexiOPaRZd8',       // Remplacez par l'ID de votre BDD
    ID_CONFIG: '1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ'  // Remplacez par l'ID de votre fichier de config
  };
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

/**
 * Détermine le profil final pour les tests r&K basés sur des seuils.
 */
function _determinerProfilFinalParSeuils_rK(scoresData, typeTest, langue) {
  const profilsMap = _chargerProfils(typeTest, langue);
  let profilFinal = 'DEFAUT'; // Un profil par défaut si aucun seuil n'est atteint
  let plusHautSeuilAtteint = -1;

  for (const codeProfil in profilsMap) {
    const profilInfo = profilsMap[codeProfil];
    const seuil = parseFloat(profilInfo.Seuil);
    const scoreAxe = scoresData[profilInfo.Axe];

    if (!isNaN(seuil) && scoreAxe >= seuil && seuil > plusHautSeuilAtteint) {
      profilFinal = codeProfil;
      plusHautSeuilAtteint = seuil;
    }
  }
  return { profilFinal: profilFinal };
}

function _chargerProfils(typeTest, langue) {
  try {
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const nomFeuil = `Profils_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuil);
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    const headers = data.shift().map(h => String(h || '').trim());

    const profilsMap = {};
    let codeColIndex = headers.indexOf('Code_Profil');
    if (codeColIndex === -1) {
      codeColIndex = headers.indexOf('Profil'); // Fallback
      if (codeColIndex === -1) {
        _log(DEBUG_DATA_LOADING, "ERREUR _chargerProfils : Colonne 'Code_Profil' ou 'Profil' introuvable.");
        return {};
      }
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
    const nomFeuil = `Questions_${typeTest}_${langue}`;
    const sheet = bdd.getSheetByName(nomFeuil);
    if (!sheet) throw new Error(`Feuille introuvable: ${nomFeuil}`);
    const data = sheet.getDataRange().getValues();
    const headersRaw = data.shift();
    const headers = (headersRaw || []).map(h => String(h || '').replace(/^\uFEFF/, '').replace(/^"|"$/g, '').trim());
    const idCol = headers.indexOf('ID');
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
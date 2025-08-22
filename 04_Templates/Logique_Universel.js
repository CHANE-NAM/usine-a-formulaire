// =================================================================================
// == FICHIER : Logique_Universel.gs
// == VERSION : 7.0 (Chargement dynamique de toutes les colonnes de profil)
// == RÔLE  : Moteur de calcul universel capable de traiter n'importe quel test.
// =================================================================================

/**
 * Fonction principale du moteur universel. Gère la logique multilingue.
 * @param {object} reponsesUtilisateur - Les réponses lues depuis la feuille.
 * @param {string} langueCible - La langue demandée pour le résultat (ex: 'EN').
 * @param {object} config - La configuration du test.
 * @param {string} langueOrigine - La langue dans laquelle l'utilisateur a répondu (ex: 'FR').
 * @returns {object} Un objet contenant toutes les données du résultat, y compris les colonnes du profil.
 */
function calculerResultats(reponsesUtilisateur, langueCible, config, langueOrigine) {
  // Initialisation simple de l'objet de résultats
  let resultats = {
    scoresData: {},
    sousTotauxParMode: {}
  };

  const profilsMap = _chargerProfils(config.Type_Test, langueCible);
  const questionsMapCible = _chargerQuestions(config.Type_Test, langueCible);
  if (!questionsMapCible) return {}; // Sécurité

  if (langueOrigine === langueCible) {
    _executerCalcul(reponsesUtilisateur, questionsMapCible, resultats);
  } else {
    // La logique de traduction complexe reste inchangée...
    const questionsMapOrigine = _chargerQuestions(config.Type_Test, langueOrigine);
    if (!questionsMapOrigine) return {};

    for (const enTeteComplet in reponsesUtilisateur) {
      if (!enTeteComplet.includes(':')) continue;
      const idQuestion = enTeteComplet.split(':')[0].trim();
      const questionConfigCible = questionsMapCible[idQuestion];

      if (questionConfigCible) {
        if (questionConfigCible.parametres.mode === 'ECHELLE_NOTE') {
          _aiguillerCalcul(questionConfigCible.parametres.mode, reponsesUtilisateur[enTeteComplet], questionConfigCible.parametres, resultats);
        } else {
          const questionConfigOrigine = questionsMapOrigine[idQuestion];
          if (questionConfigOrigine && questionConfigOrigine.parametres.options) {
            const reponseTexte = reponsesUtilisateur[enTeteComplet];
            const reponsesArray = String(reponseTexte).split(',').map(r => r.trim());

            reponsesArray.forEach(reponseSimple => {
              const optionIndex = questionConfigOrigine.parametres.options.findIndex(opt => opt.libelle === reponseSimple);
              if (optionIndex !== -1 && questionConfigCible.parametres.options && questionConfigCible.parametres.options[optionIndex]) {
                const optionCible = questionConfigCible.parametres.options[optionIndex];
                _aiguillerCalcul(questionConfigCible.parametres.mode, optionCible.libelle, questionConfigCible.parametres, resultats);
              }
            });
          }
        }
      }
    }
  }

  // ==================== DÉBUT DE LA MODIFICATION FINALE ====================
  // La détermination du profil final se fait une seule fois à la fin.
  if (Object.keys(resultats.scoresData).length > 0) {
    resultats.profilFinal = _determinerProfilFinal(resultats.scoresData, config.Type_Test, langueCible);
    
    // On récupère TOUTES les données du profil final
    const toutesLesDonneesDuProfil = profilsMap[resultats.profilFinal];

    if (toutesLesDonneesDuProfil) {
      // On fusionne les données du profil (Titre_Profil, Description_Profil, ConseilCarriere, etc.)
      // avec l'objet de résultats existant.
      resultats = { ...resultats, ...toutesLesDonneesDuProfil };
    }
    
    // On garde cette map pour la brique Ligne_Score si besoin
    resultats.mapCodeToName = _creerMapCodeVersNom(profilsMap);
  }
  // ===================== FIN DE LA MODIFICATION FINALE =====================

  Logger.log("Calculs terminés. Résultats : " + JSON.stringify(resultats));
  return resultats;
}


/**
 * Exécute la logique de calcul sur un jeu de réponses et une configuration de questions.
 */
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


/**
 * Aiguille le calcul vers la bonne sous-fonction en fonction du mode de traitement.
 */
function _aiguillerCalcul(mode, reponse, parametres, resultats) {
    switch (mode) {
      case 'QCU_DIRECT': _traiterQCU_DIRECT(reponse, parametres, resultats); break;
      case 'QCU_CAT': _traiterQCU_CAT(reponse, parametres, resultats); break;
      case 'QRM_CAT': _traiterQRM_CAT(reponse, parametres, resultats); break;
      case 'ECHELLE_NOTE': _traiterECHELLE_NOTE(reponse, parametres, resultats); break;
      default:
        Logger.log(`Mode de traitement inconnu ou non implémenté : ${mode}`);
        break;
    }
}

// =================================================================================
// == SOUS-FONCTIONS DE CALCUL (une par mode de traitement)
// =================================================================================
// CES FONCTIONS RESTENT INCHANGÉES...
function _traiterQCU_DIRECT(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.profil) return;
  resultats.scoresData[parametres.profil] = reponseUtilisateur;
}

function _traiterQCU_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;
  const optionTrouvee = parametres.options.find(opt => opt.libelle === reponseUtilisateur);
  if (optionTrouvee && optionTrouvee.profil) {
    const profil = optionTrouvee.profil;
    const valeur = (typeof optionTrouvee.valeur === 'number') ? optionTrouvee.valeur : 1;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
  }
}

function _traiterQRM_CAT(reponseUtilisateur, parametres, resultats) {
  if (!reponseUtilisateur || !parametres || !parametres.options) return;
  const reponsesArray = String(reponseUtilisateur).split(',').map(r => r.trim());
  reponsesArray.forEach(reponse => {
    const optionTrouvee = parametres.options.find(opt => opt.libelle === reponse);
    if (optionTrouvee && optionTrouvee.profil && typeof optionTrouvee.valeur === 'number') {
      const profil = optionTrouvee.profil;
      const valeur = optionTrouvee.valeur;
      resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeur;
    }
  });
}

function _traiterECHELLE_NOTE(reponseUtilisateur, parametres, resultats) {
  if (!parametres || !parametres.profil) return;
  const valeurNumerique = parseInt(reponseUtilisateur, 10);
  if (!isNaN(valeurNumerique)) {
    const profil = parametres.profil;
    resultats.scoresData[profil] = (resultats.scoresData[profil] || 0) + valeurNumerique;
  }
}

// =================================================================================
// == FONCTIONS UTILITAIRES INTERNES
// =================================================================================
// CES FONCTIONS RESTENT INCHANGÉES...
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
      if (codeProfilSeuil.toUpperCase() !== profilMajoritaire.toUpperCase()) continue;
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
  if (testsAvecSeuils.some(t => typeTest.toUpperCase() === t.toUpperCase())) {
    return _determinerProfilFinalParSeuils(scoresData, typeTest, langue);
  }
  if (typeTest.toUpperCase() === 'MBTI') {
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
        // MODIFIÉ : On cherche la clé 'Titre_Profil' ou une clé similaire
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
    const headers = data.shift();
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

// =================================================================================
// == DÉBUT DE LA SECTION PRINCIPALE MODIFIÉE
// =================================================================================

/**
 * Charge les profils depuis la BDD et retourne un objet où chaque clé est un code de profil
 * et chaque valeur est un objet contenant TOUTES les données de la ligne correspondante.
 * @param {string} typeTest Le type de test.
 * @param {string} langue La langue des profils à charger.
 * @returns {object} Une carte des profils avec toutes leurs données.
 */
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

    // Trouve la colonne qui sert de clé (Code_Profil ou Profil)
    const codeColIndex = headers.indexOf('Code_Profil') > -1 ? headers.indexOf('Code_Profil') : headers.indexOf('Profil');
    if (codeColIndex === -1) throw new Error("Colonne 'Code_Profil' ou 'Profil' introuvable.");

    data.forEach(row => {
      const codeProfil = row[codeColIndex];
      if (codeProfil) {
        profilsMap[codeProfil] = {}; // Crée un objet pour ce profil
        // Boucle sur toutes les colonnes pour remplir l'objet
        headers.forEach((header, index) => {
          if (header) {
            profilsMap[codeProfil][header] = row[index];
          }
        });
      }
    });
    
    return profilsMap;
  } catch (e) {
      Logger.log("Erreur critique _chargerProfils: " + e.message + "\n" + e.stack);
      return {};
  }
}
// =================================================================================
// == FIN DE LA SECTION MODIFIÉE
// =================================================================================
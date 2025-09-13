/**
 * =================================================================================
 * == FICHIER : TEMPLATE_Moteur_Standard.gs
 * == RÔLE    : Moteur de calcul pour les tests standards (MBTI, Couleurs, etc.).
 * =================================================================================
 */

function _executerCalculStandard(reponses, questionsMap, resultats, nbQuestionsLimite) {
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
    case 'QCU_CAT':
      _traiterQCU_CAT(reponse, parametres, resultats);
      break;
    case 'ECHELLE_NOTE':
    case 'LIKERT_5':
      _traiterECHELLE_NOTE(reponse, parametres, resultats);
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
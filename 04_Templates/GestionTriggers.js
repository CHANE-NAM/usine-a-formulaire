// =================================================================================
// == FICHIER : GestionTriggers.gs
// == VERSION : 1.0 (Création initiale)
// == RÔLE  : Gère la création et l'exécution des envois d'e-mails différés.
// =================================================================================

/**
 * Calcule le délai en millisecondes à partir d'une chaîne de caractères (ex: "4h", "1j").
 * @param {string} valeurDelai - La chaîne de caractères représentant le délai.
 * @returns {number} Le délai en millisecondes. Retourne 0 si le format est invalide.
 */
function _calculerDelaiEnMs(valeurDelai) {
  if (!valeurDelai || typeof valeurDelai !== 'string') return 0;

  const valeurNumerique = parseInt(valeurDelai.replace(/[^0-9]/g, ''), 10);
  if (isNaN(valeurNumerique)) return 0;

  if (valeurDelai.includes('h')) {
    return valeurNumerique * 60 * 60 * 1000; // Heures en millisecondes
  } else if (valeurDelai.includes('j')) {
    return valeurNumerique * 24 * 60 * 60 * 1000; // Jours en millisecondes
  } else if (valeurDelai.includes('min')) {
    return valeurNumerique * 60 * 1000; // Minutes en millisecondes
  }

  return 0; // Format non reconnu
}


/**
 * Programme l'envoi différé de l'e-mail de résultats.
 * Crée un déclencheur unique et sauvegarde les informations nécessaires.
 */
function programmerEnvoiResultats(rowIndex, langueCible, delai) {
  try {
    const delaiEnMs = _calculerDelaiEnMs(delai);
    if (delaiEnMs <= 0) {
      Logger.log(`Délai invalide ou nul (${delai}). Annulation de la programmation.`);
      return;
    }

    // Identifiant unique pour ce déclencheur et ses données
    const proprieteId = `envoiDiffere_${new Date().getTime()}_${rowIndex}`;

    // 1. Sauvegarder les informations nécessaires pour l'envoi
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(proprieteId, JSON.stringify({
      rowIndex: rowIndex,
      langueCible: langueCible
    }));

    // 2. Créer le déclencheur qui s'exécutera après le délai
    ScriptApp.newTrigger('envoyerEmailProgramme')
      .timeBased()
      .after(delaiEnMs)
      .create();

    Logger.log(`Envoi programmé avec succès pour la ligne ${rowIndex}. Délai : ${delai}. ID de propriété : ${proprieteId}`);

  } catch (e) {
    Logger.log(`ERREUR lors de la programmation de l'envoi pour la ligne ${rowIndex}: ${e.toString()}\n${e.stack}`);
  }
}

/**
 * Fonction exécutée par le déclencheur pour envoyer l'e-mail de résultats.
 * @param {object} e - L'objet événement passé par le déclencheur.
 */
function envoyerEmailProgramme(e) {
  const properties = PropertiesService.getScriptProperties();
  const toutesLesProps = properties.getProperties();

  // On cherche la première propriété correspondant à un envoi différé
  const proprieteId = Object.keys(toutesLesProps).find(key => key.startsWith('envoiDiffere_'));

  if (!proprieteId) {
    Logger.log("Déclencheur d'envoi programmé exécuté, mais aucune propriété de tâche trouvée. Annulation.");
    return;
  }

  try {
    const donnees = JSON.parse(properties.getProperty(proprieteId));
    const { rowIndex, langueCible } = donnees;

    Logger.log(`Exécution de l'envoi programmé pour la ligne ${rowIndex} (ID: ${proprieteId})`);

    // Reconstituer le contexte nécessaire
    const config = getTestConfiguration(); // Assurez-vous que cette fonction est accessible
    const reponse = _creerObjetReponse(rowIndex); // Et celle-ci aussi
    const langueOrigine = getOriginalLanguage(reponse);
    const resultats = calculerResultats(reponse, langueCible, config, langueOrigine);

    // Envoyer l'e-mail
    assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, {});

    // Nettoyage : supprimer la propriété
    properties.deleteProperty(proprieteId);
    Logger.log(`Nettoyage de la propriété ${proprieteId} terminé.`);

  } catch (err) {
    Logger.log(`ERREUR FATALE lors de l'exécution de l'envoi programmé (ID: ${proprieteId}): ${err.toString()}\n${err.stack}`);
    // On supprime quand même la propriété pour éviter des erreurs en boucle
    properties.deleteProperty(proprieteId);
  } finally {
    // Nettoyage : supprimer le déclencheur qui vient de s'exécuter
    if (e && e.triggerUid) {
      const allTriggers = ScriptApp.getProjectTriggers();
      for (const trigger of allTriggers) {
        if (trigger.getUniqueId() === e.triggerUid) {
          ScriptApp.deleteTrigger(trigger);
          Logger.log(`Déclencheur ${e.triggerUid} auto-détruit avec succès.`);
          break;
        }
      }
    }
  }
}
/**
 * @OnlyCurrentDoc
 * Fichier : TriggerHandler.gs
 * Rôle : Gérer la réception des données du formulaire via un déclencheur onFormSubmit.
 */

/**
 * Fonction appelée par le déclencheur onFormSubmit créé par le [MOTEUR].
 * Écrit les réponses dans la feuille, puis lance le traitement complet.
 * @param {Object} e L'objet événement de la soumission du formulaire.
 */
function onFormSubmitTrigger(e) {
  try {
    if (!e || !e.response) {
      Logger.log("Déclencheur reçu sans événement de réponse valide.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // On assume que les réponses doivent être écrites dans la première feuille du classeur.
    const sheet = ss.getSheets()[0]; 
    
    const itemResponses = e.response.getItemResponses();
    
    // Si la feuille est vide (première réponse), on écrit les en-têtes.
    if (sheet.getLastRow() === 0) {
      const headers = ["Timestamp"];
      itemResponses.forEach(itemResponse => {
        headers.push(itemResponse.getItem().getTitle());
      });
      sheet.appendRow(headers);
    }
    
    // On construit la nouvelle ligne avec les données de la réponse.
    const newRow = [e.response.getTimestamp()];
    itemResponses.forEach(itemResponse => {
      // Gère les réponses multiples (cases à cocher) en les joignant.
      const response = itemResponse.getResponse();
      if (Array.isArray(response)) {
        newRow.push(response.join(', '));
      } else {
        newRow.push(response);
      }
    });
    
    // On écrit la ligne dans la feuille.
    sheet.appendRow(newRow);
    const newRowIndex = sheet.getLastRow();
    Logger.log(`Nouvelle réponse écrite sur la ligne ${newRowIndex}.`);

    // On lance la logique de traitement principale (calculs, email, etc.).
    if (typeof traiterLigne === "function") {
      Logger.log(`Lancement du traitement complet pour la ligne ${newRowIndex}...`);
      traiterLigne(newRowIndex);
    } else {
      Logger.log("ERREUR : La fonction 'traiterLigne' est introuvable pour lancer le traitement des résultats.");
    }

  } catch (error) {
    Logger.log(`ERREUR critique dans onFormSubmitTrigger : ${error.toString()}\n${error.stack}`);
  }
}
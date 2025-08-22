/**
 * Ce script est un outil de diagnostic à usage unique.
 * Il va créer un formulaire et inspecter l'objet retourné pour
 * comprendre pourquoi la fonction .getShortUrl() n'est pas trouvée.
 */
function testCreationFormulaire() {
  try {
    Logger.log("--- Début du test de diagnostic de création de formulaire ---");
    
    // Étape 1 : On crée un formulaire de test.
    const form = FormApp.create("Test de Diagnostic Ultime");
    Logger.log("Objet 'form' créé.");

    // Étape 2 : On vérifie si la fonction qui pose problème existe VRAIMENT sur cet objet.
    if (form && typeof form.getShortUrl === 'function') {
      Logger.log("--> RÉSULTAT POSITIF : La fonction .getShortUrl() a été trouvée !");
      Logger.log("    Lien court obtenu : " + form.getShortUrl());
    } else {
      Logger.log("--> RÉSULTAT NÉGATIF : La fonction .getShortUrl() est INTROUVABLE sur l'objet 'form'.");
    }
    
    // Étape 3 : On liste toutes les propriétés et méthodes que l'on trouve sur l'objet.
    // Cela nous dira ce qu'est réellement l'objet 'form'.
    let properties = [];
    for (var name in form) {
      properties.push(name);
    }
    Logger.log("Liste de toutes les propriétés trouvées sur l'objet : " + properties.join(', '));

    // On supprime le formulaire de test pour ne pas polluer votre Drive.
    DriveApp.getFileById(form.getId()).setTrashed(true);
    Logger.log("Formulaire de test supprimé.");

  } catch (e) {
    Logger.log("ERREUR CATASTROPHIQUE lors du test de diagnostic : " + e.toString());
    Logger.log(e.stack);
  }
  Logger.log("--- Fin du test de diagnostic ---");
}

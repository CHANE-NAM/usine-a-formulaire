// Remplacez cette variable par l'ID de votre feuille de calcul [CONFIG]V2 Usine à Tests.
// const ID_FEUILLE_CONFIG = "1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ";

/**
 * Fonction à usage unique pour convertir toutes les URLs de formulaires existantes
 * dans l'onglet 'Paramètres Généraux' en leurs versions courtes (forms.gle).
 */
function convertirLiensExistantsEnCourts() {
  const nomOnglet = "Paramètres Généraux";
  
  try {
    const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const sheet = ss.getSheetByName(nomOnglet);
    
    if (!sheet) {
      throw new Error(`L'onglet "${nomOnglet}" est introuvable.`);
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    
    // Trouve automatiquement la colonne contenant les liens
    const linkColumnIndex = headers.indexOf("Lien_Formulaire_Public");
    if (linkColumnIndex === -1) {
      throw new Error("La colonne 'Lien_Formulaire_Public' est introuvable.");
    }

    // Boucle sur chaque ligne (en sautant l'en-tête)
    for (let i = 1; i < values.length; i++) {
      const longUrl = values[i][linkColumnIndex];
      
      // Ne traite que les URLs longues et non vides
      if (longUrl && typeof longUrl === 'string' && longUrl.includes("docs.google.com/forms")) {
        // Extrait l'ID du formulaire à partir de l'URL longue
        const formId = longUrl.split('/d/')[1].split('/')[0];
        
        if (formId) {
          // Ouvre le formulaire par son ID et obtient l'URL courte
          const form = FormApp.openById(formId);
          const shortUrl = form.getShortUrl();
          
          // Met à jour la cellule avec la nouvelle URL courte
          // Les indices de range commencent à 1, donc i+1 et linkColumnIndex+1
          sheet.getRange(i + 1, linkColumnIndex + 1).setValue(shortUrl);
          Logger.log(`Ligne ${i + 1}: URL convertie pour le formulaire ${formId}`);
        }
      }
    }
    
    SpreadsheetApp.getUi().alert("Conversion terminée avec succès !");
    
  } catch (e) {
    Logger.log(`Erreur lors de la conversion : ${e.toString()}`);
    SpreadsheetApp.getUi().alert(`Une erreur est survenue : ${e.message}`);
  }
}
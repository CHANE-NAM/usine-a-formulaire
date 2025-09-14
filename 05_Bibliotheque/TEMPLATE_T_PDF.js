/**
 * =================================================================================
 * == FICHIER : TEMPLATE_T_PDF.gs
 * == VERSION : 2.0 - Version robustifiée
 * == RÔLE    : Gère la génération de fichiers PDF à partir de modèles Google Docs.
 * =================================================================================
 */

/**
 * Génère un fichier PDF à partir d'un modèle Google Doc en remplaçant des variables.
 * Cette version améliorée vérifie le type de chaque variable avant de tenter un remplacement
 * pour éviter les erreurs avec des objets complexes, et assure un nettoyage fiable des fichiers temporaires.
 *
 * @param {string} templateId L'ID du fichier Google Doc servant de modèle.
 * @param {Object} variables Un objet où les clés sont les noms des placeholders.
 * @param {string} nomFichier Le nom du nouveau fichier PDF (sans l'extension).
 * @returns {Blob|null} Le fichier PDF sous forme de Blob, ou null en cas d'erreur.
 */
function genererPdfDepuisModele(templateId, variables, nomFichier) {
  let newDocId = null; // Déclarer l'ID ici pour qu'il soit accessible dans le 'finally'
  try {
    const tempFile = DriveApp.getFileById(templateId);
    const tempFolder = tempFile.getParents().next(); 
    newDocId = tempFile.makeCopy(nomFichier, tempFolder).getId();
    const newDoc = DocumentApp.openById(newDocId);
    const body = newDoc.getBody();

    // Remplacement des variables textuelles et numériques
    for (const key in variables) {
      const placeholder = `{{${key}}}`;
      const valeur = variables[key];

      // --- CORRECTION CRUCIALE ---
      // On vérifie que la valeur est bien un type simple (texte, nombre, booléen)
      // avant de tenter de l'insérer dans le document.
      if (typeof valeur === 'string' || typeof valeur === 'number' || typeof valeur === 'boolean') {
        // La valeur est simple, on peut la remplacer sans risque.
        body.replaceText(placeholder, valeur.toString());
      }
      // Si la valeur est un objet (comme 'scoresData'), on l'ignore simplement.
    }

    // Cas particulier pour le graphique radar (si utilisé)
    if (variables.Graphique_Radar_Blob) {
      const placeholderGraphique = '{{Graphique_Radar_Blob}}';
      const rangeElement = body.findText(placeholderGraphique);
      if (rangeElement) {
        const element = rangeElement.getElement();
        const parent = element.getParent();
        if (parent.asParagraph) {
            parent.asParagraph().clear().insertInlineImage(0, variables.Graphique_Radar_Blob);
        }
      }
    }

    newDoc.saveAndClose();
    const pdfBlob = DriveApp.getFileById(newDocId).getBlob().getAs('application/pdf');
    pdfBlob.setName(nomFichier + '.pdf');
    
    // Le nettoyage du fichier temporaire est géré dans le bloc 'finally' pour plus de sécurité.
    
    return pdfBlob;

  } catch (e) {
    Logger.log(`Erreur critique lors de la génération du PDF depuis le modèle ${templateId}: ${e.message}\n${e.stack}`);
    return null;

  } finally {
    // S'assure que le fichier Google Doc temporaire est supprimé, même en cas d'erreur.
    if (newDocId) {
      try { 
        DriveApp.getFileById(newDocId).setTrashed(true); 
      } catch (e2) {
        // Ignore les erreurs si le fichier a déjà été supprimé ou est inaccessible.
      }
    }
  }
}
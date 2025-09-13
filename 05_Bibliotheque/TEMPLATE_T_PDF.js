/**
 * T_PDF.gs
 * Génère un fichier PDF à partir d'un modèle Google Doc en remplaçant des variables.
 * @param {string} templateId L'ID du fichier Google Doc servant de modèle.
 * @param {Object} variables Un objet où les clés sont les noms des variables à remplacer (sans les accolades).
 * @param {string} nomFichier Le nom du nouveau fichier PDF (sans l'extension).
 * @returns {Blob|null} Le fichier PDF sous forme de Blob, ou null en cas d'erreur.
 */
function genererPdfDepuisModele(templateId, variables, nomFichier) {
  try {
    const tempFile = DriveApp.getFileById(templateId);
    const tempFolder = DriveApp.getRootFolder();
    const newDocId = tempFile.makeCopy(nomFichier, tempFolder).getId();
    const newDoc = DocumentApp.openById(newDocId);
    const body = newDoc.getBody();
    
    // Insertion du graphique si le blob est fourni
    if (variables.Graphique_Radar_Blob) {
      const placeholder = '{{Graphique_Radar_Blob}}';
      const rangeElement = body.findText(placeholder);
      if (rangeElement) {
        const element = rangeElement.getElement();
        const parent = element.getParent();
        parent.asParagraph().clear().insertInlineImage(0, variables.Graphique_Radar_Blob);
      } else {
        Logger.log("Avertissement : Le placeholder {{Graphique_Radar_Blob}} n'a pas été trouvé dans le modèle de document.");
      }
    }

    // Remplacement des autres variables textuelles
    for (const key in variables) {
      if (key !== 'Graphique_Radar_Blob') { 
        body.replaceText(`{{${key}}}`, variables[key] || '');
      }
    }

    newDoc.saveAndClose();
    const pdfBlob = DriveApp.getFileById(newDocId).getBlob().getAs('application/pdf');
    pdfBlob.setName(nomFichier + '.pdf');
    DriveApp.getFileById(newDocId).setTrashed(true);
    return pdfBlob;
    
  } catch (e) {
    Logger.log(`Erreur lors de la génération du PDF depuis le modèle ${templateId}: ${e.message}`);
    return null;
  }
}
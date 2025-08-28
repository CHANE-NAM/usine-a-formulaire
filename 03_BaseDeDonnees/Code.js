/**
 * @OnlyCurrentDoc
 * Crée un menu personnalisé dans l'interface utilisateur de la feuille de calcul
 * pour lancer les fonctions utilitaires.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚙️ Utilitaires BDD')
      .addItem('Lister les fichiers d\'un dossier Drive', 'listFilesFromFolder')
      .addToUi();
}

/**
 * Demande à l'utilisateur l'ID d'un dossier Drive, puis liste tous les fichiers
 * de ce dossier (et optionnellement des sous-dossiers) à la suite des données.
 */
function listFilesFromFolder() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Demander l'ID du dossier à l'utilisateur
  const result = ui.prompt(
      'Lister les Fichiers Drive',
      'Veuillez coller l\'ID du dossier Google Drive contenant vos rapports :',
      ui.ButtonSet.OK_CANCEL);

  if (result.getSelectedButton() !== ui.Button.OK || !result.getResponseText()) {
    return;
  }
  
  const folderId = result.getResponseText().trim();

  // NOUVEAU : Demander si l'on doit inclure les sous-dossiers
  const recursiveSearchResponse = ui.alert(
      'Recherche approfondie',
      'Voulez-vous inclure les fichiers des sous-dossiers ?',
      ui.ButtonSet.YES_NO);
      
  const shouldRecurse = (recursiveSearchResponse === ui.Button.YES);
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    const filesToAdd = [];

    // NOUVEAU : Lancer la recherche simple ou récursive en fonction de la réponse
    if (shouldRecurse) {
        // Lancer la recherche récursive
        getFilesRecursive(folder, filesToAdd);
    } else {
        // Lancer la recherche simple (uniquement le dossier racine)
        const files = folder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            filesToAdd.push([file.getName(), file.getId()]);
        }
    }
    
    if (filesToAdd.length === 0) {
      ui.alert('Information', `Aucun fichier n'a été trouvé dans le dossier "${folder.getName()}" (et ses sous-dossiers, si l'option était choisie).`, ui.ButtonSet.OK);
      return;
    }
    
    // 3. Écrire les résultats dans la feuille de calcul
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let outputSheet = ss.getSheetByName('Liste_Fichiers_Drive');
    
    if (!outputSheet) {
      outputSheet = ss.insertSheet('Liste_Fichiers_Drive', 0);
    }
    
    const lastRow = outputSheet.getLastRow();
    let startRow;

    if (lastRow === 0) {
      outputSheet.getRange(1, 1, 1, 2).setValues([['Nom du Fichier', 'ID du Fichier']]);
      startRow = 2;
    } else {
      startRow = lastRow + 1;
    }
    
    outputSheet.getRange(startRow, 1, filesToAdd.length, 2).setValues(filesToAdd);
    outputSheet.autoResizeColumns(1, 2);
    outputSheet.activate();
    
    ui.alert('Opération terminée', `${filesToAdd.length} nouveau(x) fichier(s) ont été ajoutés dans l'onglet "Liste_Fichiers_Drive".`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(e.toString());
    ui.alert('Erreur', 'Impossible d\'accéder au dossier. Veuillez vérifier que l\'ID est correct et que vous avez les droits d\'accès.', ui.ButtonSet.OK);
  }
}

/**
 * Fonction auxiliaire récursive pour lister les fichiers.
 * @param {Folder} folder - Le dossier à parcourir.
 * @param {Array} fileList - Le tableau où ajouter les fichiers trouvés.
 */
function getFilesRecursive(folder, fileList) {
    // Ajouter les fichiers du dossier courant
    const files = folder.getFiles();
    while (files.hasNext()) {
        const file = files.next();
        fileList.push([file.getName(), file.getId()]);
    }

    // Parcourir les sous-dossiers et s'appeler soi-même
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        const subFolder = subFolders.next();
        getFilesRecursive(subFolder, fileList);
    }
}
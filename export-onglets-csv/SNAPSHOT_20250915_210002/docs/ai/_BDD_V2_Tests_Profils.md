# _BDD_V2_Tests_Profils

> Généré automatiquement depuis **scripts__BDD_V2_Tests_Profils.txt** — snapshot: **SNAPSHOT_20250915_210002**.

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\03_BaseDeDonnées\appsscript.json

```json

{
  "timeZone": "Indian/Mauritius",
  "dependencies": {
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

## G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm\03_BaseDeDonnées\Code.js

```javascript

/**
Â * @OnlyCurrentDoc
Â * CrÃ©e un menu personnalisÃ© dans l'interface utilisateur de la feuille de calcul
Â * pour lancer les fonctions utilitaires.
Â */
function onOpen() {
Â  SpreadsheetApp.getUi()
Â  Â  Â  .createMenu('âš™ï¸ Utilitaires BDD')
Â  Â  Â  .addItem('Lister les fichiers d\'un dossier Drive', 'listFilesFromFolder')
Â  Â  Â  .addToUi();
}

/**
Â * Demande Ã  l'utilisateur l'ID d'un dossier Drive, puis liste tous les fichiers
Â * de ce dossier (et optionnellement des sous-dossiers) Ã  la suite des donnÃ©es.
Â */
function listFilesFromFolder() {
Â  const ui = SpreadsheetApp.getUi();
Â  
Â  // 1. Demander l'ID du dossier Ã  l'utilisateur
Â  const result = ui.prompt(
Â  Â  Â  'Lister les Fichiers Drive',
Â  Â  Â  'Veuillez coller l\'ID du dossier Google Drive contenant vos rapports :',
Â  Â  Â  ui.ButtonSet.OK_CANCEL);

Â  if (result.getSelectedButton() !== ui.Button.OK || !result.getResponseText()) {
Â  Â  return;
Â  }
Â  
Â  const folderId = result.getResponseText().trim();

  // NOUVEAU : Demander si l'on doit inclure les sous-dossiers
  const recursiveSearchResponse = ui.alert(
      'Recherche approfondie',
      'Voulez-vous inclure les fichiers des sous-dossiers ?',
      ui.ButtonSet.YES_NO);
      
  const shouldRecurse = (recursiveSearchResponse === ui.Button.YES);
Â  
Â  try {
    const folder = DriveApp.getFolderById(folderId);
    const filesToAdd = [];

    // NOUVEAU : Lancer la recherche simple ou rÃ©cursive en fonction de la rÃ©ponse
    if (shouldRecurse) {
        // Lancer la recherche rÃ©cursive
        getFilesRecursive(folder, filesToAdd);
    } else {
        // Lancer la recherche simple (uniquement le dossier racine)
        const files = folder.getFiles();
        while (files.hasNext()) {
            const file = files.next();
            filesToAdd.push([file.getName(), file.getId()]);
        }
    }
Â  Â  
Â  Â  if (filesToAdd.length === 0) {
Â  Â  Â  ui.alert('Information', `Aucun fichier n'a Ã©tÃ© trouvÃ© dans le dossier "${folder.getName()}" (et ses sous-dossiers, si l'option Ã©tait choisie).`, ui.ButtonSet.OK);
Â  Â  Â  return;
Â  Â  }
Â  Â  
Â  Â  // 3. Ã‰crire les rÃ©sultats dans la feuille de calcul
Â  Â  const ss = SpreadsheetApp.getActiveSpreadsheet();
Â  Â  let outputSheet = ss.getSheetByName('Liste_Fichiers_Drive');
Â  Â  
Â  Â  if (!outputSheet) {
Â  Â  Â  outputSheet = ss.insertSheet('Liste_Fichiers_Drive', 0);
Â  Â  }
    
    const lastRow = outputSheet.getLastRow();
    let startRow;

    if (lastRow === 0) {
      outputSheet.getRange(1, 1, 1, 2).setValues([['Nom du Fichier', 'ID du Fichier']]);
      startRow = 2;
    } else {
      startRow = lastRow + 1;
    }
Â  Â  
Â  Â  outputSheet.getRange(startRow, 1, filesToAdd.length, 2).setValues(filesToAdd);
    outputSheet.autoResizeColumns(1, 2);
    outputSheet.activate();
Â  Â  
Â  Â  ui.alert('OpÃ©ration terminÃ©e', `${filesToAdd.length} nouveau(x) fichier(s) ont Ã©tÃ© ajoutÃ©s dans l'onglet "Liste_Fichiers_Drive".`, ui.ButtonSet.OK);

Â  } catch (e) {
Â  Â  Logger.log(e.toString());
Â  Â  ui.alert('Erreur', 'Impossible d\'accÃ©der au dossier. Veuillez vÃ©rifier que l\'ID est correct et que vous avez les droits d\'accÃ¨s.', ui.ButtonSet.OK);
Â  }
}

/**
 * Fonction auxiliaire rÃ©cursive pour lister les fichiers.
 * @param {Folder} folder - Le dossier Ã  parcourir.
 * @param {Array} fileList - Le tableau oÃ¹ ajouter les fichiers trouvÃ©s.
 */
function getFilesRecursive(folder, fileList) {
    // Ajouter les fichiers du dossier courant
    const files = folder.getFiles();
    while (files.hasNext()) {
        const file = files.next();
        fileList.push([file.getName(), file.getId()]);
    }

    // Parcourir les sous-dossiers et s'appeler soi-mÃªme
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        const subFolder = subFolders.next();
        getFilesRecursive(subFolder, fileList);
    }
}
```


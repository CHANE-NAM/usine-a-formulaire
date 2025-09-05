function forcerAutorisation() {
  // Cette simple ligne est suffisante pour demander les autorisations Drive.
  DriveApp.getRootFolder(); 
  SpreadsheetApp.getUi().alert('Autorisation accordée ! Vous pouvez maintenant retourner à votre feuille de calcul et relancer le déploiement.');
}
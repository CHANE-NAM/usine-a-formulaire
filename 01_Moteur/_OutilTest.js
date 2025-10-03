/**
 * Outil temporaire pour générer un code de test chiffré pour le [HANDLER].
 */
function genererCodePourTest() {
  // --- MODIFIEZ LE NUMÉRO DE LIGNE ICI ---
  const rowIndex = 19; // <--- Mettez ici le numéro de la ligne à tester
  // -----------------------------------------

  const SECRET_KEY = "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC";
  const payload = { rowIndex: rowIndex };
  
  // Chiffrement du payload
  const encryptedCode = CryptoJS.AES.encrypt(JSON.stringify(payload), SECRET_KEY).toString();
  
  // Affichage du code dans les journaux d'exécution
  Logger.log("--- CODE DE TEST POUR LA LIGNE " + rowIndex + " ---");
  Logger.log("Copiez le code ci-dessous (sans les guillemets) :");
  Logger.log(encryptedCode);
  Logger.log("-----------------------------------------");
}
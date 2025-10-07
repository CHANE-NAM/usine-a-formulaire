/**
 * Clé secrète pour le déchiffrement. DOIT être la même que dans le MOTEUR.
 */
const SECRET_KEY = "FELIX QUI POTUIT RERUM COGNOCERE CAUSA VIC";

/**
 * ID de la feuille de calcul de configuration [CONFIG].
 */
const ID_FEUILLE_CONFIG = '1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ';

/**
 * Fonction principale qui s'exécute lorsqu'un utilisateur accède à l'URL de l'application Web.
 * @param {Object} e - L'objet d'événement contenant les paramètres de la requête.
 */
function doGet(e) {
  try {
    // 1. Vérifier et déchiffrer le code
    if (!e.parameter.code) {
      return HtmlService.createHtmlOutput('<h1>Erreur : Paramètre d\'identification du test manquant.</h1>');
    }
    const encryptedCode = e.parameter.code.replace(/ /g, '+');

    let payload;
    try {
      const decryptedBytes = CryptoJS.AES.decrypt(encryptedCode, SECRET_KEY);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
      if (!decryptedText) throw new Error("Le déchiffrement a produit un texte vide.");
      payload = JSON.parse(decryptedText);
    } catch (err) {
      Logger.log('Erreur de déchiffrement : ' + err.message);
      return HtmlService.createHtmlOutput('<h1>Erreur : Le code du test est invalide ou a expiré.</h1>');
    }

    // 2. Récupérer la configuration complète du test
    const config = getConfigFromRow(payload.rowIndex);
    if (!config) {
      return HtmlService.createHtmlOutput('<h1>Erreur : Impossible de trouver la configuration pour ce test.</h1>');
    }

    // 3. Aiguillage en fonction du parcours utilisateur
    switch (config.Mode_Acces_Test) {
      case 'Payant':
        return afficherPagePaiement(config, encryptedCode);

      case 'B2B':
        return afficherPageB2B(config, encryptedCode);

      case 'CLOS': {
        const userEmail = Session.getActiveUser().getEmail();
        const isAdmin = estAdmin(userEmail);
        if (!isAdmin) {
          return HtmlService.createHtmlOutput('<h1>Ce test est actuellement fermé.</h1>');
        }
        return afficherPageIdentificationGratuite(config, encryptedCode);
      }

      case 'Gratuit':
      default:
        return afficherPageIdentificationGratuite(config, encryptedCode);
    }

  } catch (error) {
    Logger.log("ERREUR FATALE dans doGet(e) : " + error.toString());
    return HtmlService.createHtmlOutput("<h1>Une erreur interne est survenue.</h1><p>Veuillez contacter l'administrateur.</p>");
  }
}

// ===============================================================
// == FONCTIONS D'AFFICHAGE DES PAGES
// ===============================================================

/**
 * Affiche la page d'identification standard pour un accès gratuit.
 */
function afficherPageIdentificationGratuite(config, encryptedCode) {
  const htmlTemplate = HtmlService.createTemplateFromFile('Index.html');
  htmlTemplate.config = config;
  htmlTemplate.encryptedCode = encryptedCode;
  return htmlTemplate.evaluate()
    .setTitle(config['Titre_Formulaire_Utilisateur'])
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Affiche la page d'identification pour un accès payant.
 */
function afficherPagePaiement(config, encryptedCode) {
  const template = HtmlService.createTemplateFromFile('Paiement.html');
  template.encryptedCode = encryptedCode;
  template.titreTest = config.Titre_Formulaire_Utilisateur;
  return template.evaluate().setTitle("Identification pour le test");
}

/**
 * Affiche la page d'identification avec code d'accès pour le B2B.
 */
function afficherPageB2B(config, encryptedCode) {
  const template = HtmlService.createTemplateFromFile('B2B.html');
  template.encryptedCode = encryptedCode;
  template.titreTest = config.Titre_Formulaire_Utilisateur;
  return template.evaluate().setTitle("Accès Entreprise");
}

// ===============================================================
// == FONCTIONS UTILITAIRES
// ===============================================================

/**
 * Récupère la configuration complète pour une ligne donnée.
 */
function getConfigFromRow(rowIndex) {
  if (!rowIndex || rowIndex < 2) throw new Error("rowIndex invalide (attendu: numéro de ligne >= 2).");

  const ss = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const sh = ss.getSheetByName('Paramètres Généraux');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const config = {};
  headers.forEach((h, i) => config[h] = row[i]);
  return config;
}

/**
 * Vérifie si un utilisateur est un administrateur.
 */
function estAdmin(userEmail) {
  const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
  const optionsSheet = configSS.getSheetByName('sys_Options_Parametres');
  const lastRow = optionsSheet.getLastRow();
  if (lastRow < 2) return false;

  const adminEmailsRange = optionsSheet.getRange(2, 26, lastRow - 1, 1).getValues();
  const ADMIN_EMAILS = adminEmailsRange
    .map(r => r[0])
    .filter(v => v != null && String(v).trim() !== '')
    .map(v => String(v).trim().toLowerCase());

  const email = (userEmail || '').trim().toLowerCase();
  return email && ADMIN_EMAILS.includes(email);
}

/**
 * Traite les données soumises depuis le formulaire HTML d'identification.
 * @param {Object} formObject - (nom, email, formType, code, encryptedCode)
 * @returns {Object} { redirectUrl } pour rediriger côté client.
 */
function processFormSubmission(formObject) {

  Logger.log("--- DÉBUT DE VÉRIFICATION ---");
  Logger.log("Données reçues du formulaire (formObject) : " + JSON.stringify(formObject));
  
  try {
    // --- ÉTAPE 1 : Valider et déchiffrer le code pour retrouver la configuration ---
    const encryptedCode = formObject.encryptedCode;
    if (!encryptedCode) {
      throw new Error("Le code de session est manquant. Impossible de continuer.");
    }

    let payload;
    try {
      const decryptedBytes = CryptoJS.AES.decrypt(encryptedCode.replace(/ /g, '+'), SECRET_KEY);
      const decryptedText = decryptedBytes.toString(CryptoJS.enc.Utf8);
      payload = JSON.parse(decryptedText);
    } catch (e) {
      throw new Error("Le code de session est invalide.");
    }

    const rowIndex = Number(payload.rowIndex);
    const config = getConfigFromRow(rowIndex);

    // Ouvre UNE SEULE FOIS la feuille CONFIG pour tout le traitement
    const configSS = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);

    // --- CONTRÔLE ONE-TIME: refuser toute réutilisation du même code chiffré ---
    let tokensSheet = configSS.getSheetByName("Tokens");
    if (!tokensSheet) {
      tokensSheet = configSS.insertSheet("Tokens");
      tokensSheet.appendRow(["Timestamp", "CodeHash", "RowIndex", "Email", "Status"]); // Status: USED | BLOCKED
    }

    // On hash le code chiffré pour ne pas stocker le code en clair
        // On crée une clé unique basée sur le test ET l'email de l'utilisateur
        const userEmail = formObject.email || 'no-email-provided';
        const uniqueKey = `rowIndex:${rowIndex}|email:${userEmail.toLowerCase()}`;

        // On hash cette clé unique pour la stocker de manière sécurisée
        const codeHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, uniqueKey)
          .map(b => (('0' + (b & 0xFF).toString(16)).slice(-2))).join('');
    // Vérifie si déjà consommé
    const lastRowTokens = tokensSheet.getLastRow();
    if (lastRowTokens >= 2) {
      const range = tokensSheet.getRange(2, 2, lastRowTokens - 1, 1).getValues(); // Col B = CodeHash
      const used = range.some(r => String(r[0]) === codeHash);
      if (used) throw new Error("Lien déjà utilisé (sécurité one-time). Contactez l’administrateur.");
    }

    // Marque comme utilisé AVANT de rediriger (anti double-clic / refresh)
    tokensSheet.appendRow([new Date(), codeHash, rowIndex, (formObject.email || ''), "USED"]);
    SpreadsheetApp.flush();

    // --- Logiques spécifiques par type de formulaire ---
    let detailsPourOrders = "";
    if (formObject.formType === 'b2b') {
      const codeAcces = formObject.code; // name="code" dans B2B.html
      Logger.log(`Parcours B2B détecté. Code d'accès fourni : ${codeAcces}`);
      detailsPourOrders = `Code: ${codeAcces}`;
      // NOTE FUTURE : ajouter ici la validation du code d'accès si nécessaire.
    } else if (formObject.formType === 'paiement') {
      Logger.log('Parcours Payant détecté.');
      detailsPourOrders = 'Parcours Payant';
      // NOTE FUTURE : déclencher ici la redirection vers module de paiement.
    }

    // --- ÉTAPE 2 : Enregistrer les informations dans la feuille "Orders" ---
    let ordersSheet = configSS.getSheetByName("Orders");
    if (!ordersSheet) {
      ordersSheet = configSS.insertSheet("Orders");
      ordersSheet.appendRow(["Timestamp", "Nom", "Email", "Test Row Index", "Statut", "Détails"]);
    }
    ordersSheet.appendRow([new Date(), formObject.nom, formObject.email, rowIndex, "IDENTIFIED", detailsPourOrders]);

    // --- ÉTAPE 3 : Récupérer l'URL publique fiable du formulaire final ---
    const finalFormId = config['ID_Formulaire_Cible'];
    if (!finalFormId) {
      throw new Error("Impossible de trouver l'ID du formulaire cible (colonne 'ID_Formulaire_Cible').");
    }

    let redirectUrl;
    try {
      // Chemin fiable : demander l’URL publiée au service Forms
      const form = FormApp.openById(finalFormId);
      // === DÉBUT DE LA MODIFICATION ===

      // 1. Récupère l'URL de base du formulaire
      let baseUrl = form.getPublishedUrl();

      // 2. Prépare le paramètre de langue
      const langCode = formObject.langue || 'FR'; // Le code FR/EN choisi par l'utilisateur
      const langFullName = {
        'FR': 'Français',
        'EN': 'English',
        'ES': 'Español',
        'DE': 'Deutsch'
      }[langCode] || langCode; // Convertit le code en nom complet (ex: "Français")

      // L'identifiant de votre question "Langue" a été intégré ici
      const languageEntryId = 'entry.32207297'; 

      // 3. Construit l'URL finale avec le paramètre de langue prérempli
      redirectUrl = `${baseUrl}?${languageEntryId}=${encodeURIComponent(langFullName)}`;

      // === FIN DE LA MODIFICATION ===


    } catch (e) {
      // Repli : /d/<id>/viewform
      redirectUrl = `https://docs.google.com/forms/d/${finalFormId}/viewform`;
    }

    // Option future si tu stockes l’URL publique en CONFIG :
    // if (config['Lien_Formulaire_Questions']) redirectUrl = String(config['Lien_Formulaire_Questions']);

    Logger.log(`Redirection de l'utilisateur vers le formulaire final : ${redirectUrl}`);
    try {
  let logsSheet = configSS.getSheetByName('Handler_Logs');
  if (!logsSheet) {
    logsSheet = configSS.insertSheet('Handler_Logs');
    logsSheet.appendRow(['Timestamp','RowIndex','Key','Value']);
  }
  logsSheet.appendRow([new Date(), rowIndex, 'redirectUrl', redirectUrl]);
    SpreadsheetApp.flush(); // (optionnel) force l’écriture

} catch (e) {
  Logger.log('Handler_Logs error: ' + e);
}

    return { redirectUrl: redirectUrl };

  } catch (error) {
    Logger.log("Erreur critique dans processFormSubmission: " + error.toString());
    throw new Error("Une erreur interne est survenue. Veuillez réessayer. Détails : " + error.message);
  }
}

/**
 * Fonction de compatibilité, non utilisée dans le flux principal.
 */
function handleFormSubmit(e) {
  try {
    const configSpreadsheet = SpreadsheetApp.openById(ID_FEUILLE_CONFIG);
    const ordersSheet = configSpreadsheet.getSheetByName("Orders");

    if (!ordersSheet) {
      Logger.log("ERREUR CRITIQUE: L'onglet 'Orders' est introuvable.");
      return;
    }

    const itemResponses = e.response.getItemResponses();
    let emailRepondant = '';
    let nomRepondant = '';

    for (let i = 0; i < itemResponses.length; i++) {
      const question = itemResponses[i].getItem().getTitle().toLowerCase();
      const reponse = itemResponses[i].getResponse();

      if (question.includes('mail')) {
        emailRepondant = reponse;
      } else if (question.includes('nom')) {
        nomRepondant = reponse;
      }
    }

    ordersSheet.appendRow([new Date(), emailRepondant, nomRepondant]);
    Logger.log('SUCCÈS (via handleFormSubmit) : Ligne ajoutée à l\'onglet Orders. Email: ' + emailRepondant + ', Nom: ' + nomRepondant);

  } catch (err) {
    Logger.log('ERREUR dans handleFormSubmit: ' + err.toString() + "\n" + err.stack);
  }
}

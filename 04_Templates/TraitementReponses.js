// =================================================================================
// == FICHIER : TraitementReponses.gs
// == VERSION : 19.4 (Nettoyage automatique des ID de pièces jointes)
// == RÔLE  : Gère la logique de traitement des réponses et aiguille vers le bon moteur.
// =================================================================================

/**
 * Nettoie une chaîne de caractères pour la rendre utilisable comme nom de variable/placeholder.
 */
function _nettoyerEnTete(enTete) {
  if (!enTete) return "";
  const accents = 'ÀÁÂÃÄÅàáâãäåÒÓÔÕÖØòóôõöøÈÉÊËèéêëÇçÌÍÎÏìíîïÙÚÛÜùúûüÿÑñ';
  const sansAccents = 'AAAAAAaaaaaaOOOOOOooooooEEEEeeeeCcIIIIiiiiUUUUuuuuyNn';
  return enTete.toString().split('').map((char) => {
    const accentIndex = accents.indexOf(char);
    return accentIndex !== -1 ? sansAccents[accentIndex] : char;
  }).join('')
  .replace(/[^a-zA-Z0-9_]/g, '_');
}

/**
 * Crée un objet de réponse standardisé à partir d'un numéro de ligne.
 */
function _creerObjetReponse(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const reponse = {};

  headers.forEach((header, i) => {
    let cle = header;
    if (header && !header.includes(':')) {
      cle = _nettoyerEnTete(header);
    }
    if (cle) {
      reponse[cle] = rowValues[i];
    }
  });
  return reponse;
}

/**
 * Point d'entrée principal pour traiter une soumission de formulaire.
 */
function onFormSubmit(e) {
  try {
    const rowIndex = e.range.getRow();
    traiterLigne(rowIndex, {});
  } catch (err) {
    Logger.log(`Erreur critique dans onFormSubmit pour la ligne ${e.range.getRow()}: ${err.toString()}\n${err.stack}`);
  }
}

/**
 * Envoie un e-mail de confirmation.
 */
function _envoyerEmailDeConfirmation(config, reponse, langueCible) {
    try {
        const nomColonneOverride = `ID_Gabarit_Email_Confirmation_${langueCible}`;
        let idGabaritConfirmation = config[nomColonneOverride];
        if (!idGabaritConfirmation || String(idGabaritConfirmation).trim() === '') {
            const systemIds = getSystemIds();
            const nomCleDefaut = `ID_GABARIT_CONFIRMATION_${langueCible}`;
            idGabaritConfirmation = systemIds[nomCleDefaut];
            Logger.log(`Utilisation du gabarit de confirmation PAR DÉFAUT pour la langue ${langueCible}.`);
        } else {
            Logger.log(`Utilisation du gabarit de confirmation SPÉCIFIQUE pour la langue ${langueCible}.`);
        }
        const emailRepondant = reponse.Votre_adresse_e_mail || reponse.Adresse_e_mail || reponse.emailRepondant;
        if (!idGabaritConfirmation || String(idGabaritConfirmation).trim() === '' || !emailRepondant) {
            Logger.log(`Aucun e-mail de confirmation à envoyer pour la langue ${langueCible}.`);
            return;
        }
        const doc = DocumentApp.openById(idGabaritConfirmation);
        let sujet = doc.getName();
        const url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + idGabaritConfirmation + "&exportFormat=html";
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
        let corpsHtml = response.getContentText();
        for (const key in reponse) {
            if (reponse.hasOwnProperty(key)) {
                const placeholder = `{{${key}}}`;
                const valeur = reponse[key] || '';
                const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
                sujet = sujet.replace(regex, valeur);
                corpsHtml = corpsHtml.replace(regex, valeur);
            }
        }
        const mailOptions = { to: emailRepondant, subject: sujet, htmlBody: corpsHtml };
        if (config.Email_Alias && config.Email_Alias.trim() !== '') {
            mailOptions.from = config.Email_Alias;
        }
        GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", mailOptions);
        Logger.log(`E-mail de confirmation [${langueCible}] envoyé avec succès à ${emailRepondant}.`);
    } catch (e) {
        Logger.log(`ERREUR lors de l'envoi de l'e-mail de confirmation : ${e.toString()}\n${e.stack}`);
    }
}

/**
 * COEUR LOGIQUE : Aiguille le traitement vers le moteur.
 */
function traiterLigne(rowIndex, optionsSurcharge = {}) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex);
    
    const langueOrigine = getOriginalLanguage(reponse);
    const langueCible = optionsSurcharge.langue || langueOrigine;
    
    if (!optionsSurcharge.isRetraitement) {
      _envoyerEmailDeConfirmation(config, reponse, langueCible);
    }
    
    const resultats = calculerResultats(reponse, langueCible, config, langueOrigine);

    if (optionsSurcharge.isRetraitement || config.Repondant_Quand === 'Immediat') {
      if (config.Moteur_Calcul === 'Universel') {
        Logger.log("Moteur UNIVERSEL détecté pour la ligne " + rowIndex + ". Envoi immédiat (retraitement ou configuration).");
        assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge);
      } else {
        // Moteur legacy...
      }
    } else {
      Logger.log(`Envoi différé détecté. Valeur : "${config.Repondant_Quand}". Lancement de la programmation.`);
      programmerEnvoiResultats(rowIndex, langueCible, config.Repondant_Quand);
    }

  } catch (err) {
    Logger.log("ERREUR FATALE dans traiterLigne: " + err.toString() + "\n" + err.stack);
  }
}

// =================================================================================
// == MOTEUR UNIVERSEL - FONCTION D'ENVOI D'EMAIL
// =================================================================================
/**
 * Lit la BDD, assemble et envoie l'e-mail de résultats.
 */
function assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge = {}) {
  const typeTest = config.Type_Test;
  let codeNiveauEmail = config.ID_Gabarit_Email_Repondant.replace('RESULTATS_', '');

  if (optionsSurcharge && optionsSurcharge.niveau && optionsSurcharge.niveau !== '') {
    codeNiveauEmail = optionsSurcharge.niveau;
  }

  const profilFinal = resultats.profilFinal;
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const compoSheet = bdd.getSheetByName("sys_Composition_Emails");
  const compoData = compoSheet.getDataRange().getValues();
  const compoHeaders = compoData.shift();
  const idx = {
    typeTest: compoHeaders.indexOf('Type_Test'),
    langue: compoHeaders.indexOf('Code_Langue'),
    niveau: compoHeaders.indexOf('Code_Niveau_Email'),
    profil: compoHeaders.indexOf('Code_Profil'),
    element: compoHeaders.indexOf('Element'),
    ordre: compoHeaders.indexOf('Ordre'),
    contenu: compoHeaders.indexOf('Contenu / ID_Document')
  };

  let briquesDeContenu = compoData.filter(row => {
    const typeLigne = (row[idx.typeTest] || '').trim();
    const typeCible = (typeTest || '').trim();
    const typeMatch = (typeLigne === typeCible || typeLigne === '');
    const langMatch = row[idx.langue] === langueCible;
    const levelMatch = row[idx.niveau].includes(codeNiveauEmail);
    const profilLigne = (row[idx.profil] || '').trim();
    const profilCible = (profilFinal || '').trim();
    const profileMatch = (profilLigne === profilCible || profilLigne === '');
    return typeMatch && langMatch && levelMatch && profileMatch;
  });

  briquesDeContenu.sort((a, b) => a[idx.ordre] - b[idx.ordre]);
  
  let contenuInfoCopie = null;
  const indexInfoCopie = briquesDeContenu.findIndex(brique => (brique[idx.element] || '').trim() === 'Info_Copie');
  if (indexInfoCopie > -1) {
    contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu];
    briquesDeContenu.splice(indexInfoCopie, 1);
  }

  let sujet = `Résultats de votre test ${typeTest}`;
  let corpsHtml = "";
  const piecesJointesIds = new Set();

  for (const brique of briquesDeContenu) {
    const elementType = (brique[idx.element] || '').trim();
    const contenu = brique[idx.contenu];
    switch (elementType) {
      case 'Sujet_Email': sujet = contenu; break;
      case 'Introduction': case 'Corps_Texte': corpsHtml += contenu + "<br>"; break;
      // ==================== DÉBUT DE LA MODIFICATION (ROBUSTESSE ID) ====================
      case 'Document':
        // On s'assure que l'ID est une chaîne de caractères et on nettoie les espaces/sauts de ligne
        if (contenu && String(contenu).trim()) {
          piecesJointesIds.add(String(contenu).trim());
        }
        break;
      // ===================== FIN DE LA MODIFICATION (ROBUSTESSE ID) =====================
      case 'Ligne_Score':
        Object.entries(resultats.scoresData).sort((a, b) => b[1] - a[1]).forEach(([code, score]) => {
          let ligneScore = contenu.replace(/{{nom_profil}}/g, resultats.mapCodeToName[code] || code).replace(/{{score}}/g, score);
          corpsHtml += ligneScore + "<br>";
        });
        break;
    }
  }

  const donneesPourEmail = { ...reponse, ...resultats };
  for (const key in donneesPourEmail) {
    if (donneesPourEmail.hasOwnProperty(key)) {
      const placeholder = `{{${key}}}`;
      const valeur = donneesPourEmail[key] || '';
      const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
      sujet = sujet.replace(regex, valeur);
      corpsHtml = corpsHtml.replace(regex, valeur);
      if (contenuInfoCopie) {
        contenuInfoCopie = contenuInfoCopie.replace(regex, valeur);
      }
    }
  }

  const piecesJointes = Array.from(piecesJointesIds).map(id => {
    let blob = null;
    for (let i = 0; i < 3; i++) { // Tentative de 3 essais
      try {
        blob = DriveApp.getFileById(id).getBlob();
        break; // Succès, on sort de la boucle
      } catch (e) {
        Logger.log(`Tentative ${i + 1} échouée pour la PJ : ${id}. Erreur: ${e.message}`);
        if (i < 2) { // Si ce n'est pas la dernière tentative
          Utilities.sleep(1000); // On attend 1 seconde avant de réessayer
        } else {
          Logger.log(`ERREUR FINALE PJ : Impossible d'attacher le fichier ${id} après 3 tentatives.`);
        }
      }
    }
    return blob;
  }).filter(Boolean);

  const T = loadTraductions(langueCible);
  const adressesUniques = new Set();
  const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Adresse_e_mail || reponse.emailRepondant;
  
  const destinatairesSurcharge = optionsSurcharge.destinataires || {};
  const useSurcharge = Object.keys(destinatairesSurcharge).length > 0;

  if (useSurcharge) {
    if (destinatairesSurcharge.repondant && emailRepondantPrincipal) { adressesUniques.add(emailRepondantPrincipal); }
    if (destinatairesSurcharge.formateur && destinatairesSurcharge.formateurEmail) { adressesUniques.add(destinatairesSurcharge.formateurEmail); }
    if (destinatairesSurcharge.patron && destinatairesSurcharge.patronEmail) { adressesUniques.add(destinatairesSurcharge.patronEmail); }
    if (destinatairesSurcharge.test && destinatairesSurcharge.test.trim() !== '') {
      destinatairesSurcharge.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email));
    }
  } else {
    if (config.Repondant_Email_Actif === 'Oui' && emailRepondantPrincipal) { adressesUniques.add(emailRepondantPrincipal); }
    if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) { adressesUniques.add(config.Patron_Email); }
    if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) { adressesUniques.add(config.Formateur_Email); }
  }
  
  if (config.Developpeur_Email) { adressesUniques.add(config.Developpeur_Email); }

  adressesUniques.forEach(adresse => {
    try {
      let sujetFinal = sujet;
      let corpsHtmlFinal = corpsHtml;
      if (adresse.toLowerCase() !== (emailRepondantPrincipal || "").toLowerCase()) {
        sujetFinal = (T.PREFIXE_COPIE_EMAIL || "Copie : ") + sujet;
        if (contenuInfoCopie) { corpsHtmlFinal = contenuInfoCopie + corpsHtml; }
      }
      const mailOptions = { to: adresse, subject: sujetFinal, htmlBody: corpsHtmlFinal, attachments: piecesJointes };
      const aliasExpediteur = optionsSurcharge.alias || config.Email_Alias;
      if (aliasExpediteur && aliasExpediteur.trim() !== '') {
        mailOptions.from = aliasExpediteur;
      }
      GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", mailOptions);
      Logger.log(`E-mail de RÉSULTATS [${langueCible}] envoyé avec succès à ${adresse}.`);
    } catch (e) {
      Logger.log(`Echec de l'envoi des résultats à ${adresse}. Erreur: ${e.message}`);
    }
  });
}

// =================================================================================
// == SECTION INTERFACE UTILISATEUR (UI)
// =================================================================================

function getDonneesPourRetraitement(rowIndex) {
  try {
    const config = getTestConfiguration();
    const reponse = _creerObjetReponse(rowIndex);

    const nomRepondant = reponse.Votre_nom_et_prenom || reponse.Nom_et_prenom || '';
    const emailRepondant = reponse.Votre_adresse_e_mail || reponse.Adresse_e_mail || '';
    
    return {
      nomRepondant: nomRepondant,
      emailRepondant: emailRepondant,
      langueOrigine: getOriginalLanguage(reponse),
      repondantActif: config.Repondant_Email_Actif === 'Oui',
      formateurActif: config.Formateur_Email_Actif === 'Oui',
      patronActif: config.Patron_Email_Mode === 'Oui',
      emailAlias: config.Email_Alias || ''
    };
  } catch (e) {
    Logger.log(`ERREUR dans getDonneesPourRetraitement(${rowIndex}): ${e.toString()}`);
    throw new Error("Impossible de récupérer les données pour la ligne " + rowIndex + ". " + e.message);
  }
}

/**
 * Lance le retraitement depuis l'interface utilisateur.
 */
function lancerRetraitementDepuisUI(options) {
  if (!options || !options.rowIndex) {
    throw new Error("Les options de retraitement sont invalides.");
  }

  options.isRetraitement = true; 
  
  traiterLigne(options.rowIndex, options);
  
  return "Retraitement pour la ligne " + options.rowIndex + " lancé avec succès !";
}
/**
 * T_Mail.gs
 * @fileoverview Gère la composition et l'envoi des e-mails de résultats.
 * @version 2.2 - Correction de la logique de sélection des destinataires et des variables de fusion lors du retraitement.
 */

function normalizeAndDedupeCompositionEmailsRows_(rows, idx) {
  const seen = new Set();
  if (!Array.isArray(rows)) {
    Logger.log("AVERTISSEMENT: normalizeAndDedupeCompositionEmailsRows_ a reçu une valeur non-array.");
    return [];
  }
  return rows.filter(r => {
    const key = [
      r[idx.typeTest] || '', r[idx.langue] || '', r[idx.niveau] || '',
      r[idx.profil] || '', r[idx.element] || '', r[idx.ordre] || ''
    ].join('|');
    if (seen.has(key)) { return false; }
    seen.add(key);
    return true;
  });
}

function _enrichirDonneesPourEmail_(reponse, resultats) {
  const nomPrenom = (reponse.Votre_nom_et_prenom || reponse.Nom_et_prenom || "Participant");
  const email = (reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || "");
  const base = {
    Nom_et_prenom: nomPrenom,
    Votre_nom_et_prenom: nomPrenom,
    Email_du_repondant: email,
    Votre_adresse_e_mail: email, // Correction : Ajout de la clé pour la fusion
    Date_du_jour: new Date().toLocaleDateString('fr-FR')
  };
  return { ...base, ...resultats };
}

function assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge = {}) {
  const typeTest = (config.Type_Test || '').toString().trim();
  let codeNiveauEmail = (config.ID_Gabarit_Email_Repondant || '').toString().replace('RESULTATS_', '').trim();
  if (optionsSurcharge && optionsSurcharge.niveau && optionsSurcharge.niveau !== '') codeNiveauEmail = optionsSurcharge.niveau;
  const profilFinal = (resultats.profilFinal || '').toString().trim();
  
  const systemIds = getSystemIds();
  const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
  const compoSheet = bdd.getSheetByName("sys_Composition_Emails");
  const compoData = compoSheet.getDataRange().getValues();
  const compoHeaders = compoData.shift().map(h => String(h || '').trim());
  const idx = { typeTest: compoHeaders.indexOf('Type_Test'), langue: compoHeaders.indexOf('Code_Langue'), niveau: compoHeaders.indexOf('Code_Niveau_Email'), profil: compoHeaders.indexOf('Code_Profil'), element: compoHeaders.indexOf('Element'), ordre: compoHeaders.indexOf('Ordre'), contenu: compoHeaders.indexOf('Contenu / ID_Document') };
  const compoRows = normalizeAndDedupeCompositionEmailsRows_(compoData, idx);
  
  let briquesDeContenu = compoRows.filter((row) => {
    const typeLigne = (row[idx.typeTest] || '').toString().trim();
    const langLigne = (row[idx.langue] || '').toString().trim();
    const levelValue = (row[idx.niveau] || '').toString();
    const profilLigne = (row[idx.profil] || '').toString().trim();
    const typeMatch = (typeLigne === typeTest || typeLigne === '');
    const langMatch = (langLigne === langueCible || langLigne === '');
    const levelList = levelValue.split(',').map(s => s.trim()).filter(Boolean);
    const levelMatch = levelList.length > 0 ? levelList.includes(codeNiveauEmail) : levelValue.includes(codeNiveauEmail);
    const profileMatch = (profilLigne === profilFinal || profilLigne === '');
    return typeMatch && langMatch && levelMatch && profileMatch;
  });

  if (briquesDeContenu.length === 0) { Logger.log("ERREUR ❌: Aucune brique de contenu n'a été trouvée pour construire l'e-mail."); }
  briquesDeContenu.sort((a, b) => (Number(a[idx.ordre]) || 0) - (Number(b[idx.ordre]) || 0));
  
  const donneesPourEmail = _enrichirDonneesPourEmail_(reponse, resultats);
  let sujet = `Résultats de votre test ${typeTest}`;
  let corpsHtml = "";
  let contenuInfoCopie = null;
  const piecesJointesIds = new Set();
  const indexInfoCopie = briquesDeContenu.findIndex(b => (b[idx.element] || '').toString().trim() === 'Info_Copie');
  if (indexInfoCopie > -1) {
    contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu];
    briquesDeContenu.splice(indexInfoCopie, 1);
  }

  for (const brique of briquesDeContenu) {
    const elementType = (brique[idx.element] || '').toString().trim();
    const contenu = brique[idx.contenu];
    switch (elementType) {
      case 'Sujet_Email': sujet = contenu; break;
      case 'Introduction': case 'Corps_Texte': corpsHtml += (contenu || "") + "<br>"; break;
      case 'Champ_Profil': if (contenu && donneesPourEmail[contenu]) { corpsHtml += donneesPourEmail[contenu] + "<br>"; } break;
      case 'Document': if (contenu && String(contenu).trim()) piecesJointesIds.add(String(contenu).trim()); break;
      case 'Ligne_Score':
        const scoresAAfficher = resultats.scoresData;
        if (scoresAAfficher) {
          Object.entries(scoresAAfficher).sort((a, b) => b[1] - a[1]).forEach(([code, score]) => {
            let scoreArrondi = (typeof score === 'number') ? score.toFixed(1) : score;
            const nomProfil = resultats.mapCodeToName[code] || code;
            const totalPossible = resultats.scoresMaxPossible ? (resultats.scoresMaxPossible[code] || '') : '';
            let ligneScore = (contenu || `- {{nom_profil}} : {{score}}`).replace(/{{nom_profil}}/g, nomProfil).replace(/{{score}}/g, scoreArrondi).replace(/{{total_possible}}/g, totalPossible);
            corpsHtml += ligneScore + "<br>";
          });
        }
        break;
    }
  }
  
  const variablesFusion = { ...donneesPourEmail, ...resultats };
  for (const key in variablesFusion) {
    const placeholder = `{{${key}}}`;
    const valeur = variablesFusion[key] || '';
    const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
    sujet = sujet.replace(regex, valeur);
    corpsHtml = corpsHtml.replace(regex, valeur);
    if (contenuInfoCopie) contenuInfoCopie = contenuInfoCopie.replace(regex, valeur);
  }

  const piecesJointes = [];
  if (resultats.Graphique_Radar_Blob) {
    piecesJointes.push(resultats.Graphique_Radar_Blob.setName('Profil_Resilience.png'));
  }

  for (const contenuDoc of Array.from(piecesJointesIds)) {
    let candidateId = contenuDoc;
    if (candidateId.startsWith("{{") && candidateId.endsWith("}}")) {
      const cle = candidateId.slice(2, -2);
      candidateId = variablesFusion[cle] || "";
    }
    if (/^[a-zA-Z0-9_-]{20,}$/.test(candidateId)) {
      try {
        const file = DriveApp.getFileById(candidateId);
        const mimeType = file.getMimeType();

        if (mimeType === MimeType.GOOGLE_DOCS) {
          Logger.log(`Détection d'un Google Doc (ID: ${candidateId}). Lancement de la génération PDF...`);
          const nomRapport = (resultats.Titre_Profil || resultats.profilFinal || "Rapport");
          const pdf = genererPdfDepuisModele(candidateId, variablesFusion, nomRapport);
          if (pdf) { piecesJointes.push(pdf); }
        } else {
          Logger.log(`Détection d'un fichier statique (ID: ${candidateId}, Type: ${mimeType}). Ajout direct.`);
          piecesJointes.push(file.getBlob());
        }
      } catch (e) {
        Logger.log(`Impossible de traiter la pièce jointe avec l'ID ${candidateId} : ${e.message}`);
      }
    } else {
      Logger.log("Ignoré (Document) : valeur non reconnue comme un ID de fichier valide : " + candidateId);
    }
  }
  
  const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email;
  const override = optionsSurcharge.overrideRecipients === true;
  const ignoreDev = optionsSurcharge.ignoreDeveloppeurEmail === true;
  const dryRun = optionsSurcharge.dryRun === true;
  const destS = optionsSurcharge.destinataires || {};
  const adressesUniques = new Set();
  
  if (override) {
    if (destS.repondant === true && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
    if (destS.formateur === true && destS.formateurEmail) adressesUniques.add(destS.formateurEmail);
    if (destS.patron === true && destS.patronEmail) adressesUniques.add(destS.patronEmail);
    if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); }
  } else {
    if (config.Repondant_Email_Actif === 'Oui' && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
    if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) adressesUniques.add(config.Patron_Email);
    if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) adressesUniques.add(config.Formateur_Email);
    if (config.Developpeur_Email && !ignoreDev) adressesUniques.add(config.Developpeur_Email);
  }

  if (dryRun) {
    Logger.log('— DRY-RUN — AUCUN EMAIL ENVOYÉ —');
    Logger.log('Destinataires simulés : ' + Array.from(adressesUniques).join(', '));
    Logger.log('Sujet (après remplacements) : ' + sujet);
    Logger.log('Corps (aperçu 400c) : ' + (corpsHtml || '').slice(0, 400));
    Logger.log('Pièces jointes (nb) : ' + piecesJointes.length);
    return;
  }
  
  adressesUniques.forEach(adresse => {
    try {
      let sujetFinal = sujet;
      let corpsHtmlFinal = corpsHtml;
      if (adresse.toLowerCase() !== (emailRepondantPrincipal || "").toLowerCase()) {
        sujetFinal = "Copie : " + sujet;
        if (contenuInfoCopie) corpsHtmlFinal = contenuInfoCopie + corpsHtml;
      }
      const mailOptions = {
        to: adresse,
        subject: sujetFinal,
        htmlBody: corpsHtmlFinal,
        attachments: piecesJointes,
        from: (optionsSurcharge.alias || config.Email_Alias || null)
      };
      
      GmailApp.sendEmail(mailOptions.to, mailOptions.subject, "", {
        htmlBody: mailOptions.htmlBody,
        attachments: mailOptions.attachments,
        from: mailOptions.from
      });
      Logger.log(`E-mail de RÉSULTATS envoyé à ${adresse}.`);
    } catch (e) {
      Logger.log(`Echec de l'envoi des résultats à ${adresse}. Erreur: ${e.message}`);
    }
  });
}
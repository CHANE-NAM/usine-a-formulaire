/**
 * @fileoverview Gère la composition et l'envoi des e-mails de résultats.
 * @version 1.1 - Corrigé
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
    Date_du_jour: new Date().toLocaleDateString('fr-FR')
  };
  return { ...base, ...resultats };
}

function _envoyerEmailDeConfirmation(config, reponse, langueCible) {
    // Placeholder
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
  const compoRows = normalizeAndDedupeCompositionEmailsRows_(compoData, idx);

  let briquesDeContenu = compoRows.filter((row) => {
    const typeLigne = (row[idx.typeTest] || '').toString().trim();
    const langLigne = (row[idx.langue] || '').toString().trim();
    const levelValue = (row[idx.niveau] || '').toString();
    const profilLigne = (row[idx.profil] || '').toString().trim();

    const typeMatch = (typeLigne === typeTest || typeLigne === '');
    const langMatch = (langLigne === langueCible || langLigne === '');
    const levelList = levelValue.split(',').map(s => s.trim()).filter(Boolean);
    
    // CORRECTION APPLIQUÉE ICI
    const levelMatch = levelList.length > 0 ? levelList.includes(codeNiveauEmail) : levelValue.includes(codeNiveauEmail);

    const profileMatch = (profilLigne === profilFinal || profilLigne === '');

    return typeMatch && langMatch && levelMatch && profileMatch;
  });

  if (briquesDeContenu.length === 0) {
    Logger.log("ERREUR ❌: Aucune brique de contenu n'a été trouvée pour construire l'e-mail.");
  }

  briquesDeContenu.sort((a, b) => (Number(a[idx.ordre]) || 0) - (Number(b[idx.ordre]) || 0));

  const donneesPourEmail = _enrichirDonneesPourEmail_(reponse, resultats);
  let contenuInfoCopie = null;
  const indexInfoCopie = briquesDeContenu.findIndex(b => (b[idx.element] || '').toString().trim() === 'Info_Copie');
  if (indexInfoCopie > -1) {
    contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu];
    briquesDeContenu.splice(indexInfoCopie, 1);
  }

  let sujet = `Résultats de votre test ${typeTest}`;
  let corpsHtml = "";
  const piecesJointesIds = new Set();

  for (const brique of briquesDeContenu) {
    const elementType = (brique[idx.element] || '').toString().trim();
    const contenu = brique[idx.contenu];
    switch (elementType) {
      case 'Sujet_Email':
        sujet = contenu;
        break;
      case 'Introduction':
      case 'Corps_Texte':
        corpsHtml += (contenu || "") + "<br>";
        break;
      case 'Champ_Profil':
        if (contenu && donneesPourEmail[contenu]) {
          corpsHtml += donneesPourEmail[contenu] + "<br>";
        }
        break;
      case 'Document':
        if (contenu && String(contenu).trim()) {
          piecesJointesIds.add(String(contenu).trim());
        }
        break;
      case 'Ligne_Score':
        const scoresAAfficher = resultats.scoresData;
        if (scoresAAfficher) {
          Object.entries(scoresAAfficher)
            .sort((a, b) => b[1] - a[1])
            .forEach(([code, score]) => {
              let scoreArrondi = (typeof score === 'number') ? score.toFixed(1) : score;
              const nomProfil = resultats.mapCodeToName[code] || code;
              let ligneScore = (contenu || `- {{nom_profil}} : {{score}}`)
                .replace(/{{nom_profil}}/g, nomProfil).replace(/{{score}}/g, scoreArrondi);
              corpsHtml += ligneScore + "<br>";
            });
        }
        break;
    }
  }

  for (const key in donneesPourEmail) {
    const placeholder = `{{${key}}}`;
    const valeur = donneesPourEmail[key] || '';
    const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
    sujet = sujet.replace(regex, valeur);
    corpsHtml = corpsHtml.replace(regex, valeur);
    if (contenuInfoCopie) contenuInfoCopie = contenuInfoCopie.replace(regex, valeur);
  }

  const variablesFusion = { ...donneesPourEmail, ...resultats };
  const piecesJointes = [];
  if (resultats.Graphique_Radar_Blob) {
    piecesJointes.push(resultats.Graphique_Radar_Blob.setName('Profil_Resilience.png'));
  }

  for (const contenuDoc of Array.from(piecesJointesIds)) {
    let candidate = contenuDoc;
    if (candidate.startsWith("{{") && candidate.endsWith("}}")) {
      const cle = candidate.slice(2, -2);
      candidate = variablesFusion[cle] || "";
    }
    if (/^[a-zA-Z0-9_-]{20,}$/.test(candidate)) {
      try {
        const nomRapport = (resultats.Titre_Profil || resultats.profilFinal || config.Type_Test || "Rapport");
        const pdf = genererPdfDepuisModele(candidate, variablesFusion, nomRapport);
        if (pdf) {
          piecesJointes.push(pdf);
        }
      } catch (e) {
        Logger.log("Fusion Doc->PDF échouée pour " + candidate + " : " + e.message);
        try {
          piecesJointes.push(DriveApp.getFileById(candidate).getBlob());
        } catch (_) {}
      }
    } else {
      Logger.log("Ignoré (Document) : valeur non reconnue " + candidate);
    }
  }

  const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email;
  const override = optionsSurcharge.overrideRecipients === true;
  const ignoreDev = optionsSurcharge.ignoreDeveloppeurEmail === true;
  const dryRun = optionsSurcharge.dryRun === true;
  const destS = optionsSurcharge.destinataires || {};
  const adressesUniques = new Set();

  if (override) {
    if (destS.repondant && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
    if (destS.formateur && destS.formateurEmail) adressesUniques.add(destS.formateurEmail);
    if (destS.patron && destS.patronEmail) adressesUniques.add(destS.patronEmail);
    if (destS.test && destS.test.trim() !== '') {
      destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email));
    }
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
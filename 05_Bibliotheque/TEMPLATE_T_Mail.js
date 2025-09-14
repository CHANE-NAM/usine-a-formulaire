/**
 * =================================================================================
 * == FICHIER : TEMPLATE_T_Mail.gs
 * == VERSION : 4.1 - Version corrigée et robustifiée
 * == RÔLE    : Gère la composition et l'envoi des e-mails de résultats.
 * =================================================================================
 */

// ======================= SECTION DE DÉBOGAGE (ESPIONS) =======================
const DEBUG_MODE_MAIL = true; // INTERRUPTEUR GÉNÉRAL : Mettre à false pour désactiver les espions de ce fichier.

/**
 * Fonction utilitaire pour l'affichage conditionnel des logs de débogage pour ce module.
 */
function _log_mail(flag, ...args) {
  if (DEBUG_MODE_MAIL && flag) {
    const message = args.map(arg => typeof arg === 'object' ? JSON.stringify(arg, null, 2) : arg).join(' ');
    Logger.log(`[ESPION Mail] ${message}`);
  }
}
// =================================================================================


/**
 * Crée un objet unique contenant toutes les données nécessaires à la fusion,
 * en combinant les informations de base du répondant avec tous les résultats calculés.
 * @param {Object} reponse - L'objet contenant les réponses du formulaire.
 * @param {Object} resultats - L'objet contenant les résultats des calculs du moteur.
 * @returns {Object} Un objet complet prêt pour la fusion.
 */
function _enrichirDonneesPourEmail_(reponse, resultats) {
  const nomPrenom = (reponse.Votre_nom_et_prenom || reponse.Nom_et_prenom || "Participant");
  const email = (reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email || "");
  const base = {
    Nom_et_prenom: nomPrenom,
    Votre_nom_et_prenom: nomPrenom,
    Email_du_repondant: email,
    Votre_adresse_e_mail: email,
    Date_du_jour: new Date().toLocaleDateString('fr-FR')
  };
  // Fusionne les données de base avec l'intégralité de l'objet de résultats
  return { ...base, ...resultats };
}


/**
 * Assemble et envoie un e-mail universel basé sur les briques de contenu de la BDD.
 * @param {Object} config - La configuration du test.
 * @param {Object} reponse - L'objet de réponse de la ligne traitée.
 * @param {Object} resultats - Les résultats calculés par le moteur de test.
 * @param {string} langueCible - Le code de la langue pour l'e-mail ('FR', 'EN'...).
 * @param {Object} optionsSurcharge - Options pour le retraitement (destinataires, dryRun...).
 * @param {Spreadsheet} kitSpreadsheet - L'objet Spreadsheet du kit actif.
 */
function assemblerEtEnvoyerEmailUniversel(config, reponse, resultats, langueCible, optionsSurcharge = {}, kitSpreadsheet) {
  _log_mail(true, '--- DÉBUT ASSEMBLAGE EMAIL (VERSION CORRIGÉE) ---');
  try {
    const typeTest = (config.Type_Test || '').toString().trim();
    let codeNiveauEmail = (config.ID_Gabarit_Email_Repondant || '').toString().replace('RESULTATS_', '').trim();
    if (optionsSurcharge && optionsSurcharge.niveau && optionsSurcharge.niveau !== '') {
      codeNiveauEmail = optionsSurcharge.niveau;
    }
    const profilFinal = (resultats.profilFinal || '').toString().trim();

    // 1. CRÉATION DE L'OBJET DE FUSION CENTRALISÉ
    // C'est l'étape clé : toutes les variables sont réunies ici.
    const variablesFusion = _enrichirDonneesPourEmail_(reponse, resultats);
    _log_mail(true, 'Variables de fusion disponibles :', Object.keys(variablesFusion));

    // 2. RÉCUPÉRATION DES BRIQUES DE CONTENU DE L'EMAIL
    _log_mail(true, `Critères de recherche: Type_Test="${typeTest}", Langue="${langueCible}", Niveau="${codeNiveauEmail}", Profil="${profilFinal}"`);
    const systemIds = getSystemIds();
    const bdd = SpreadsheetApp.openById(systemIds.ID_BDD);
    const compoSheet = bdd.getSheetByName("sys_Composition_Emails");
    const compoData = compoSheet.getDataRange().getValues();
    const compoHeaders = compoData.shift().map(h => String(h || '').trim());
    const idx = { typeTest: compoHeaders.indexOf('Type_Test'), langue: compoHeaders.indexOf('Code_Langue'), niveau: compoHeaders.indexOf('Code_Niveau_Email'), profil: compoHeaders.indexOf('Code_Profil'), element: compoHeaders.indexOf('Element'), ordre: compoHeaders.indexOf('Ordre'), contenu: compoHeaders.indexOf('Contenu / ID_Document') };
    
    let briquesDeContenu = compoData.filter((row) => {
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
    
    _log_mail(true, `${briquesDeContenu.length} briques de contenu trouvées après filtrage.`);
    if (briquesDeContenu.length === 0) {
        _log_mail(true, 'ERREUR : Aucune brique de contenu trouvée. L\'e-mail ne sera pas construit.');
        return; 
    }
    briquesDeContenu.sort((a, b) => (Number(a[idx.ordre]) || 0) - (Number(b[idx.ordre]) || 0));

    // 3. CONSTRUCTION DU SUJET ET DU CORPS DE L'EMAIL
    let sujet = `Résultats de votre test ${typeTest}`;
    let corpsHtml = "";
    let contenuInfoCopie = null;
    const piecesJointesIds = new Set();
    
    const indexInfoCopie = briquesDeContenu.findIndex(b => (b[idx.element] || '').toString().trim() === 'Info_Copie');
    if (indexInfoCopie > -1) {
        contenuInfoCopie = briquesDeContenu[indexInfoCopie][idx.contenu];
        briquesDeContenu.splice(indexInfoCopie, 1);
    }

    briquesDeContenu.forEach(brique => {
        const elementType = (brique[idx.element] || '').toString().trim();
        const contenu = brique[idx.contenu];
        switch (elementType) {
            case 'Sujet_Email': sujet = contenu; break;
            case 'Introduction': case 'Corps_Texte': corpsHtml += (contenu || "") + "<br>"; break;
            case 'Champ_Profil': if (contenu && variablesFusion[contenu]) { corpsHtml += variablesFusion[contenu] + "<br>"; } break;
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
    });
    
    // 4. REMPLACEMENT DES PLACEHOLDERS EN UTILISANT L'OBJET DE FUSION COMPLET
    for (const key in variablesFusion) {
        const placeholder = `{{${key}}}`;
        const valeur = variablesFusion[key] || '';
        const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        sujet = sujet.replace(regex, valeur);
        corpsHtml = corpsHtml.replace(regex, valeur);
        if (contenuInfoCopie) contenuInfoCopie = contenuInfoCopie.replace(regex, valeur);
    }
    _log_mail(true, 'Sujet final (après fusion) :', sujet);

    // 5. PRÉPARATION DES PIÈCES JOINTES (PDF inclus)
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
                    _log_mail(true, `Détection d'un Google Doc (ID: ${candidateId}). Lancement de la génération PDF...`);
                    const nomRapport = (resultats.Titre_Profil || resultats.profilFinal || "Rapport");
                    const pdf = genererPdfDepuisModele(candidateId, variablesFusion, nomRapport);
                    if (pdf) { 
                        piecesJointes.push(pdf);
                        _log_mail(true, `PDF "${nomRapport}.pdf" généré avec succès.`);
                    }
                } else {
                    _log_mail(true, `Détection d'un fichier statique (ID: ${candidateId}, Type: ${mimeType}). Ajout direct.`);
                    piecesJointes.push(file.getBlob());
                }
            } catch (e) {
                _log_mail(true, `Impossible de traiter la pièce jointe avec l'ID ${candidateId} : ${e.message}`);
            }
        }
    }
    
    // 6. DÉTERMINATION DES DESTINATAIRES ET ENVOI
    const emailRepondantPrincipal = reponse.Votre_adresse_e_mail || reponse.Votre_adresse_email;
    const override = optionsSurcharge.overrideRecipients === true;
    const ignoreDev = optionsSurcharge.ignoreDeveloppeurEmail === true;
    const dryRun = optionsSurcharge.dryRun === true;
    const destS = optionsSurcharge.destinataires || {};
    const adressesUniques = new Set();
    
    if (override) {
        if (destS.repondant === true && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
        if (destS.formateur === true && config.Formateur_Email) adressesUniques.add(config.Formateur_Email);
        if (destS.patron === true && config.Patron_Email) adressesUniques.add(config.Patron_Email);
        if (destS.test && destS.test.trim() !== '') { destS.test.split(',').map(e => e.trim()).forEach(email => adressesUniques.add(email)); }
    } else {
        if (config.Repondant_Email_Actif === 'Oui' && emailRepondantPrincipal) adressesUniques.add(emailRepondantPrincipal);
        if (config.Patron_Email_Mode === 'Oui' && config.Patron_Email) adressesUniques.add(config.Patron_Email);
        if (config.Formateur_Email_Actif === 'Oui' && config.Formateur_Email) adressesUniques.add(config.Formateur_Email);
        if (config.Developpeur_Email && !ignoreDev) adressesUniques.add(config.Developpeur_Email);
    }

    _log_mail(true, 'Adresses de destination finales :', Array.from(adressesUniques));
    if (dryRun) {
        _log_mail(true, '— DRY-RUN — AUCUN EMAIL ENVOYÉ —');
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
            _log_mail(true, `E-mail de RÉSULTATS envoyé à ${adresse}.`);
        } catch (e) {
            _log_mail(true, `Echec de l'envoi des résultats à ${adresse}. Erreur: ${e.message}`);
        }
    });

  } catch(e) {
      Logger.log("ERREUR FATALE dans assemblerEtEnvoyerEmailUniversel: " + e.message + "\n" + e.stack);
  } finally {
      _log_mail(true, '--- FIN ASSEMBLAGE EMAIL ---');
  }
}
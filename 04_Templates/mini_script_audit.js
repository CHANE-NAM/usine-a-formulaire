/*******************************
 * AUDIT RAPIDE DU KIT (V1)
 * À coller dans le projet du KIT (Sheet de réponses)
 * Utilisation:
 *   1) Mets l'ID d'un modèle Google Doc accessible ci-dessous (DOC_MODELE_ID)
 *   2) Exécute auditKit() (menu Exécuter)
 *******************************/
function auditKit() {
  const DOC_MODELE_ID = '1F-vPh9xhtWlF2eAHEfzwgwo3cmGbIyJXrMgmCePaDKQ'; // <-- remplace si besoin
  const out = {
    horodatage: new Date().toISOString(),
    kitSpreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    effectiveUser: null,
    checks: {
      sheets: null,
      driveWrite: null,
      docsOpenModel: null,
      docsCopyOpenExport: null,
      gmailAliases: null,
      formLink: null,
      triggerInstallable: null
    },
    notes: [],
    errors: []
  };

  // Qui exécute ?
  try {
    out.effectiveUser = Session.getEffectiveUser().getEmail();
  } catch (e) {
    out.errors.push('Session.getEffectiveUser: ' + e.message);
  }

  // 1) Sheets
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    out.checks.sheets = { ok: true, name: ss.getName() };
  } catch (e) {
    out.checks.sheets = { ok: false, error: e.message };
    out.errors.push('Sheets: ' + e.message);
  }

  // 2) Drive (écriture)
  let tmpId = null;
  try {
    tmpId = DriveApp.createFile(Utilities.newBlob('ok', 'text/plain', 'audit_temp.txt')).getId();
    out.checks.driveWrite = { ok: true, createdFileId: tmpId };
  } catch (e) {
    out.checks.driveWrite = { ok: false, error: e.message };
    out.errors.push('Drive write: ' + e.message);
  } finally {
    if (tmpId) { try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (_) {} }
  }

  // 3) Docs (ouvrir le modèle)
  try {
    const doc = DocumentApp.openById(DOC_MODELE_ID);
    out.checks.docsOpenModel = { ok: true, title: doc.getName() };
  } catch (e) {
    out.checks.docsOpenModel = { ok: false, error: e.message };
    out.errors.push('Docs open model: ' + e.message);
  }

  // 4) Docs + Drive (copie → open → export PDF → poubelle)
  let copyId = null;
  try {
    const src = DriveApp.getFileById(DOC_MODELE_ID);
    const parent = src.getParents().hasNext() ? src.getParents().next() : DriveApp.getRootFolder();
    const name = 'AUDIT_COPY_' + new Date().toISOString().slice(0,19).replace(/[:T]/g,'-');
    copyId = src.makeCopy(name, parent).getId();

    const newDoc = DocumentApp.openById(copyId);
    newDoc.getBody().replaceText('{{AUDIT_PLACEHOLDER}}', new Date().toLocaleString('fr-FR'));
    newDoc.saveAndClose();

    const pdf = DriveApp.getFileById(copyId).getBlob().getAs('application/pdf');
    out.checks.docsCopyOpenExport = { ok: true, pdfBytes: pdf.getBytes().length };
  } catch (e) {
    out.checks.docsCopyOpenExport = { ok: false, error: e.message };
    out.errors.push('Docs copy/open/export: ' + e.message);
  } finally {
    if (copyId) { try { DriveApp.getFileById(copyId).setTrashed(true); } catch (_) {} }
  }

  // 5) Gmail (alias)
  try {
    const aliases = GmailApp.getAliases();
    out.checks.gmailAliases = { ok: true, aliases: aliases };
    if (!aliases || aliases.length === 0) out.notes.push('Aucun alias Gmail configuré pour ce compte.');
  } catch (e) {
    out.checks.gmailAliases = { ok: false, error: e.message };
    out.errors.push('Gmail aliases: ' + e.message);
  }

  // 6) Lien Form (si associé)
  try {
    const url = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
    out.checks.formLink = { ok: !!url, formUrl: url || null };
    if (!url) out.notes.push('Aucun Form lié détecté (getFormUrl() renvoie null).');
  } catch (e) {
    out.checks.formLink = { ok: false, error: e.message };
    out.errors.push('Form link: ' + e.message);
  }

  // 7) Déclencheur installable (handleFormSubmit)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const list = ScriptApp.getUserTriggers(ss) || [];
    const hasInstallable = list.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'handleFormSubmit');
    out.checks.triggerInstallable = { ok: hasInstallable, found: list.map(t => t.getHandlerFunction && t.getHandlerFunction()) };
    if (!hasInstallable) out.notes.push('Aucun déclencheur installable "handleFormSubmit" détecté—relance le menu "Activer le traitement automatique".');
  } catch (e) {
    out.checks.triggerInstallable = { ok: false, error: e.message };
    out.errors.push('Triggers: ' + e.message);
  }

  // Affichage
  const summary =
    '=== AUDIT KIT ===\n' +
    'Effective user: ' + (out.effectiveUser || 'inconnu') + '\n' +
    'Sheets: ' + JSON.stringify(out.checks.sheets) + '\n' +
    'Drive write: ' + JSON.stringify(out.checks.driveWrite) + '\n' +
    'Docs open model: ' + JSON.stringify(out.checks.docsOpenModel) + '\n' +
    'Docs copy/open/export: ' + JSON.stringify(out.checks.docsCopyOpenExport) + '\n' +
    'Gmail aliases: ' + JSON.stringify(out.checks.gmailAliases) + '\n' +
    'Form link: ' + JSON.stringify(out.checks.formLink) + '\n' +
    'Trigger installable: ' + JSON.stringify(out.checks.triggerInstallable) + '\n' +
    (out.notes.length ? ('Notes: ' + out.notes.join(' | ') + '\n') : '') +
    (out.errors.length ? ('Errors: ' + out.errors.join(' | ') + '\n') : '');

  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert('Audit terminé.\nConsulte le journal pour le détail.'); } catch (_) {}

  return out; // pratique si tu veux consommer le JSON ailleurs
}

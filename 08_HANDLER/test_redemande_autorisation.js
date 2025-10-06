

function warmUpScopes() {
  // Spreadsheet scope
  const ss = SpreadsheetApp.openById("1kLBqIHZWbHrb4SsoSQcyVsLOmqKHkhSA4FttM5hZtDQ");
  ss.getSheets()[0].getRange(1,1).getValue();

  // Form scope (mettre un ID de formulaire valide pour tester)
  try {
    const TEST_FORM_ID = '13X7ZbByK_XvxaKvDVn6XZYOS6AjNJ9RsEVyC_L2yMrE'; // <- ton ID d'édition
    FormApp.openById(TEST_FORM_ID).getTitle();
  } catch (e) {}
}

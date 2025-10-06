function testFormsApiCreate() {
  const f = Forms.Forms.create({ info: { title: "Test API Forms (auto)" } });
  Logger.log("formId = " + f.formId);
}

function testFormsApiREST() {
  const form = {
    info: { title: "Formulaire via Google Forms API" }
  };
  const url = "https://forms.googleapis.com/v1/forms";
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(form),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

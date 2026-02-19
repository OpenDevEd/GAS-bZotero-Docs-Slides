function showAccessDeniedWindow() {
  const template = HtmlService.createTemplateFromFile('API Forest access denied html.html');

  const audit = googleDocAudit();
  template.auditData = JSON.stringify(audit);

  const htmlOutput = template.evaluate()
    .setWidth(500)
    .setHeight(450);

  const ui = getUi();
  ui.showModalDialog(htmlOutput, 'Forest API access denied');
}
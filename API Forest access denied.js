function showAccessDeniedWindow() {
  const template = HtmlService.createTemplateFromFile('API Forest access denied html.html');

  const htmlOutput = template.evaluate()
    .setWidth(500)
    .setHeight(380);

  const ui = getUi();
  ui.showModalDialog(htmlOutput, 'Forest API access denied');
}
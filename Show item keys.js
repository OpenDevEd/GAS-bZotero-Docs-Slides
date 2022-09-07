function showItemKeys() {
  const ui = getUi();
  let bibReferences, text = '', itemKeys;
  let result = validateLinks(false, false, true);
  if (result.status == 'ok') {
    if (result.bibReferences.length > 0) {
      bibReferences = result.bibReferences;

      for (let i in bibReferences) {
        itemKeys = bibReferences[i].split(':');
        text += 'zotero://select/groups/' + itemKeys[0] + '/items/' + itemKeys[1] + '<br>';
      }
      let html = `<!DOCTYPE html>
                  <html>
                    <head>
                      <base target="_top">
                    </head>
                    <body>
                      ${text}
                    </body>
                  </html>`;
      html = HtmlService.createHtmlOutput(html).setWidth(800).setHeight(800);
      ui.showModalDialog(html, 'Item keys');
    } else {
      ui.alert('Links for bibliography weren\'t found.');
    }
  } else {
    // Error in validateLinks!
    // We don't show alert here because validateLinks has own alerts about errors.
    return 0;
  }
}

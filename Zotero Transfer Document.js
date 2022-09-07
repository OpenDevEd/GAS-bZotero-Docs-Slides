function zoteroTransferDoc() {
  const ui = getUi();
  try {
    let zoteroCitation, start, end, zoteroString, zoteroLongString;
    const counter = { value: 0 };
    const clsCitationRegex = 'ITEM CSL_CITATION.*?}}}\\]}';

    if (HOST_APP == 'docs') {
      // Doc part
      const doc = DocumentApp.getActiveDocument();
      const body = doc.getBody();
      zoteroCitation = body.findText(clsCitationRegex);
      while (zoteroCitation) {
        start = zoteroCitation.getStartOffset();
        end = zoteroCitation.getEndOffsetInclusive();

        zoteroText = zoteroCitation.getElement().asText().getText();

        zoteroLongString = zoteroText.substring(start, end + 1);
        zoteroString = zoteroText.substring(start + 18, end + 1);

        let { linkText, zoteroLink, startLink, endLink } = zoteroTransferDocHelper(zoteroString);

        // change this to text with url:
        newZoteroText = zoteroText.replace(zoteroLongString, linkText);
        zoteroCitation.getElement().asText().setText(newZoteroText).setLinkUrl(startLink, endLink, zoteroLink);

        zoteroCitation = body.findText(clsCitationRegex);
        counter.value++;
      }
      // End. Doc part
    } else {
      // Slides part
      const slides = SlidesApp.getActivePresentation().getSlides();
      for (let i in slides) {
        slides[i].getPageElements().forEach(function (pageElement) {
          if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
            rangeElementStart = pageElement.asShape().getText().find(clsCitationRegex);
            zoteroTransferSlidesHelper(rangeElementStart, zoteroString, counter);
          } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {
            const table = pageElement.asTable();
            numRows = table.getNumRows();
            numCols = table.getNumColumns();
            for (let m = 0; m < numRows; m++) {
              for (let k = 0; k < numCols; k++) {
                cell = table.getRow(m).getCell(k);
                if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
                  rangeElementStart = cell.getText().find(clsCitationRegex);
                  zoteroTransferSlidesHelper(rangeElementStart, zoteroString, counter);
                }
              }
            }
          }
        });
      }
      //End. Slides part
    }

    ui.alert(' Number of Zotero links: ' + counter.value);
  }
  catch (error) {
    ui.alert('Error in zoteroTransferDoc. ' + error);
  }
}

function zoteroTransferSlidesHelper(rangeElementStart, zoteroString, counter) {
  for (let j in rangeElementStart) {
    zoteroText = rangeElementStart[j].asString();
    //Logger.log(zoteroText);
    zoteroString = zoteroText.substring(18, zoteroText.length);
    let { linkText, zoteroLink, startLink, endLink } = zoteroTransferDocHelper(zoteroString);

    rangeElementStart[j].replaceAllText(zoteroText, linkText);
    rangeElementStart[j].getTextStyle().setLinkUrl(zoteroLink).setUnderline(false);
    counter.value++;
  }
}

function zoteroTransferDocHelper(zoteroString) {
  let zoteroJson, linkText, groupItem, zoteroLink, startLink, endLink, url;
  const groupItemRegex = /groups\/([^\:\|]*)\/items\/([^\:\|]*)$/;
  zoteroJson = JSON.parse(zoteroString);

  linkText = zoteroJson.properties.plainCitation;
  url = String(zoteroJson.citationItems[0].uris);

  groupItem = url.match(groupItemRegex);
  groupId = groupItem[1];
  itemKey = groupItem[2];

  zoteroLink = 'https://ref.opendeved.net/zo/zg/' + groupId + '/7/' + itemKey + '/';
  zoteroLink = replaceAddParameter(zoteroLink, 'openin', 'zoteroapp');

  if (/^\(.*\)$/.test(linkText)) {
    linkText = linkText.substr(0, 1) + "⇡" + linkText.substr(1);
    startLink = 1;
    endLink = linkText.length - 2;
  } else {
    linkText = "⇡" + linkText;
    startLink = 0;
    endLink = linkText.length - 1;
  }
  return { 'linkText': linkText, 'zoteroLink': zoteroLink, 'startLink': startLink, 'endLink': endLink };
}
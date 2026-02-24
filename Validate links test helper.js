function validateLinksTestHelper() {
  const ui = getUi();
  const linksArray = [];

  if (HOST_APP == 'docs') {
    // Doc part
    const doc = DocumentApp.getActiveDocument();

    const element = doc.getBody();
    testingFindAllLinks(element, 'body', linksArray);

    const footnotes = doc.getFootnotes();
    let footnote, numChildren;
    for (let i in footnotes) {
      footnote = footnotes[i].getFootnoteContents();
      if (footnote == null) {
        alertSuggestedFootnoteBug(i);
        continue;
      }
      numChildren = footnote.getNumChildren();
      for (let j = 0; j < numChildren; j++) {
        testingFindAllLinks(footnote.getChild(j), 'footnotes', linksArray);
      }
    }
    // End. Doc part
  } else {
    // Slides part
    const slides = SlidesApp.getActivePresentation().getSlides();
    let url, linkText, runs;
    for (let i in slides) {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
          runs = pageElement.asShape().getText().getRuns();
          for (let j in runs) {
            // Logger.log(runs[j].getStartIndex());
            // Logger.log(runs[j].asString());
            // Logger.log(runs[j].getLinks());

            links = runs[j].getLinks();
            for (let m in links) {
              //linksArray.push(links[j].getTextStyle().getLink().getUrl());
              url = links[m].getTextStyle().getLink().getUrl();
              // linkText = links[m].getTextStyle().getLink();
              linkText = runs[j].asString();
              linksArray.push({ link: url, linkText: linkText });
            }
          }
        }
      });
    }
    //End. Slides part
  }

  //Logger.log(linksArray);
  let allLinks = '';
  for (let i in linksArray) {
    allLinks += `
  <br>
      ${escapeHtml(linksArray[i].linkText)}
  <br>
      <a target="_blank" href="${escapeHtml(linksArray[i].link)}">${escapeHtml(linksArray[i].link)}</a>
  <br>
  `;
  }

  let html = `<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
  ${allLinks}
  </body>
</html>`;
  html = HtmlService.createHtmlOutput(html).setWidth(800).setHeight(800);
  ui.showModalDialog(html, 'Links');
  addUsageTrackingRecord('validateLinksTestHelper');
}

function testingFindAllLinks(element, source, linksArray) {

  let text, end, indices, partAttributes, numChildren;
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    //Logger.log(indices);
    for (let i = 0; i < indices.length; i++) {
      partAttributes = element.getAttributes(indices[i]);
      //Logger.log(partAttributes);
      if (partAttributes.LINK_URL) {

        text = element.getText();
        if (i == indices.length - 1) {
          end = text.length - 1;
        } else {
          end = indices[i + 1] - 1;
        }

        linksArray.push({ link: partAttributes.LINK_URL, linkText: text.substr(indices[i], end - indices[i] + 1), source: source });
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        testingFindAllLinks(element.getChild(i), source, linksArray);
      }
    }
  }
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
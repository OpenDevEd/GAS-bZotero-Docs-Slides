function unpackZoteroSlides(warning) {
  const ui = SlidesApp.getUi();

  try {
    let links, link, groupId, itemKey, linkText, regEx, advice = '', showAdviseFlag = { value: false };

    let counter = { value: 0 };
    if (warning === true) {
      regEx = new RegExp('https?://ref.opendeved.net/zo/zg/([0-9]+)/7/([^/\?]+)/([^/\?]+)', 'i');
    } else {
      regEx = new RegExp('https?://ref.opendeved.net/zo/zg/([0-9]+)/7/([^/\?]+)/?', 'i');
    }

    const slides = SlidesApp.getActivePresentation().getSlides();
    for (let i in slides) {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
          links = pageElement.asShape().getText().getLinks();
          Logger.log(links);
          unpackZoteroSlidesHelper(warning, links, showAdviseFlag, regEx, counter);
        } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {
          const table = pageElement.asTable();
          numRows = table.getNumRows();
          numCols = table.getNumColumns();
          for (let m = 0; m < numRows; m++) {
            for (let k = 0; k < numCols; k++) {
              cell = table.getRow(m).getCell(k);
              if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
                links = cell.getText().getLinks();
                unpackZoteroSlidesHelper(warning, links, showAdviseFlag, regEx, counter);
              }
            }
          }
        }
      });
    }

    if (showAdviseFlag === true) {
      advice = ' Following unpacking, check the document for symbols "+UL+", "+UR+", ⇡ and 《warning》 and then run cleanup.';
    }
    ui.alert('Number of citations that were Zotero-unpacked: ' + counter.value + advice);
  }
  catch (error) {
    ui.alert('Error in unpackZoteroSlides: ' + error);
  }
}

function unpackZoteroSlidesHelper(warning, links, showAdviseFlag, regEx, counter) {
  let link, groupId, itemKey, linkText;
  //for (let j in links) {
  for (let j = links.length - 1; j >= 0; j--) {
    warningFlag = false;
    link = links[j].getTextStyle().getLink().getUrl();
    result = regEx.exec(link);
    if (result != null) {
      itemKey = result[2];
      groupId = result[1];
      linkTextHelper = result[3];
      linkText = links[j].asString().replace(/^⇡/, '');
      if (warning === true && linkTextHelper != linkText) {
        warningText = `《warning:${linkTextHelper}》`;
        warningTextLength = warningText.length;
        newText = `⟦zg:${groupId}:${itemKey}|${linkText}${warningText}⟧`;
        warningFlag = true;
        showAdviseFlag.value = true;
      } else {
        newText = `⟦zg:${groupId}:${itemKey}|${linkText}⟧`;
      }
      pinkLength = groupId.length + itemKey.length + 6;
      Logger.log(newText);
      links[j].setText(newText).getTextStyle().setLinkUrl(null).setBackgroundColorTransparent();
      links[j].getRange(0, pinkLength).getTextStyle().setFontSize(6).setForegroundColor('#fe01dc');
      links[j].getRange(newText.length - 1, newText.length).getTextStyle().setFontSize(6).setForegroundColor('#fe01dc');
      links[j].getRange(pinkLength, newText.length - 1).getTextStyle().setForegroundColor('#0123dd');
      if (warningFlag === true) {
        links[j].getRange(newText.length - warningTextLength - 1, newText.length - 1).getTextStyle().setFontSize(6).setForegroundColor('#fe01dc').setBackgroundColor('#dedede');
      }
      counter.value++;
    }
  }
}

function packZoteroSlides(toDo) {
  const ui = SlidesApp.getUi();
  try {
    const zoteroLinkRegEx = '⟦.*?⟧';
    const counter = { value: 0 };

    const slides = SlidesApp.getActivePresentation().getSlides();
    for (let i in slides) {
      slides[i].getPageElements().forEach(pageElement => {
        pageElementHelper(toDo, pageElement, zoteroLinkRegEx, counter);
      });
    }

    if (toDo == 'packZotero') {
      ui.alert('Number of citations that were Zotero-packed: ' + counter.value);
    }
  }
  catch (error) {
    ui.alert('Error in packZoteroSlides ' + error);
  }
}

function packZoteroSelectiveCallSlides() {
  const ui = SlidesApp.getUi();
  try {
    let ranges;
    const toDo = 'packZotero';
    const zoteroLinkRegEx = '⟦.*?⟧';
    const counter = { value: 0 };
    const selection = SlidesApp.getActivePresentation().getSelection();
    const selectionType = selection.getSelectionType();
    Logger.log(selectionType);

    if (selectionType == 'PAGE') {

      const pageRange = selection.getPageRange();
      pageRange.getPages().forEach(page => {
        page.getPageElements().forEach(pageElement => {
          pageElementHelper(toDo, pageElement, zoteroLinkRegEx, counter);
        });
      });

    } else if (selectionType == 'CURRENT_PAGE') {

      var currentPage = selection.getCurrentPage();
      currentPage.getPageElements().forEach(pageElement => {
        pageElementHelper(toDo, pageElement, zoteroLinkRegEx, counter);
      });

    } else if (selectionType == 'PAGE_ELEMENT') {

      selection.getPageElementRange().getPageElements().forEach(pageElement => {
        pageElementHelper(toDo, pageElement, zoteroLinkRegEx, counter);
      });

    } else if (selectionType == 'TEXT') {

      ranges = selection.getTextRange().find(zoteroLinkRegEx);
      zoteroToLinkSlides(toDo, ranges, counter);

    } else if (selectionType == 'TABLE_CELL') {

      const tableCells = selection.getTableCellRange().getTableCells();
      tableCells.forEach(cell => {
        if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
          ranges = cell.getText().find(zoteroLinkRegEx);
          zoteroToLinkSlides(toDo, ranges, counter);
        }
      });

    }

    ui.alert('Number of citations that were Zotero-unpacked: ' + counter.value);
  }
  catch (error) {
    ui.alert('Error in packZoteroSelectiveCallSlides ' + error);
  }
}

function pageElementHelper(toDo, pageElement, zoteroLinkRegEx, counter) {
  let cell, numRows, numCols, ranges;
  const pageElementType = pageElement.getPageElementType();
  if (pageElementType == SlidesApp.PageElementType.SHAPE) {
    ranges = pageElement.asShape().getText().find(zoteroLinkRegEx);
    zoteroToLinkSlides(toDo, ranges, counter);
  } else if (pageElementType == SlidesApp.PageElementType.TABLE) {
    const table = pageElement.asTable();
    numRows = table.getNumRows();
    numCols = table.getNumColumns();
    for (let m = 0; m < numRows; m++) {
      for (let k = 0; k < numCols; k++) {
        cell = table.getRow(m).getCell(k);
        if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
          ranges = cell.getText().find(zoteroLinkRegEx);
          zoteroToLinkSlides(toDo, ranges, counter);
        }
      }
    }
  }
}

function zoteroToLinkSlides(toDo, ranges, counter) {
  let text, found, newUrl, linkText = '';
  const regex = /^⟦(zg)?\:([^\:\|]*)\:([^\:\|]+)\|(.*?)⟧$/;
  //for (let j in ranges) {
  for (let j = ranges.length - 1; j >= 0; j--) {

    text = ranges[j].asRenderedString();
    found = text.match(regex);

    //Logger.log(found[0] + '-' + found[1] + '-' + found[2] + '-' + found[3] + '-' + found[4]);
    if (found) {
      counter.value++;
      if (!found[1]) {
        found[1] = 'zg';
      };
      if (!found[2]) {
        found[2] = '2249382';
      };
      urlpart = found[1] + '/' + found[2] + '/7/' + found[3] + '/' + found[4]; // +"//"+text;
      linkText = found[4];
    } else {
      urlpart = 'NA/' + text;
    };

    if (toDo == 'packZotero') {
      // Adds to URL zoteroItemKey and zoteroCollectionKey of current Slide
      newUrl = addCurrentKeys(urlpart);
      ranges[j].setText('⇡' + found[4]).getTextStyle().setLinkUrl(newUrl).setFontSize(12);
    } else if (toDo == 'minifyCitations') {
      ranges[j].getTextStyle().setFontSize(6).setForegroundColor('#fe01dc');
      ranges[j].getRange(text.length - linkText.length - 1, text.length - 1).getTextStyle().setForegroundColor('#0123dd').setFontSize(12);
    } else if (toDo == 'maxifyCitations') {
      ranges[j].getTextStyle().setFontSize(12).setForegroundColor('#fe01dc');
      ranges[j].getRange(text.length - linkText.length - 1, text.length - 1).getTextStyle().setForegroundColor('#0123dd');
    } else if (toDo == 'unfyCitations') {
      ranges[j].getTextStyle().setForegroundColor('#000000').setFontSize(12);
    }

  }
}

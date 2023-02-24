function clearLinkMarkers() {
  const ui = getUi();
  try {
    collectLinkMarks();
    if (HOST_APP == 'docs') {
      const doc = DocumentApp.getActiveDocument();
      for (let mark in LINK_MARK_OBJ) {
        doc.replaceText(LINK_MARK_OBJ[mark], '');
      }
      /* doc.replaceText(ORPHANED_LINK_MARK, '');
       doc.replaceText(URL_CHANGED_LINK_MARK, '');
       doc.replaceText(BROKEN_LINK_MARK, '');
       //doc.replaceText(UNKNOWN_LIBRARY_MARK, '');
       doc.replaceText(NORMAL_LINK_MARK, '');
       doc.replaceText(NORMAL_REDIRECT_LINK_MARK, '');
 */
      const footnotes = doc.getFootnotes();
      let footnote;
      for (let i in footnotes) {
        footnote = footnotes[i].getFootnoteContents();
        for (let mark in LINK_MARK_OBJ) {
          footnote.replaceText(LINK_MARK_OBJ[mark], '');
        }
        // footnote.replaceText(ORPHANED_LINK_MARK, '');
        // footnote.replaceText(URL_CHANGED_LINK_MARK, '');
        // footnote.replaceText(BROKEN_LINK_MARK, '');
        // //footnote.replaceText(UNKNOWN_LIBRARY_MARK, '');
        // footnote.replaceText(NORMAL_LINK_MARK, '');
        // footnote.replaceText(NORMAL_REDIRECT_LINK_MARK, '');
      }
    } else {
      // Slides part
      const slides = SlidesApp.getActivePresentation().getSlides();
      for (let i in slides) {
        slides[i].getPageElements().forEach(function (pageElement) {
          if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
            for (let mark in LINK_MARK_OBJ) {
              pageElement.asShape().getText().replaceAllText(LINK_MARK_OBJ[mark], '');
            }
            // pageElement.asShape().getText().replaceAllText(ORPHANED_LINK_MARK, '');
            // pageElement.asShape().getText().replaceAllText(URL_CHANGED_LINK_MARK, '');
            // pageElement.asShape().getText().replaceAllText(BROKEN_LINK_MARK, '');
            // //pageElement.asShape().getText().replaceAllText(UNKNOWN_LIBRARY_MARK, '');
            // pageElement.asShape().getText().replaceAllText(NORMAL_LINK_MARK, '');
            // pageElement.asShape().getText().replaceAllText(NORMAL_REDIRECT_LINK_MARK, '');
          } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {

            const table = pageElement.asTable();
            numRows = table.getNumRows();
            numCols = table.getNumColumns();
            for (let m = 0; m < numRows; m++) {
              for (let k = 0; k < numCols; k++) {
                cell = table.getRow(m).getCell(k);
                if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
                  for (let mark in LINK_MARK_OBJ) {
                    cell.getText().replaceAllText(LINK_MARK_OBJ[mark], '');
                  }
                  // cell.getText().replaceAllText(ORPHANED_LINK_MARK, '');
                  // cell.getText().replaceAllText(URL_CHANGED_LINK_MARK, '');
                  // cell.getText().replaceAllText(BROKEN_LINK_MARK, '');
                  // //cell.getText().replaceAllText(UNKNOWN_LIBRARY_MARK, '');
                  // cell.getText().replaceAllText(NORMAL_LINK_MARK, '');
                  // cell.getText().replaceAllText(NORMAL_REDIRECT_LINK_MARK, '');
                }
              }
            }
          }
        });
      }
      // End. Slides part
    }

  }
  catch (error) {
    ui.alert('Error in clearLinkMarkers: ' + error);
  }
}

// Instead of clearZwarnings
function clearWarningMarkers() {
  const ui = getUi();
  try {
    if (HOST_APP == 'docs') {
      const doc = DocumentApp.getActiveDocument();
      const regEx = "《warning:[^《》]*?》";
      const regEx2 = UL + "|" + UR;
      doc.replaceText(regEx, '');
      doc.replaceText(regEx2, '');

      const footnotes = doc.getFootnotes();
      let footnote;
      for (let i in footnotes) {
        footnote = footnotes[i].getFootnoteContents();
        footnote.replaceText(regEx, '');
        footnote.replaceText(regEx2, '');
      }
    } else {
      removeMarkersSlides('removeWarningMarkers');
    }
  }
  catch (error) {
    ui.alert('Error in clearWarningMarkers: ' + error);
  }
}

function removeCountryMarkers() {
  const ui = getUi();
  try {
    if (HOST_APP == 'docs') {
      singleReplace("⇡[^⇡:]+: ?", "⇡", true, false, false);
    } else {
      removeMarkersSlides('removeCountryMarkers');
    }
  }
  catch (error) {
    ui.alert('Error in removeCountryMarkers: ' + error);
  }
}


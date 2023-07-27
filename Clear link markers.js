function alertSuggestedFootnoteBug(i) {
  alert("Footnote has no contents. Footnote number = " + (+i + 1) + ". This appears to be a GDocs bug that happens if the footnote is suggested text only.");
}

function clearLinkMarkers() {
  const ui = getUi();
  try {
    collectLinkMarks();
    // Converts link markers to regular explessions
    for (let mark in LINK_MARK_OBJ) {
      if (['valid_ambiguous', 'redirect_ambiguous', 'importable', 'importable_ambiguous', 'importable_redirect'].includes(mark.replace('_LINK_MARK', '').toLowerCase())) {
        LINK_MARK_OBJ[mark] = LINK_MARK_OBJ[mark].replace('>', ':?[^<>]*>');
      }
    }
    // End. Converts link markers to regular explessions

    if (HOST_APP == 'docs') {
      const doc = DocumentApp.getActiveDocument();
      for (let mark in LINK_MARK_OBJ) {
        doc.replaceText(LINK_MARK_OBJ[mark], '');
      }
      const footnotes = doc.getFootnotes();
      let footnote;
      for (let i in footnotes) {
        footnote = footnotes[i].getFootnoteContents();
        if (footnote == null) {
          //alertSuggestedFootnoteBug(i);
          continue;
        }
        for (let mark in LINK_MARK_OBJ) {
          footnote.replaceText(LINK_MARK_OBJ[mark], '');
        }
      }
    } else {
      // Slides part
      const slides = SlidesApp.getActivePresentation().getSlides();
      for (let i in slides) {
        slides[i].getPageElements().forEach(function (pageElement) {
          if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
            for (let mark in LINK_MARK_OBJ) {
              const textRangesArray = pageElement.asShape().getText().find(LINK_MARK_OBJ[mark]);
              if (textRangesArray.length > 0) {
                textRangesArray.forEach(textRange => textRange.clear());
              }
            }
          } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {

            const table = pageElement.asTable();
            numRows = table.getNumRows();
            numCols = table.getNumColumns();
            for (let m = 0; m < numRows; m++) {
              for (let k = 0; k < numCols; k++) {
                cell = table.getRow(m).getCell(k);
                if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
                  for (let mark in LINK_MARK_OBJ) {
                    const textRangesArray = cell.getText().find(LINK_MARK_OBJ[mark]);
                    if (textRangesArray.length > 0) {
                      textRangesArray.forEach(textRange => textRange.clear());
                    }
                  }
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
        if (footnote == null) {
          alertSuggestedFootnoteBug(i);
          continue;
        }
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


function removeMarkersSlides(toDo) {
  let regEx;
  if (toDo == 'removeCountryMarkers') {
    regEx = '⇡[^⇡:]+: ?';
  } else if (toDo == 'removeWarningMarkers') {
    regEx = '《warning:[^《》]*?》';
  }
  const slides = SlidesApp.getActivePresentation().getSlides();
  for (let i in slides) {
    slides[i].getPageElements().forEach(function (pageElement) {
      if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
        rangeElement = pageElement.asShape().getText().find(regEx);
        removeMarkersSlidesHelper(toDo, rangeElement);
      } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {
        const table = pageElement.asTable();
        numRows = table.getNumRows();
        numCols = table.getNumColumns();
        for (let m = 0; m < numRows; m++) {
          for (let k = 0; k < numCols; k++) {
            cell = table.getRow(m).getCell(k);
            if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
              rangeElement = cell.getText().find(regEx);
              removeMarkersSlidesHelper(toDo, rangeElement);
            }
          }
        }
      }
    });
  }
}

function removeMarkersSlidesHelper(toDo, rangeElement) {
  for (let j in rangeElement) {
    textToReplace = rangeElement[j].asString();
    if (toDo == 'removeCountryMarkers') {
      rangeElement[j].insertText(textToReplace.length, '⇡');
      rangeElement[j].replaceAllText(textToReplace, '');
    } else if (toDo == 'removeWarningMarkers') {
      rangeElement[j].replaceAllText(textToReplace, '');
    }
  }
}

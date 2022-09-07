function changeAllSlidesLinks(toDo) {
  const slides = SlidesApp.getActivePresentation().getSlides();
  for (let i in slides) {
    slides[i].getPageElements().forEach(function (pageElement) {
      if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
        links = pageElement.asShape().getText().getLinks();
        changeAllSlidesLinksHelper(toDo, links);
      } else if (pageElement.getPageElementType() == SlidesApp.PageElementType.TABLE) {
        const table = pageElement.asTable();
        numRows = table.getNumRows();
        numCols = table.getNumColumns();
        for (let m = 0; m < numRows; m++) {
          for (let k = 0; k < numCols; k++) {
            cell = table.getRow(m).getCell(k);
            if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {
              links = cell.getText().getLinks();
              changeAllSlidesLinksHelper(toDo, links);
            }
          }
        }
      }
    });
  }
}

function changeAllSlidesLinksHelper(toDo, links) {
  for (let j in links) {
    if (toDo == 'removeUnderlineFromHyperlinks' && links[j].getTextStyle().isUnderline()) {
      links[j].getTextStyle().setUnderline(false);
    } else if (toDo == 'removeOpeninZoteroapp') {
      linkUrl = links[j].getTextStyle().getLink().getUrl();

      checkOpenin = /openin=zoteroapp&?/i.exec(linkUrl);
      if (checkOpenin != null) {
        links[j].getTextStyle().setLinkUrl(removeOpeninZoteroappFromUrl(linkUrl, checkOpenin));
      }
    }
  }
}
function removeOpeninZoteroapp() {
  const ui = getUi();
  try {
    const toDo = 'removeOpeninZoteroapp';
    if (HOST_APP == 'docs') {
      const doc = DocumentApp.getActiveDocument();
      changeAllBodyLinks(toDo, doc);
      changeAllFootnotesLinks(toDo, doc);
    } else {
      changeAllSlidesLinks(toDo);
    }
  }
  catch (error) {
    ui.alert('Error in removeOpeninZoteroapp: ' + error);
  }
}

function removeUnderlineFromHyperlinks() {
  const ui = getUi();
  try {
    const toDo = 'removeUnderlineFromHyperlinks';
    if (HOST_APP == 'docs') {
      const doc = DocumentApp.getActiveDocument();
      changeAllBodyLinks(toDo, doc);
      changeAllFootnotesLinks(toDo, doc);
    } else {
      changeAllSlidesLinks(toDo);
    }
  }
  catch (error) {
    ui.alert('Error in removeUnderlineFromHyperlinks: ' + error);
  }
}

function changeAllBodyLinks(toDo, doc) {
  const element = doc.getBody();
  changeAllLinks(element, toDo);
}

// Copy of part of function addSrcToURL
function removeOpeninZoteroappFromUrl(url, checkOpenin) {
  url = url.replace(checkOpenin[0], '');

  const lastChar = url.charAt(url.length - 1);
  if (lastChar == '&' || lastChar == '?') {
    url = url.slice(0, -1);
  }

  return url;
}

function changeAllLinks(element, toDo) {

  let text, end, indices, partAttributes, numChildren, getIndexFlag, checkOpenin;
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    //Logger.log(indices);
    for (let i = 0; i < indices.length; i++) {
      partAttributes = element.getAttributes(indices[i]);
      //Logger.log(partAttributes);
      if (partAttributes.LINK_URL) {

        getIndexFlag = false;

        if (toDo == 'removeOpeninZoteroapp') {
          checkOpenin = /openin=zoteroapp&?/i.exec(partAttributes.LINK_URL);
          if (checkOpenin != null) {
            getIndexFlag = true;
          }
        } else if (toDo == 'removeUnderlineFromHyperlinks' && partAttributes.UNDERLINE) {
          getIndexFlag = true;
        }

        if (getIndexFlag === true) {
          if (i == indices.length - 1) {
            text = element.getText();
            end = text.length - 1;
          } else {
            end = indices[i + 1] - 1;
          }
          if (toDo == 'removeOpeninZoteroapp') {
            element.setLinkUrl(indices[i], end, removeOpeninZoteroappFromUrl(partAttributes.LINK_URL, checkOpenin));
            //Logger.log(partAttributes.LINK_URL + ' change');
          }
          if (toDo == 'removeUnderlineFromHyperlinks') {
            element.setUnderline(indices[i], end, false);
          }
        }
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        changeAllLinks(element.getChild(i), toDo);
      }
    }
  }
}


function changeAllFootnotesLinks(toDo, doc) {
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
      changeAllLinks(footnote.getChild(j), toDo);
    }
  }
}

function modifyLinkParenthesis() {
  if (HOST_APP == 'docs') {
    modifyLinkParenthesisDoc();
  } else {
    alert('Sorry! The modifyLink function has not been implemented yet for Google Slides.');
  }
}

function modifyLinkParenthesisDoc() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const cursor = doc.getCursor();
  let start, end, cursorPos;
  const urlRegEx = refOpenDevEdLinksRegEx();

  if (selection) {
    const elements = selection.getRangeElements();
    elements.forEach(element => {
      const el = element.getElement();
      const text = el.getText();
      if (element.isPartial()) {
        start = element.getStartOffset();
        end = element.getEndOffsetInclusive();
      } else {
        start = 0;
        end = text.length;
      }
      if (el.getType() === DocumentApp.ElementType.TEXT) {
        processTextElement('selection', el, start, end, cursorPos, urlRegEx);
      } else {
        const paraout = [];
        let paragraph, paragraphs;
        if (el.getType() == DocumentApp.ElementType.LIST_ITEM || el.getType() == DocumentApp.ElementType.PARAGRAPH) {
          // paragraph = el.asListItem();
          paragraph = el;
          paraout.push(paragraph);
        } else if (el.getType() == DocumentApp.ElementType.TABLE_CELL) {
          paragraphs = paragraphsFromTableCell(el);
          paraout.push(...paragraphs);
        } else if (el.getType() == DocumentApp.ElementType.TABLE_ROW) {
          paragraphs = paragraphsFromTableRow(el);
          paraout.push(...paragraphs);
        } else if (el.getType() == DocumentApp.ElementType.TABLE) {
          paragraphs = paragraphsFromTable(el);
          paraout.push(...paragraphs);
        }
        paraout.forEach(paragraph => {
          processTextElement('selection', paragraph.asText().asText(), 0, paragraph.getText().length, cursorPos, urlRegEx);
        });
      }
    });
  } else if (cursor) {
    cursorPos = cursor.getSurroundingTextOffset();
    const element = cursor.getSurroundingText();
    const elementType = element.getType();
    if (elementType === DocumentApp.ElementType.TEXT || elementType === DocumentApp.ElementType.PARAGRAPH || elementType === DocumentApp.ElementType.LIST_ITEM) {
      processTextElement('cursor', element, start, end, cursorPos, urlRegEx);
    }
  }
}


function processTextElement(type, element, start, end, cursorPos, urlRegEx) {
  //Logger.log('%s %s %s %s %s', type, element, start, end, cursorPos);

  const text = element.getText();
  const linkUrls = [];
  const indices = element.getTextAttributeIndices();

  for (let i = 0; i < indices.length; i++) {
    const offset = indices[i];
    const url = element.getLinkUrl(offset);
    if (url) {
      const linkStart = offset;
      const linkEnd = (i < indices.length - 1) ? indices[i + 1] : text.length;
      const linkText = text.substr(linkStart, linkEnd - linkStart)
      linkUrls.push({ url: url, linkStart: linkStart, linkEnd: linkEnd, linkText: linkText });
    }
  }

  // Logger.log(linkUrls);


  const linksInParenthesisRegEx = /\(\s*⇡[A-Za-z\s&.,|]+\s*\d{4}(?:\s*;\s*⇡[A-Za-z\s&.,]+\s*\d{4})*\s*\)/g;
  const linksWithYearsInParenthesisRegEx = /(?:⇡\s*[A-Za-z\s&.,]+\s*\(\s*\d{4}\s*\)\s*,\s*)*(?:⇡\s*[A-Za-z\s&.,]+\s*\(\s*\d{4}\s*\)\s*(?:, and | and ))?⇡\s*[A-Za-z\s&.,]+\s*\(\s*\d{4}\s*\)/g;;

  let textsToTurnArray = [];
  const linksInParenthesis = findReferences(textsToTurnArray, text, linksInParenthesisRegEx, 'linksInParenthesis', type, start, end, cursorPos);
  //Logger.log(linksInParenthesis);
  const linksWithYearsInParenthes = findReferences(textsToTurnArray, text, linksWithYearsInParenthesisRegEx, 'linksWithYearsInParenthes', type, start, end, cursorPos);
  //Logger.log(linksWithYearsInParenthes);

  textsToTurnArray = linksInParenthesis.concat(linksWithYearsInParenthes);
  textsToTurnArray = sortByStart(textsToTurnArray);
  //Logger.log(textsToTurnArray);
  textsToTurnArray = processLinks(textsToTurnArray, linkUrls);
  //Logger.log(textsToTurnArray);

  let resultTxt = '';
  let nonOpenDevEdLinksFlag = false;
  let nonOpenDevEdLinksList = '';
  for (let i = textsToTurnArray.length - 1; i >= 0; i--) {
    let allOpenDevEdLinksFlag = true;
    //Logger.log('%s, %s', textsToTurnArray[i].start, textsToTurnArray[i].end);

    let newText = '';
    let startNewText = textsToTurnArray[i].start;
    if (textsToTurnArray[i].linksType === 'linksWithYearsInParenthes') {
      startNewText += 1;
    }

    for (let j = 0; j < textsToTurnArray[i].linksArray.length; j++) {
      let year, firstPart, fullMatch, linkText, newStart, newEnd, oneNewLinkText, semicolonOrComma;
      linkText = textsToTurnArray[i].linksArray[j].linkText;
      if (urlRegEx.test(textsToTurnArray[i].linksArray[j].url) === false) {
        allOpenDevEdLinksFlag = false;
        nonOpenDevEdLinksFlag = true;
        nonOpenDevEdLinksList += '\n\n' + linkText + ' ' + textsToTurnArray[i].linksArray[j].url;
        break;
      }


      if (textsToTurnArray[i].linksType === 'linksWithYearsInParenthes') {
        semicolonOrComma = ';';
      } else {
        semicolonOrComma = ',';
      }

      resultTxt = extractYear(textsToTurnArray[i].linksType, linkText);
      if (resultTxt == null) {
        continue;
      }
      year = resultTxt.year;
      fullMatch = resultTxt.fullMatch;
      firstPart = linkText.replace(fullMatch, '').trim();
      if (textsToTurnArray[i].linksType === 'linksWithYearsInParenthes') {
        oneNewLinkText = firstPart + ', ' + year;
      } else {
        oneNewLinkText = firstPart + ' (' + year + ')';
      }
      newText += oneNewLinkText;
      newStart = startNewText;
      newEnd = newStart + oneNewLinkText.length - 1;
      textsToTurnArray[i].linksArray[j].newStart = newStart;
      textsToTurnArray[i].linksArray[j].newEnd = newEnd;
      startNewText += oneNewLinkText.length;
      if (textsToTurnArray[i].linksType === 'linksWithYearsInParenthes') {
        // if (j != textsToTurnArray[i].linksArray.length - 1 && j < textsToTurnArray[i].linksArray.length - 1) {
        if (j < textsToTurnArray[i].linksArray.length - 1) {
          newText += semicolonOrComma + ' ';
          startNewText += 2;
        }
      } else {
        if (textsToTurnArray[i].linksArray.length === 2 && j === 0) {
          newText += ' and ';
          startNewText += 5;
        } else if (textsToTurnArray[i].linksArray.length > 2 && j === textsToTurnArray[i].linksArray.length - 2) {
          newText += ', and ';
          startNewText += 6;
        } else if (j !== textsToTurnArray[i].linksArray.length - 1) {
          newText += semicolonOrComma + ' ';
          startNewText += 2;
        }
      }
    }

    if (allOpenDevEdLinksFlag === true) {
      element.deleteText(textsToTurnArray[i].start, textsToTurnArray[i].end - 1);
      if (textsToTurnArray[i].linksType === 'linksWithYearsInParenthes') {
        newText = '(' + newText + ')';
      }

      element.insertText(textsToTurnArray[i].start, newText);
      for (let j = 0; j < textsToTurnArray[i].linksArray.length; j++) {
        if (textsToTurnArray[i].linksArray[j].newStart >= 0 && textsToTurnArray[i].linksArray[j].newEnd >= 0) {
          element.setLinkUrl(textsToTurnArray[i].linksArray[j].newStart, textsToTurnArray[i].linksArray[j].newEnd, textsToTurnArray[i].linksArray[j].url).setUnderline(false);
        }
      }

    }
  }
  if (nonOpenDevEdLinksFlag === true) {
    alert('Non-OpenDevEd links:' + nonOpenDevEdLinksList);
  }
}

function processLinks(textsArray, linksArray) {
  return textsArray.map(textObj => {
    textObj.linksArray = linksArray.filter(linkObj =>
      linkObj.linkStart >= textObj.start && linkObj.linkEnd <= textObj.end
    );
    return textObj;
  });
}

function sortByStart(arr) {
  return arr.sort((a, b) => a.start - b.start);
}

function findReferences(textsToTurnArray, text, regex, linksType, type, startSelection, endSelection, cursorPos) {
  let match, start, end;
  let results = [];

  // Use regex.exec to find all matches and their positions
  while ((match = regex.exec(text)) !== null) {
    start = match.index;
    end = match.index + match[0].length;
    if ((type === 'cursor' && cursorPos >= start && cursorPos <= end) ||
      (type === 'selection' && (
        startSelection >= start && startSelection <= end
        || endSelection >= start && endSelection <= end
        || startSelection <= start && endSelection >= start
        || startSelection <= end && endSelection >= end
      ))) {
      results.push({
        start: start,
        end: end,
        substring: match[0],
        linksType: linksType
      });
    }
  }
  return results;
}

function paragraphsFromTableCell(element) {
  const paragraphs = [];
  const cellParagraphs = element.getNumChildren();
  for (let m = 0; m < cellParagraphs; m++) {
    const child = element.getChild(m);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      paragraphs.push(child);
    }
  }
  return paragraphs;
}

function paragraphsFromTableRow(element) {
  const paragraphs = [];
  const cells = element.getNumCells();
  for (let k = 0; k < cells; k++) {
    const cell = element.getCell(k);
    paragraphs.push(...paragraphsFromTableCell(cell));
  }
  return paragraphs;
}

function paragraphsFromTable(element) {
  const paragraphs = [];
  const rows = element.getNumRows();
  for (let k = 0; k < rows; k++) {
    const row = element.getRow(k);
    paragraphs.push(...paragraphsFromTableRow(row));
  }
  return paragraphs;
}

function extractYear(type, text){
    const regex = type === 'linksWithYearsInParenthes' ? /\(\s*(\d{4})\s*\)/ : /,\s*(\d{4})\s*/;
    const match = text.match(regex);

    if (match) {
      return {
        fullMatch: match[0],  // e.g., "(2022)", "( 2022)", etc.
        year: match[1]        // e.g., "2022"
      };
    }
    return null;
}
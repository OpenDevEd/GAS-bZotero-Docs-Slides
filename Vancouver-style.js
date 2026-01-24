function applyVancouverStyle() {
  vancouverStyle('Vancouver');
  addUsageTrackingRecord('applyVancouverStyle');
}

function applyAPA7Style() {
  vancouverStyle('APA7');
  addUsageTrackingRecord('applyAPA7Style');
}

function vancouverStyle(toDo) {
  const linksArray = [];
  const linkNumbers = new Object();
  const urlRegEx = refOpenDevEdLinksRegEx();

  if (HOST_APP == 'docs') {
    // Doc part
    const doc = DocumentApp.getActiveDocument();

    const element = doc.getBody();
    vancouverStyleFindAllLinks(toDo, element, 'body', linksArray, linkNumbers, urlRegEx);

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
        vancouverStyleFindAllLinks(toDo, footnote.getChild(j), 'footnotes', linksArray, linkNumbers, urlRegEx);
      }
    }
    // End. Doc part
  } else {
    // Slides part
    const slides = SlidesApp.getActivePresentation().getSlides();
    let url, linkText, runs, color, fontFamily, fontSize;
    let toDoArray = [];
    for (let i in slides) {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
          toDoArray = [];
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
              vancouverOrAPA7(toDo, toDoArray, url, start = 0, end = linkText.length - 1, urlRegEx, partAttributes = links[m], text = linkText, linkNumbers, linksArray, source = 'slides');
            }
          }

          //Logger.log(toDoArray);
          if (toDoArray.length > 0) {
            for (let i in toDoArray) {
              link = toDoArray[i].partAttributes;
              color = link.getTextStyle().getForegroundColor();
              fontFamily = link.getTextStyle().getFontFamily();
              fontSize = link.getTextStyle().getFontSize();
              link.setText(toDoArray[i].number).getTextStyle().setLinkUrl(toDoArray[i].url).setForegroundColor(color).setFontFamily(fontFamily).setFontSize(fontSize);
            }
          }
        }
      });
    }
    //End. Slides part
  }

  // Logger.log(linksArray);
  // Logger.log(linkNumbers);
  return 0;
}

function vancouverStyleFindAllLinks(toDo, element, source, linksArray, linkNumbers, urlRegEx) {
  let text, end, indices, partAttributes, numChildren, url, groupIdItemKey, number, linkText, reNumbering, resultLinkTextFromUrl, skip;
  const toDoArray = [];
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {

    indices = element.getTextAttributeIndices();
    for (let i = 0; i < indices.length; i++) {
      skip = false;
      partAttributes = element.getAttributes(indices[i]);
      //Logger.log(partAttributes);
      if (partAttributes.LINK_URL) {

        text = element.getText();
        start = indices[i];

        if (i == indices.length - 1) {
          end = text.length - 1;
        } else {
          end = indices[i + 1] - 1;
        }

        url = partAttributes.LINK_URL;
        vancouverOrAPA7(toDo, toDoArray, url, start, end, urlRegEx, partAttributes, text, linkNumbers, linksArray, source);
      }
    }

    if (toDoArray.length > 0) {
      for (let i in toDoArray) {
        element.deleteText(toDoArray[i].start, toDoArray[i].end);
        element.insertText(toDoArray[i].start, toDoArray[i].number).setAttributes(toDoArray[i].start, toDoArray[i].newEnd, toDoArray[i].partAttributes);
        element.setLinkUrl(toDoArray[i].start, toDoArray[i].newEnd, toDoArray[i].url).setUnderline(false);
      }
    }

  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        vancouverStyleFindAllLinks(toDo, element.getChild(i), source, linksArray, linkNumbers, urlRegEx);
      }
    }
  }
}


function addTextParameter(url, linkText) {
  linkText = linkText.replace('⇡', '');
  linkText = encodeURIComponent(linkText);
  linkText = linkText.replaceAll('%20', ' ');
  const resultLinkTextFromUrl = /text=([^&]+)/i.exec(url);
  if (resultLinkTextFromUrl == null) {
    if (/\?+\&?/.test(url)) {
      qmOrAmp = '&';
    } else {
      qmOrAmp = '?';
    }
    url = url + qmOrAmp + 'text=' + linkText;
  } else {
    url = url.replace(resultLinkTextFromUrl[0], 'text=' + linkText);
  }
  return url;
}

function removeTextParameter(url) {
  const resultLinkTextFromUrl = /text=([^&]+)/i.exec(url);
  if (resultLinkTextFromUrl != null) {
    url = url.replace(resultLinkTextFromUrl[0], '');
    if (/(\?|\&)$/.test(url)) {
      url = url.slice(0, -1);
    }
  }
  return url;
}

function vancouverOrAPA7(toDo, toDoArray, url, start, end, urlRegEx, partAttributes, text, linkNumbers, linksArray, source) {
  let linkText, reNumbering, resultLinkTextFromUrl, number, skip = false;
  if (urlRegEx.test(url)) {
    if (source == 'slides') {
      linkText = text;
    } else {
      linkText = text.substr(start, end - start + 1);
    }
    if (/⇡[0-9]+$/.test(linkText)) {
      // Now it's Vancouver style
      reNumbering = true;
      //Logger.log('reNumbering = true');
    } else {
      // Now it's APA7 style
      reNumbering = false;
      //Logger.log('reNumbering = false');
    }

    if (toDo == 'APA7' && reNumbering === true) {
      // Vancouver -> APA7
      resultLinkTextFromUrl = /text=([^&]+)/i.exec(url);
      if (resultLinkTextFromUrl == null) {
        skip = true;
      } else {
        number = '⇡' + decodeURIComponent(resultLinkTextFromUrl[1]);
        url = removeTextParameter(url);
      }
      // End. Vancouver -> APA7
    } else if (toDo == 'Vancouver') {
      // APA7 -> Vancouver
      const resultGroupIdItemKeyIn = getGroupIdItemKey(url);
      if (resultGroupIdItemKeyIn.status != 'ok') {
        throw new Error(resultGroupIdItemKeyIn.message);
      }
      const groupIdItemKey = resultGroupIdItemKeyIn.groupId + ':' + resultGroupIdItemKeyIn.itemKey;
      //Logger.log(resultGroupIdItemKeyIn);
      if (linkNumbers.hasOwnProperty(groupIdItemKey)) {
        number = linkNumbers[groupIdItemKey];
        if (reNumbering === false) {
          // linkText = linkText.replace('⇡', '');
          // linkText = encodeURIComponent(linkText);
          // linkText = linkText.replaceAll('%20', ' ');
          url = addTextParameter(url, linkText);
        }
      } else {
        number = linksArray.length + 1;
        linkNumbers[groupIdItemKey] = number;

        if (reNumbering === false) {
          // linkText = linkText.replace('⇡', '');
          // linkText = encodeURIComponent(linkText);
          // linkText = linkText.replaceAll('%20', ' ');
          url = addTextParameter(url, linkText);
        }
        linksArray.push({ link: url, linkText: linkText, source: source });
      }
      number = '⇡' + number.toString();
      // End. APA7 -> Vancouver
    } else {
      skip = true;
    }

    if (skip === false && linkText != number) {
      const newEnd = start + number.length - 1;
      toDoArray.unshift({ start: start, end: end, newEnd: newEnd, url: url, partAttributes: partAttributes, number: number });
    } else {
      //Logger.log("No need to change the link")
    }
  } else {
    //Logger.log('Other links');
  }
}
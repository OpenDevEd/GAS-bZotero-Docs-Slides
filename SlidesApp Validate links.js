function validateSlides(bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, getparams, markorphanedlinks, analysekerkolinks, flagsObject, onlyLinks, newForestAPIjson, newForestAPI) {
  //Logger.log('1 ' + JSON.stringify(newForestAPIjson) + ' ' + newForestAPI);
  const slides = SlidesApp.getActivePresentation().getSlides();
  let bibSlideFlag, shapes, pageElement;

  for (let i in slides) {
    bibSlideFlag = false;
    shapes = slides[i].getShapes();
    for (let j in shapes) {
      pageElement = shapes[j];
      rangeElementStart = pageElement.getText().find(TEXT_TO_DETECT_START_BIB);
      rangeElementEnd = pageElement.getText().find(TEXT_TO_DETECT_END_BIB);
      if ((rangeElementStart.length > 0 && rangeElementEnd.length > 0)) {
        //Logger.log('Bibliography slide. Don\'t collect links!');
        bibSlideFlag = true;
        break;
      } else {

        if (markorphanedlinks === true) {
          links = pageElement.getText().getLinks();
          orphanedChangedLinksHelper(links, flagsObject);
        }

        links = pageElement.getText().getLinks();
        validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks, newForestAPIjson, newForestAPI);
      }
    }

    if (bibSlideFlag === true) {
      break;
    }

    slides[i].getTables().forEach(table => {
      numRows = table.getNumRows();
      numCols = table.getNumColumns();
      for (let m = 0; m < numRows; m++) {
        for (let k = 0; k < numCols; k++) {
          cell = table.getRow(m).getCell(k);
          if (cell.getMergeState() == 'HEAD' || cell.getMergeState() == 'NORMAL') {

            if (markorphanedlinks === true) {
              links = pageElement.getText().getLinks();
              orphanedChangedLinksHelper(links, flagsObject);
            }

            links = cell.getText().getLinks();
            validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks, newForestAPIjson, newForestAPI);
          }
        }
      }
    });
  }
}


function orphanedChangedLinksHelper(links, flagsObject) {
  previousLinks = [];
  for (let m in links) {
    url = links[m].getTextStyle().getLink().getUrl();
    linkText = links[m].asRenderedString();
    insertOrphanedChangedLinksMarkers(links[m], previousLinks, flagsObject, links[m].getStartIndex(), links[m].getEndIndex(), linkText, url);
  }
}

function validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks, newForestAPIjson, newForestAPI) {
  let linkMarkNormal;
  const ui = SlidesApp.getUi();


  for (let j = links.length - 1; j >= 0; j--) {
    // Logger.log('TXT=' + links[j].asRenderedString());
    // Logger.log(links[j]);
    // Logger.log(links[j].getTextStyle());
    // Logger.log(links[j].getTextStyle().getLink());
    // Logger.log(links[j].getTextStyle().getLink().getUrl());
    link = links[j].getTextStyle().getLink().getUrl();
    if (link == null || REF_OPENDEVED_LINKS_REGEX.test(link) === false) {
      continue;
    }
    result = checkHyperlinkSlides(bibReferences, alreadyCheckedLinks, link, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, newForestAPIjson, newForestAPI);
    if (result.status == 'error') {
      ui.alert(result.message);
      return 0;
    }

    if (validate || getparams) {
      realLinkText = links[j].asRenderedString();

      if (analysekerkolinks) {
        if (validate && (result.type == 'NORMAL LINK' || result.forestType == 'valid' || result.forestType == 'redirect')) {
          const validType = newForestAPI === true ? result.forestType : result.normalLinkType;
          const VALID_OR_REDIRECT_MARK = LINK_MARK_OBJ[validType.toString().toUpperCase() + '_LINK_MARK'];
          insertMarkSlides(links[j], realLinkText, VALID_OR_REDIRECT_MARK, link);
          flagsObject['notiText' + '_' + validType].status = true;
        }
      } else {
        // Normal link. No need to set any mark
        if (result.url != link) {
          links[j].getTextStyle().setLinkUrl(result.url);
        }
        // End. Normal link. No need to set any mark        
      }

      if (validate && (result.type == 'BROKEN' || (newForestAPI === true && result.forestType != 'redirect' && result.forestType != 'valid'))) {
        const NOT_VALID_NOT_REDIRECT_MARK = getNotValidNotRedirectMark(result, flagsObject);
        insertMarkSlides(links[j], realLinkText, NOT_VALID_NOT_REDIRECT_MARK, link);
      }
    }
  }
}

function insertMarkSlides(linkElement, realLinkText, linkMark, url) {
  const markPlusLinkText = linkMark + realLinkText;
  linkElement.setText(markPlusLinkText);
  linkElement.getRange(markPlusLinkText.length - realLinkText.length, markPlusLinkText.length).getTextStyle().setLinkUrl(url).setBackgroundColorTransparent();
  linkElement.getRange(0, linkMark.length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BACKGROUND_COLOR']).setForegroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_FOREGROUND_COLOR']).setUnderline(false).setBold(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BOLD']);
}


function checkHyperlinkSlides(bibReferences, alreadyCheckedLinks, url, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, newForestAPIjson, newForestAPI) {
  let result, urlWithParameters;

  if (alreadyCheckedLinks.hasOwnProperty(url)) {
    result = alreadyCheckedLinks[url];
  } else {
    result = checkLink(url, validationSite, validate, newForestAPIjson, newForestAPI);
    if (result.status == 'error') {
      return result;
    }
    alreadyCheckedLinks[url] = result;
  }

  if (bibReferences.indexOf(result.bibRef) == -1) {
    bibReferences.push(result.bibRef);
  }

  let updatedUrl;
  if (validate && ((result.type == 'NORMAL LINK' && result.normalLinkType == 'REDIRECT') || (result.type == 'NEW FOREST API LINK' && result.forestType == 'redirect'))) {
    updatedUrl = result.url;
  } else {
    updatedUrl = url;
  }
  urlWithParameters = addSrcToURL(updatedUrl, targetRefLinks, zoteroItemKeyParameters, zoteroCollectionKey);
  return { status: 'ok', type: result.type, normalLinkType: result.normalLinkType, url: urlWithParameters, forestType: result.forestType };
}


function scanForItemKeySlides(targetRefLinks) {

  let rangeElementStart, tableText, libLink, result;
  let foundFlag = false;

  const slides = SlidesApp.getActivePresentation().getSlides();
  for (let i in slides) {
    if (foundFlag) {
      break;
    } else {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {

          rangeElementStart = pageElement.asShape().getText().find('docs.edtechhub.org/lib/[^/]+|docs.opendeved.net/lib/[^/]+');
          if (rangeElementStart.length > 0) {
            tableText = rangeElementStart[0].asRenderedString();
            libLink = /docs.edtechhub.org\/lib\/[a-zA-Z0-9]+|docs.opendeved.net\/lib\/[a-zA-Z0-9]+/.exec(tableText);
            if (libLink != null) {
              result = detectZoteroItemKeyType('https://' + libLink);
              if (result.status == 'error') {
                result = addZoteroItemKey('', false, false, targetRefLinks);
                return result;
              }
              foundFlag = true;
            }
          }
        }
      });
    }
  }

  if (!foundFlag) {
    result = addZoteroItemKey('', false, false, targetRefLinks);
  }
  return result;
}

function insertOrphanedChangedLinksMarkers(element, previousLinks, flagsObject, start, end, linkText, url) {
  let previousLinkIndex, flagMarkOrphanedLinks = false;

  // If there is an isolated whitespace (or more whitespaces) with a link
  if (linkText.match(/^\s+$/)) {
    flagMarkOrphanedLinks = true;
    //Logger.log('flagMarkOrphanedLinks = true;');
  }
  // End. If there is an isolated whitespace (or more whitespaces) with a link

  previousLinkIndex = previousLinks.length - 1;
  if (previousLinks.length > 0) {
    if (!flagMarkOrphanedLinks && previousLinks[previousLinkIndex].end == start && previousLinks[previousLinkIndex].url != url && !previousLinks[previousLinkIndex].linkText.match(/^\s+$/) && !previousLinks[previousLinkIndex].linkText[0].match(/^\s+$/)) {
      flagsObject.notiTextURLChanged = true;
      element.insertText(0, LINK_MARK_OBJ['URL_CHANGED_LINK_MARK']);
      element.getRange(0, LINK_MARK_OBJ['URL_CHANGED_LINK_MARK'].length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BACKGROUND_COLOR']).setForegroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_FOREGROUND_COLOR']).setUnderline(false).setBold(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BOLD']);
    }
  }

  previousLinks.push({ url: url, start: start, end: end, linkText: linkText });

  if (flagMarkOrphanedLinks) {
    flagsObject.notiTextOrphaned = true;
    element.insertText(0, LINK_MARK_OBJ['ORPHANED_LINK_MARK']);
    element.getRange(0, LINK_MARK_OBJ['ORPHANED_LINK_MARK'].length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BACKGROUND_COLOR']).setForegroundColor(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_FOREGROUND_COLOR']).setUnderline(false).setBold(LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BOLD']);
  }

}
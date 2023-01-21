function validateSlides(bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, getparams, markorphanedlinks, analysekerkolinks, flagsObject) {
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
        Logger.log('Bibliography slide. Don\'t collect links!');
        bibSlideFlag = true;
        break;
      } else {

        if (markorphanedlinks === true) {
          links = pageElement.getText().getLinks();
          orphanedChangedLinksHelper(links, flagsObject);
        }

        links = pageElement.getText().getLinks();
        validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks);
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
            validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks);
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

function validateSlidesHelper(links, bibReferences, alreadyCheckedLinks, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate, flagsObject, getparams, analysekerkolinks) {
  let linkMarkNormal;
  const ui = SlidesApp.getUi();


  for (let j = links.length - 1; j >= 0; j--) {
    // Logger.log('TXT=' + links[j].asRenderedString());
    // Logger.log(links[j]);
    // Logger.log(links[j].getTextStyle());
    // Logger.log(links[j].getTextStyle().getLink());
    // Logger.log(links[j].getTextStyle().getLink().getUrl());
    link = links[j].getTextStyle().getLink().getUrl();
    //Logger.log('JJ ' + links[j].getStartIndex() + ' ' + links[j].getEndIndex() + ' ' + link + ' ' + links[j].asRenderedString());
    //orphanedChangedlinksHelper(links[j], previousLinks, flagsObject, links[j].getStartIndex(), links[j].getEndIndex(), links[j].asRenderedString(), link);
    if (link == null) {
      continue;
    }
    result = checkHyperlinkSlides(bibReferences, alreadyCheckedLinks, link, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate);
    if (result.status == 'error') {
      ui.alert(result.message);
      return 0;
    }

    if (validate || getparams) {
      realLinkText = links[j].asRenderedString();

      if (result.type == 'NORMAL LINK') {
        if (analysekerkolinks) {
          if (result.normalLinkType == 'NORMAL') {
            linkMarkNormal = NORMAL_LINK_MARK;
            flagsObject.notiTextNormalLink = true;
          } else {
            linkMarkNormal = NORMAL_REDIRECT_LINK_MARK;
            flagsObject.notiTextNormalRedirectLink = true;
          }
          insertMarkSlides(links[j], realLinkText, linkMarkNormal, link);
        } else {
          // Normal link. No need to set any mark
          if (result.url != link) {
            links[j].getTextStyle().setLinkUrl(result.url);
          }
          // End. Normal link. No need to set any mark
        }
      }

      // The link is broken
      if (result.type == 'BROKEN LINK' && validate) {
        flagsObject.notiTextBroken = true;
        insertMarkSlides(links[j], realLinkText, BROKEN_LINK_MARK, link);
      }
      // End. The link is broken

      // The library isn't permitted
      /* if (result.permittedLibrary == false) {
        flagsObject.notiTextUnknownLibrary = true;
        insertMarkSlides(links[j], realLinkText, UNKNOWN_LIBRARY_MARK, link);
      } */
      // End. The library isn't permitted

    }
  }
}

function insertMarkSlides(linkElement, realLinkText, linkMark, url) {
  const markPlusLinkText = linkMark + realLinkText;
  linkElement.setText(markPlusLinkText);
  linkElement.getRange(markPlusLinkText.length - realLinkText.length, markPlusLinkText.length).getTextStyle().setLinkUrl(url).setBackgroundColorTransparent();
  linkElement.getRange(0, linkMark.length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_BACKGROUND_COLOR).setForegroundColor(LINK_MARK_STYLE_FOREGROUND_COLOR).setUnderline(false);
}


function checkHyperlinkSlides(bibReferences, alreadyCheckedLinks, url, validationSite, zoteroItemKeyParameters, targetRefLinks, zoteroCollectionKey, validate) {
  let result, urlWithParameters;

  const urlRegEx = new RegExp('https?://ref.opendeved.net/zo/zg/[0-9]+/7/[^/]+/?|https?://docs.(edtechhub.org|opendeved.net)/lib(/[^/\?]+/?|.*id=[A-Za-z0-9]+)', 'i');
  if (url.search(urlRegEx) == 0) {
    //Logger.log('Yes----------------------');

    if (alreadyCheckedLinks.hasOwnProperty(url)) {
      result = alreadyCheckedLinks[url];
    } else {
      result = checkLink(url, validationSite, validate);
      if (result.status == 'error') {
        return result;
      }
      alreadyCheckedLinks[url] = result;
    }


    if (bibReferences.indexOf(result.bibRef) == -1) {
      bibReferences.push(result.bibRef);
    }

    if (!validate || result.type == 'BROKEN LINK') {
      urlWithParameters = addSrcToURL(url, targetRefLinks, zoteroItemKeyParameters, zoteroCollectionKey);
    } else {
      urlWithParameters = addSrcToURL(result.url, targetRefLinks, zoteroItemKeyParameters, zoteroCollectionKey);
    }
    return { status: 'ok', type: result.type, normalLinkType: result.normalLinkType, url: urlWithParameters};
  }
  return { status: 'ok' };
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
              //Logger.log('libLink ' + libLink);
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
      //Logger.log('flagsObject.notiTextURLChanged = true;');
      element.insertText(0, URL_CHANGED_LINK_MARK);
      element.getRange(0, URL_CHANGED_LINK_MARK.length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_BACKGROUND_COLOR).setForegroundColor(LINK_MARK_STYLE_FOREGROUND_COLOR).setUnderline(false);
    }
  }

  previousLinks.push({ url: url, start: start, end: end, linkText: linkText });

  if (flagMarkOrphanedLinks) {
    flagsObject.notiTextOrphaned = true;
    element.insertText(0, ORPHANED_LINK_MARK);
    element.getRange(0, ORPHANED_LINK_MARK.length).getTextStyle().setBackgroundColor(LINK_MARK_STYLE_BACKGROUND_COLOR).setForegroundColor(LINK_MARK_STYLE_FOREGROUND_COLOR).setUnderline(false);
  }

}
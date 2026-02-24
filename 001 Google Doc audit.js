// ── Mode constants ────────────────────────────────────────────────────────────

const AUDIT_MODE = {
  AUDIT: 'audit',                                  // Mode 1: returns audit result object
  CONVERT_TEXT_URLS: 'convertTextUrls',            // Mode 2: converts plain-text URLs to hyperlinks
  SET_ZOTERO_HIGHLIGHT: 'setZoteroHighlight',      // Mode 3: yellow background on Zotero links
  REMOVE_ZOTERO_HIGHLIGHT: 'removeZoteroHighlight' // Mode 4: removes yellow background
};

// ── Entry point ───────────────────────────────────────────────────────────────

function testGoogleDocAudit() {
  Logger.log(googleDocAudit(AUDIT_MODE.AUDIT));
}

function convertPlainTextUrls() {
  googleDocAudit(AUDIT_MODE.CONVERT_TEXT_URLS);
  addUsageTrackingRecord('convertPlainTextUrls');
}

function highlightZoteroLinks() {
  googleDocAudit(AUDIT_MODE.SET_ZOTERO_HIGHLIGHT);
  addUsageTrackingRecord('highlightZoteroLinks');
}

function removeZoteroLinkHighlights() {
  googleDocAudit(AUDIT_MODE.REMOVE_ZOTERO_HIGHLIGHT);
  addUsageTrackingRecord('removeZoteroLinkHighlights');
}

function googleDocAudit(mode) {
  mode = mode || AUDIT_MODE.AUDIT;

  const result = {
    itemCslCitation: 0,
    cslBibliography: 0,
    zoteroUrl: 0,
    textUrl: 0,
    ourUrl: 0,
    otherUrl: 0,
    numberOfTabs: 0
  };

  const urlRegEx = refOpenDevEdLinksRegEx();

  // linksArray items:
  //   Docs:   { link, linkText, element, start, end }
  //   Slides: { link, linkText, slideRange }
  // slideRange is the TextRange object from Slides API, used for highlight modes.
  const linksArray = [];

  // ── 1. Collect links and scan text ─────────────────────────────────────────

  if (HOST_APP == 'docs') {
    const doc = DocumentApp.getActiveDocument();
    result.numberOfTabs = _countAllTabs(doc.getTabs());

    const body = doc.getBody();
    let convertCount = 0;

    if (mode === AUDIT_MODE.CONVERT_TEXT_URLS) {
      convertCount += _convertTextUrlsInElement(body);
    } else {
      _auditFindAllLinks(body, linksArray);
      _auditScanText(body, result);
    }

    const footnotes = doc.getFootnotes();
    for (let i in footnotes) {
      const footnoteContents = footnotes[i].getFootnoteContents();
      if (footnoteContents == null) continue;
      const numChildren = footnoteContents.getNumChildren();
      for (let j = 0; j < numChildren; j++) {
        const child = footnoteContents.getChild(j);
        if (mode === AUDIT_MODE.CONVERT_TEXT_URLS) {
          convertCount += _convertTextUrlsInElement(child);
        } else {
          _auditFindAllLinks(child, linksArray);
          _auditScanText(child, result);
        }
      }
    }

    // Mode 2 (Docs): save then alert
    if (mode === AUDIT_MODE.CONVERT_TEXT_URLS) {
      doc.saveAndClose();
      _alertConverted(convertCount);
      return;
    }

    // Modes 3 & 4 (Docs): apply, save, then alert
    if (mode === AUDIT_MODE.SET_ZOTERO_HIGHLIGHT || mode === AUDIT_MODE.REMOVE_ZOTERO_HIGHLIGHT) {
      const highlightCount = _applyZoteroHighlight(linksArray, mode);
      doc.saveAndClose();
      _alertHighlight(highlightCount, mode);
      return;
    }

  } else {
    // ── Slides ────────────────────────────────────────────────────────────────
    // Note: Slides auto-saves — no explicit save call needed.
    const slides = SlidesApp.getActivePresentation().getSlides();
    let convertCount = 0;

    for (let i in slides) {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;

        const shape = pageElement.asShape();
        const textRange = shape.getText();

        if (mode === AUDIT_MODE.CONVERT_TEXT_URLS) {
          convertCount += _convertTextUrlsInSlideShape(shape);
        } else {
          _auditScanSlideTextRange(textRange, result);

          // Collect links, storing slideRange for highlight modes
          const runs = textRange.getRuns();
          for (let j in runs) {
            const links = runs[j].getLinks();
            for (let m in links) {
              const url = links[m].getTextStyle().getLink().getUrl();
              linksArray.push({
                link: url,
                linkText: runs[j].asString(),
                slideRange: links[m] // TextRange — used for highlight modes
              });
            }
          }
        }
      });
    }

    // Mode 2 (Slides): alert (Slides auto-saves)
    if (mode === AUDIT_MODE.CONVERT_TEXT_URLS) {
      _alertConverted(convertCount);
      return;
    }

    // Modes 3 & 4 (Slides): apply then alert (Slides auto-saves)
    if (mode === AUDIT_MODE.SET_ZOTERO_HIGHLIGHT || mode === AUDIT_MODE.REMOVE_ZOTERO_HIGHLIGHT) {
      const highlightCount = _applyZoteroHighlight(linksArray, mode);
      _alertHighlight(highlightCount, mode);
      return;
    }
  }

  // ── 2. Mode 1: classify links and return audit result ──────────────────────

  for (let i in linksArray) {
    const url = linksArray[i].link;
    const linkText = linksArray[i].linkText || '';
    if (!url) continue;

    if (/^https?:\/\/www\.zotero\.org\/(google-docs|styles)\//.test(url)) {
      const isZoteroFieldCode =
        linkText.indexOf('ITEM CSL_CITATION') !== -1 ||
        linkText.indexOf('CSL_BIBLIOGRAPHY') !== -1 ||
        linkText.indexOf('DOCUMENT_PREFERENCES') !== -1;
      if (!isZoteroFieldCode) {
        result.zoteroUrl++;
      }
    } else if (urlRegEx.test(url)) {
      result.ourUrl++;
    } else {
      result.otherUrl++;
    }
  }

  return result;
}

// ── Shared action helpers ─────────────────────────────────────────────────────

/**
 * Applies or removes yellow highlight on Zotero links in linksArray.
 * Works for both Docs (uses element/start/end) and Slides (uses slideRange).
 * Returns the number of links affected.
 */
function _applyZoteroHighlight(linksArray, mode) {
  const YELLOW = '#FFFF00';
  let count = 0;

  for (let i in linksArray) {
    const item = linksArray[i];
    if (!item.link) continue;
    if (!/^https?:\/\/www\.zotero\.org\/(google-docs|styles)\//.test(item.link)) continue;

    const isZoteroFieldCode =
      (item.linkText || '').indexOf('ITEM CSL_CITATION') !== -1 ||
      (item.linkText || '').indexOf('CSL_BIBLIOGRAPHY') !== -1 ||
      (item.linkText || '').indexOf('DOCUMENT_PREFERENCES') !== -1;
    if (isZoteroFieldCode) continue;

    const color = mode === AUDIT_MODE.SET_ZOTERO_HIGHLIGHT ? YELLOW : null;

    if (item.element) {
      // Docs
      item.element.setBackgroundColor(item.start, item.end, color);
    } else if (item.slideRange) {
      // Slides
      if (color) {
        item.slideRange.getTextStyle().setBackgroundColor(color);
      } else {
        item.slideRange.getTextStyle().setBackgroundColorTransparent();
      }
    }

    count++;
  }

  return count;
}

/**
 * Shows an alert after Convert Text URLs to Hyperlinks completes.
 */
function _alertConverted(count) {
  if (count === 0) {
    getUi().alert(
      'Convert Text URLs into Hyperlinks',
      'No plain-text URLs were found.',
      getUi().ButtonSet.OK
    );
    return;
  }
  const noun = count === 1 ? 'plain-text URL was' : 'plain-text URLs were';
  getUi().alert(
    'Convert Text URLs into Hyperlinks',
    count + ' ' + noun + ' converted into hyperlinks.',
    getUi().ButtonSet.OK
  );
}

/**
 * Shows an alert after a highlight mode completes.
 */
function _alertHighlight(count, mode) {
  const title = mode === AUDIT_MODE.SET_ZOTERO_HIGHLIGHT
    ? 'Highlight Zotero links'
    : 'Remove Zotero link highlights';

  if (count === 0) {
    getUi().alert(title, 'No Zotero links were found.', getUi().ButtonSet.OK);
    return;
  }

  const message = mode === AUDIT_MODE.SET_ZOTERO_HIGHLIGHT
    ? count + ' ' + (count === 1 ? 'Zotero link was' : 'Zotero links were') + ' highlighted.\n\nTo undo the highlighting, run bZotBib → Additional functions → Remove Zotero link highlights or simply press Ctrl+Z (Cmd+Z on Mac) after closing this alert.'
    : count + ' ' + (count === 1 ? 'Zotero link highlight was' : 'Zotero link highlights were') + ' removed.';
  getUi().alert(title, message, getUi().ButtonSet.OK);
}

// ── Docs helpers ──────────────────────────────────────────────────────────────

/**
 * Recursively walks a Docs element tree and pushes every hyperlink into linksArray.
 * Stores element reference and character offsets for use by highlight modes.
 */
function _auditFindAllLinks(element, linksArray) {
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {
    const indices = element.getTextAttributeIndices();
    const text = element.getText();
    for (let i = 0; i < indices.length; i++) {
      const partAttributes = element.getAttributes(indices[i]);
      if (partAttributes.LINK_URL) {
        const start = indices[i];
        const end = (i == indices.length - 1) ? text.length - 1 : indices[i + 1] - 1;
        linksArray.push({
          link: partAttributes.LINK_URL,
          linkText: text.substr(start, end - start + 1),
          element: element,
          start: start,
          end: end
        });
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        _auditFindAllLinks(element.getChild(i), linksArray);
      }
    }
  }
}

/**
 * Walks a Docs element tree and counts CSL markers and bare (unlinked) URLs in text.
 */
function _auditScanText(element, result) {
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {
    const fullText = element.getText();
    const indices = element.getTextAttributeIndices();

    for (let i = 0; i < indices.length; i++) {
      const start = indices[i];
      const end = (i == indices.length - 1) ? fullText.length : indices[i + 1];
      const segment = fullText.substring(start, end);
      const partAttributes = element.getAttributes(start);

      _auditCountCslInText(segment, result);

      if (!partAttributes.LINK_URL) {
        _auditCountBareUrlInText(segment, result);
      }
    }
  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        _auditScanText(element.getChild(i), result);
      }
    }
  }
}

/**
 * Recursively walks a Docs element tree and converts bare plain-text URLs
 * into proper hyperlinks. Processes in reverse order to preserve offsets.
 * Returns the number of URLs converted.
 */
function _convertTextUrlsInElement(element) {
  const elementType = String(element.getType());
  let count = 0;

  if (elementType == 'TEXT') {
    const fullText = element.getText();
    const indices = element.getTextAttributeIndices();
    const matches = [];

    for (let i = 0; i < indices.length; i++) {
      const segStart = indices[i];
      const segEnd = (i == indices.length - 1) ? fullText.length : indices[i + 1];
      const partAttributes = element.getAttributes(segStart);

      if (partAttributes.LINK_URL) continue;

      const segment = fullText.substring(segStart, segEnd);
      const urlPattern = /https?:\/\/\S+/g;
      let match;
      while ((match = urlPattern.exec(segment)) !== null) {
        matches.push({
          url: match[0],
          start: segStart + match.index,
          end: segStart + match.index + match[0].length - 1
        });
      }
    }

    for (let j = matches.length - 1; j >= 0; j--) {
      element.setLinkUrl(matches[j].start, matches[j].end, matches[j].url);
      count++;
    }

  } else {
    const arrayTypes = ['BODY_SECTION', 'PARAGRAPH', 'LIST_ITEM', 'TABLE', 'TABLE_ROW', 'TABLE_CELL'];
    if (arrayTypes.includes(elementType)) {
      const numChildren = element.getNumChildren();
      for (let i = 0; i < numChildren; i++) {
        count += _convertTextUrlsInElement(element.getChild(i));
      }
    }
  }

  return count;
}

// ── Slides helpers ────────────────────────────────────────────────────────────

/**
 * For Slides: counts CSL markers in full shape text, and bare URLs only in unlinked runs.
 */
function _auditScanSlideTextRange(textRange, result) {
  const fullText = textRange.asString();
  _auditCountCslInText(fullText, result);

  const runs = textRange.getRuns();
  for (let j in runs) {
    const links = runs[j].getLinks();
    if (!links || links.length === 0) {
      _auditCountBareUrlInText(runs[j].asString(), result);
    }
  }
}

/**
 * For Slides: converts bare plain-text URLs in a shape into proper hyperlinks.
 * Uses absolute character indices from getRuns() to set links via getRange().
 * In the Slides API, getRange(start, end) end index is exclusive.
 * Returns the number of URLs converted.
 */
function _convertTextUrlsInSlideShape(shape) {
  const textRange = shape.getText();
  const runs = textRange.getRuns();
  const matches = [];

  for (let j in runs) {
    const links = runs[j].getLinks();
    if (links && links.length > 0) continue; // already linked — skip

    const runText = runs[j].asString();
    const runStart = runs[j].getStartIndex();
    const urlPattern = /https?:\/\/\S+/g;
    let match;
    while ((match = urlPattern.exec(runText)) !== null) {
      matches.push({
        url: match[0],
        start: runStart + match.index,
        end: runStart + match.index + match[0].length // exclusive end for Slides getRange()
      });
    }
  }

  // Apply in reverse order to keep indices stable
  for (let j = matches.length - 1; j >= 0; j--) {
    textRange.getRange(matches[j].start, matches[j].end).getTextStyle().setLinkUrl(matches[j].url);
  }

  return matches.length;
}

// ── Shared text-counting helpers ──────────────────────────────────────────────

/**
 * Counts ITEM CSL_CITATION and CSL_BIBLIOGRAPHY occurrences in a text string.
 */
function _auditCountCslInText(text, result) {
  if (!text) return;
  const itemCslMatches = text.match(/ITEM CSL_CITATION/g);
  if (itemCslMatches) result.itemCslCitation += itemCslMatches.length;

  const cslBibMatches = text.match(/CSL_BIBLIOGRAPHY/g);
  if (cslBibMatches) result.cslBibliography += cslBibMatches.length;
}

/**
 * Counts tokens that start with http/https (bare, unlinked URLs).
 */
function _auditCountBareUrlInText(text, result) {
  if (!text) return;
  const tokens = text.split(/\s+/);
  for (let i in tokens) {
    if (/^https?:\/\/.+/.test(tokens[i])) {
      result.textUrl++;
    }
  }
}

// ── Docs tab helper ───────────────────────────────────────────────────────────

/**
 * Recursively counts all tabs including nested child tabs.
 */
function _countAllTabs(tabs) {
  let count = 0;
  for (let i in tabs) {
    count++;
    const childTabs = tabs[i].getChildTabs();
    if (childTabs && childTabs.length > 0) {
      count += _countAllTabs(childTabs);
    }
  }
  return count;
}
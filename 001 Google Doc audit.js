function testGoogleDocAudit() {
  Logger.log(googleDocAudit());
}

function googleDocAudit() {
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
  const linksArray = [];

  // ── 1. Collect all links and scan text ──────────────────────────────────────

  if (HOST_APP == 'docs') {
    const doc = DocumentApp.getActiveDocument();

    // Count all tabs including nested child tabs
    result.numberOfTabs = _countAllTabs(doc.getTabs());

    // Body
    const body = doc.getBody();
    _auditFindAllLinks(body, linksArray);
    _auditScanText(body, result);

    // Footnotes
    const footnotes = doc.getFootnotes();
    for (let i in footnotes) {
      const footnoteContents = footnotes[i].getFootnoteContents();
      if (footnoteContents == null) continue;
      const numChildren = footnoteContents.getNumChildren();
      for (let j = 0; j < numChildren; j++) {
        const child = footnoteContents.getChild(j);
        _auditFindAllLinks(child, linksArray);
        _auditScanText(child, result);
      }
    }

  } else {
    // Slides
    const slides = SlidesApp.getActivePresentation().getSlides();
    for (let i in slides) {
      slides[i].getPageElements().forEach(function (pageElement) {
        if (pageElement.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
          const textRange = pageElement.asShape().getText();

          // Scan text (CSL markers + unlinked bare URLs)
          _auditScanSlideTextRange(textRange, result);

          // Collect proper hyperlinks
          const runs = textRange.getRuns();
          for (let j in runs) {
            const links = runs[j].getLinks();
            for (let m in links) {
              const url = links[m].getTextStyle().getLink().getUrl();
              linksArray.push({ link: url, linkText: runs[j].asString() });
            }
          }
        }
      });
    }
  }

  // ── 2. Classify every collected link ────────────────────────────────────────

  for (let i in linksArray) {
    const url = linksArray[i].link;
    const linkText = linksArray[i].linkText;
    if (!url) continue;

    if (/^https?:\/\/www\.zotero\.org\/(google-docs|styles)\//.test(url)) {
      // if (/^https:\/\/www\.zotero\.org\/google-docs\//.test(url)) {
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

  // ── 3. Scan link *text* for bare http/https URLs (not proper Doc links) ─────
  // These are already counted during text scanning above (see _auditCountInText).

  return result;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Recursively walks a Docs element tree and pushes every hyperlink into linksArray.
 */
function _auditFindAllLinks(element, linksArray) {
  const elementType = String(element.getType());

  if (elementType == 'TEXT') {
    const indices = element.getTextAttributeIndices();
    const text = element.getText();
    for (let i = 0; i < indices.length; i++) {
      const partAttributes = element.getAttributes(indices[i]);
      if (partAttributes.LINK_URL) {
        const end = (i == indices.length - 1) ? text.length - 1 : indices[i + 1] - 1;
        linksArray.push({
          link: partAttributes.LINK_URL,
          linkText: text.substr(indices[i], end - indices[i] + 1)
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

      // Count CSL markers in every segment (linked or not)
      _auditCountCslInText(segment, result);

      // Only count bare URL if this segment is NOT a proper hyperlink
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
 * For Slides: counts CSL markers in full shape text, and bare URLs only in unlinked runs.
 */
function _auditScanSlideTextRange(textRange, result) {
  const fullText = textRange.asString();

  // CSL markers — scan full text
  _auditCountCslInText(fullText, result);

  // Bare URLs — only in runs that carry no link
  const runs = textRange.getRuns();
  for (let j in runs) {
    const links = runs[j].getLinks();
    if (!links || links.length === 0) {
      // This run has no hyperlink attached — check if it looks like a bare URL
      _auditCountBareUrlInText(runs[j].asString(), result);
    }
  }
}

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
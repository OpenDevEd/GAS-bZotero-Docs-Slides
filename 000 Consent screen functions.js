/**
 * ConsentScreen Wrapper Functions for bZotBib Menu
 * 
 * These functions are called when the add-on is in AuthMode.NONE (no authorization yet).
 * They check email preferences before executing the actual function.
 * 
 * Pattern:
 * function [originalFunctionName]ConsentScreen(){
 *   const result = checkEmailPreferences('[Menu Label]', '[originalFunctionName]');
 *   if (result.stopExecution) {
 *     return 0;
 *   }
 *   [originalFunctionName]();
 * }
 */


function showSidebarMenuConsentScreen() {
  const result = checkEmailPreferences('Insert/update bibliography', 'insertUpdateBibliography');
  if (result.stopExecution) {
    return 0;
  }
  showSidebarMenu();
}

// ============================================================================
// Main Menu Items
// ============================================================================

function insertUpdateBibliographyConsentScreen() {
  const result = checkEmailPreferences('Insert/update bibliography', 'insertUpdateBibliography');
  if (result.stopExecution) {
    return 0;
  }
  insertUpdateBibliography();
}

function bibliographySidebarConsentScreen() {
  const result = checkEmailPreferences('Bibliography sidebar', 'bibliographySidebar');
  if (result.stopExecution) {
    return 0;
  }
  bibliographySidebar();
}

function modifyLinkParenthesisConsentScreen() {
  const result = checkEmailPreferences('Author (Year) ↔ Author, Year', 'modifyLinkParenthesis');
  if (result.stopExecution) {
    return 0;
  }
  modifyLinkParenthesis();
}

function validateLinksConsentScreen() {
  const result = checkEmailPreferences('Update/validate document links', 'validateLinks');
  if (result.stopExecution) {
    return 0;
  }
  validateLinks();
}

function clearLinkMarkersConsentScreen() {
  const result = checkEmailPreferences('Clear validation markers', 'clearLinkMarkers');
  if (result.stopExecution) {
    return 0;
  }
  clearLinkMarkers();
}

function removeUnderlineFromHyperlinksConsentScreen() {
  const result = checkEmailPreferences('Remove underlines from hyperlinks', 'removeUnderlineFromHyperlinks');
  if (result.stopExecution) {
    return 0;
  }
  removeUnderlineFromHyperlinks();
}

// ============================================================================
// Configure and Publish Submenu
// ============================================================================

function prepareForPublishingConsentScreen() {
  const result = checkEmailPreferences('Prepare for publishing', 'prepareForPublishing');
  if (result.stopExecution) {
    return 0;
  }
  prepareForPublishing();
}

function collectCitedItemsConsentScreen() {
  const result = checkEmailPreferences('Collect cited items', 'collectCitedItems');
  if (result.stopExecution) {
    return 0;
  }
  collectCitedItems();
}

function addZoteroItemKeyConsentScreen() {
  const result = checkEmailPreferences('Add/Change Zotero item key for this doc', 'addZoteroItemKey');
  if (result.stopExecution) {
    return 0;
  }
  addZoteroItemKey();
}

function addZoteroCollectionKeyConsentScreen() {
  const result = checkEmailPreferences('Add/Change Zotero collection key for this doc', 'addZoteroCollectionKey');
  if (result.stopExecution) {
    return 0;
  }
  addZoteroCollectionKey();
}

function targetReferenceLinksConsentScreen() {
  const result = checkEmailPreferences('Toggle target', 'targetReferenceLinks');
  if (result.stopExecution) {
    return 0;
  }
  targetReferenceLinks();
}

function removeOpeninZoteroappConsentScreen() {
  const result = checkEmailPreferences('Remove openin=zoteroapp from hyperlinks', 'removeOpeninZoteroapp');
  if (result.stopExecution) {
    return 0;
  }
  removeOpeninZoteroapp();
}

function enterValidationSiteConsentScreen() {
  const result = checkEmailPreferences('Enter validation site', 'enterValidationSite');
  if (result.stopExecution) {
    return 0;
  }
  enterValidationSite();
}

// ============================================================================
// Additional Functions Submenu
// ============================================================================

function analyseKerkoLinksConsentScreen() {
  const result = checkEmailPreferences('Analyse Kerko links', 'analyseKerkoLinks');
  if (result.stopExecution) {
    return 0;
  }
  analyseKerkoLinks();
}

function analyseKerkoLinksV1ConsentScreen() {
  const result = checkEmailPreferences('Analyse Kerko links V1', 'analyseKerkoLinksV1');
  if (result.stopExecution) {
    return 0;
  }
  analyseKerkoLinksV1();
}

function validateLinksV1ConsentScreen() {
  const result = checkEmailPreferences('Update/validate document links V1', 'validateLinksV1');
  if (result.stopExecution) {
    return 0;
  }
  validateLinksV1();
}

function showItemKeysConsentScreen() {
  const result = checkEmailPreferences('Show item keys', 'showItemKeys');
  if (result.stopExecution) {
    return 0;
  }
  showItemKeys();
}

function validateLinksTestHelperConsentScreen() {
  const result = checkEmailPreferences('Show links & urls', 'validateLinksTestHelper');
  if (result.stopExecution) {
    return 0;
  }
  validateLinksTestHelper();
}

function zoteroTransferDocConsentScreen() {
  const result = checkEmailPreferences('Convert ZoteroTransfer markers to bZotBib', 'zoteroTransferDoc');
  if (result.stopExecution) {
    return 0;
  }
  zoteroTransferDoc();
}

function removeCountryMarkersConsentScreen() {
  const result = checkEmailPreferences('Remove country markers', 'removeCountryMarkers');
  if (result.stopExecution) {
    return 0;
  }
  removeCountryMarkers();
}

function applyVancouverStyleConsentScreen() {
  const result = checkEmailPreferences('Convert to numbered references (Vancouver)', 'applyVancouverStyle');
  if (result.stopExecution) {
    return 0;
  }
  applyVancouverStyle();
}

function applyAPA7StyleConsentScreen() {
  const result = checkEmailPreferences('Convert to text references (APA7)', 'applyAPA7Style');
  if (result.stopExecution) {
    return 0;
  }
  applyAPA7Style();
}

// ============================================================================
// Text Citation Functions Submenu
// ============================================================================

function packZoteroCallConsentScreen() {
  const result = checkEmailPreferences('zpack Turn Zotero text citations into links', 'packZoteroCall');
  if (result.stopExecution) {
    return 0;
  }
  packZoteroCall();
}

function unpackCombinedConsentScreen() {
  const result = checkEmailPreferences('zunpack Turn Zotero links into text', 'unpackCombined');
  if (result.stopExecution) {
    return 0;
  }
  unpackCombined();
}

function packZoteroSelectiveCallConsentScreen() {
  const result = checkEmailPreferences('zpacks Turn selected Zotero text citations into links', 'packZoteroSelectiveCall');
  if (result.stopExecution) {
    return 0;
  }
  packZoteroSelectiveCall();
}

function unpackCombinedWarningConsentScreen() {
  const result = checkEmailPreferences('zunpackWarning Turn Zotero links into text, with warnings', 'unpackCombinedWarning');
  if (result.stopExecution) {
    return 0;
  }
  unpackCombinedWarning();
}

// ============================================================================
// Format Text Citations (Nested Submenu)
// ============================================================================

function minifyCitationsConsentScreen() {
  const result = checkEmailPreferences('zminify - with small text', 'minifyCitations');
  if (result.stopExecution) {
    return 0;
  }
  minifyCitations();
}

function maxifyCitationsConsentScreen() {
  const result = checkEmailPreferences('zmaxify - with coloured text', 'maxifyCitations');
  if (result.stopExecution) {
    return 0;
  }
  maxifyCitations();
}

function unfyCitationsConsentScreen() {
  const result = checkEmailPreferences('zunfy - plain text', 'unfyCitations');
  if (result.stopExecution) {
    return 0;
  }
  unfyCitations();
}

function clearWarningMarkersConsentScreen() {
  const result = checkEmailPreferences('zclear warnings', 'clearWarningMarkers');
  if (result.stopExecution) {
    return 0;
  }
  clearWarningMarkers();
}
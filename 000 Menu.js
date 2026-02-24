/**
 * Universal bZotBib Menu Function
 * Generates both traditional Google Docs menu and data structure for Vue.js sidebar
 */

function universal_bZotBib_menu(e, returnType = 'menu') {
  let targetMenuString, kerkoValidationSite, zoteroItemKeyAction, zoteroCollectionKeyAction, opendevedUser = false;
  let targetRefLinks;

  const activeUser = Session.getEffectiveUser().getEmail();
  opendevedUser = getStyleValue('local_show_advanced_menu');

  let consentSuffix = '';
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    consentSuffix = 'ConsentScreen';
    targetMenuString = 'Toggle target';
    kerkoValidationSite = '<Enter validation site>';
    zoteroItemKeyAction = 'Add/change';
    zoteroCollectionKeyAction = 'Add/change';
    targetRefLinks = 'zotero';
  } else {
    kerkoValidationSite = getDocumentPropertyString('kerko_validation_site');
    if (kerkoValidationSite == null) {
      if (activeUser.search(/edtechhub.org/i) != -1) {
        kerkoValidationSite = 'https://docs.edtechhub.org/lib/';
      } else if (opendevedUser) {
        kerkoValidationSite = 'https://docs.opendeved.net/lib/';
      } else {
        kerkoValidationSite = '<Enter validation site>';
      }
    }

    targetRefLinks = getDocumentPropertyString('target_ref_links');
    if (targetRefLinks == 'kerko') {
      targetMenuString = 'Toggle target (current: Kerko)';
    } else {
      targetMenuString = 'Toggle target (current: Zotero)';
      targetRefLinks = 'zotero';
    }

    let currentZoteroCollectionKey = getDocumentPropertyString('zotero_collection_key');
    zoteroCollectionKeyAction = currentZoteroCollectionKey == null ? 'Add' : 'Change';

    let currentZoteroItemKey = getDocumentPropertyString('zotero_item');
    zoteroItemKeyAction = currentZoteroItemKey == null ? 'Add' : 'Change';
  }

  const where = ' via ' + targetRefLinks;
  const whereForest = targetRefLinks == 'kerko' ? ' via Forest API' : ' via zotero';

  // Build the menu structure as a data object
  const menuStructure = {
    title: 'bZotBib',
    items: [
      { type: 'item', label: 'Insert/update bibliography', functionName: 'insertUpdateBibliography' + consentSuffix },
      { type: 'item', label: 'Bibliography sidebar [experimental]', functionName: 'bibliographySidebar' + consentSuffix },
      { type: 'separator' },
      { type: 'item', label: 'Author (Year) ↔ Author, Year', functionName: 'modifyLinkParenthesis' + consentSuffix },
      { type: 'item', label: 'Update/validate document links' + whereForest, functionName: 'validateLinks' + consentSuffix },
      { type: 'item', label: 'Clear validation markers', functionName: 'clearLinkMarkers' + consentSuffix },
      { type: 'item', label: 'Convert text URLs into hyperlinks', functionName: 'convertPlainTextUrls' + consentSuffix },
      { type: 'item', label: 'Remove underlines from hyperlinks', functionName: 'removeUnderlineFromHyperlinks' + consentSuffix },
      { type: 'separator' },
      {
        type: 'submenu',
        label: 'Configure and publish',
        items: [
          { type: 'item', label: 'Prepare for publishing', functionName: 'prepareForPublishing' + consentSuffix },
          { type: 'item', label: 'Collect cited items', functionName: 'collectCitedItems' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: zoteroItemKeyAction + ' Zotero item key for this doc', functionName: 'addZoteroItemKey' + consentSuffix },
          { type: 'item', label: zoteroCollectionKeyAction + ' Zotero collection key for this doc', functionName: 'addZoteroCollectionKey' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: targetMenuString, functionName: 'targetReferenceLinks' + consentSuffix },
          { type: 'item', label: 'Remove openin=zoteroapp from hyperlinks', functionName: 'removeOpeninZoteroapp' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'Enter validation site', functionName: 'enterValidationSite' + consentSuffix }
        ]
      },
      {
        type: 'submenu',
        label: 'Additional functions',
        items: [
          { type: 'item', label: 'Analyse Kerko links', functionName: 'analyseKerkoLinks' + consentSuffix },
          { type: 'item', label: 'Analyse Kerko links V1', functionName: 'analyseKerkoLinksV1' + consentSuffix },
          { type: 'item', label: 'Update/validate document links' + where + ' V1', functionName: 'validateLinksV1' + consentSuffix },
          { type: 'item', label: 'Show item keys', functionName: 'showItemKeys' + consentSuffix },
          { type: 'item', label: 'Show links & urls', functionName: 'validateLinksTestHelper' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'Convert ZoteroTransfer markers to bZotBib', functionName: 'zoteroTransferDoc' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'Remove country markers (⇡Country: )', functionName: 'removeCountryMarkers' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'Convert to numbered references (\'Vancouver\')', functionName: 'applyVancouverStyle' + consentSuffix },
          { type: 'item', label: 'Convert to text references (\'APA7\')', functionName: 'applyAPA7Style' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'Highlight Zotero links', functionName: 'highlightZoteroLinks' + consentSuffix },
          { type: 'item', label: 'Remove Zotero link highlights', functionName: 'removeZoteroLinkHighlights' + consentSuffix }
        ]
      },
      { type: 'separator' },
      {
        type: 'submenu',
        label: 'Text citation functions',
        items: [
          { type: 'item', label: 'zpack Turn Zotero text citations into links', functionName: 'packZoteroCall' + consentSuffix },
          { type: 'item', label: 'zunpack Turn Zotero links into text', functionName: 'unpackCombined' + consentSuffix },
          { type: 'separator' },
          { type: 'item', label: 'zpacks Turn selected Zotero text citations into links', functionName: 'packZoteroSelectiveCall' + consentSuffix },
          { type: 'item', label: 'zunpackWarning Turn Zotero links into text, with warnings where citation text has changed', functionName: 'unpackCombinedWarning' + consentSuffix },
          { type: 'separator' },
          {
            type: 'submenu',
            label: 'Format text citations',
            items: [
              { type: 'item', label: 'zminify - with small text', functionName: 'minifyCitations' + consentSuffix },
              { type: 'item', label: 'zmaxify - with coloured text', functionName: 'maxifyCitations' + consentSuffix },
              { type: 'item', label: 'zunfy - plain text', functionName: 'unfyCitations' + consentSuffix },
              { type: 'separator' },
              { type: 'item', label: 'zclear warnings 《warning:...》and ❲ and ❳', functionName: 'clearWarningMarkers' + consentSuffix }
            ]
          }
        ]
      },
      {
        type: 'submenu',
        label: 'Settings',
        items: [
          { type: 'item', label: 'Email preferences', functionName: 'openEmailPreferencesSettings' }
        ]
      }
    ]
  };

  // Return based on the requested type
  if (returnType === 'data') {
    return menuStructure;
  } else {
    return buildUIMenu(menuStructure, consentSuffix);
  }
}


/**
 * Helper function to build Google Docs UI menu from menu structure
 * Handles recursive nested submenus
 */
function buildUIMenu(menuStructure, consentSuffix) {
  const ui = getUi();

  // Add sidebar launcher at the top of the menu
  menuStructure.items.unshift(
    {
      type: "item",
      label: "Open menu in sidebar 🚀",
      functionName: "showSidebarMenu" + consentSuffix
    },
    { type: 'separator' }
  );

  const menu = ui.createMenu(menuStructure.title);

  // Recursive helper function to build menu items at any depth
  function buildMenuItems(parentMenu, items) {
    items.forEach(item => {
      if (item.type === 'separator') {
        parentMenu.addSeparator();
      } else if (item.type === 'item') {
        // Handle items with parameters (nested object methods)
        if (item.functionParams && item.functionParams.length > 0) {
          parentMenu.addItem(item.label, item.functionName + '.' + item.functionParams[0] + '.run');
        } else {
          parentMenu.addItem(item.label, item.functionName);
        }
      } else if (item.type === 'submenu') {
        const submenu = ui.createMenu(item.label);
        // RECURSIVE CALL: Handle nested submenus
        buildMenuItems(submenu, item.items);
        parentMenu.addSubMenu(submenu);
      }
    });
  }

  // Build all menu items recursively
  buildMenuItems(menu, menuStructure.items);

  return menu;
}


/**
 * Show the sidebar menu
 */
function showSidebarMenu() {
  const ui = getUi();
  const template = HtmlService.createTemplateFromFile('000 Menu sidebar');
  const menuStructure = universal_bZotBib_menu(null, 'data');
  template.menuStructureJson = JSON.stringify(menuStructure);

  const html = template.evaluate()
    .setTitle(menuStructure.title + ' menu');

  ui.showSidebar(html);
}


/**
 * Updated onOpen function that uses the universal menu
 */
function onOpen(e) {
  const menu = universal_bZotBib_menu(e, 'menu');
  menu.addToUi();
}
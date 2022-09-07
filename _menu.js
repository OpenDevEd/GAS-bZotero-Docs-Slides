function onOpen(e) {
  let targetMenuString, kerkoValidationSite, zoteroItemKeyAction, zoteroCollectionKeyAction, opendevedUser = false;
  // https://developers.google.com/workspace/add-ons/concepts/editor-auth-lifecycle#the_complete_lifecycle
  let targetRefLinks;

  const activeUser = Session.getEffectiveUser().getEmail();
  opendevedUser = getStyleValue('local_show_advanced_menu');

  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    targetMenuString = 'Target: Zotero; change to Kerko';
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
      targetMenuString = 'Target: Kerko; change to Zotero';
    } else {
      targetMenuString = 'Target: Zotero; change to Kerko';
      targetRefLinks = 'zotero';
    }
    let currentZoteroCollectionKey = getDocumentPropertyString('zotero_collection_key');
    zoteroCollectionKeyAction = currentZoteroCollectionKey == null ? 'Add' : 'Change';

    let currentZoteroItemKey = getDocumentPropertyString('zotero_item');
    zoteroItemKeyAction = currentZoteroItemKey == null ? 'Add' : 'Change';
  }

  const ui = getUi();

  const menu = ui.createMenu('bZotero');
  const where = ' via ' + targetRefLinks;
  menu.addItem('Insert/update bibliography', 'insertUpdateBibliography');
  menu.addItem('Bibliography sidebar [experimental]', 'bibliographySidebar');
  menu.addSeparator();
  menu.addItem('Update/validate document links' + where, 'validateLinks');
  menu.addItem('Clear validation markers', 'clearLinkMarkers');
  menu.addItem('Remove underlines from hyperlinks', 'removeUnderlineFromHyperlinks');
  menu.addItem('Remove openin=zoteroapp from hyperlinks', 'removeOpeninZoteroapp');
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Configure and publish')
    .addItem('Prepare for publishing', 'prepareForPublishing')
    .addSeparator()
    .addItem(zoteroItemKeyAction + ' Zotero item key for this doc', 'addZoteroItemKey')
    .addItem(zoteroCollectionKeyAction + ' Zotero collection key for this doc', 'addZoteroCollectionKey')
    .addSeparator()
    .addItem(targetMenuString, 'targetReferenceLinks')
    .addItem('Enter validation site', 'enterValidationSite')
  );
  menu.addSubMenu(ui.createMenu('Additional functions')
    .addItem('Show item keys', 'showItemKeys')
    .addItem('Show links & urls', 'validateLinksTestHelper')
    .addSeparator()
    .addItem('Convert ZoteroTransfer markers to BZotero', 'zoteroTransferDoc')
    .addSeparator()
    .addItem('Remove country markers (⇡Country: )', 'removeCountryMarkers')
    // Remove for now:  
    //.addItem('znocountry highlight missing country info', 'highlightMissingCountryMarker')
    // Remove for now
    // .addItem('zsuper make superscripts for citations', 'makeSuperscriptsForCitations')
  );

  menu.addSeparator()
    .addSubMenu(ui.createMenu('Text citation functions')
      .addItem('zpack Turn Zotero text citations into links', 'packZoteroCall')
      .addItem('zunpack Turn Zotero links into text', 'unpackCombined')
      .addSeparator()
      .addItem('zpacks Turn selected Zotero text citations into links', 'packZoteroSelectiveCall')
      .addItem('zunpackWarning Turn Zotero links into text, with warnings where citation text has changed', 'unpackCombinedWarning')
      .addSeparator()
      .addSubMenu(ui.createMenu('Format text citations')
        .addItem('zminify - with small text', 'minifyCitations')
        .addItem('zmaxify - with coloured text', 'maxifyCitations')
        .addItem('zunfy - plain text', 'unfyCitations')
        .addSeparator()
        // clearWarningMarkers is new version of clearZwarnings
        .addItem('zclear warnings 《warning:...》and ❲ and ❳', 'clearWarningMarkers')
      )
    )

  menu.addToUi();
}


function minifyCitations() {
  if (HOST_APP == 'docs') {
    minifyCitationsDocs();
  } else {
    packZoteroSlides('minifyCitations');
  }
};

function maxifyCitations() {
  if (HOST_APP == 'docs') {
    maxifyCitationsDocs();
  } else {
    packZoteroSlides('maxifyCitations');
  }
};

function unfyCitations() {
  if (HOST_APP == 'docs') {
    unfyCitationsDocs();
  } else {
    packZoteroSlides('unfyCitations');
  }
};

function unpackCombined() {
  if (HOST_APP == 'docs') {
    restoreZoteroLinks();
    zoteroUnpackCall(false);
  } else {
    unpackZoteroSlides(false);
  }
};

function unpackCombinedWarning() {
  if (HOST_APP == 'docs') {
    restoreZoteroLinks();
    zoteroUnpackCall(true);
  } else {
    unpackZoteroSlides(true);
  }
};

function packZoteroCall() {
  if (HOST_APP == 'docs') {
    zoteroPackUnpack(true, false, null);
  } else {
    packZoteroSlides('packZotero');
  }
};

function zoteroUnpackCall(warning) {
  zoteroPackUnpack(false, false, warning);
};

function packZoteroSelectiveCall() {
  if (HOST_APP == 'docs') {
    zoteroPackUnpack(true, true, null);
  } else {
    packZoteroSelectiveCallSlides();
  }
};
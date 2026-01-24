// For testing
function deleteDocumentPropertyStringZoteroItemKey() {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty('zotero_item');
  documentProperties.deleteProperty('zotero_collection_key');
  documentProperties.deleteProperty('kerko_validation_site');
}

function setDocumentPropertyString(property_name, value) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(property_name, value);
}

function getDocumentPropertyString(property_name) {
  var documentProperties = PropertiesService.getDocumentProperties();
  var value = documentProperties.getProperty(property_name);
  return value;
}

function enterValidationSite(errorText = '', trackUsage = true) {
  // UserProperty -> DocumentProperty Update
  const kerkoValidationSite = getDocumentPropertyString('kerko_validation_site');
  if (kerkoValidationSite != null) {
    errorText += '\nCurrent Validation Site:\n' + kerkoValidationSite + '\n';
  }

  const ui = getUi();
  const response = ui.prompt('Validation site', errorText + '\nPlease enter the validation site, e.g. in the form of https://docs.edtechhub.org/lib/', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let url = response.getResponseText().trim();
    if (url.charAt(url.length - 1) != '/' && url != '-') url += '/';
    const urlRegEx = new RegExp('https?://', 'i');
    if (url.search(urlRegEx) == 0 || url == '-') {
      // UserProperty -> DocumentProperty Update
      setDocumentPropertyString('kerko_validation_site', url);
      updateStyle();
      onOpen();
    } else {
      enterValidationSite('Error!\n', false);
    }
  }
  if (trackUsage === true) {
    addUsageTrackingRecord('enterValidationSite');
  }
}

function scanForItemKey(targetRefLinks) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();
  let rangeElement, tableText, libLink, result, tablesLimit = 2;
  let foundFlag = false;
  // How many tables to check
  if (tables.length < 2) {
    tablesLimit = tables.length;
  }
  if (tablesLimit > 0) {
    const domains = getCustomDomainsArray();
    for (let i = 0; i < tablesLimit; i++) {
      // Example of regEx in the line below: '(docs.edtechhub.org/lib|docs.opendeved.net/lib|maths.educationevidence.io/lib)/[^/]+'
      rangeElement = tables[i].findText('(' + domains.join('|') + ')/[^/]+');
      if (rangeElement) {
        tableText = rangeElement.getElement().asText().getText();
        // Example of regEx in the line below: /docs.edtechhub.org\/lib\/[a-zA-Z0-9]+|docs.opendeved.net\/lib\/[a-zA-Z0-9]+|maths.educationevidence.io\/lib\/[a-zA-Z0-9]+/
        libLink = new RegExp(domains.join('/[a-zA-Z0-9]+|') + '/[a-zA-Z0-9]+').exec(tableText);
        if (libLink != null) {
          result = detectZoteroItemKeyType('https://' + libLink);
          if (result.status == 'error') {
            result = addZoteroItemKey('', false, false, targetRefLinks, false);
            return result;
          }
          foundFlag = true;
        }
        break;
      }
    }
  }

  if (!foundFlag) {
    result = addZoteroItemKey('', false, false, targetRefLinks, false);
  }
  return result;
}

function targetReferenceLinks() {
  const ui = getUi();
  let result;
  let currentTargetRefLinks = getDocumentPropertyString('target_ref_links');
  // UserProperty -> DocumentProperty Update
  const kerkoValidationSite = getDocumentPropertyString('kerko_validation_site');


  if (currentTargetRefLinks == null) {
    // UserProperty -> DocumentProperty Update
    setDocumentPropertyString('target_ref_links', 'zotero');
    currentTargetRefLinks = 'zotero';
  }

  const proposedTargetRefLinks = (currentTargetRefLinks == 'zotero') ? 'kerko' : 'zotero';

  let currentZoteroItemKey = getDocumentPropertyString('zotero_item');
  if (currentZoteroItemKey == null) {

    if (HOST_APP == 'docs') {
      result = scanForItemKey(proposedTargetRefLinks);
    } else {
      result = scanForItemKeySlides(proposedTargetRefLinks);
    }
    if (result.status == 'error') {
      ui.alert(result.message);
      return 0;
    }
    if (result.status == 'stop') {
      return 0;
    }
  }


  const response = ui.alert('Advanced option: Target for reference links',
    `Chose whether the reference links in this document should open in Zotero or on ${kerkoValidationSite} (kerko). 
  
If you are preparing this document for publication, please set target to ‘kerko’. If this is currently an internal document, please set target to 'zotero'.
  
Current target: ‘${currentTargetRefLinks}’.

Do you wish to change the target to ‘${proposedTargetRefLinks}’?`, ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    setDocumentPropertyString('target_ref_links', proposedTargetRefLinks);
    onOpen();
    const response2 = ui.alert('You have changed the target for reference links', 'The links will now be reconfigured.', ui.ButtonSet.OK_CANCEL);
    if (response2 == ui.Button.OK) {
      const newForestAPI = userCanCallForestAPI();
      validateLinks(false, true, false, false, newForestAPI, false);
    }
  }
  addUsageTrackingRecord('targetReferenceLinks');
}


function addZoteroCollectionKey(errorText = '', optional = false, bibliography = false, trackUsage = true) {
  // Logger.log('Function addZoteroCollectionKey ' + ' optional = ' + optional + ' bibliography = ' + bibliography);
  let currentZoteroCollectionKey = getDocumentPropertyString('zotero_collection_key');
  if (currentZoteroCollectionKey == null) {
    currentZoteroCollectionKey = '(you haven\'t set up Zotero collection key)';
  }

  const optionalText = optional ? "Optional: " : "";
  const bibText = bibliography ? "If you do not have the key, press ‘cancel’ to proceed to inserting the bibliography anyway.\n\n" : "";

  const ui = getUi();
  const response = ui.prompt(optionalText + 'Zotero Collection Key', errorText + 'Current collection key: ' + currentZoteroCollectionKey + '.\n\nWhile this document is internal, the reference links open the corresponding item in Zotero. If you have a Zotero collection for your items, the reference links can open the Zotero items in the context of the Zotero collection. To enable this, please enter the Zotero collection key here.\n\nYou can obtain the Zotero collection key using Zutilo; In the Zutlio preferences (“User Interface”), enable ”‘Zotero collection menu -> Get select collection link”; then right click on the collection, copy the collection link, and paste it below.\n\n' + bibText + 'Please enter Zotero collection key as a zotero://select-link.', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const zoteroCollectionItem = response.getResponseText();
    const zoteroCollectionRegEx = new RegExp('zotero://select/groups/[0-9]+/collections/[^/]+', 'i');
    if (zoteroCollectionRegEx.test(zoteroCollectionItem)) {
      setDocumentPropertyString('zotero_collection_key', zoteroCollectionItem);
    } else {
      addZoteroCollectionKey('Error! Wrong Zotero collection key.\n', optional, bibliography, false);
    }
  }
  onOpen();
  if (trackUsage === true) {
    addUsageTrackingRecord('addZoteroCollectionKey');
  }
}

function addZoteroItemKey(errorText = '', optional = false, bibliography = false, targetRefLinks, trackUsage = true) {
  if (trackUsage === true) {
    addUsageTrackingRecord('addZoteroItemKey');
  }
  const currentZoteroItemKey = getDocumentPropertyString('zotero_item');
  if (currentZoteroItemKey != null) {
    errorText += '\nCurrent Zotero Item Key:\n' + currentZoteroItemKey + '\n';
  }
  const optionalText = optional ? "Optional: " : "";
  const bibText = bibliography ? "\nIf you do not have the key, press ‘cancel’ to proceed to inserting the bibliography anyway." : "";

  const ui = getUi();
  const response = ui.prompt(optionalText + 'Enter the Zotero item key', errorText + '\nPlease enter Zotero item key as a zotero://select link (which you can copy with Zutilo in the Zotero app) or use the url to this item in the evidence library. If this document has a document citation (e.g., on the second page), you can find the link there.\n\nPlease enter the link and click ok.' + bibText, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    let zotero_item = response.getResponseText();

    let res = detectZoteroItemKeyType(zotero_item);
    if (res.status == 'error') {
      res = addZoteroItemKey(res.message, optional, bibliography, targetRefLinks, false);
    }
    //return res;
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
    if (targetRefLinks == 'kerko') {
      ui.alert('For validation against the evidence library a Zotero item key is required.', 'Please obtain the key from Zotero or your reference, and try again.', ui.ButtonSet.OK);
      return { status: 'stop' };
    }
  }
  return { status: 'ok' };
}

// addZoteroItemKey, scanForItemKey use the function
function detectZoteroItemKeyType(zotero_item) {
  zotero_item = zotero_item.trim();
  const zoteroItemRegEx = new RegExp('zotero://select/groups/[0-9]+/items/[^/]+', 'i');

  // Previous
  // const itemKey = zotero_item.split('/')[4];
  // New 2021-04-24
  let itemKey;
  const checkId = /id=[a-zA-Z0-9]+&?/.exec(zotero_item);

  if (checkId == null) {
    itemKey = zotero_item.split('/')[4];
  } else {
    if (checkId[0].charAt(checkId[0].length - 1) == '&') {
      checkId[0] = checkId[0].slice(0, -1);
    }
    itemKey = checkId[0].slice(3);
  }
  // End. New 2021-04-24

  let setDP = false;

  if (/^[0-9]+:[A-Z0-9]+$/.test(zotero_item)) {
    groupIdItemKey = zotero_item.split(':');
    zotero_item = 'zotero://select/groups/' + groupIdItemKey[0] + '/items/' + groupIdItemKey[1];
    setDP = true;
  } else if (zoteroItemRegEx.test(zotero_item)) {
    setDP = true;
  } else {
    const domains = getCustomDomainsArray();
    for (let i in domains) {
      const textRegEx = 'https://' + domains[i] + '/[^/]+';
      const regEx = new RegExp(textRegEx, 'i');
      if (regEx.test(zotero_item)) {
        const thisStyle = Object.keys(styles).find(key => {
          return styles[key]['kerkoValidationSite'] && styles[key]['kerkoValidationSite'].includes(domains[i]);
        });
        if (!thisStyle) {
          throw new Error('Style containing kerkoValidationSite ' + domains[i] + ' wasn\'t found.\nAsk admin to check object styles in config_public.gs file.');
        }
        const groupId = styles[thisStyle]['group_id'];
        if (!groupId) {
          throw new Error('Key group_id wasn\'t found in settings of ' + domains[i] + '\nAsk admin to check object styles in config_public.gs file.');
        }
        zotero_item = 'zotero://select/groups/' + groupId + '/items/' + itemKey;
        setDP = true;
        break;
      }
    }
  }

  if (setDP) {
    setDocumentPropertyString('zotero_item', zotero_item);
    onOpen();
    return { status: 'ok' };
  } else {
    //addZoteroItemKey('Error! Wrong Zotero Item Key\n', optional, bibliography);
    return { status: 'error', message: 'Error! Wrong Zotero Item Key ' + zotero_item + '\n' };
  }

}


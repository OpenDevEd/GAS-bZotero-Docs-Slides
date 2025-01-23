const styles = {
  "default":
  {
    "name": "ZoteroDocs (default)",
    "default_everybody": true,
    "permitted_libraries": [],
    "local_show_advanced_menu": false,
    "AUTO_PROMPT_COLLECTION": false,
    "ORPHANED_LINK_MARK": "<ORPHANED_LINK>",
    "URL_CHANGED_LINK_MARK": "<URL_CHANGED_LINK>",
    "BROKEN_LINK_MARK": "<BROKEN_LINK>",
    "NORMAL_LINK_MARK": "<NORMAL_LINK>",
    "NORMAL_REDIRECT_LINK_MARK": "<NORMAL_REDIRECT_LINK>",
    //  "UNKNOWN_LIBRARY_MARK": "<UNKNOWN_LIBRARY>",

    // New Forest API markers
    "VALID_LINK_MARK": "<VALID_LINK>",
    "VALID_AMBIGUOUS_LINK_MARK": "<VALID_AMBIGUOUS_LINK>",
    "REDIRECT_LINK_MARK": "<VALID_REDIRECT_LINK>",
    "REDIRECT_AMBIGUOUS_LINK_MARK": "<REDIRECT_AMBIGUOUS_LINK>",
    "IMPORTABLE_LINK_MARK": "<IMPORTABLE_LINK>",
    "IMPORTABLE_AMBIGUOUS_LINK_MARK": "<IMPORTABLE_AMBIGUOUS_LINK>",
    "IMPORTABLE_REDIRECT_LINK_MARK": "<IMPORTABLE_REDIRECT_LINK>",
    "UNKNOWN_LINK_MARK": "<UNKNOWN_LINK>",
    "INVALID_SYNTAX_LINK_MARK": "<INVALID_SYNTAX_LINK>",
    // End. New Forest API markers

    "TEXT_TO_DETECT_START_BIB": "⁅bibliography:start⁆",
    "TEXT_TO_DETECT_END_BIB": "⁅bibliography:end⁆",
    "LINK_MARK_STYLE_FOREGROUND_COLOR": "#ff0000",
    "LINK_MARK_STYLE_BACKGROUND_COLOR": "#ffffff",
    "LINK_MARK_STYLE_BOLD": true,
    "kerkoValidationSite": null,
    "group_id": null
  },
  "opendeved":
  {
    "name": "ZoteroDocs (OpenDevEd)",
    "default_everybody": false,
    "default_for": "opendeved.net",
    //  "permitted_libraries": ["2129771", "2405685", "2486141", "2447227"],
    "local_show_advanced_menu": true,
    "kerkoValidationSite": 'https://docs.opendeved.net/lib/',
    "group_id": "2129771",
  },
  "edtechhub":
  {
    "name": "ZoteroDocs (EdTech Hub)",
    "default_everybody": false,
    "default_for": "edtechhub.org",
    //  "permitted_libraries": ["2405685", "2339240", "2129771"],
    "LINK_MARK_STYLE_BACKGROUND_COLOR": "#dddddd",
    "kerkoValidationSite": 'https://docs.edtechhub.org/lib/',
    "group_id": "2405685"
  },
  "educationevidence":
  {
    "name": "ZoteroDocs (maths.educationevidence.io)",
    "default_everybody": false,
    "default_for": "maths.educationevidence.io",
    "LINK_MARK_STYLE_BACKGROUND_COLOR": "#dddddd",
    "kerkoValidationSite": 'https://maths.educationevidence.io/lib/',
    "group_id": "5168324"
  },
  "unlockingdata":
  {
    "name": "bZotBibDocs (docs.unlockingdata.africa)",
    "default_everybody": false,
    "default_for": "docs.unlockingdata.africa",
    "LINK_MARK_STYLE_BACKGROUND_COLOR": "#dddddd",
    "kerkoValidationSite": 'https://docs.unlockingdata.africa/lib/',
    "group_id": "5578897"
  }
};

// Gets styleName based on email of active user or owner's domain
// validateLinks, getDefaultStyle use the function
function detectDefaultForStyle(emailOrDomain) {
  for (let styleName in styles) {
    if (styles[styleName]['default_for'] && emailOrDomain.search(new RegExp(styles[styleName]['default_for'], 'i')) != -1) {
      return styleName;
    }
  }
  return null;
}

// Gets default style based on user's domain
function getDefaultStyle() {
  const activeUser = Session.getEffectiveUser().getEmail();

  const defaultForStyle = detectDefaultForStyle(activeUser);
  if (defaultForStyle != null) {
    return defaultForStyle;
  }

  // If user's domain isn't presented in styles object, find style that is suitable for everybody
  return getDefaultEverybodyStyleName();
}

// The variable contains style of current doc
// Initially, it is default style but function updateStyle can change it to another style based on DocumentProperties kerko_validation_site
let ACTIVE_STYLE = getDefaultStyle();
//Logger.log('Test 1' + ACTIVE_STYLE);

let TEXT_TO_DETECT_START_BIB, TEXT_TO_DETECT_END_BIB;

const LINK_MARK_STYLE_NEW = new Object();



// Changes a value of ACTIVE_STYLE to style that is default for DocumentProperties kerko_validation_site
function updateStyle() {
  HOST_APP = 'docs';
  // try {
  //   const ui = DocumentApp.getUi();
  //   HOST_APP = 'docs';
  // }
  // catch (error) {
  //   //Logger.log(error);
  // }
  // if (HOST_APP == null) {
  //   try {
  //     const ui = SlidesApp.getUi();
  //     HOST_APP = 'slides';
  //   }
  //   catch (error) {
  //     //Logger.log(error);
  //   }
  // }

  try {
    const kerkoValidationSite = getDocumentPropertyString('kerko_validation_site');
    //Logger.log('kerkoValidationSite ' + kerkoValidationSite);
    if (kerkoValidationSite != null) {
      for (let styleName in styles) {
        if (styles[styleName]['default_for'] && kerkoValidationSite.search(new RegExp(styles[styleName]['default_for'], 'i')) != -1) {
          ACTIVE_STYLE = styleName;
          //Logger.log('Test 2' + ACTIVE_STYLE);
          break;
        }
      }
    }
  }
  catch (error) {
    Logger.log(error);
  }

  //PERMITTED_LIBRARIES = getStyleValue('permitted_libraries');
  // AUTO_PROMPT_COLLECTION = getStyleValue('AUTO_PROMPT_COLLECTION');
  //ORPHANED_LINK_MARK = getStyleValue('ORPHANED_LINK_MARK');
  //URL_CHANGED_LINK_MARK = getStyleValue('URL_CHANGED_LINK_MARK');

  // BROKEN_LINK_MARK = getStyleValue('BROKEN_LINK_MARK');
  // NORMAL_LINK_MARK = getStyleValue('NORMAL_LINK_MARK');
  // NORMAL_REDIRECT_LINK_MARK = getStyleValue('NORMAL_REDIRECT_LINK_MARK');
  // REDIRECT_AMBIGUOUS_LINK_MARK = getStyleValue('REDIRECT_AMBIGUOUS_LINK_MARK');


  //UNKNOWN_LIBRARY_MARK = getStyleValue('UNKNOWN_LIBRARY_MARK');
  // TEXT_TO_DETECT_START_BIB = getStyleValue('TEXT_TO_DETECT_START_BIB');
  // TEXT_TO_DETECT_END_BIB = getStyleValue('TEXT_TO_DETECT_END_BIB');

  // LINK_MARK_STYLE_FOREGROUND_COLOR = getStyleValue('LINK_MARK_STYLE_FOREGROUND_COLOR');
  // LINK_MARK_STYLE_BACKGROUND_COLOR = getStyleValue('LINK_MARK_STYLE_BACKGROUND_COLOR');
  // LINK_MARK_STYLE_BOLD = getStyleValue('LINK_MARK_STYLE_BOLD');

  // if (HOST_APP == 'docs') {
  //   LINK_MARK_STYLE_NEW[DocumentApp.Attribute.FOREGROUND_COLOR] = LINK_MARK_STYLE_FOREGROUND_COLOR;
  //   LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BACKGROUND_COLOR] = LINK_MARK_STYLE_BACKGROUND_COLOR;
  //   LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BOLD] = LINK_MARK_STYLE_BOLD;
  // }
}

updateStyle();
//Logger.log('Test 3' + ACTIVE_STYLE);

function prepareBibMarkers() {
  TEXT_TO_DETECT_START_BIB = getStyleValue('TEXT_TO_DETECT_START_BIB');
  TEXT_TO_DETECT_END_BIB = getStyleValue('TEXT_TO_DETECT_END_BIB');
}

const LINK_MARK_OBJ = new Object();
const LINK_MARK_STYLE_OBJ = new Object();
function collectLinkMarks() {
  LINK_MARK_OBJ['ORPHANED_LINK_MARK'] = getStyleValue('ORPHANED_LINK_MARK');
  LINK_MARK_OBJ['URL_CHANGED_LINK_MARK'] = getStyleValue('URL_CHANGED_LINK_MARK');

  LINK_MARK_OBJ['BROKEN_LINK_MARK'] = getStyleValue('BROKEN_LINK_MARK');
  LINK_MARK_OBJ['NORMAL_LINK_MARK'] = getStyleValue('NORMAL_LINK_MARK');
  LINK_MARK_OBJ['NORMAL_REDIRECT_LINK_MARK'] = getStyleValue('NORMAL_REDIRECT_LINK_MARK');

  LINK_MARK_OBJ['VALID_LINK_MARK'] = getStyleValue('VALID_LINK_MARK');
  LINK_MARK_OBJ['VALID_AMBIGUOUS_LINK_MARK'] = getStyleValue('VALID_AMBIGUOUS_LINK_MARK');
  LINK_MARK_OBJ['REDIRECT_LINK_MARK'] = getStyleValue('REDIRECT_LINK_MARK');
  LINK_MARK_OBJ['REDIRECT_AMBIGUOUS_LINK_MARK'] = getStyleValue('REDIRECT_AMBIGUOUS_LINK_MARK');
  LINK_MARK_OBJ['IMPORTABLE_LINK_MARK'] = getStyleValue('IMPORTABLE_LINK_MARK');
  LINK_MARK_OBJ['IMPORTABLE_AMBIGUOUS_LINK_MARK'] = getStyleValue('IMPORTABLE_AMBIGUOUS_LINK_MARK');
  LINK_MARK_OBJ['IMPORTABLE_REDIRECT_LINK_MARK'] = getStyleValue('IMPORTABLE_REDIRECT_LINK_MARK');

  LINK_MARK_OBJ['UNKNOWN_LINK_MARK'] = getStyleValue('UNKNOWN_LINK_MARK');
  LINK_MARK_OBJ['INVALID_SYNTAX_LINK_MARK'] = getStyleValue('INVALID_SYNTAX_LINK_MARK');

  LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_FOREGROUND_COLOR'] = getStyleValue('LINK_MARK_STYLE_FOREGROUND_COLOR');
  LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BACKGROUND_COLOR'] = getStyleValue('LINK_MARK_STYLE_BACKGROUND_COLOR');
  LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BOLD'] = getStyleValue('LINK_MARK_STYLE_BOLD');

  if (HOST_APP == 'docs') {
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.FOREGROUND_COLOR] = LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_FOREGROUND_COLOR'];
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BACKGROUND_COLOR] = LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BACKGROUND_COLOR'];
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BOLD] = LINK_MARK_STYLE_OBJ['LINK_MARK_STYLE_BOLD'];
  }
}


// Finds style that is suitable for everybody
function getDefaultEverybodyStyleName() {
  for (let styleName in styles) {
    if (styles[styleName]['default_everybody'] === true) {
      return styleName;
    }
  }
}

function getStyleValue(property) {
  if (styles[ACTIVE_STYLE].hasOwnProperty(property)) {
    return styles[ACTIVE_STYLE][property];
  } else {
    const defaultStyle = getDefaultEverybodyStyleName();
    return styles[defaultStyle][property];
  }
}

function getUi() {
  return HOST_APP == 'docs' ? DocumentApp.getUi() : SlidesApp.getUi();
}

function getEditors() {
  if (HOST_APP == 'docs') {
    return DocumentApp.getActiveDocument().getEditors();
  } else {
    return SlidesApp.getActivePresentation().getEditors();
  }
}

// Retrieves domain and folder from kerkoValidationSite
// For example, gets https://docs.opendeved.net/lib/, returns docs.opendeved.net/lib
function retrieveDomain(kerkoValidationSite) {
  const result = new RegExp('https?://(.+)/?', 'i').exec(kerkoValidationSite);
  if (!result?.[1]) {
    throw new Error('Unexpected kerkoValidationSite ' + kerkoValidationSite + '\nAsk admin to check object styles in config_public.gs file.');
  } else {
    return result[1].endsWith('/') ? result[1].slice(0, -1) : result[1];
  }
}

let CUSTOM_DOMAINS_ARRAY;
// Retrieves kerko validation sites from the object style
// Returns array, something like [docs.opendeved.net/lib, docs.edtechhub.org/lib, maths.educationevidence.io/lib]
function getCustomDomainsArray() {
  if (CUSTOM_DOMAINS_ARRAY != null){
    return CUSTOM_DOMAINS_ARRAY;
  }  
  const domains = [];
  for (let styleName in styles) {
    if (styles[styleName]['kerkoValidationSite']) {
      domains.push(retrieveDomain(styles[styleName]['kerkoValidationSite']));
    }
  }
  CUSTOM_DOMAINS_ARRAY = domains;
  return domains;
}

let REF_OPENDEVED_LINKS_REGEX;
function refOpenDevEdLinksRegEx() {
  let customDomainsPart = '';
  const domains = getCustomDomainsArray();
  if (domains.length > 0) {
    customDomainsPart = '|https?://(' + domains.join('|') + ')(/[^/\?]+/?|.*id=[A-Za-z0-9]+)';
  }
  const regExpText = 'https?://ref.opendeved.net/(g/[0-9]+/[^/]+|zo/zg/[0-9]+/7/[^/]+)/?' + customDomainsPart + '|https?://(www.|)zotero.org/groups/[0-9]+/[^/]+/items/[^/]+/library';

  // Example of regular expression returned by the function refOpenDevEdLinksRegEx:
  // new RegExp('https?://ref.opendeved.net/(g/[0-9]+/[^/]+|zo/zg/[0-9]+/7/[^/]+)/?|https?://(docs.edtechhub.org|docs.opendeved.net|maths.educationevidence.io)/lib(/[^/\?]+/?|.*id=[A-Za-z0-9]+)|https?://(www.|)zotero.org/groups/[0-9]+/[^/]+/items/[^/]+/library', 'i');
  return new RegExp(regExpText, 'i');
}

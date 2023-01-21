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
    "NORMAL_LINK_MARK": "<VALID_LINK>",
    "NORMAL_REDIRECT_LINK_MARK": "<VALID_REDIRECT_LINK>",
    //"UNKNOWN_LIBRARY_MARK": "<UNKNOWN_LIBRARY>",
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
    // "permitted_libraries": ["2129771", "2405685", "2486141", "2447227"],
    "local_show_advanced_menu": true,
    "kerkoValidationSite": 'https://docs.opendeved.net/lib/',
    "group_id": "2129771",
  },
  "edtechhub":
  {
    "name": "ZoteroDocs (EdTech Hub)",
    "default_everybody": false,
    "default_for": "edtechhub.org",
    // "permitted_libraries": ["2405685", "2339240", "2129771"],
    "LINK_MARK_STYLE_BACKGROUND_COLOR": "#dddddd",
    "kerkoValidationSite": 'https://docs.edtechhub.org/lib/',
    "group_id": "2405685"
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

let AUTO_PROMPT_COLLECTION, ORPHANED_LINK_MARK, URL_CHANGED_LINK_MARK, BROKEN_LINK_MARK, UNKNOWN_LIBRARY_MARK, TEXT_TO_DETECT_START_BIB, TEXT_TO_DETECT_END_BIB,
  LINK_MARK_STYLE_FOREGROUND_COLOR, LINK_MARK_STYLE_BACKGROUND_COLOR, LINK_MARK_STYLE_BOLD, HOST_APP,
  NORMAL_LINK_MARK, NORMAL_REDIRECT_LINK_MARK;

const LINK_MARK_STYLE_NEW = new Object();

// Changes a value of ACTIVE_STYLE to style that is default for DocumentProperties kerko_validation_site
function updateStyle() {

  try {
    const ui = DocumentApp.getUi();
    HOST_APP = 'docs';
  }
  catch (error) {
    //Logger.log(error);
  }
  if (HOST_APP == null) {
    try {
      const ui = SlidesApp.getUi();
      HOST_APP = 'slides';
    }
    catch (error) {
      //Logger.log(error);
    }
  }

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
  AUTO_PROMPT_COLLECTION = getStyleValue('AUTO_PROMPT_COLLECTION');
  ORPHANED_LINK_MARK = getStyleValue('ORPHANED_LINK_MARK');
  URL_CHANGED_LINK_MARK = getStyleValue('URL_CHANGED_LINK_MARK');
  BROKEN_LINK_MARK = getStyleValue('BROKEN_LINK_MARK');
  NORMAL_LINK_MARK = getStyleValue('NORMAL_LINK_MARK');
  NORMAL_REDIRECT_LINK_MARK = getStyleValue('NORMAL_REDIRECT_LINK_MARK');
  //UNKNOWN_LIBRARY_MARK = getStyleValue('UNKNOWN_LIBRARY_MARK');
  TEXT_TO_DETECT_START_BIB = getStyleValue('TEXT_TO_DETECT_START_BIB');
  TEXT_TO_DETECT_END_BIB = getStyleValue('TEXT_TO_DETECT_END_BIB');

  LINK_MARK_STYLE_FOREGROUND_COLOR = getStyleValue('LINK_MARK_STYLE_FOREGROUND_COLOR');
  LINK_MARK_STYLE_BACKGROUND_COLOR = getStyleValue('LINK_MARK_STYLE_BACKGROUND_COLOR');
  LINK_MARK_STYLE_BOLD = getStyleValue('LINK_MARK_STYLE_BOLD');

  if (HOST_APP == 'docs') {
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.FOREGROUND_COLOR] = LINK_MARK_STYLE_FOREGROUND_COLOR;
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BACKGROUND_COLOR] = LINK_MARK_STYLE_BACKGROUND_COLOR;
    LINK_MARK_STYLE_NEW[DocumentApp.Attribute.BOLD] = LINK_MARK_STYLE_BOLD;
  }
}

updateStyle();
//Logger.log('Test 3' + ACTIVE_STYLE);

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

function refOpenDevEdLinksRegEx(){
  return new RegExp('https?://ref.opendeved.net/zo/zg/[0-9]+/7/[^/]+/?|https?://docs.(edtechhub.org|opendeved.net)/lib(/[^/\?]+/?|.*id=[A-Za-z0-9]+)', 'i');
}

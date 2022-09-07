/*

The functions below offer style-based reformatting.

*/

//var style = {};
//style[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF'; // null
//style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000'; // null
//style[DocumentApp.Attribute.BOLD] = true;
//style[DocumentApp.Attribute.ITALIC] = true;
//style[DocumentApp.Attribute.STRIKETHROUGH] = true;
//style[DocumentApp.Attribute.UNDERLINE] = true;
//style[DocumentApp.Attribute.FONT_SIZE] = 12;
//style[DocumentApp.Attribute.FONT_FAMILY] = "...";
//style[DocumentApp.Attribute.GLYPH_TYPE] = "...";

function regexpRestyle(target,mystyle,elementsIn) {
  regexpRestyleOffset(target,mystyle,0,0,elementsIn);
};

function regexpRestyleOffset(target,mystyle,startOff,endOff,elementsIn) {
  // If no search parameter was provided, ask for one
  if (arguments.length == 0) {
    var ui = DocumentApp.getUi();
    var result = ui.prompt('Text Highlighter',
      'Enter text to highlight:', ui.ButtonSet.OK_CANCEL);
    // Exit if user hit Cancel.
    if (result.getSelectedButton() !== ui.Button.OK) return;
    // else
    target = result.getResponseText();
  }
  var elements = [];
  if (!elementsIn) {
    elements = getParagraphsInBodyAndFootnotesExtended(false,true,true); 
  } else {
    elements = elementsIn;
  };
  for (var i=0; i<elements.length; i++) {
    var thisEl = elements[i];
    var searchResult = thisEl.findText(target);
    while (searchResult !== null) {
      var thisElement = searchResult.getElement();
      var thisElementText = thisElement.asText();
      // thisElementText.setAttributes(mystyle);
      thisElementText.setAttributes(searchResult.getStartOffset()+startOff, searchResult.getEndOffsetInclusive()+endOff,mystyle);    
      searchResult = thisEl.findText(target, searchResult);
    };      
  }
  
}

/*
This function is obsolete: It's a demonstration of how to do this operation all in one function. However it operates on the body only
*/
function regexpRestyleBodyOnly(target,mystyle,elementsIn) {
  // If no search parameter was provided, ask for one
  if (arguments.length == 0) {
    var ui = DocumentApp.getUi();
    var result = ui.prompt('Text Highlighter',
      'Enter text to highlight:', ui.ButtonSet.OK_CANCEL);
    // Exit if user hit Cancel.
    if (result.getSelectedButton() !== ui.Button.OK) return;
    // else
    target = result.getResponseText();
  }
  var elements = [];
  if (!elementsIn) {
    var doc = DocumentApp.getActiveDocument();
    var bodyElement = DocumentApp.getActiveDocument().getBody();
    elements[0] = bodyElement;
  } else {
    elements = elementsIn;
  };
  for (var i=0; i<elements.length; i++) {
    var thisEl = elements[i];
    var searchResult = thisEl.findText(target);
    while (searchResult !== null) {
      var thisElement = searchResult.getElement();
      var thisElementText = thisElement.asText();
      // thisElementText.setAttributes(mystyle);
      thisElementText.setAttributes(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive(),mystyle);    
      searchResult = thisEl.findText(target, searchResult);
    };      
  }
  
}

   
/**
 * Create custom menu when document is opened.

function onOpen() {
  DocumentApp.getUi().createMenu('Custom')
      .addItem('CreateFormats', 'myFormats')
      .addToUi();
}
 */

/*
https://stackoverflow.com/questions/12064972/can-i-color-certain-words-in-google-document-using-google-apps-script
function highlightTextTwo() {
  var doc  = DocumentApp.openById('<your document id');
  var textToHighlight = 'dusty death';
  var highlightStyle = {};
  highlightStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FF0000';
  var paras = doc.getParagraphs();
  var textLocation = {};
  var i;

  for (i=0; i<paras.length; ++i) {
    textLocation = paras[i].findText(textToHighlight);
    if (textLocation != null && textLocation.getStartOffset() != -1) {
      textLocation.getElement().setAttributes(textLocation.getStartOffset(),textLocation.getEndOffsetInclusive(), highlightStyle);
    }
  }
}
*/

/*

Better would be:

setParaForegound
setParaBackground
setParaBold
setParaItalics
setParaUnderline
setParaStrikethrough
setParaFontSize
setParaFontFamily (does this always request offset?)
setParaStyle

function(regexp,paraMatch,[color|flag])


setTextForegound
setTextBackground
setTextBold
setTextItalics
setTextUnderline
setTextFontSize

function(regexp,[color|flag])


// text.setBackgroundColor("rgba(0, 0, 0, 0)");
// text.setBackgroundColor(null);


*/
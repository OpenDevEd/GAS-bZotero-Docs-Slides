function setPropertyObj(str, styleSet) {
  var sty = JSON.stringify(styleSet);
  // alert("setProperty "+str+": "+sty);
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(str, sty);
}

function getPropertyObj(str) {
  var sty = null;
  try {
    var userProperties = PropertiesService.getUserProperties();
    var proo = userProperties.getProperty(str);
    // alert("getProperty "+str+": "+proo);
    sty = JSON.parse(proo);
  } catch (e) {
  };
  return sty;
}

// ---- Get selected paragraphs --------

// If a selection is made, this function gets the paragraphs that include the selection (i.e. more than the section)
// If no selection is made AND onePara=true, the paragraph in which the cursor is, is returned.
// If no selection is made AND onePara=false, all paragraphs are returned.
function getParagraphs(onePara) {
  return getParagraphsInBodyAndFootnotes(onePara,false);
};

function getParagraphsInBodyAndFootnotes(onePara,getFootnoteParas) {
  return getParagraphsInBodyAndFootnotesExtended(onePara,true,getFootnoteParas);
};

/*
// Example: 
getParagraphsInBodyAndFootnotesExtended(onePara,getBodyParas,getFootnoteParas)

// We want to be sure that only selected (or in-cursor) paragraphs are returned, and we only want body paras
getParagraphsInBodyAndFootnotesExtended(true,true,false);
*/

function getParagraphsInBodyAndFootnotesExtended(onePara,getBodyParas,getFootnoteParas) {
  var paraout = [];
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    // If there's a selection, getBodyParas is ignored.
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only modify elements that can be edited as text; skip images and other non-text elements.
      if (element.getElement().editAsText) {
        var elem = element.getElement();
        // var text = elem.editAsText();
        if (elem.getType() == DocumentApp.ElementType.TEXT) {
          elem = elem.getParent();
        }
        if (elem.getType() == DocumentApp.ElementType.PARAGRAPH) {
          var paragraph = elem.asParagraph();
          paraout.push(paragraph);
        } else if (elem.getType() == DocumentApp.ElementType.LIST_ITEM) {
          var paragraph = elem.asListItem();
          paraout.push(paragraph);
        } else {
          DocumentApp.getUi().alert("Cursor is in object that is not paragraph or list item:" + element.getElement().getType() );
        }
      }
    }
  } else {
    if (onePara) {
      // if onePara is true, ignore getBodyParas
      var cursor = DocumentApp.getActiveDocument().getCursor();
      var element = cursor.getElement();
      var paragraph;
      if (element.getParent().getType() == DocumentApp.ElementType.PARAGRAPH) {
        paragraph = element.getParent().asParagraph();
        paraout.push(paragraph);
      } else if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
        paragraph = element.asParagraph();
        paraout.push(paragraph);
      } else if (element.getParent().getType() == DocumentApp.ElementType.LIST_ITEM) {
        paragraph = element.getParent().asListItem();
        paraout.push(paragraph);
      } else if (element.getType() == DocumentApp.ElementType.LIST_ITEM) {
        paragraph = element.asListItem();
        paraout.push(paragraph);
      } else {      
        DocumentApp.getUi().alert("Cursor is in object that is not paragraph or list item: " + element.getParent().getType() + ", parent of " + element.getType());
      }
    } else {
      var doc = DocumentApp.getActiveDocument();
      var paraout = [];
      if (getBodyParas) {
        try {
          var body = doc.getBody();
          paraout = doc.getParagraphs();
        } catch (e) {
          alert("Error in getParagraphsInBodyAndFootnotesExtended: " + e);
        };
      };
      if (getFootnoteParas) {
        var footnote = doc.getFootnotes();
        if (footnote) {
          // alert("Getting fn: "+footnote.length);
          for(var i in footnote){
            if (footnote[i].getFootnoteContents()) {
              var paragraphs = footnote[i].getFootnoteContents().getParagraphs();
              if (paragraphs) {
                //alert("Getting paras: "+paragraphs.length);
                for (var i = 0; i < paragraphs.length; i++) {
                  var element = paragraphs[i];
                  paraout.push(element);
                };
              };
            } else {
              var j = i+1;
              alert("Footnote has no contents. Footnote number= "+j+". This appears to be a GDocs bug that happens if the footnote is suggested text only.");
            };
          }
        } else {
          alert("There are no footnotes.");
        };
      }
    };
  };
  // DocumentApp.getUi().alert(paraout.length);
  // alert(paraout.length);
  return paraout;
}


function getValueFromUser(title,text, defaultOK, defaultCancel, defaultClose) {
  text = text || "Please enter a value.";
  defaultOK = defaultOK || "";
  defaultCancel = defaultCancel || null;
  defaultClose = defaultClose || null;
  if (!text) {
    //text = text || title;
    title = text;
    title = "BUtils";
  };
  var result = DocumentApp.getUi().prompt(title,text, DocumentApp.getUi().ButtonSet.OK_CANCEL);
  // Process the user's response:
  if (result.getSelectedButton() == DocumentApp.getUi().Button.OK) {
    var res = result.getResponseText();
    // DocumentApp.getUi().alert('Result: '+res);
    if (res == "" && defaultOK) {
      return defaultOK;
    } else {      
      return res;
    };
  } else if (result.getSelectedButton() == DocumentApp.getUi().Button.CANCEL) {
    //DocumentApp.getUi().alert('The user didn\'t want to provide a value.');
    return defaultCancel;
  } else if (result.getSelectedButton() == DocumentApp.getUi().Button.CLOSE) {
    //DocumentApp.getUi().alert('The user clicked the close button in the dialog\'s title bar.');
    return defaultClose;
  }
  DocumentApp.getUi().alert('Unknown action.');
  return null;
}

function getConfirmationFromUser(text) {
  // Display a dialog box with a message and "Yes" and "No" buttons.
  var ui = DocumentApp.getUi();
  var response = ui.alert(text, ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  if (response == ui.Button.OK) {
    return true;
  } else {
    return false;
  }
}


function alert(text) {
  DocumentApp.getUi().alert(text);
};

function alertv(text,myvariable) {
  DocumentApp.getUi().alert(text+" = "+JSON.stringify(myvariable));
};

function onlyUnique(value, index, self) { 
    return self.indexOf(value) === index;
};

function getColourByName(name) {
  if (name) {
    if (htmlColourNames[name]) {
      return htmlColourNames[name];
    } else if (standardColorObjectGoogle[name]) {
      return standardColorObjectGoogle[name];
    } else {
      return null;
    };
  } else {
    return null;
  };
};
// replace regexp
function enterRegexpAndReplace() {
  var find = DocumentApp.getUi().prompt('Please enter regular expression to find.','Regular expression:', DocumentApp.getUi().ButtonSet.OK_CANCEL);
  if (find.getSelectedButton() == DocumentApp.getUi().Button.OK) {
    var replace = DocumentApp.getUi().prompt('Please enter regular expression to find.','Regular expression:', DocumentApp.getUi().ButtonSet.OK_CANCEL);
    if (replace.getSelectedButton() == DocumentApp.getUi().Button.OK) {
      DocumentApp.getUi().alert('Regular expression /' + find.getResponseText() + '/' + replace.getResponseText()+'/' );
      var body = DocumentApp.getActiveDocument().getBody(); 
      body.replaceText(find.getResponseText(),replace.getResponseText()); 
//      body.replaceText(/(xbx)/,'$1-$1');
    } else if (replace.getSelectedButton() == DocumentApp.getUi().Button.CANCEL) {
    } else if (replace.getSelectedButton() == DocumentApp.getUi().Button.CLOSE) {
    }
  } else if (find.getSelectedButton() == DocumentApp.getUi().Button.CANCEL) {
    //DocumentApp.getUi().alert('The user didn\'t want to provide a name.');
  } else if (find.getSelectedButton() == DocumentApp.getUi().Button.CLOSE) {
    //DocumentApp.getUi().alert('The user clicked the close button in the dialog\'s title bar.');
  }
}

function multiReplaceX() {
  // var body = DocumentApp.getActiveDocument().getBody(); 
  // body.replaceText(/(xbx)/,'$1-$1');
  // body.replaceText(/txt/,'FF');
  singleReplace(":::: txt/(\\d+)","FF$1.",true,true,null);
}

function multiReplace() {
  var iostring = getValueFromUser("Multi Replace","Please paste json consisting of [ [A1,B1], [A2,B2], ...]", null);
  if (iostring) {
    var iostr = JSON.parse(iostring);
    alert('Programmed multi-replace. Read '+iostr.length+ " replacements.");
    // alert('hello '+iostr[0][0]+"->"+iostr[0][1]);
    for (var i=0; i<iostr.length; i++) {
      if (iostr[i][0] && iostr[i][1] && iostr[i][0].match(/\S/)) {
        // alert('hello '+iostr[i][0]+"->"+iostr[i][1]);
        singleReplacePartial(iostr[i][0],iostr[i][1],false,false,null, true, true);
      };
    };
  };
};




//(?=a)
//(?!a)
//(?<!b)
//(?<=b)

function singleReplace(re,str,isRegExp,useJS,flags) {
  singleReplacePartial(re,str,isRegExp,useJS,flags,true,true);
};
  
function singleReplacePartial(re,str,isRegExp,useJS,flags,bodyFlag,fnFlag) {
  // isRegExp: Is the string passed in a regexp or a literal string?
  // useJS: 
  // Seems that the find function doesn't take a regexp, but just a string...
  // https://developers.google.com/apps-script/reference/document/text#replaceText(String,String)
  // Therefore, you can switch on useJS, to use the JS engine. Because this operates on text, you would lose formatting.
  // Hence, the expressing is 'padded', see repad below, and only the matching text is retained. Then GAS replace is used.
  // useJS may have unexpected side effects and should only be used if you require features that are not in re2 but are in JS (suchas '$1')
  if (!isRegExp) {
    re = re.replace(/[|\\{}()[\]^$+*?.]/g, '\\$&');
    //    re = re.replace(/[\|\\\{\}\(\)\[\]^$\+\*\?\.]/g, '\\$&');
  };
  //var reg = new RegExp(re);
  // string - for use in findtext and replacetext
  var regu = re;
  // regexp - for use in replace
  var regexp = null;
  // why this?
  // var repad = ".*?(?:"+re+").$";
  var repad = re;
  if (flags) { 
    regexp = new RegExp(repad, flags);
  } else {
    regexp = new RegExp(repad);
  };
  // DocumentApp.getUi().alert("Search: " + regu.toString() );
/*
  var bodyElement = DocumentApp.getActiveDocument().getBody();
*/  
  var p;
  try {
    p = getParagraphsInBodyAndFootnotesExtended(false, bodyFlag, fnFlag);
  } catch (e) {
    alert('Error in singleReplacePartial calling getParagraphsInBodyAndFootnotesExtended: ' + e);
  };
  if (p) {
    for (var i=0; i < p.length; i++) {
      var bodyElement = p[i];
      var rangeElement = bodyElement.findText(regu);
      var replacement;
      while (rangeElement !== null) {
        //DocumentApp.getUi().alert("Found" );
        var thisElement = rangeElement.getElement();
        //var thisElementText = thisElement.asText();
        // work-around to be able to use JS regexp engine
        // https://stackoverflow.com/questions/30968419/replacetext-regex-not-followed-by/33528920#33528920 
        if (useJS) {
          
           alert("useJS not working");
          /*
                  // Determine the input text:
          var mytext =  rangeElement.getElement().getText();
          var elem = rangeElement.getElement().copy().editAsText();
          var elemlength = elem.getText().length;
          if (rangeElement.isPartial()) {
            if (rangeElement.getEndOffsetInclusive()+1 < elemlength-1) {
              elem.deleteText(rangeElement.getEndOffsetInclusive()+1,elemlength-1);
            }
            if (rangeElement.getStartOffset()-1>0) {
              elem.deleteText(0,rangeElement.getStartOffset()-1);
            };
            elemlength = elem.getText().length;
          };
          var url = null;
          // var text = elem.getText();                                        
          alert(regexp+ "\n" + str + "\n" + thisElement.getText() + "\n" + elem.getText());
          replacement = thisElement.getText().replace(regexp,str);
          
          var eat = rangeElement.getElement().editAsText();
          eat.deleteText(rangeElement.getStartOffset(),rangeElement.getEndOffsetInclusive());
          eat.insertText(rangeElement.getStartOffset(),replacement);      
          */
          
          alert(regexp+ "\n" + str + "\n" + thisElement.getText());
          replacement = thisElement.getText().replace(regexp,str);
        } else {
          replacement = str;
        }
        thisElement = thisElement.replaceText(regu, replacement);
        //.setBackgroundColor(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive(),backgroundColor);
        // search for next match
        rangeElement = bodyElement.findText(regu, rangeElement);
      };
    };
  } else {
    alert("singleReplacePartial: No paragraphs found.");
  };
};
  
function singleReplaceKeepUrl(regu,replacement,isRegExp,offsetStart, offsetEnd) {
  // isRegExp: Is the string passed in a regexp or a literal string?
  try {
    if (!isRegExp) {
      regu = regu.replace(/[|\\{}()[\]^$+*?.]/g, '\\$&');
    };
  } catch(e) {
    alert("Error in regu: "+e);
  };
  /*
  // This is just a basic replacement:
  var bodyElement = DocumentApp.getActiveDocument().getBody();
  var searchResult = bodyElement.findText(regu);
  var replacement = str;
  while (searchResult !== null) {
    var thisElement = searchResult.getElement();
    thisElement = thisElement.replaceText(regu, replacement);
    searchResult = bodyElement.findText(regu, searchResult);
  }
  */
  try {
    var doc = DocumentApp.getActiveDocument();
    var counter = 0;
    var p = getParagraphsInBodyAndFootnotes(false,true);
  } catch(e) {
    alert("Error in singleReplaceKeepUrl>getParagraphsInBodyAndFootnotes: "+e);
  };
  try {
  for (var i=0; i<p.length; i++) {
    try {
      var mybody = p[i];
      // Track through range elements
      var rangeElement = mybody.findText(regu);
      while (rangeElement !== null) {
        counter++;
        // Determine the input text:
        var mytext =  rangeElement.getElement().getText();
        var elem = rangeElement.getElement().copy().editAsText();
        var elemlength = elem.getText().length;
        if (rangeElement.isPartial()) {
          if (rangeElement.getEndOffsetInclusive()+1 < elemlength-1) {
            elem.deleteText(rangeElement.getEndOffsetInclusive()+1,elemlength-1);
          }
          if (rangeElement.getStartOffset()-1>0) {
            elem.deleteText(0,rangeElement.getStartOffset()-1);
          };
          elemlength = elem.getText().length;
        };
        var url = null;
        var text = elem.getText();
        // Find URL by tracking through the input text - take first available url
        for (var i=0; i < text.length; i++) {
          if (!url) {
            url = elem.getLinkUrl(i);
          };
        };
        // replacement = "{"+replacement+"}"
        eat = rangeElement.getElement().editAsText();
        eat.deleteText(rangeElement.getStartOffset(),rangeElement.getEndOffsetInclusive());
        eat.insertText(rangeElement.getStartOffset(),replacement);
        var pin = rangeElement.getStartOffset();
        var pout = rangeElement.getStartOffset() + replacement.length - 1;
        eat.setLinkUrl(pin, pout, null);
        pin += offsetStart;
        pout +=  offsetEnd;
        if (pout >= pin ) {
          eat.setLinkUrl(pin, pout, url);
        } else {
          //alert("Error setting link on: "+text);
        };
        rangeElement = mybody.findText(regu,rangeElement );
        rangeElement = null;
      };
    } catch(e) {
      alert(i+ " " + e);
    };
  }
  } catch(e) {
    alert("Error in singleReplaceKeepUrl: "+e);
  };
  return counter;
};
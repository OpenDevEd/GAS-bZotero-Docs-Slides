// URL unpack markers: left and right
// This will clash with the bib markers
// const UL = "⁅";
// const UR = "⁆";
const UL = "❲";
const UR = "❳";

var redirectTarget = "https://bjohas.de/zo/";
redirectTarget = "https://ref.opendeved.net/zo/";

// https://bjohas.de/zo/zg/2317526/7/XAXNF68P/BBSSA-F
var unpackRe = "https\:\/\/(bjohas\.de|ref\.opendeved\.net)\/zo\/";

function minifyCitationsDocs() {
  var p = getParagraphsInBodyAndFootnotes(false, true);
  var mystyle = {};
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0123DD';
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 60;
  var regexp = "⟦(zg)?";
  regexpRestyle(regexp, mystyle, p);
  mystyle = {};
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0123DD';
  /*  var regexp = "⟦([\\d\\:\\w]+?\\|)?";
  regexpRestyle(regexp,mystyle);
  regexp = "(\\|[\\d\\:\\w]+?)?⟧";
  regexpRestyle(regexp,mystyle); */
  regexp = "⟦(zg)?\\:\\d*\\:[\\d\\w]+\\|.*?⟧";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 6;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FE01DC';
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  regexp = "⟦((zg)?\\:\\d*\\:[\\d\\w]+\\|)?";
  regexpRestyle(regexp, mystyle, p);
  regexp = "⟧";
  regexpRestyle(regexp, mystyle, p);
  regexp = "</?pn>";
  regexpRestyle(regexp, mystyle, p);
  // regexp = "(\@|\#)";
  // regexpRestyle(regexp,mystyle,p); 
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#DEDEDE';
  regexp = "《warning:.*?》";
  regexpRestyle(regexp, mystyle, p);
};

// Old version clearZwarnings
function OLDclearZwarnings() {
  singleReplace("《warning: [^《》]*?》", "", true, false, null);
  //  singleReplace("<\/?pn>","",true,false, null);
  //  singleReplace("\\|\\:","|",true,false, null);
  singleReplace(UL+"|"+UR, "", true, false, null);
};

function maxifyCitationsDocs() {
  var p = getParagraphsInBodyAndFootnotes(false, true);

  // deal with remnants of reformatting
  mystyle = {};
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  mystyle[DocumentApp.Attribute.ITALIC] = false;
  regexpRestyle("\\(⟦|⟧\\)", mystyle, p);

  var mystyle = {};
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#0123DD';
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  mystyle[DocumentApp.Attribute.ITALIC] = false;
  /*  
  var regexp = "⟦([\\d\\:\\w]+?\\|)?";
  regexpRestyle(regexp,mystyle);
  // Country marker/
  regexp = "(\\|[\\d\\:\\w]+?)?⟧";
  regexpRestyle(regexp,mystyle); */
  var regexp = "⟦(zg)?\\:\\d*\\:[\\d\\w]+\\|.*?⟧";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FE01DC';
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  regexp = "⟦((zg)?\\:\\d*\\:[\\d\\w]+\\|)?";
  regexpRestyle(regexp, mystyle, p);
  regexp = "⟧";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#DEDEDE';
  regexp = "《warning:.*?》";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#88FF88';
  //regexp = "⟦(zg)?\\:2249382\\:[\\d\\w]+\\|";
  //regexpRestyle(regexp,mystyle,p);
  //highlightMissingCountryMarker();



};

function unfyCitationsDocs() {
  var p = getParagraphsInBodyAndFootnotes(false, true);
  var mystyle = {};
  //  mystyle[DocumentApp.Attribute.FONT_SIZE] = 6;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  /*  var regexp = "⟦([\\d\\:\\w]+?\\|)?";
  regexpRestyle(regexp,mystyle);
  regexp = "(\\|[\\d\\:\\w]+?)?⟧";
  regexpRestyle(regexp,mystyle); */
  var regexp = "⟦(zg)?\\:\\d*\\:[\\d\\w]+\\|.*?⟧";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  regexp = "⟦((zg)?\\:\\d*\\:[\\d\\w]+\\|)?";
  regexpRestyle(regexp, mystyle, p);
  regexp = "⟧";
  regexpRestyle(regexp, mystyle, p);
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  regexp = "《warning:.*?》";
  regexpRestyle(regexp, mystyle, p);
};

function highlightMissingCountryMarker() {
  var p = getParagraphsInBodyAndFootnotes(false, true);
  var mystyle = {};
  mystyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  mystyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FE01DC';
  mystyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#00FFFF';
  regexp = "\\|\\:";
  regexpRestyle(regexp, mystyle, p);
};


function zoteroPackUnpack(pack, promptForLibrary, showWarningsWhenUnpacking, vancouverStyle) {
  var number = 1;
  var total = 0;
  var string = "";
  var rounds = -1;
  var value;
  if (promptForLibrary) {
    value = getValueFromUser("Convert Zotero text to links", "Please enter the library number you wish to process. Only references matching that library number will be converted. Press cancel to convert all.", "", "", "");
  };
  while (number > 0) {
    rounds++;
    if (pack) {
      number = packZotero(value, vancouverStyle, value);
    } else {
      if (rounds == 0) {
        number = zoteroUnpack(true, showWarningsWhenUnpacking);
      } else {
        number = zoteroUnpack(false, showWarningsWhenUnpacking);
      };
    };
    total += number;
  };
  if (pack) {
    string = "packed";
  } else {
    minifyCitationsDocs();
    string = "unpacked and minified";
  };
  var advice = "";
  if (!pack) {
    advice = " Following unpacking, check the document for symbols "+UL+", "+UR+", ⇡ and 《warning》 and then run cleanup.";
  };
  alert("Number of citations that were Zotero-" + string + ": " + total + " (" + rounds + " rounds)." + advice);
};


function packZotero(promptForLibrary, vancouverStyle, value) {
  var vancouver;
  if (vancouverStyle) {
    vancouver = vancouverStyle;
  } else {
    vancouver = false;
  };
  var counter = 0;
  var marker1 = '⟦';
  var marker2 = '⟧';
  var vancnum = {};
  var vancord = [];
  var vancmax = 0;
  var doc = DocumentApp.getActiveDocument();
  // var body = doc.getBody();
  var regu = "⟦.*?⟧";
  if (promptForLibrary) {
    if (value) {
      regu = "⟦zg:" + value + ":.*?⟧";
    }
  }
  var p = getParagraphsInBodyAndFootnotes(false, true);
  if (!p) {
    alert("No paragraphs!");
    return;
  };
  for (var i = 0; i < p.length; i++) {
    var mybody = p[i];
    var searchResult = mybody.findText(regu);
    while (searchResult !== null) {
      counter++;
      startx = searchResult.getStartOffset();
      endx = searchResult.getEndOffsetInclusive();
      searchResult.getElement().asText();
      var eat = searchResult.getElement().editAsText();
      var text = eat.getText();
      text = text.substring(startx, endx + 1);
      var marker1 = "";
      var marker2 = "";
      marker1 = "⇡";
      //marker2 = "";
      //marker1 = "⸢";
      // marker2 = "";
      //marker1 = "⸤";
      //marker2 = "⸥";
      //marker2 = "⸣";
      marker2 = "";
      var offset = marker1.toString().length;
      var endoffset = marker2.toString().length;
      eat.insertText(startx, marker1);
      eat.insertText(endx + offset, marker2);
      // eat.insertText(startx,"⇡");
      //var offset = 1;
      //var endoffset = 0;      
      //eat.insertText(startx,"⸢");
      //var offset = 1;
      //eat.insertText(endx+offset,"⸣");
      //var endoffset = 1;
      var urlpart = "";
      var regex = /^⟦(zg)?\:([^\:\|]*)\:([^\:\|]+)\|(.*?)⟧$/;
      var found = text.match(regex);
      if (found) {
        if (!found[1]) {
          found[1] = "zg";
        };
        if (!found[2]) {
          found[2] = "2249382";
        };
        // found[4] = encodeURIComponent(found[4]);
        urlpart = found[1] + "/" + found[2] + "/7/" + found[3] + "/" + found[4]; // +"//"+text;
      } else {
        urlpart = "NA/" + text;
      };

// Adds to URL zoteroItemKey and zoteroCollectionKey of current Doc
let newUrl = addCurrentKeys(urlpart);
      // // New 2021-04-24
      // let newUrl = redirectTarget + urlpart;

      // let currentZoteroItemKey = getDocumentPropertyString('zotero_item');
      // if (currentZoteroItemKey != null) {
      //   const zoteroItemKeyParts = currentZoteroItemKey.split('/');
      //   const zoteroItemKeyParameters = zoteroItemKeyParts[4] + ':' + zoteroItemKeyParts[6];
      //   newUrl = replaceAddParameter(newUrl, 'src', zoteroItemKeyParameters);
      // }

      // const currentZoteroCollectionKey = getDocumentPropertyString('zotero_collection_key');
      // if (currentZoteroCollectionKey != null) {
      //   const zoteroCollectionKeyParts = currentZoteroCollectionKey.split('/');
      //   zoteroCollectionKey = zoteroCollectionKeyParts[6];
      //   newUrl = replaceAddParameter(newUrl, 'collection', zoteroCollectionKey);
      // }

      // newUrl = replaceAddParameter(newUrl, 'openin', 'zoteroapp');
      // // End. New 2021-04-24

      //eat.setLinkUrl(startx,endx+offset,redirectTarget+urlpart);
      eat.setLinkUrl(startx, endx + offset, newUrl);
      // Delete the closing bracket:
      try {
        eat.deleteText(endx + offset + endoffset, endx + offset + endoffset);
      } catch (e) {
      };
      // Find the zotero marker:
      var regex = /^(⟦\S+?\|)/;
      var found = text.match(regex);
      if (found) {
        var len = found[1].length;
        if (vancouver) {
          // Adjust the middle text if needed
          var insertPoint = startx + len - 1 + offset + 1;
          if (vancnum[found[0]] != null) {
          } else {
            vancmax++;
            vancnum[found[0]] = vancmax;
            vancord.push(found[0]);
          };
          var insStr = String(vancnum[found[0]])
          eat.insertText(insertPoint, insStr);
          eat.deleteText(insertPoint + insStr.length, endx + offset + insStr.length - 1);
        };
        // Remove the Zotero code marker:
        eat.deleteText(startx + offset, startx + len - 1 + offset);
      };
      /*    Possible things to do in the report:
      What references are cited where? Make a list of references and which sections they are cited in. In an ideal world, one would be able to click from any reference down to the list, and then backup
      S Sandirasegarane, S. S., A. Gill…. (2016). Context-driven entrepreneurial education in vocational schools. ↑6.4 ↑6.9 ↑9.1
      What would be even more amazing is if various citations were linked… as well as countries… 
      Ghana↶⇊↷
      Sandirasegarane et al. 2016↶⇊↷
      ↶Sandirasegarane et al. 2016↷
      ⤤|
      ⇈↟⇞⍐⇡
      */
      searchResult = mybody.findText(regu, searchResult);
    };
  };
  return counter;
}

// ⇡...⸣
// The problem with this approach is that it leads to ⇡...⇡...⸣ which then cannot be unpacked.
// Would be better to unpack the links directly - or to use different symbol.
// NOW IMPLEMENTED
// unpackCombined, unpackCombinedWarning use the function 
function restoreZoteroLinks() {
  element = DocumentApp.getActiveDocument().getBody();
  //linkMarker(element,true);
  markLinksWithSymbol(element, UL, UR);
  var footnote = DocumentApp.getActiveDocument().getFootnotes();
  for (var i in footnote) {
    // alert(i);
    var paragraphs = footnote[i].getFootnoteContents().getParagraphs();
    if (paragraphs) {
      for (var i = 0; i < paragraphs.length; i++) {
        var element = paragraphs[i];
        markLinksWithSymbol(element, UL, UR);
      };
    };
  }
  // Make some fixes
  try {
    fixZoteroLinkPresentation();
  } catch (e) {
    alert("Fixes error = " + e);
  };
  // var returnx = getConfirmationFromUser("Now use zunpack to turn links into Zotero.");
};


function fixZoteroLinkPresentation() {
  // correct
  // ("+UL+"⇡ ... "+UR+")
  // Incorrect
  // (⇡"+UL+" or ⇡"+UL+"  
  // "+UL+"(⇡
  // ...)"+UR+"
  // Incorrect, but doesn't occur
  // ⇡("+UL+"
  /* singleReplaceKeepUrl("⇡"+UL,UL+"⇡",false);
  singleReplaceKeepUrl(UL+"(⇡","("+UL+"⇡",false);
  singleReplaceKeepUrl(")"+UR, UR+")",false); */
  /*
  singleReplaceKeepUrl("(⇡"+UL,"("+UL+"⇡",false,1,0);
  singleReplaceKeepUrl("⇡"+UL, UL+"⇡",false,0,0);
  singleReplaceKeepUrl(UL+"(⇡","("+UL+"⇡",false,1,0);
  singleReplaceKeepUrl(")"+UR, UR+")",false,0,-1);
  */
  // The following singleReplaceKeepUrl work on footnotes too
  // Fix wrong sequences:
  singleReplaceKeepUrl("(⇡"+UL, "("+UL, false, 1, 0);
  singleReplaceKeepUrl("⇡"+UL, UL, false, 0, 0);
  singleReplaceKeepUrl(UL+"(⇡", "("+UL, false, 1, 0);
  singleReplaceKeepUrl(UL+"⇡", UL, false, 0, 0);
  // This may or may not be correct: We dont know!
  // singleReplaceKeepUrl(")"+UR, UR+")",false,0,-1);
  // remove ⇡ in correct sequences:
  singleReplaceKeepUrl("("+UL+"⇡", "("+UL, false, 1, 0);
  singleReplaceKeepUrl(UL+"⇡", UL, false, 0, 0);
  // actually - this cannot occur! The error was in 06c...       eat.setLinkUrl(pin, pout, null);
  // actuqlly, seems it can.
  // This is correct sequence, but the url placement may not be.
  // Need to do this in two steps, otherwise infinite loop
  // singleReplaceKeepUrl(""+UR+")",""+UR+")",false,0,-1);
  // singleReplaceKeepUrl(""+UR+")",""+UR+"《》)",false,0,-2);
  // singleReplaceKeepUrl(""+UR+"《》)",""+UR+"-)",false,0,-3
};

// This is based on matching ⇡...⸣
// The problem with this approach is that it fails on ⇡...⇡...⸣ which then cannot be unpacked.
// Would be better to unpack links directly - or to use a unique combination...

function zoteroUnpack(promptUser, showWarningsWhenUnpacking) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  // Should first search for ""+UL+"",""+UR+""
  var rangeElem = body.findText(UL+"|"+UR);
  if (1 == 1) {
    // !promptUser || rangeElem === null || getConfirmationFromUser("There are already "+UL+"..."+UR+" markers in the document. Proceed?")) {
    // Mark links
    // markLinksWithSymbol(body,""+UL+"",""+UR+"");
    // fixZoteroLinkPresentation();
    // Turn links into Zotero expressions.
    // var returnx = getConfirmationFromUser("Press OK to convert links to Zotero experssions.");
    var counter = 0;
    var regu = UL+"⇡?[^"+UL+UR+"]*?"+UR;
    var p = getParagraphsInBodyAndFootnotes(false, true);
    for (var i = 0; i < p.length; i++) {
      var mybody = p[i];
      var rangeElement = mybody.findText(regu);
      while (rangeElement !== null) {
        counter++;
        var mytext = rangeElement.getElement().getText();
        // get a copy      
        var elem = rangeElement.getElement().copy().editAsText();
        var elemlength = elem.getText().length;
        // find the text that has the link
        if (rangeElement.isPartial()) {
          //alert(elemlength+" "+elem.getText());
          if (rangeElement.getEndOffsetInclusive() + 1 < elemlength) {
            elem.deleteText(rangeElement.getEndOffsetInclusive() + 1, elemlength - 1);
          }
          if (rangeElement.getStartOffset() - 1 > 0) {
            elem.deleteText(0, rangeElement.getStartOffset() - 1);
          };
          elemlength = elem.getText().length;
          //alert(elemlength+" "+elem.getText());
        } else {
        };
        var str = null;
        var text = elem.getText();
        // Find URL.
        for (var i = 0; i < text.length; i++) {
          if (!str) {
            str = elem.getLinkUrl(i);
            // New 2021-04-24
            if (str != null) {
              //Logger.log('str = ' + str);
              let qMarkPos = str.indexOf('?')
              if (qMarkPos != -1) {
                str = str.slice(0, qMarkPos);
                //Logger.log('str = ' + str);
              }
            }
            // End. New 2021-04-24
          };
        };
        if (str && str.match(unpackRe)) {
          // This is a zotero url that needs to be processed.
          if (text) {
            text = text.slice(0, -1).slice(1);
          };
          // var elemt = elem.getText();
          var strarr = [0, 0, 0, 0, 0, 0, 0, 0, 0];
          if (str) {
            strarr = str.split("/");
          };
          var strarr8 = strarr[8];
          for (var i = 9; i < strarr.length; i++) {
            strarr8 += "/" + strarr[i];
          };
          var vancmode = false;
          //alert(text);
          if (text.match(/^⇡?\d+$/)) {
            vancmode = true;
          };
          if (strarr8) {
            strarr8 = strarr8.replace(/\%20/g, " ");
            if (!vancmode) {
              if (strarr8 == text || (!showWarningsWhenUnpacking && strarr8.length < 30)) {
                strarr8 = "";
              } else {
                strarr8 = "《warning:" + strarr8 + "》";
              };
            };
          };
          // Now replace the text
          eat = rangeElement.getElement().editAsText();
          //eat.insertText(rangeElement.getEndOffsetInclusive()+1, "⟧");
          eat.deleteText(rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive());
          if (strarr.length > 2 && vancmode) {
            text = "";
          };
          eat.insertText(rangeElement.getStartOffset(), "⟦" + strarr[4] + ":" + strarr[5] + ":" + strarr[7] + "|" + text + strarr8 + "⟧");
        } else {
          // This element has no url... (unlikely)
          if (!str) {
            alert("Did the document has markers already? No url found in text: " + text);
          };
          // or (regular):
          // This is not a Zotero url, that should be left... delete the start and end marker so we don't end up in a loop
          eat = rangeElement.getElement().editAsText();
          // delete at end first
          eat.deleteText(rangeElement.getEndOffsetInclusive(), rangeElement.getEndOffsetInclusive());
          eat.deleteText(rangeElement.getStartOffset(), rangeElement.getStartOffset());
        };
        // "http://oer.educ.cam.ac.uk/zotero/select/"+item_code+"?r="+encodeURI(mytext);
        //      var mynewtext = rangeElement.getElement().editAsText().getLinkUrl(rangeElement.getStartOffset());
        //      mynewtext = mynewtext.replace(/^.*?\?r=/,"");
        //      mynewtext = decodeURI(mynewtext);
        //    DocumentApp.getUi().alert(mynewtext);
        //alert("No URL in: "+ text);
        rangeElement = mybody.findText(regu, rangeElement);
        rangeElement = null;
      };
    }
    return counter;
  } else {
    // user cancelled    
    // alert("cancelled");
  }
}





/*  // Parse the text iteratively to find the start and end indices for each link
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var links = [];
    var string = element.getText();
    var previousUrl = null; // The URL of the previous character
    var currentLink = null; // The latest link being built
    for (var charIndex = 0; charIndex < string.length; charIndex++) {
      var currentUrl = element.getLinkUrl(charIndex);
      // New URL means create a new link
      if (currentUrl !== null && previousUrl !== currentUrl) {
        if (currentLink !== null) links.push(currentLink);
        currentLink = {};
        currentLink.url = String(currentUrl);
        currentLink.startOffset = charIndex;
      }
      // In a URL means extend the end of the current link
      if (currentUrl !== null) {
        currentLink.endOffsetInclusive = charIndex;
      }
      // Not in a URL means close and push the link if ready
      if (currentUrl === null) {
        if (currentLink !== null) links.push(currentLink);
        currentLink = null;
      }
      // End the loop and go again
      previousUrl = currentUrl;
    }
    // Handle the end case when final character is a link
    if (currentLink !== null) links.push(currentLink);
    // Convert the links into a range before returning
    links.forEach(function(link) {
      rangeBuilder.addElement(element, link.startOffset, link.endOffsetInclusive);
    });
  }

  // If not a text element then recursively get links from child elements
  else if (element.getNumChildren) {
    for (var i = 0; i < element.getNumChildren(); i++) {
      links.push(getAllLinks(element.getChild(i)));
    }
  }

  return links;
} */
function markLinksWithSymbol(element,startMarker,endMarker) {
  markLinksWithSymbolReport(element,startMarker,endMarker,true);
};
  
function markLinksWithSymbolReport(element,startMarker,endMarker,showAlert) {
  // Parse the text iteratively to find the start and end indices for each link
  var startMarkersInserted = 0;
  var endMarkersInserted = 0;
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var urlstack=0;
    var string = element.getText();
    var previousUrl = null; // The URL of the previous character 
    var currentChar = null;
    var prevChar = null;
    var max = string.length
    for (var charIndex = 0; charIndex < max; charIndex++) {
      var currentUrl = element.getLinkUrl(charIndex);
      currentChar = string.slice(charIndex,charIndex+1);
      // First check whether we are coming out of a URL OR transitioning between urls
      if ( 
        endMarker && (
        (currentUrl === null && previousUrl !== null) 
        || 
        (currentUrl !== null && previousUrl !== null && previousUrl !== currentUrl)
      )) {
        // We have come out of a url - so now need to insert "end of url" character
        urlstack--;
        /* Block 2 = Block 1 repeated */
        var eou = endMarker;
        if (currentChar !== eou && prevChar !== eou) {
          element.editAsText().insertText(charIndex, eou);
          // alert(prevChar+currentChar);
          element.editAsText().setLinkUrl(charIndex, charIndex, previousUrl);
          element.editAsText().setForegroundColor(charIndex, charIndex, "#103DC1");
          element.editAsText().setUnderline(charIndex, charIndex, false);
          charIndex++;
          max++;
          string = element.getText();
          prevChar = currentChar;
          currentChar = eou;       
          endMarkersInserted++;
        }
      }
      // Now check whether we are going into a URL OR transitioning between urls
      if ( 
        startMarker && (
        (currentUrl !== null && previousUrl === null) 
        || 
        (currentUrl !== null && previousUrl !== null && previousUrl !== currentUrl)
      )) {
        var sou = startMarker;
        if (currentChar !== sou && prevChar !== sou) {
          /* Block 1 */
          urlstack++;
          element.editAsText().insertText(charIndex, sou);
          element.editAsText().setLinkUrl(charIndex, charIndex, currentUrl);
          charIndex++;
          max++;
          string = element.getText();
          prevChar = currentChar;
          currentChar = sou;
          startMarkersInserted++;
        } else if (currentChar !== sou && prevChar === sou) {
          // link the previous one.
          element.editAsText().setLinkUrl(charIndex-1, charIndex-1, currentUrl);
        } else {
          // current char is ⇡ and we are in a link - all good
        };
      }
      // End the loop and go again
      previousUrl = currentUrl;
      prevChar = currentChar;
    } // end of loop
    // Handle the end case when final character is a link
    if (endMarker && currentUrl !== null) {
      // alert("End of text insertion");
      urlstack--;
      /* Block 3 = Block 2 repeated */
      var eou = endMarker;
      if (currentChar !== eou && prevChar !== eou) {
        element.editAsText().insertText(charIndex, eou);
        element.editAsText().setLinkUrl(charIndex, charIndex, currentUrl);
        endMarkersInserted++;
        /* not needed as we're done.
        charIndex++;
        max++;
        string = element.getText();
        prevChar = currentChar;
        currentChar = eou;          
        */
      };
    };
    if (urlstack != 0) {
      //alert("urlstack: "+urlstack);
    };
  }  else if (element.getNumChildren) {
    // If not a text element then recursively get links from child elements
    for (var i = 0; i < element.getNumChildren(); i++) {
      markLinksWithSymbolReport(element.getChild(i), startMarker, endMarker,false);
    }
  }
  if (showAlert) {
    // For this to work, we'd have to pass those numbers out or make them global?
    // alert("Done! startMarkersInserted="+startMarkersInserted+"; endMarkersInserted="+endMarkersInserted+"; urlstack="+urlstack);
  };
}

function linkMarker(element,insertEndMarker) {
  // Parse the text iteratively to find the start and end indices for each link
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var urlstack=0;
    var string = element.getText();
    var previousUrl = null; // The URL of the previous character 
    //var currentLink = null; // The latest link being built
    var currentChar = null;
    var prevChar = null;
    var max = string.length
    for (var charIndex = 0; charIndex < max; charIndex++) {
      var currentUrl = element.getLinkUrl(charIndex);
      currentChar = string.slice(charIndex,charIndex+1);
      // First check whether we are coming out of a URL OR transitioning between urls
      if ( 
        (currentUrl === null && previousUrl !== null) 
        || 
        (currentUrl !== null && previousUrl !== null && previousUrl !== currentUrl)
      ) {
        if (insertEndMarker) {
          // We have come out of a url - so now need to insert "end of url" character
          urlstack--;
          /* Block 2 = Block 1 repeated */
          var eou = "⸣";
          if (currentChar !== eou && prevChar !== eou) {
            element.editAsText().insertText(charIndex, eou);
            // alert(prevChar+currentChar);
            element.editAsText().setLinkUrl(charIndex, charIndex, previousUrl);
            element.editAsText().setForegroundColor(charIndex, charIndex, "#103DC1");
            element.editAsText().setUnderline(charIndex, charIndex, false);
            charIndex++;
            max++;
            string = element.getText();
            prevChar = currentChar;
            currentChar = eou;          
          }
        };
      }
      // Now check whether we are going into a URL OR transitioning between urls
      if ( 
        (currentUrl !== null && previousUrl === null) 
        || 
        (currentUrl !== null && previousUrl !== null && previousUrl !== currentUrl)
      ) {
        if (currentChar !== "⇡" && prevChar !== "⇡") {
          /* Block 1 */
          // We have entered a url, but the previous or current char is not ⇡ -> insert ⇡ at current position and link it
          urlstack++;
          element.editAsText().insertText(charIndex, "⇡");
          //            alert(prevChar+currentChar);
          element.editAsText().setLinkUrl(charIndex, charIndex, currentUrl);
          element.editAsText().setForegroundColor(charIndex, charIndex, "#103DC1");
          element.editAsText().setUnderline(charIndex, charIndex, false);
          charIndex++;
          max++;
          string = element.getText();
          prevChar = currentChar;
          currentChar = "⇡";
        } else if (currentChar !== "⇡" && prevChar === "⇡") {
          // link the previous one.
          element.editAsText().setLinkUrl(charIndex-1, charIndex-1, currentUrl);
          element.editAsText().setForegroundColor(charIndex-1, charIndex-1, "#103DC1");
          element.editAsText().setUnderline(charIndex-1, charIndex-1, false);
        } else {
          // current char is ⇡ and we are in a link - all good
        };
        //if (currentLink !== null) links.push(currentLink);
        //currentLink = {};
        //currentLink.url = String(currentUrl);
        //currentLink.startOffset = charIndex;
        // -
        // In a URL means extend the end of the current link
        // This tracks the end of the url.
        // currentLink.endOffsetInclusive = charIndex;
        element.editAsText().setForegroundColor(charIndex, charIndex, "#103DC1");
        element.editAsText().setUnderline(charIndex, charIndex, false);
      }
      //if (element.editAsText().getForegroundColor(charIndex) !== "#000000") {
      element.editAsText().setForegroundColor(charIndex, charIndex, null);
      //};        
      // End the loop and go again
      previousUrl = currentUrl;
      prevChar = currentChar;
    } // end of loop
    // Handle the end case when final character is a link
    // if (currentLink !== null) links.push(currentLink);
    if (insertEndMarker && currentUrl !== null) {
      // alert("End of text insertion");
      urlstack--;
      /* Block 3 = Block 2 repeated */
      var eou = "⸣";
      if (currentChar !== eou && prevChar !== eou) {
        element.editAsText().insertText(charIndex, eou);
        // alert(prevChar+currentChar);
        element.editAsText().setLinkUrl(charIndex, charIndex, currentUrl);
        element.editAsText().setForegroundColor(charIndex, charIndex, "#103DC1");
        element.editAsText().setUnderline(charIndex, charIndex, false);
        /* not needed as we're done.
        charIndex++;
        max++;
        string = element.getText();
        prevChar = currentChar;
        currentChar = eou;          
        */
      };
    };
    if (urlstack != 0) {
      //alert("urlstack: "+urlstack);
    };
  }  else if (element.getNumChildren) {
    // If not a text element then recursively get links from child elements
    for (var i = 0; i < element.getNumChildren(); i++) {
      linkMarker(element.getChild(i),insertEndMarker);
    }
  }
// return links;
}

// https://stackoverflow.com/questions/18727341/get-all-links-in-a-document
function getAllLinks(element) {
  var rangeBuilder = DocumentApp.getActiveDocument().newRange();
  element = element || DocumentApp.getActiveDocument().getBody();

  // Parse the text iteratively to find the start and end indices for each link
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
      rangeBuilder.addRange(getAllLinks(element.getChild(i)));
    }
  }

  return rangeBuilder.build();
}

/*
// https://stackoverflow.com/questions/18727341/get-all-links-in-a-document
function getAllLinkUrls(element) {
  element = element || DocumentApp.getActiveDocument().getBody();
  var urls = [];
  // Parse the text iteratively to find the start and end indices for each link
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
        if (currentLink !== null) urls.push(String(currentLink.url));
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
        if (currentLink !== null) urls.push(String(currentLink.url));
        currentLink = null;
      }
      // End the loop and go again
      previousUrl = currentUrl;
    }
    // Handle the end case when final character is a link
    if (currentLink !== null) links.push(currentLink);
    if (currentLink !== null) urls.push(String(currentLink.url));
    // Convert the links into a range before returning

    //links.forEach(function(link) {
    //  rangeBuilder.addElement(element, link.startOffset, link.endOffsetInclusive);
    //});
    
  }

  // If not a text element then recursively get links from child elements
  else if (element.getNumChildren) {
    for (var i = 0; i < element.getNumChildren(); i++) {
      // rangeBuilder.addRange(getAllLinks(element.getChild(i)));
       urls.push(getAllLinkUrls(element.getChild(i)));
    }
  }

  // return rangeBuilder.build();
  return urls;
}
*/
//  https://gist.github.com/mogsdad/6518632
/**
 * Get an array of all LinkUrls in the document. The function is
 * recursive, and if no element is provided, it will default to
 * the active document's Body element.
 *
 * @param {Element} element The document element to operate on. 
 * .
 * @returns {Array}         Array of objects, vis
 *                              {element,
 *                               startOffset,
 *                               endOffsetInclusive, 
 *                               url}
 */
function getAllLinkUrls(element) {
  var links = [];
  element = element || DocumentApp.getActiveDocument().getBody();
  
  if (element.getType() === DocumentApp.ElementType.TEXT) {
    var textObj = element.editAsText();
    var text = element.getText();
    var inUrl = false;
    for (var ch=0; ch < text.length; ch++) {
      var url = textObj.getLinkUrl(ch);
      // if (url != null) {
      if (url != null && ch != text.length-1) {
        if (!inUrl) {
          // We are now!
          inUrl = true;
          var curUrl = {};
          curUrl.element = element;
          curUrl.url = String( url ); // grab a copy
          curUrl.startOffset = ch;
        }
        else {
          curUrl.endOffsetInclusive = ch;
        }          
      }
      else {
        if (inUrl) {
          // Not any more, we're not.
          inUrl = false;
          links.push(curUrl);  // add to links
          curUrl = {};
        }
      }
    }
  }
  else {
    // Get number of child elements, for elements that can have child elements. 
    try {
      var numChildren = element.getNumChildren();
    }
    catch (e) {
      numChildren = 0;
    }
    for (var i=0; i<numChildren; i++) {
      links = links.concat(getAllLinkUrls(element.getChild(i)));
    }
  }

  return links;
}
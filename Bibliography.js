function insertUpdateBibliography() {
    const result = checkEmailPreferences('Insert/update bibliography', 'insertUpdateBibliography');
    if (result.stopExecution) {                                                                                                                        
      return 0;             
    }    
  if (HOST_APP == 'docs') {
    universalInsertUpdateBibliography(false, false, false);
  } else {
    insertUpdateBibliographySlides(false, false, false);
  }
  addUsageTrackingRecord('insertUpdateBibliography');
}

// insertUpdateBibliography and prepareForPublishing use the function
function universalInsertUpdateBibliography(validate, getparams, newForestAPI) {
  const ui = DocumentApp.getUi();

  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();

    let result = validateLinks(validate, getparams, true, false, newForestAPI, false);
    let validationSite, zoteroItemKey, zoteroItemGroup, bibLink, zoteroItemKeyParameters, biblTexts = [];
    const textToDetectStartBib = TEXT_TO_DETECT_START_BIB;
    const textToDetectEndBib = TEXT_TO_DETECT_END_BIB;
    let bibReferences = [];
    if (result.status == 'ok') {
      validationSite = result.validationSite;
      zoteroItemKey = result.zoteroItemKey;
      zoteroItemGroup = result.zoteroItemGroup;
      //      bibLink = validationSite + zoteroItemKey;
      zoteroItemKeyParameters = result.zoteroItemKeyParameters;
      targetRefLinks = result.targetRefLinks;
      //Logger.log('targetRefLinks=' + targetRefLinks);

      if (result.bibReferences.length > 0) {
        bibReferences = result.bibReferences;
        // API task
        // Logger.log(bibReferences);
        resultBiblTexts = forestAPIcall(validationSite, zoteroItemKey, zoteroItemGroup, bibReferences, documentId, targetRefLinks);
        if (resultBiblTexts.status == 'error') {
          if (resultBiblTexts.modalWindow === true) {
            showAccessDeniedWindow();
          } else {
            ui.alert(resultBiblTexts.message);
          }
          return 0;
        }
        biblTexts = resultBiblTexts.biblTexts;
        // End. API task
      } else {
        ui.alert('Links for bibliography weren\'t found.');
      }
    } else {
      // Error in validateLinks!
      return 0;
    }

    const body = doc.getBody();

    //If you use usual square brackets, line below will be     let rangeElementStart = body.findText(textToDetectStartBib.replace('[', '\\[').replace(']', '\\]'));
    let rangeElementStart = body.findText(textToDetectStartBib);
    let rangeElementEnd = body.findText(textToDetectEndBib);

    rangeElementStartSecond = body.findText(textToDetectStartBib, rangeElementStart);
    if (rangeElementStartSecond != null) {
      ui.alert('Doc contains more than one ' + textToDetectStartBib);
      return 0;
    }

    rangeElementEndSecond = body.findText(textToDetectEndBib, rangeElementEnd);
    if (rangeElementEndSecond != null) {
      ui.alert('Doc contains more than one ' + textToDetectEndBib);
      return 0;
    }

    let styleStartEnd = new Object();
    styleStartEnd[DocumentApp.Attribute.FONT_FAMILY] = 'Courier';
    if (targetRefLinks == 'zotero') {
      styleStartEnd[DocumentApp.Attribute.FONT_SIZE] = 10;
      styleStartEnd[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    } else {
      styleStartEnd[DocumentApp.Attribute.FONT_SIZE] = 1;
      styleStartEnd[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ffffff';
    }



    if (!rangeElementStart && !rangeElementEnd) {

      body.appendParagraph('Bibliography').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      // body.appendParagraph('This bibliography is available digitally in our evidence library at ').setHeading(DocumentApp.ParagraphHeading.NORMAL)
      //   .appendText(bibLink).setLinkUrl(bibLink);
      body.appendParagraph(textToDetectStartBib).setAttributes(styleStartEnd);
      body.appendParagraph(textToDetectEndBib).setAttributes(styleStartEnd);

      rangeElementStart = body.findText(textToDetectStartBib);
      rangeElementEnd = body.findText(textToDetectEndBib);
    }
    if (!rangeElementStart && rangeElementEnd) {
      ui.alert('Script found  ' + textToDetectEndBib + ' but didn\'t find ' + textToDetectStartBib + '. Add ' + textToDetectStartBib + ' at the begining of existing bibliography.');
      return 0;
    }

    if (rangeElementStart && !rangeElementEnd) {
      ui.alert('Script found ' + textToDetectStartBib + ' but didn\'t find ' + textToDetectEndBib + '. Add ' + textToDetectEndBib + ' to the end of existing bibliography.');
      return 0;
    }

    const startElText = rangeElementStart.getElement();
    const startElParagraph = startElText.getParent().setAttributes(styleStartEnd);
    const grandParent = startElParagraph.getParent();
    const chIndex = grandParent.getChildIndex(startElParagraph);

    rangeElementEnd.getElement().getParent().setAttributes(styleStartEnd);

    // Check [bibliography:start][bibliography:end] situation
    const startEndTheSameEl = startElParagraph.findText(textToDetectEndBib);
    if (startEndTheSameEl != null) {
      // Logger.log('Yes, the same!');
      startElParagraph.asParagraph().setText(textToDetectStartBib);
      body.insertParagraph(chIndex + 1, textToDetectEndBib);
    }
    // End. Check [bibliography:start][bibliography:end] situation

    // Delete paragraphs between [bibliography:start] and [bibliography:end]
    let nextEl = startElParagraph.getNextSibling();
    let notBibPars = false;
    while (nextEl && !notBibPars) {
      nextEl = nextEl.getNextSibling();
      if (nextEl != null) {

        // UNSUPPORTED can't be cast to TEXT
        if (nextEl.getType() != 'UNSUPPORTED') {
          if (nextEl.asText().getText().trim() == textToDetectEndBib) {
            notBibPars = true;
          }
        }

        // Don't remove section break
        if (nextEl.getPreviousSibling().getType() != 'UNSUPPORTED') {
          nextEl.getPreviousSibling().removeFromParent();
        }
      }
    }
    // End. Delete paragraphs between [bibliography:start] and [bibliography:end]

    let endText;
    let startParagraph;
    let workWithPar = true;
    let parIndex = chIndex + 1;

    // Task 5 2021-04-13
    let bibIntroText;
    if (targetRefLinks == 'zotero') {
      bibIntroText = 'The bibliography entries in this document link to our internal Zotero library. In the final version of this document, the links will be replaced with links to ';
      bibLink = validationSite;
    } else {
      bibIntroText = 'This bibliography is available digitally in our evidence library at ';
      bibLink = validationSite + zoteroItemKey;
    }

    const bibParagraph = body.insertParagraph(parIndex, bibIntroText).setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .appendText(bibLink).setLinkUrl(bibLink);

    // Default foreground colour
    const attributes = body.getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);

    // Get the foreground color
    let foregroundColor = attributes[DocumentApp.Attribute.FOREGROUND_COLOR];

    // If foregroundColor is null, it means the default color (usually black) is being used
    if (foregroundColor === null) {
      foregroundColor = '#000000'; // Default black color
    }
    // End. Default foreground colour



    if (targetRefLinks == 'zotero') {
      let additionalText = '. Do not edit the bibliography entries below - edit them in Zotero instead. Any changes that you make to bibliography entries below will be overwritten.';
      startText = bibParagraph.getText().length;
      endText = startText + additionalText.length - 1;
      bibParagraph.appendText(additionalText).setLinkUrl(startText, endText, null).setForegroundColor(startText, endText, foregroundColor);
    }
    parIndex++;
    // End. Task 5 2021-04-13

    let bibEntriesCounter = 0;
    for (let i = 0; i < biblTexts.length; i++) {
      if (biblTexts[i].text == '\n') {
        workWithPar = true;
        bibEntriesCounter++;
      } else {

        if (workWithPar === true) {
          startParagraph = body.insertParagraph(parIndex, biblTexts[i].text).setHeading(DocumentApp.ParagraphHeading.NORMAL).setIndentStart(28.34645669291339);
          startText = startParagraph.getText().length;
          parIndex++;
        } else {
          endText = startText + biblTexts[i].text.length - 1;

          if (biblTexts[i].name == 'a') {
            startParagraph.editAsText().appendText(biblTexts[i].text).setLinkUrl(startText, endText, biblTexts[i].link);
          } else if (biblTexts[i].name == 'i') {
            startParagraph.editAsText().appendText(biblTexts[i].text).setItalic(startText, endText, true).setLinkUrl(startText, endText, null).setForegroundColor(startText, endText, foregroundColor);
          } else {
            startParagraph.editAsText().appendText(biblTexts[i].text).setItalic(startText, endText, false).setLinkUrl(startText, endText, null).setForegroundColor(startText, endText, foregroundColor);
          }
          startText = startText + biblTexts[i].text.length;
        }

        workWithPar = false;
      }
    }
    doc.saveAndClose();

    if (resultBiblTexts.errorsInSomeKeys === true) {
      alert('bZotBib sent ' + bibReferences.length + ' ' + getWordForm(bibReferences.length, 'key') + '  to Forest API. Forest API succesfully returned bibliography ' + getWordForm(bibEntriesCounter, 'entry', 'entries') + ' for ' + bibEntriesCounter + ' ' + getWordForm(bibEntriesCounter, 'key') + ' but there were errors with the remaining ' + getWordForm(bibReferences.length - bibEntriesCounter, 'key') + ':\n' + resultBiblTexts.errorsInSomeKeysMessage);
    }

    return 0;
  }
  catch (error) {
    ui.alert('Error in insertUpdateBibliography. ' + error);
  }
}

function getWordForm(number, singular, plural) {
  // Handle special case where plural is not provided
  if (!plural) {
    plural = singular + 's';
  }
  // Return appropriate form
  return number === 1 ? singular : plural;
}

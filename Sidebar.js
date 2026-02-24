function bibliographySidebar() {
  addUsageTrackingRecord('bibliographySidebar');
  if (userCanCallForestAPI() === false) {
    showAccessDeniedWindow();
    return 0;
  }
  const ui = getUi();
  const html = HtmlService.createHtmlOutputFromFile('Sidebar html').setTitle('Bibliography');
  ui.showSidebar(html);
}

function bibliographyForSidebar(trackUsage = true) {
  const ui = getUi();
  let documentId;
  if (HOST_APP == 'docs') {
    const doc = DocumentApp.getActiveDocument();
    documentId = doc.getId();
  } else {
    documentId = SlidesApp.getActivePresentation().getId();
  }
  let result = validateLinks(validate = false, getparams = false, false, false, true, false);
  let validationSite, zoteroItemKey, zoteroItemGroup, zoteroItemKeyParameters, biblTexts;
  let bibReferences = [];
  if (!result.hasOwnProperty('status')) {
    throw new Error('Error in validateLinks!');
  }

  if (result.status == 'ok') {
    validationSite = result.validationSite;
    zoteroItemKey = result.zoteroItemKey;
    zoteroItemGroup = result.zoteroItemGroup;
    zoteroItemKeyParameters = result.zoteroItemKeyParameters;
    targetRefLinks = result.targetRefLinks;
    //Logger.log('targetRefLinks=' + targetRefLinks);

    if (result.bibReferences.length > 0) {
      bibReferences = result.bibReferences;
      resultBiblTexts = forestAPIcall(validationSite, zoteroItemKey, zoteroItemGroup, bibReferences, documentId, targetRefLinks, mode = 'sidebar');
      if (resultBiblTexts.status == 'error') {
        if (resultBiblTexts.modalWindow === true) {
          showAccessDeniedWindow();
        } else {
          ui.alert(resultBiblTexts.message);
        }
        return 0;
      }
      biblTexts = resultBiblTexts.biblTexts;
    } else {
      // ui.alert('Links for bibliography weren\'t found.');
      //return 0;
      throw new Error('Links for bibliography weren\'t found.');
    }
  } else {
    throw new Error('Error in validateLinks!');
  }

  let workWithPar = true;
  let parIndex = 0;
  let html = '';
  let bibEntriesCounter = 0;
  for (let i = 0; i < biblTexts.length; i++) {
    if (biblTexts[i].text == '\n') {
      workWithPar = true;
      bibEntriesCounter++;
    } else {

      if (workWithPar === true) {
        if (parIndex == 0) {
          html += '<p>';
        } else {
          html += '</p><p>';
        }
        html += biblTexts[i].text;
        parIndex++;
      } else {
        if (biblTexts[i].name == 'a') {
          html += `<a href="${biblTexts[i].link}" target="_blank">${biblTexts[i].text}</a>`;
        } else if (biblTexts[i].name == 'i') {
          html += `<i>${biblTexts[i].text}</i>`;
        } else {
          html += biblTexts[i].text;
        }
      }

      workWithPar = false;
    }
  }

  if (resultBiblTexts.errorsInSomeKeys === true) {
    alert('bZotBib sent ' + bibReferences.length + ' ' + getWordForm(bibReferences.length, 'key') + '  to Forest API. Forest API succesfully returned bibliography ' + getWordForm(bibEntriesCounter, 'entry', 'entries') + ' for ' + bibEntriesCounter + ' ' + getWordForm(bibEntriesCounter, 'key') + ' but there were errors with the remaining ' + getWordForm(bibReferences.length - bibEntriesCounter, 'key') + ':\n' + resultBiblTexts.errorsInSomeKeysMessage);
  }

  if (trackUsage === true) {
    addUsageTrackingRecord('bibliographyForSidebar');
  }
  return html + '</p>';
}
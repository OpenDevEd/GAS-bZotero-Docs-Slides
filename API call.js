function forestAPIcall(validationSite, zoteroItemKey, zoteroItemGroup, bibReferences, docOrPresoId, target, mode = 'bib') {
  try {
    // Logger.log(bibReferences);
    const groupkeys = bibReferences.join(',');
    const activeUser = Session.getActiveUser().getEmail();
    const token = BIBAPI_TOKEN;

    if (userCanCallForestAPI() === false) {
      return { status: 'error', message: 'Access denied! You can\'t use Forest API. Please visit https://opendeved.net/our-tools/bZotero to find out how to use bZotBib.' };
    }

    /* const apiCall = 'https://forest.opendeved.net/api/bib/';
    const options = {
      'method': 'post',
      'payload': JSON.stringify({
        'user': activeUser,
        'zkey': zoteroItemKey,
        'zgroup': zoteroItemGroup,
        'gdoc': docOrPresoId,
        'token': token,
        'groupkeys': groupkeys,
        'target': target,
        'mode': mode
      }),
      'muteHttpExceptions': true
    };
    */

    const apiCall = 'https://forest.opendeved.net/api/v1/getbib/';
    const options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'user': activeUser,
        'zkey': zoteroItemKey,
        'zgroup': zoteroItemGroup,
        'gdoc': docOrPresoId,
        'groupkeys': groupkeys,
        'target': target,
        'mode': mode
      }),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(apiCall, options);
    const code = response.getResponseCode();
    let biblTexts = [];

    if (code == 200) {
      let jsonResponse = JSON.parse(response.getContentText());
      if (jsonResponse.status == 0) {
        /*
        Björn
        jsonResponse = {
          "status": 0,
          "count": 25,
          "duration": 2.645,
          "data": { }
        }
      */
        workOnBibApiElement(1, biblTexts, jsonResponse.data.elements);
        const errorsInSomeKeys = jsonResponse.errors.length > 0 ? true : false;
let errorsInSomeKeysMessage;
        if (errorsInSomeKeys){
          errorsInSomeKeysMessage = JSON.stringify(jsonResponse.errors);
        }
        return { status: 'ok', biblTexts: biblTexts, errorsInSomeKeys, errorsInSomeKeysMessage};
      } else {
        /*
        or 
        {
        "status": 1,
        "message": "message ; on 2021-04-23T22:02:46.242Z",
        }
  
        or
  
        {
        "status": 2,
        "message": "caught error in response ; on 2021-04-23T22:02:46.242Z",
        "error": "TypeError: Cannot read property 'replace' of undefined"
        }
        */
      }
      const messagestring = "Status: " + jsonResponse.status + ". Message: " + jsonResponse.message + ". Error: " + jsonResponse.error;
      return { status: 'error', message: 'Failed to retrieve data from Zotero. ' + messagestring + ". Please let your admins know about this error." };
    } else {
      return { status: 'error', message: 'Call to forestAPI failed. Response Code = ' + code + '. ' + response + ' Please let your admins know about this error.' };
    }
  }
  catch (error) {
    return { status: 'error', message: 'Error in function forestAPIcall. Error: ' + error };
  }
}

function workOnBibApiElement(level, biblTexts, elements, name, link) {

  if (level == 3) {
    biblTexts.push({ text: '\n' });
  }

  elements.forEach(item => {
    if (item.type == 'element') {
      if (!item.attributes) {
        item.attributes = '-';
      } else {
        if (!item.attributes.href) {
          item.attributes = '-';
        }
      }
      workOnBibApiElement(level + 1, biblTexts, item.elements, item.name, item.attributes.href)
    } else {
      //Logger.log(item);
      biblTexts.push({ level: level, text: item.text, name: name, link: link });
    }
  });
}

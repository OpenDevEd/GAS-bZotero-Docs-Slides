function userCanCallForestAPI() {
  const activeUser = Session.getActiveUser().getEmail();
  const activeUserDomain = String(activeUser).split('@')[1];
  if (activeUserDomain != 'edtechhub.org' && activeUserDomain != 'opendeved.net') {
    return false;
  }else{
    return true;
  }
}

function forestAPIcallGetRedirects(validationSite, bibReferences, docOrPresoId) {
  try {

    let groupId, validationSiteRegEx;
    for (let style in styles) {
      if (styles[style]['kerkoValidationSite'] != null) {
        validationSiteRegEx = new RegExp(styles[style]['kerkoValidationSite'], 'i');
        if (validationSiteRegEx.test(validationSite)) {
          groupId = styles[style]['group_id'];
          break;
        }
      }
    }

    const groupkeys = bibReferences.join(',');
    const activeUser = Session.getActiveUser().getEmail();
    const token = BIBAPI_TOKEN;

    if (userCanCallForestAPI() === false) {
      return { status: 'error', message: 'Access denied! You can\'t use Forest API. Please visit https://opendeved.net/our-tools/bZotero to find out how to use bZotBib.', modalWindow: true};
    }

    const apiCall = 'https://forest.opendeved.net/api/v1/getRedirects/';
    const options = {
      'method': 'post',
      'headers': {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'user': activeUser,
        'gdoc': docOrPresoId,
        'group': groupId,
        'keys': groupkeys
      }),
      'muteHttpExceptions': true
    };
    //Logger.log(options);

    const response = UrlFetchApp.fetch(apiCall, options);
    const code = response.getResponseCode();

    if (code == 200) {
      let jsonResponse = JSON.parse(response.getContentText());
      if (jsonResponse.status == 0) {
        //Logger.log(jsonResponse);
        return { status: 'ok', json: jsonResponse };
      }
      const messagestring = "Status: " + jsonResponse.status + ". Message: " + jsonResponse.message + ". Error: " + jsonResponse.error;
      return { status: 'error', message: 'Failed to retrieve data from Zotero. ' + messagestring + ". Please let your admins know about this error." };
    } else {
      return { status: 'error', message: 'Call to forestAPI failed. Response Code = ' + code + '. ' + response + ' Please let your admins know about this error.' };
    }
  }
  catch (error) {
    return { status: 'error', message: 'Error in function forestAPIcallGetRedirects. Error: ' + error };
  }
}
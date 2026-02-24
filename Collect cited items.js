function collectCitedItems() {
  addUsageTrackingRecord('collectCitedItems');
  if (userCanCallForestAPI() === false) {
    showAccessDeniedWindow();
    return 0;
  }

  const ui = getUi();
  try {

    if (userCanCallForestAPI() === false) {
      showAccessDeniedWindow();
      return 0;
    }

    const result = validateLinks(false, false, false, false, false, false);
    if (result.status != 'ok') {
      return 0;
    }
    // Logger.log(result);

    const { validationSite, bibReferences, docOrPresoId, docOrPresoTitle } = result;
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

    const apiCall = 'https://forest.opendeved.net/api/v1/collectciteditems/';
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
        'keys': groupkeys,
        'documenttitle': docOrPresoTitle
      }),
      'muteHttpExceptions': true
    };
    //Logger.log(options);

    const response = UrlFetchApp.fetch(apiCall, options);
    const code = response.getResponseCode();

    if (code == 200) {
      const jsonResponse = JSON.parse(response.getContentText());

      const errorDetails = {
        status: jsonResponse.status || 'Unknown',
        message: jsonResponse.message || 'Malformed message.',
        error: jsonResponse.error || 'No error details'
      };

      if (jsonResponse.message) {
        ui.alert(`Item collection\n${jsonResponse.message}`);
        return 0;
      }
      throw Error(`Unable to collect cited items. Status: ${errorDetails.status} Message: ${errorDetails.message} Error: ${errorDetails.error} Please contact admins.`);
    }
    throw Error(`Unable to collect cited items. Server error. Response code: ${code}`);
  }
  catch (error) {
    ui.alert('Error in function collectCitedItems. ' + error);
  }
}

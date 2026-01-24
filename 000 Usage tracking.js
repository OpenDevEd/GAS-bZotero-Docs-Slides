function addUsageRecord(data) {
  const { utUrl, utKey } = utCredentials();
  const response = UrlFetchApp.fetch(
    utUrl,
    {
      method: 'post',
      headers: {
        'X-Addon-Key': utKey
      },
      contentType: 'application/json',
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    }
  );
  const responseCode = response.getResponseCode();
  if (response.getResponseCode() != 204) {
    throw new Error(responseCode + ' ' + response.getContentText());
  }
}

function addUsageTrackingRecord(functionName) {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const data = {
      email: activeEmail,
      functionName: functionName,
      app: HOST_APP,
      collection: 'usageTracking'
    };
    addUsageRecord(data);
  }
  catch (e) {
    console.error(e);
  }
}

function addInstallationRecord() {
  try {
    let activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail == null || activeEmail === ''){
      activeEmail = 'unknown email';
    }
    const data = {
      email: activeEmail,
      app: HOST_APP,
      collection: 'installations'
    };
    addUsageRecord(data);
  }
  catch (e) {
    console.error(e);
  }
}
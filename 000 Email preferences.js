/**
 * Email preferences management
 * Handles email subscription preferences and syncs with Firebase
 */

const EMAIL_PREFERENCES_KEY = 'EMAIL_PREFERENCES';

/**
 * Get subscription preferences from user properties
 * @returns {Object|null} Subscription preferences object or null if not set
 */
function getEmailPreferences() {
  const userProperties = PropertiesService.getUserProperties();
  const emailPreferencesData = userProperties.getProperty(EMAIL_PREFERENCES_KEY);
  if (emailPreferencesData == null) {
    return null;
  }
  try {
    return JSON.parse(emailPreferencesData);
  } catch (e) {
    console.error('Error parsing email preferences data:', e);
    return null;
  }
}

/**
 * Set subscription preferences in user properties
 * @param {Object} preferences - Subscription preferences object
 */
function setEmailPreferences(preferences) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(EMAIL_PREFERENCES_KEY, JSON.stringify(preferences));
}

/**
 * Clear subscription preferences (for testing)
 */
function clearEmailPreferences() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(EMAIL_PREFERENCES_KEY);
}

/**
 * Log subscription preferences (for testing)
 */
function logEmailPreferences() {
  const preferences = getEmailPreferences();
  Logger.log(preferences);
}

/**
 * Check if subscription preferences are set. If not, show the modal.
 * @param {string} userFriendlyName - User-friendly function name for display
 * @param {string} functionName - Real function name for re-run message
 * @returns {Object} Object with stopExecution property
 */
function checkEmailPreferences(userFriendlyName, functionName) {
  const preferences = getEmailPreferences();
  if (preferences == null) {
    showEmailPreferencesModal(userFriendlyName, functionName);
    return { stopExecution: true };
  }
  return { stopExecution: false };
}

/**
 * Show the subscription preferences modal dialog
 * @param {string} userFriendlyName - User-friendly function name for display
 * @param {string} functionName - Real function name for re-run message
 */
function showEmailPreferencesModal(userFriendlyName, functionName) {
  const template = HtmlService.createTemplateFromFile('000 Email preferences html.html');
  template.userFriendlyName = userFriendlyName || '';
  template.functionName = functionName || '';
  template.currentPreferences = getEmailPreferences();

  const htmlOutput = template.evaluate()
    .setWidth(450)
    .setHeight(450);

  const ui = getUi();
  ui.showModalDialog(htmlOutput, 'Email preferences');
  addUsageTrackingRecord('showEmailPreferencesModal');
}

/**
 * Open subscription modal from Settings menu
 */
function openEmailPreferencesSettings() {
  showEmailPreferencesModal('', 'openEmailPreferencesSettings');
}

/**
 * Handle subscription response from the modal
 * @param {Object} preferences - The subscription preferences from the form
 * @param {string} functionName - The function name that triggered the modal
 * @returns {Object} The saved preferences
 */
function handleEmailPreferencesResponse(preferences, functionName) {
  // Save to user properties
  setEmailPreferences(preferences);

  // Send to Firebase
  try {
    addEmailPreferencesRecord(preferences, functionName);
  } catch (e) {
    console.error('Error sending email preferences to Firebase:', e);
  }

  return preferences;
}

/**
 * Send subscription preferences to Firebase
 * @param {Object} preferences - The subscription preferences
 * @param {string} functionName - The function name that triggered the modal
 */
function addEmailPreferencesRecord(preferences, functionName) {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const data = {
      email: activeEmail,
      functionName: functionName,
      emailPreferences: preferences,
      app: HOST_APP,
      collection: 'emailPreferences'
    };
    addPreferencesRecord(data);
  } catch (e) {
    console.error(e);
  }
}

function addPreferencesRecord(data) {
  const { logPrefUrl, utKey } = utCredentials();
  const response = UrlFetchApp.fetch(
    logPrefUrl,
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

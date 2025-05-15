// --- Constants for Microsoft Graph OAuth ---
var CLIENT_ID; // Populated by initializeCredentials_
var CLIENT_SECRET; // Populated by initializeCredentials_

const AUTHORIZATION_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
const ONEDRIVE_SCOPES = 'Files.ReadWrite offline_access openid profile User.Read';

/**
 * Initializes client ID and secret from script properties.
 * Call this at the beginning of functions that need them if they might not be set.
 */
function initializeCredentials_() {
  // Check if already initialized to avoid redundant property reads
  if (CLIENT_ID && CLIENT_SECRET) {
    return;
  }
  var scriptProps = PropertiesService.getScriptProperties();
  CLIENT_ID = scriptProps.getProperty('MICROSOFT_CLIENT_ID'); // Store your actual Client ID here
  CLIENT_SECRET = scriptProps.getProperty('MICROSOFT_CLIENT_SECRET'); // Store your actual Client Secret here
  if (!CLIENT_ID || !CLIENT_SECRET) {
    Logger.log('CRITICAL ERROR: Client ID or Client Secret not set in Script Properties. Please go to File > Project Properties > Script Properties and add MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET.');
    throw new Error("Configuration Error: Client ID or Client Secret not set in Script Properties.");
  }
  // Logger.log('Credentials Initialized: CLIENT_ID loaded.'); // Optional: for debugging
}

/**
 * Handles GET requests to the web app.
 */
function doGet(e) {
  try {
    initializeCredentials_(); // Ensure credentials are loaded for any path
  } catch (err) {
    Logger.log('Error in doGet during credential initialization: ' + err.message);
    return HtmlService.createHtmlOutput("<b>Configuration Error:</b> " + err.message + " Please check Script Properties.");
  }

  return HtmlService.createHtmlOutputFromFile('PickerPage')
      .setTitle('OneDrive Picker')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Creates and configures the OAuth2 service for Microsoft OneDrive/Graph.
 * @return {OAuth2.Service} The configured OAuth2 service.
 * @private
 */
function getOneDriveService_() {
  initializeCredentials_();

  return OAuth2.createService('microsoftOneDrive')
      .setAuthorizationBaseUrl(AUTHORIZATION_URL)
      .setTokenUrl(TOKEN_URL)
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)
      .setCallbackFunction('authCallback')
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(ONEDRIVE_SCOPES)
      .setParam('prompt', 'select_account');
}

/**
 * Called by the client-side to get the Microsoft Authorization URL.
 * @return {string} The Microsoft Authorization URL.
 */
function getMicrosoftAuthUrl() {
  initializeCredentials_();
  var oneDriveService = getOneDriveService_();
  var mainAppUrl = ScriptApp.getService().getUrl();
  // Pass the main app URL to the callback so it can redirect back correctly
  var authorizationUrl = oneDriveService.getAuthorizationUrl({ MaintargetUrl: mainAppUrl });
  Logger.log('Providing Microsoft Auth URL to client: ' + authorizationUrl);
  return authorizationUrl;
}

/**
 * Handles the OAuth2 callback from Microsoft.
 * @param {Object} request The request data received from the OAuth2 provider.
 * @return {HtmlService.HtmlOutput} A success or failure message page.
 */
function authCallback(request) {
  initializeCredentials_();
  var oneDriveService = getOneDriveService_();
  var authorized = false;
  var lastError = "Unknown error during authorization.";
  // Retrieve the MaintargetUrl passed during the authorization request
  var mainAppUrl = request.parameter.MaintargetUrl;

  if (!mainAppUrl) {
    // Fallback if MaintargetUrl wasn't passed or retrieved, though it should be
    mainAppUrl = ScriptApp.getService().getUrl();
    Logger.log('authCallback: MaintargetUrl not found in request parameters, using default ScriptApp URL.');
  } else {
    Logger.log('authCallback: MaintargetUrl from request: ' + mainAppUrl);
  }

  try {
    authorized = oneDriveService.handleCallback(request);
  } catch (e) {
    Logger.log('Error during handleCallback: ' + e.toString());
    lastError = e.toString();
    authorized = false;
  }

  if (authorized) {
    Logger.log('authCallback: Authorization successful.');
    // Use mainAppUrl for the redirect link
    var successHtml = '<!DOCTYPE html><html><head><title>Success</title></head><body>' +
                      '<h1>Success!</h1>' +
                      '<p>Authentication complete.</p>' +
                      '<p><a href="' + mainAppUrl.replace(/"/g, '"') + '" target="_top">Click here to return to the application.</a></p>' +
                      '<p>You may need to reload the application page or click its main button again.</p>' +
                      '</body></html>';
    return HtmlService.createHtmlOutput(successHtml);
  } else {
    var serviceError = oneDriveService.getLastError();
    if (serviceError) {
        lastError = serviceError;
    }
    Logger.log('authCallback: Authorization failed. Error: ' + lastError);
    var failureHtml = '<!DOCTYPE html><html><head><title>Denied</title></head><body>' +
                      '<h1>Authentication Denied</h1>' +
                      '<p>Authentication failed: ' + lastError + '</p>' +
                      '<p><a href="' + mainAppUrl.replace(/"/g, '"') + '" target="_top">Click here to return to the application and try again.</a></p>' +
                      '</body></html>';
    return HtmlService.createHtmlOutput(failureHtml);
  }
}

/**
 * Gets the stored OneDrive access token.
 * @return {string | null} The access token, or null if not authorized or refresh fails.
 */
function getOneDriveAccessToken() {
  initializeCredentials_();
  var oneDriveService = getOneDriveService_();

  if (oneDriveService.hasAccess()) {
    try {
      var tokenObject = oneDriveService.getToken();
      Logger.log('getOneDriveAccessToken (Server): Full token object from library: ' + JSON.stringify(tokenObject));

      if (tokenObject && typeof tokenObject.access_token === 'string') {
        var accessToken = tokenObject.access_token;
        Logger.log('getOneDriveAccessToken (Server): Extracted access_token (first 30): ' + (accessToken ? accessToken.substring(0,30) : 'N/A') + '...');
        Logger.log('getOneDriveAccessToken (Server): Extracted access_token length: ' + (accessToken ? accessToken.length : 'N/A'));
        return accessToken;
      } else {
        Logger.log('getOneDriveAccessToken (Server): Token object retrieved, but access_token field is missing, not a string, or tokenObject is null. Token object: ' + JSON.stringify(tokenObject));
        return null;
      }
    } catch (e) {
      Logger.log('getOneDriveAccessToken (Server): Error processing token object: ' + e.toString());
      try {
        var rawTokenAttemptOnError = oneDriveService.getToken();
        Logger.log('getOneDriveAccessToken (Server): Raw token object on error (might be object): ' + rawTokenAttemptOnError);
      } catch (e2) {
        Logger.log('getOneDriveAccessToken (Server): Could not get raw token object on error: ' + e2.toString());
      }
      return null;
    }
  } else {
    Logger.log('getOneDriveAccessToken (Server): No access. User needs to authorize or re-authorize.');
    return null;
  }
}

/**
 * Resets the OAuth2 service for the current user.
 */
function resetOneDriveAuth() {
  initializeCredentials_();
  var oneDriveService = getOneDriveService_();
  oneDriveService.reset();
  Logger.log('OneDrive authentication has been reset for the current user.');
}

/**
 * Logs the redirect URI to be registered in Azure AD.
 */
function logOAuthRedirectUri() {
  // No need to initialize real credentials for this, just need the library's logic
  var dummyService = OAuth2.createService('microsoftTempForLog')
      .setClientId('YOUR_CLIENT_ID_PLACEHOLDER_FOR_LOGGING') // Placeholder
      .setCallbackFunction('authCallback');
  Logger.log('Register this redirect URI in Azure AD (Web platform): ' + dummyService.getRedirectUri());
}

/**
 * Exposes the web app's /exec URL to the client-side.
 * @return {string} The script's service URL.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

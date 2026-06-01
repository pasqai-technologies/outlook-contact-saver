// --- BLOCK: MSAL Authentication ---
// Purpose: Handle Microsoft authentication using MSAL.js
// Depends on: msal-browser (CDN), config.js

let msalInstance = null;

async function initMsal() {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: CONFIG.baseUrl + "/index.html",
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  });
  await msalInstance.initialize();
}

async function getToken() {
  const scopes = ["Contacts.ReadWrite", "Files.ReadWrite", "User.Read"];
  const accounts = msalInstance.getAllAccounts();

  if (accounts.length > 0) {
    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes,
        account: accounts[0],
      });
      return result.accessToken;
    } catch (error) {
      if (error instanceof msal.InteractionRequiredAuthError) {
        // Fall through to popup
      } else {
        throw new Error("Authentication failed: " + error.message);
      }
    }
  }

  try {
    const result = await msalInstance.acquireTokenPopup({ scopes });
    return result.accessToken;
  } catch (error) {
    throw new Error("Authentication failed: " + error.message);
  }
}

function isLoggedIn() {
  return msalInstance !== null && msalInstance.getAllAccounts().length > 0;
}

async function signIn() {
  await msalInstance.loginPopup({
    scopes: ["Contacts.ReadWrite", "Files.ReadWrite", "User.Read"],
  });
}

window.Auth = { initMsal, getToken, isLoggedIn, signIn };
// src/authConfig.ts
import { PublicClientApplication, AccountInfo } from "@azure/msal-browser";

/**
 * Centralized MSAL configuration used by both taskpane and dialog.
 */
export const msalConfig = {
  auth: {
    clientId: "cc4403ef-7360-4427-87f5-af7f6f236e2c", // Flowpoint App ID
    authority: "https://login.microsoftonline.com/adf44a4d-b671-4672-ba02-21fdc77f982a",
    redirectUri: "https://localhost:3000/taskpane.html",
  },
};

/**
 * Default login request scopes.
 * Exported because TaskPane.tsx imports it directly.
 */
export const loginRequest = {
  scopes: ["User.Read", "Mail.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"],
};

const msalInstance = new PublicClientApplication(msalConfig);
let initializePromise: Promise<void> | null = null;
let account: AccountInfo | null = null;

/** Ensure the MSAL instance is initialized */
async function ensureInitialized() {
  if (!initializePromise) initializePromise = msalInstance.initialize();
  await initializePromise;
}

/** Acquire (or silently refresh) an access token */
export async function getAccessToken(): Promise<string> {
  await ensureInitialized();

  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    account = accounts[0] || (await msalInstance.loginPopup(loginRequest)).account;
  }

  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return tokenResponse.accessToken;
  } catch {
    const popupResponse = await msalInstance.acquireTokenPopup(loginRequest);
    return popupResponse.accessToken;
  }
}
export const loginAndGetToken = getAccessToken;

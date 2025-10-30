// src/graphSharePoint.ts
import { PublicClientApplication, AccountInfo, AuthenticationResult } from "@azure/msal-browser";

/**
 * ──────────────────────────────────────────────────────────────
 * MSAL setup
 * ──────────────────────────────────────────────────────────────
 * Keep your IDs/URIs as-is. Redirect must match what Outlook loads.
 */
const msalConfig = {
  auth: {
    clientId: "cc4403ef-7360-4427-87f5-af7f6f236e2c",          // ✅ App (client) ID
    authority: "https://login.microsoftonline.com/adf44a4d-b671-4672-ba02-21fdc77f982a", // ✅ Tenant
    redirectUri: "https://localhost:3000/taskpane.html",
  },
};

const loginRequest = {
  scopes: ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"],
};

const pca = new PublicClientApplication(msalConfig);

let initPromise: Promise<void> | null = null;
let account: AccountInfo | null = null;

// We’ll keep token+expiry and refresh as needed
let cachedAuth: { token: string; expiresOn: number } | null = null;

// Site selection (defaults to Technology site as before)
let siteHostname = "dialecticeng.sharepoint.com";
let sitePath = "/sites/Technology";
let siteIdCache: string | null = null;

/**
 * Optionally allow other parts of the app to switch sites later
 * (safe no-op for existing code if never called).
 */
export function setActiveSite(hostname: string, path: string) {
  if (!hostname || !path) return;
  siteHostname = hostname;
  sitePath = path.startsWith("/") ? path : `/${path}`;
  siteIdCache = null; // force re-resolve
}

/* ──────────────────────────────────────────────────────────────
   Auth helpers
   ────────────────────────────────────────────────────────────── */
async function ensureMsalReady() {
  if (!initPromise) initPromise = pca.initialize();
  await initPromise;
}

async function ensureAccount(): Promise<AccountInfo> {
  await ensureMsalReady();

  if (account) return account;

  const accounts = pca.getAllAccounts();
  if (accounts.length) {
    account = accounts[0];
    return account;
  }

  const login = await pca.loginPopup(loginRequest);
  account = login.account!;
  return account!;
}

function isExpiredSoon(expiresOn: number, skewSeconds = 60): boolean {
  return Date.now() >= (expiresOn - skewSeconds * 1000);
}

export async function getAccessToken(forceRefresh = false): Promise<string> {
  const acct = await ensureAccount();

  if (!forceRefresh && cachedAuth && !isExpiredSoon(cachedAuth.expiresOn)) {
    return cachedAuth.token;
  }

  let result: AuthenticationResult;
  try {
    result = await pca.acquireTokenSilent({ ...loginRequest, account: acct });
  } catch {
    result = await pca.acquireTokenPopup(loginRequest);
  }

  cachedAuth = {
    token: result.accessToken,
    expiresOn: (result.expiresOn?.getTime?.() ?? Date.now() + 55 * 60 * 1000), // fallback ~55m
  };

  return cachedAuth.token;
}

/* ──────────────────────────────────────────────────────────────
   Graph Site & Drive helpers
   ────────────────────────────────────────────────────────────── */
export async function getSiteId(): Promise<string> {
  if (siteIdCache) return siteIdCache;

  const token = await getAccessToken();
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteHostname}:${sitePath}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) {
    throw new Error(`Failed to resolve site id: ${resp.status} ${await resp.text()}`);
  }

  const json = await resp.json();
  siteIdCache = json.id;
  return siteIdCache!;
}

export async function getDrives(): Promise<Array<{ id: string; name: string }>> {
  const token = await getAccessToken();
  const siteId = await getSiteId();

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives?$select=id,name`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) throw new Error(`Failed to get drives: ${resp.status} ${await resp.text()}`);
  const json = await resp.json();
  return json.value || [];
}

/* ──────────────────────────────────────────────────────────────
   Drive & Folder retrieval
   ────────────────────────────────────────────────────────────── */
export interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
  folder?: { childCount: number };
  file?: unknown;
}

/**
 * Generic children fetcher that **handles both "root" and any folder ID**.
 */
async function getDriveChildren(
  driveId: string,
  folderId: string,                               // "root" OR actual itemId
  select = "id,name,webUrl,folder,file"
): Promise<DriveItem[]> {
  const token = await getAccessToken();
  const siteId = await getSiteId();

  const path =
    folderId === "root"
      ? `drives/${driveId}/root/children`
      : `drives/${driveId}/items/${folderId}/children`;

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/${path}?$select=${encodeURIComponent(select)}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) {
    throw new Error(`Failed to get children for ${folderId}: ${resp.status} ${await resp.text()}`);
  }
  const json = await resp.json();
  return json.value || [];
}

export async function getDriveRootItems(driveId: string): Promise<DriveItem[]> {
  return getDriveChildren(driveId, "root");
}

export async function getDriveFolderItems(driveId: string, folderId: string): Promise<DriveItem[]> {
  // ✅ FIX: supports "root" or real folder IDs without breaking Task Pane
  return getDriveChildren(driveId, folderId);
}

/**
 * Dialog expects listFolders(token, driveId). We accept the token param
 * for signature compatibility, but we’ll still refresh if needed.
 */
export async function listFolders(
  maybeToken: string | undefined,
  driveId: string
): Promise<Array<{ id: string; name: string }>> {
  // Use provided token if it exists and looks non-empty; otherwise get one.
  const token = (maybeToken && maybeToken.length > 10) ? maybeToken : await getAccessToken();
  const siteId = await getSiteId();

  // Only folders under drive root
  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children?$select=id,name,folder`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!resp.ok) throw new Error(`Failed to list folders: ${resp.status} ${await resp.text()}`);

  const json = await resp.json();
  return (json.value || [])
    .filter((item: any) => !!item.folder)
    .map((i: any) => ({ id: i.id, name: i.name }));
}

/* ──────────────────────────────────────────────────────────────
   Uploads
   ────────────────────────────────────────────────────────────── */
export async function uploadBlobToFolderId(
  driveId: string,
  folderId: string, // can be "root" too; Graph supports /root:/path OR /items/{id}
  fileName: string,
  blob: Blob
): Promise<void> {
  const token = await getAccessToken();
  const siteId = await getSiteId();

  // If folderId is "root", we must use /root:/name:/createUploadSession
  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}`;
  const createUrl =
    folderId === "root"
      ? `${base}/root:/${encodeURIComponent(fileName)}:/createUploadSession`
      : `${base}/items/${folderId}:/${encodeURIComponent(fileName)}:/createUploadSession`;

  const sessionResp = await fetch(createUrl, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: JSON.stringify({ "@microsoft.graph.conflictBehavior": "rename" }),
  });

  if (!sessionResp.ok) {
    throw new Error(`Create upload session failed: ${sessionResp.status} ${await sessionResp.text()}`);
  }

  const session = await sessionResp.json();
  const uploadUrl: string = session.uploadUrl;

  const chunkSize = 5 * 1024 * 1024;
  const arrayBuffer = await blob.arrayBuffer();
  const total = arrayBuffer.byteLength;
  let offset = 0;

  while (offset < total) {
    const end = Math.min(offset + chunkSize, total);
    const slice = arrayBuffer.slice(offset, end);

    const putResp = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(end - offset),
        "Content-Range": `bytes ${offset}-${end - 1}/${total}`,
      },
      body: slice,
    });

    if (![200, 201, 202].includes(putResp.status)) {
      throw new Error(`Upload chunk failed: ${putResp.status} ${await putResp.text()}`);
    }

    offset = end;
  }
}

/* ──────────────────────────────────────────────────────────────
   High-level helpers (used by Dialog / Task Pane)
   ────────────────────────────────────────────────────────────── */
export async function listDocumentLibraries(): Promise<Array<{ id: string; name: string }>> {
  return getDrives();
}

/**
 * Upload an Outlook item as .eml to SharePoint.
 * Expects an Office.js item (passed from TaskPane or Dialog).
 */
export async function uploadEmailToSharePoint(
  _token: string,                 // kept for signature compatibility
  outlookItem: any,
  driveId: string,
  folderId: string
): Promise<void> {
  const safeName = (outlookItem?.subject || "message").toString().replace(/[\\/:*?"<>|]/g, "_");
  const fileName = `${safeName}.eml`;

  return new Promise<void>((resolve, reject) => {
    if (!outlookItem?.getItemAsync) {
      reject("Invalid Outlook item reference.");
      return;
    }

    outlookItem.getItemAsync(async (r: any) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        reject("Failed to read email content.");
        return;
      }

      try {
        const blob = new Blob([r.value], { type: "message/rfc822" });
        await uploadBlobToFolderId(driveId, folderId || "root", fileName, blob);
        resolve();
      } catch (e) {
        reject(e);
      }
    });
  });
}

/**
 * Favorites placeholder (wire up to real storage later)
 */
export async function getFavorites(): Promise<Array<{ id: string; name: string }>> {
  return [];
}

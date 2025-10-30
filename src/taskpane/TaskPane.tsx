// src/taskpane/TaskPane.tsx
import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  FluentProvider,
  webDarkTheme,
  Card,
  Input,
  Button,
  Tooltip,
  Spinner,
  Caption1,
  Combobox,
  Option,
} from "@fluentui/react-components";
import { StarIcon, CloudArrowUpIcon } from "@heroicons/react/24/solid";
import { PublicClientApplication, AccountInfo } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "../authConfig";

/* global Office */

// ──────────────────────────────
// 🎨 Colors
const darkOrange = "#FF8C00";
const yellowStar = "#FFD700";

// ──────────────────────────────
// 📌 SharePoint field internal names we use if present
const FIELD_FROM = "From";
const FIELD_FROM_ADDRESS = "From_x002d_Address";
const FIELD_RECEIVED = "Received";
const FIELD_ATTACHMENT = "Attachment";
const FIELD_ORIGINAL_LINK = "OriginalMessageLink"; // recommended new column
const FIELD_INTERNET_ID = "InternetMessageId"; // recommended new column

// Library mappings list (hosted on the Technology site)
const MAPPINGS_LIST_ID = "9d2d86da-237f-4628-9cf7-65723967018f";
const MAPPINGS_LIST_TITLE = "Dialectic Flowpoint Mappings";

// ──────────────────────────────
// 📌 Types
interface Drive {
  id: string;
  name: string;
}
interface DriveItem {
  id: string;
  name: string;
  webUrl: string;
  folder?: { childCount: number };
  file?: any;
}
interface FolderStackEntry {
  id: string;
  name: string;
}
interface Favorite {
  id: string;
  name: string;
  driveId: string;
  path: string;
}
interface LibraryProfile {
  DepartmentOrGroup: string;
  Label: string;
  SiteUrl: string;
  DriveId: string;
  SortOrder?: number;
  IsDefault?: boolean;
}

// ──────────────────────────────
// 🔧 Utilities
const encodeDrivePathForGraph = (path: string) =>
  path
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/");

const safeFileNameFromSubject = (subject: string) =>
  `${(subject || "Email").replace(/[^a-z0-9\-_. ]/gi, "_")}.eml`;

const msalInstance = new PublicClientApplication(msalConfig);

// generic retry for fetches
async function withRetry<T>(fn: () => Promise<T>, max = 4, baseDelayMs = 500): Promise<T> {
  let attempt = 0;
  let delay = baseDelayMs;
  while (true) {
    try {
      return await fn();
    } catch (e: any) {
      const status = e?.status ?? e?.response?.status;
      const retriable = status === 429 || (status >= 500 && status <= 599);
      attempt++;
      if (!retriable || attempt >= max) throw e;
      await new Promise((r) => setTimeout(r, delay));
      delay *= 2;
    }
  }
}

// Basic GET returning JSON with bearer and automatic retry on 429/5xx
async function graphGET<T>(url: string, token: string): Promise<T> {
  return withRetry(async () => {
    const resp = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!resp.ok) {
      const err: any = new Error(`GET ${url} failed: ${resp.status}`);
      err.status = resp.status;
      err.body = await resp.text().catch(() => "");
      throw err;
    }
    return resp.json();
  });
}

// Resolve site ID from a full SiteUrl using Graph
async function getSiteIdFromUrl(siteUrl: string, token: string): Promise<string> {
  // e.g. siteUrl: https://dialecticeng.sharepoint.com/sites/Accounting
  const u = new URL(siteUrl);
  const host = u.host; // dialecticeng.sharepoint.com
  const path = u.pathname; // /sites/Accounting
  const data = await graphGET<{ id: string }>(
    `https://graph.microsoft.com/v1.0/sites/${host}:${path}`,
    token
  );
  return data.id;
}

// Drives & folders (per-site)
async function getDrives(token: string, siteId: string): Promise<Drive[]> {
  const data = await graphGET<{ value: Array<{ id: string; name: string }> }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    token
  );
  return (data.value || []).map((d) => ({ id: d.id, name: d.name }));
}
async function getDriveRootItems(token: string, siteId: string, driveId: string): Promise<DriveItem[]> {
  const data = await graphGET<{ value: DriveItem[] }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root/children`,
    token
  );
  return data.value || [];
}
async function getDriveFolderItems(
  token: string,
  siteId: string,
  driveId: string,
  folderItemId: string
): Promise<DriveItem[]> {
  const data = await graphGET<{ value: DriveItem[] }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${folderItemId}/children`,
    token
  );
  return data.value || [];
}

// Columns for metadata patching
async function getDriveListFieldNames(
  token: string,
  siteId: string,
  driveId: string
): Promise<Set<string>> {
  const cols = await graphGET<{ value: Array<{ name: string }> }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/list/columns?$select=name`,
    token
  );
  return new Set((cols.value || []).map((c) => c.name));
}

// ──────────────────────────────
// 🧠 Component
export default function TaskPane() {
  // 📨 Email info
  const [emailFrom, setEmailFrom] = useState("");
  const [emailSubject, setEmailSubject] = useState("");

  // 📂 Drives & folders
  const [drives, setDrives] = useState<Drive[]>([]);
  const [selectedDriveName, setSelectedDriveName] = useState<string | null>(null);
  const [selectedDriveId, setSelectedDriveId] = useState<string | null>(null);
  const [driveItems, setDriveItems] = useState<DriveItem[]>([]);
  const [filteredItems, setFilteredItems] = useState<DriveItem[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [folderStack, setFolderStack] = useState<FolderStackEntry[]>([]);

  // ⭐ Favorites
  const [favorites, setFavorites] = useState<Favorite[]>([]);

  // Auth / Graph
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [token, setToken] = useState<string>("");

  // A default "Technology" site context (used for mappings fetch)
  const [techSiteId, setTechSiteId] = useState<string>("");

  // Active site for the currently-selected library (from mappings via Combobox)
  const [activeSiteId, setActiveSiteId] = useState<string>("");

  // UI
  const [loading, setLoading] = useState(false);
  const [libCollapsed, setLibCollapsed] = useState(false); // repurposed: collapse Library Selection card
  const [headerCollapsed, setHeaderCollapsed] = useState(false);
  const [statusMsg, setStatusMsg] = useState<string>("");
  const [favoritesCollapsed, setFavoritesCollapsed] = useState(false);
 
 // Bulk archive state
  const [bulkFailed, setBulkFailed] = useState<Array<{ restId: string; error: string }>>([]);
  const [bulkLog, setBulkLog] = useState<string>("");

  // Column discovery
  const [availableFields, setAvailableFields] = useState<Set<string>>(new Set());

  // Library mappings
  const [libraryProfiles, setLibraryProfiles] = useState<LibraryProfile[]>([]);
  const [mappingsLoading, setMappingsLoading] = useState<boolean>(false);
  const [mappingsError, setMappingsError] = useState<string>("");
  const [userDepartment, setUserDepartment] = useState<string>("");

  // Library selection
  const [selectedLibraryProfile, setSelectedLibraryProfile] = useState<LibraryProfile | null>(null);
  const [libraryAccessError, setLibraryAccessError] = useState<string>("");

  const currentPath = useMemo(() => folderStack.map((f) => f.name).join("/"), [folderStack]);

  // convenience: whichever siteId is relevant for file ops
  const getCurrentSiteId = () => activeSiteId || techSiteId;

  // ──────────────────────────────
// ──────────────────────────────
// Office item info (safe initialization for Classic Outlook)
useEffect(() => {
  let handlerAdded = false;

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      console.warn("⚠️ Office.js is loaded outside Outlook:", info.host);
      return;
    }

    console.log("✅ Office is ready in Outlook context");

    const tryInitializeItem = () => {
      const mailbox = Office.context?.mailbox;
      const item = mailbox?.item as Office.MessageRead | undefined;

      if (!mailbox) {
        console.warn("⚠️ Office.context.mailbox is not yet available, retrying...");
        setTimeout(tryInitializeItem, 800);
        return;
      }

      if (!item) {
        console.warn("⚠️ No active item yet, retrying...");
        setTimeout(tryInitializeItem, 800);
        return;
      }

      console.log("📬 Mailbox and item available:", item.itemId);

      // Initialize email info
      setEmailFrom(item?.from?.emailAddress || "");
      setEmailSubject(item?.subject || "");

      // Add handler only once
      if (!handlerAdded) {
        handlerAdded = true;
        mailbox.addHandlerAsync(
          Office.EventType.ItemChanged,
          () => {
            const updated = mailbox.item as Office.MessageRead;
            console.log("📨 Item changed:", updated.itemId);
            setEmailFrom(updated?.from?.emailAddress || "");
            setEmailSubject(updated?.subject || "");
          },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("✅ ItemChanged handler attached successfully");
            } else {
              console.error("❌ Failed to attach handler:", asyncResult.error);
            }
          }
        );
      }
    };

    // Kick off first attempt
    tryInitializeItem();
  });
}, []);


  // Favorites load/persist
  useEffect(() => {
    const saved = localStorage.getItem("flowpoint:favorites");
    if (saved) setFavorites(JSON.parse(saved));
  }, []);
  useEffect(() => {
    localStorage.setItem("flowpoint:favorites", JSON.stringify(favorites));
  }, [favorites]);

  // MSAL init
  useEffect(() => {
    (async () => {
      try {
        await msalInstance.initialize();
        const redirectResult = await msalInstance.handleRedirectPromise();
        if (redirectResult?.account) {
          setAccount(redirectResult.account);
        } else {
          const accounts = msalInstance.getAllAccounts();
          if (accounts.length > 0) setAccount(accounts[0]);
        }
      } catch (e) {
        console.error("MSAL init error:", e);
      }
    })();
  }, []);

  // Acquire token
  useEffect(() => {
    if (!account) return;
    (async () => {
      try {
        const result = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
        setToken(result.accessToken);
        setHeaderCollapsed(true);
      } catch {
        try {
          const result = await msalInstance.acquireTokenPopup(loginRequest);
          setAccount(result.account!);
          setToken(result.accessToken);
          setHeaderCollapsed(true);
        } catch (e) {
          console.error("Token acquisition failed:", e);
        }
      }
    })();
  }, [account]);

  // Resolve the "Technology" site id once (used for fetching the mappings list)
  useEffect(() => {
    if (!token || techSiteId) return;
    (async () => {
      try {
        const site = await graphGET<{ id: string }>(
          "https://graph.microsoft.com/v1.0/sites/dialecticeng.sharepoint.com:/sites/Technology",
          token
        );
        setTechSiteId(site.id);
      } catch (e) {
        console.error("Failed resolving Technology site id:", e);
      }
    })();
  }, [token, techSiteId]);

  // Filter search
  useEffect(() => {
    if (!searchQuery) setFilteredItems(driveItems);
    else {
      const q = searchQuery.toLowerCase();
      setFilteredItems(driveItems.filter((item) => item.name.toLowerCase().includes(q)));
    }
  }, [searchQuery, driveItems]);

  // Load library mappings once we have the Technology site
  useEffect(() => {
    if (!token || !techSiteId) return;
    (async () => {
      try {
        setMappingsLoading(true);
        setMappingsError("");
        const url = `https://graph.microsoft.com/v1.0/sites/${techSiteId}/lists/${MAPPINGS_LIST_ID}/items?expand=fields`;
        const data = await graphGET<{ value: Array<{ fields: any }> }>(url, token);
        const profiles: LibraryProfile[] = (data.value || []).map((item) => {
          const f = item.fields;
          return {
            DepartmentOrGroup: f.DepartmentOrGroup || "",
            Label: f.Label || "",
            SiteUrl: f.SiteUrl || "",
            DriveId: f.DriveId || "",
            SortOrder: f.SortOrder ? Number(f.SortOrder) : undefined,
            IsDefault: f.IsDefault === true || f.IsDefault === "true",
          };
        });
        setLibraryProfiles(profiles);
      } catch (err) {
        console.error("Error fetching library mappings:", err);
        setMappingsError("Failed to load mappings.");
      } finally {
        setMappingsLoading(false);
      }
    })();
  }, [token, techSiteId]);

  // Fetch user's department (for auto-select)
  useEffect(() => {
    if (!token) return;
    (async () => {
      try {
        const me = await graphGET<{ department?: string }>(
          "https://graph.microsoft.com/v1.0/me?$select=department",
          token
        );
        if (me.department) setUserDepartment(me.department);
      } catch (err) {
        console.error("Failed to fetch user department:", err);
      }
    })();
  }, [token]);

  // 📂 Handle selecting/opening a document library from the dropdown (auto-trigger)
  const handleLibraryOpen = async (profile: LibraryProfile | null) => {
    if (!profile || !token) return;
    try {
      setLoading(true);
      const siteId = await getSiteIdFromUrl(profile.SiteUrl, token);

      setActiveSiteId(siteId);
      setSelectedDriveId(profile.DriveId);
      setSelectedDriveName(profile.Label);
      setFolderStack([]);
      setDriveItems([]);
      setFilteredItems([]);
      setSearchQuery("");

      await refreshAvailableFieldsForDrive(profile.DriveId, siteId);
      const items = await getDriveRootItems(token, siteId, profile.DriveId);
      setDriveItems(items);
      setFilteredItems(items);

      // Keep drives in sync with the selected mapping (for consistency with previous UI)
      setDrives([{ id: profile.DriveId, name: profile.Label }]);
      setLibraryAccessError("");
    } catch (err) {
      console.error("Error accessing selected library:", err);
      setLibraryAccessError("🚫 You don’t have permission to access this document library.");
      setSelectedLibraryProfile(null);
      setSelectedDriveId(null);
      setSelectedDriveName(null);
      setActiveSiteId("");
      setDriveItems([]);
      setFilteredItems([]);
    } finally {
      setLoading(false);
    }
  };

  // Auto-select library by dept or default (and auto-open it)
  useEffect(() => {
    if (!libraryProfiles.length || !userDepartment || selectedLibraryProfile) return;
    const match = libraryProfiles.find(
      (p) => p.DepartmentOrGroup?.toLowerCase() === userDepartment.toLowerCase()
    );
    const profileToUse = match || libraryProfiles.find((p) => p.IsDefault);
    if (!profileToUse) return;

    (async () => {
      try {
        setSelectedLibraryProfile(profileToUse);
        await handleLibraryOpen(profileToUse);
      } catch (err) {
        // handleLibraryOpen logs & sets UI state already
      }
    })();
  }, [libraryProfiles, userDepartment, selectedLibraryProfile, token]);

  // ──────────────────────────────
  // Helpers
  function buildPathWithLeaf(leafName?: string) {
    const names = folderStack.map((f) => f.name);
    if (leafName) names.push(leafName);
    return names.join("/");
  }
  function setStatus(text: string) {
    setStatusMsg(text);
    if (text.toLowerCase().includes("complete") || text.toLowerCase().includes("uploaded")) {
      setTimeout(() => setStatusMsg(""), 3000);
    }
  }

  // Multi-select detection
  function getSelectedMessageRestIds(): Promise<string[]> {
    return new Promise((resolve) => {
      const mbox: any = Office?.context?.mailbox;
      const fallbackToCurrent = () => {
        const current = mbox?.item as Office.MessageRead | undefined;
        if (current?.itemId) {
          const restId = mbox.convertToRestId(current.itemId, Office.MailboxEnums.RestVersion.v2_0);
          resolve([restId]);
        } else resolve([]);
      };

      try {
        const hasApi = !!mbox && typeof mbox.getSelectedItemsAsync === "function";
        const supportsReq = !!Office?.context?.requirements?.isSetSupported?.("Mailbox", "1.13");
        if (!hasApi || !supportsReq) return fallbackToCurrent();

        mbox.getSelectedItemsAsync((asyncResult: Office.AsyncResult<any>) => {
          if (
            asyncResult?.status === Office.AsyncResultStatus.Succeeded &&
            Array.isArray(asyncResult.value) &&
            asyncResult.value.length > 1
          ) {
            const restIds = asyncResult.value
              .map((it: any) => it?.itemId)
              .filter(Boolean)
              .map((ewsId: string) =>
                mbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0)
              );
            resolve(restIds);
          } else {
            fallbackToCurrent();
          }
        });
      } catch {
        fallbackToCurrent();
      }
    });
  }

  // list fields for metadata patch
  async function refreshAvailableFieldsForDrive(driveId: string, siteIdToUse?: string) {
    try {
      const sid = siteIdToUse || getCurrentSiteId();
      const fields = await getDriveListFieldNames(token, sid, driveId);
      setAvailableFields(fields);
    } catch (e) {
      console.warn("Could not load list columns; will patch only core fields.", e);
      setAvailableFields(new Set());
    }
  }

  // Dedupe by InternetMessageId (library-wide)
  async function existsByInternetMessageId(driveId: string, internetMessageId: string): Promise<boolean> {
    try {
      if (!internetMessageId) return false;
      if (!availableFields.has(FIELD_INTERNET_ID)) return false;

      const sid = getCurrentSiteId();
      // IMPORTANT: internetMessageId can include '<', '>' — we single-quote escape
      const filterVal = encodeURIComponent(internetMessageId).replace(/'/g, "''");
      const url =
        `https://graph.microsoft.com/v1.0/sites/${sid}/drives/${driveId}/list/items` +
        `?$filter=fields/${FIELD_INTERNET_ID} eq '${filterVal}'&$top=1&$select=id`;

      const data = await graphGET<{ value: any[] }>(url, token);
      return (data.value || []).length > 0;
    } catch (e) {
      console.warn("Dedupe check failed; proceeding without blocking.", e);
      return false;
    }
  }

  // ──────────────────────────────
  // 📦 Bulk archive helpers

  // Logs bulk actions to C:\FlowpointBulkLogs\flowpoint-bulk-archive.log
  async function appendToBulkLog(logText: string) {
    try {
      const logFileName = `flowpoint-bulk-archive.log`;
      const logDir = `C:\\FlowpointBulkLogs`;

      // For browser/Outlook Add-in context we can't truly write to disk directly,
      // but we can use OfficeRuntime.storage or trigger download as a fallback.
      // This stub allows us to later hook into native agent or Graph API.
      console.log(`[BULK LOG] ${logText}`);

      // If in future we build a desktop agent, replace this stub with fs append.
    } catch (e) {
      console.warn("Bulk log write failed:", e);
    }
  }

  // Retry a single failed item
  async function retryFailedBulkItem(restId: string, driveId: string, folderPath: string) {
    try {
      await archiveMessageByRestId(restId, driveId, folderPath);
      setBulkFailed((prev) => prev.filter((f) => f.restId !== restId));
      await appendToBulkLog(`✅ Retry succeeded for message ${restId}`);
    } catch (e: any) {
      console.error("Retry failed:", e);
      await appendToBulkLog(`❌ Retry failed for message ${restId}: ${e?.message || e}`);
    }
  }

  // Sequentially archive multiple messages with retry tracking
  async function bulkArchiveMessagesSequential(
    restIds: string[],
    driveId: string,
    folderPath: string
  ) {
    let success = 0;
    let failed: Array<{ restId: string; error: string }> = [];

    for (let i = 0; i < restIds.length; i++) {
      const id = restIds[i];
      setStatus(`Uploading ${i + 1} of ${restIds.length}…`);
      try {
        await archiveMessageByRestId(id, driveId, folderPath);
        success++;
        await appendToBulkLog(`✅ Archived ${id}`);
      } catch (e: any) {
        console.error(`Bulk item ${id} failed:`, e);
        failed.push({ restId: id, error: e?.message || String(e) });
        await appendToBulkLog(`❌ Failed to archive ${id}: ${e?.message || e}`);
      }
    }

    setBulkFailed(failed);

    if (failed.length === 0) {
      setStatus(`All ${success} uploaded ✅`);
      await appendToBulkLog(`🎉 Bulk upload complete: ${success} succeeded, 0 failed`);
    } else {
      setStatus(`✅ ${success} uploaded, ❌ ${failed.length} failed`);
      await appendToBulkLog(
        `⚠️ Bulk upload finished: ${success} succeeded, ${failed.length} failed`
      );
    }
  }

  // Archive a single message by REST id
  async function archiveMessageByRestId(
    messageRestId: string,
    driveId: string,
    folderPath: string
  ): Promise<void> {
    const sid = getCurrentSiteId();

    // Get message details
    const msg = await graphGET<{
      subject?: string;
      from?: { emailAddress?: { address?: string; name?: string } };
      hasAttachments?: boolean;
      receivedDateTime?: string;
      webLink?: string;
      internetMessageId?: string;
    }>(
      `https://graph.microsoft.com/v1.0/me/messages/${messageRestId}?$select=subject,from,hasAttachments,receivedDateTime,webLink,internetMessageId`,
      token
    );

    // Optional de-dupe
    if (await existsByInternetMessageId(driveId, msg.internetMessageId || "")) {
      return; // already archived
    }

    // Download MIME (.eml)
    const emlBlob = await withRetry(async () => {
      const r = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${messageRestId}/$value`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!r.ok) {
        const err: any = new Error(`Message download failed: ${r.status}`);
        err.status = r.status;
        throw err;
      }
      return r.blob();
    });

    // Create upload session
    const fileName = safeFileNameFromSubject(msg.subject || "Email");
    const encodedPath = encodeDrivePathForGraph(folderPath ? `${folderPath}/${fileName}` : fileName);

    const session = await withRetry(async () => {
      const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${sid}/drives/${driveId}/root:/${encodedPath}:/createUploadSession`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ "@microsoft.graph.conflictBehavior": "rename" }),
        }
      );
      if (!resp.ok) {
        const err: any = new Error(`Create session failed: ${resp.status}`);
        err.status = resp.status;
        throw err;
      }
      return resp.json();
    });

    const uploadUrl: string = session.uploadUrl;

    // Upload chunks
    await uploadInChunksWithRetry(uploadUrl, emlBlob);

    // Resolve uploaded item
    const pathForFetch = folderPath?.length
      ? `/${encodeDrivePathForGraph(folderPath)}/${encodeURIComponent(fileName)}`
      : `/${encodeURIComponent(fileName)}`;

    const uploadedItem = await withRetry(async () => {
      const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${sid}/drives/${driveId}/root:${pathForFetch}`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!resp.ok) {
        const err: any = new Error(`Resolve uploaded item failed: ${resp.status}`);
        err.status = resp.status;
        throw err;
      }
      return resp.json();
    });

    // Build metadata payload
    const fieldsToPatch: Record<string, any> = {};
    fieldsToPatch[FIELD_FROM_ADDRESS] = msg?.from?.emailAddress?.address || "";
    fieldsToPatch[FIELD_RECEIVED] = msg?.receivedDateTime ?? new Date().toISOString();
    fieldsToPatch[FIELD_ATTACHMENT] = !!msg?.hasAttachments;
    fieldsToPatch[FIELD_FROM] = msg?.from?.emailAddress?.name || "";

    if (availableFields.has(FIELD_ORIGINAL_LINK) && msg.webLink) {
      fieldsToPatch[FIELD_ORIGINAL_LINK] = msg.webLink;
    }
    if (availableFields.has(FIELD_INTERNET_ID) && msg.internetMessageId) {
      fieldsToPatch[FIELD_INTERNET_ID] = msg.internetMessageId;
    }

    const allowed = new Set([
      FIELD_FROM,
      FIELD_FROM_ADDRESS,
      FIELD_RECEIVED,
      FIELD_ATTACHMENT,
      ...Array.from(availableFields),
    ]);
    const payload: Record<string, any> = {};
    Object.keys(fieldsToPatch).forEach((k) => {
      if (allowed.has(k)) payload[k] = fieldsToPatch[k];
    });

    await withRetry(async () => {
      const patchResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${sid}/drives/${driveId}/items/${uploadedItem.id}/listItem/fields`,
        {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(payload),
        }
      );
      if (!patchResp.ok) {
        const err: any = new Error(`Metadata update failed: ${patchResp.status}`);
        err.status = patchResp.status;
        throw err;
      }
      return;
    });
  }

  // Chunked upload
  async function uploadInChunksWithRetry(uploadUrl: string, blob: Blob) {
    const chunkSize = 5 * 1024 * 1024; // 5 MB
       const total = blob.size;
    let offset = 0;
    while (offset < total) {
      const slice = blob.slice(offset, Math.min(offset + chunkSize, total));
      const end = offset + slice.size - 1;
      const headers: Record<string, string> = {
        "Content-Length": String(slice.size),
        "Content-Range": `bytes ${offset}-${end}/${total}`,
      };
      await withRetry(async () => {
        const putResp = await fetch(uploadUrl, { method: "PUT", headers, body: slice });
        if (!(putResp.ok || putResp.status === 202 || putResp.status === 201)) {
          const text = await putResp.text();
          const err: any = new Error(`Chunk upload failed ${putResp.status}: ${text}`);
          err.status = putResp.status;
          throw err;
        }
        return;
      });
      offset += slice.size;
    }
  }

  // Public handler for archive actions
  const handleArchiveToPath = async (driveId: string, folderPath: string) => {
    try {
      if (!token) throw new Error("Not authenticated to Graph.");
      if (!driveId) throw new Error("No drive selected.");
      const sid = getCurrentSiteId();
      if (!sid) throw new Error("SiteId not resolved yet.");

      const restIds = await getSelectedMessageRestIds();
      const total = restIds.length;
	  console.log("Selected REST IDs:", restIds);

      if (total <= 1) {
        setStatus("Preparing email…");
        const mbox: any = Office.context.mailbox;
        const currentId = (mbox?.item as Office.MessageRead | undefined)?.itemId;
        const restId =
          restIds[0] ||
          (currentId ? mbox.convertToRestId(currentId, Office.MailboxEnums.RestVersion.v2_0) : "");
        if (!restId) throw new Error("No email selected.");
        setStatus("Uploading email…");
        await archiveMessageByRestId(restId, driveId, folderPath);
        setStatus("Upload complete ✅");
        return;
      }

      // 📦 Bulk sequential upload using new helper
      await bulkArchiveMessagesSequential(restIds, driveId, folderPath);
    } catch (err: any) {
      console.error("Archive failed:", err);
      setStatus(`Upload failed: ${err?.message || err}`);
    }
  };
// ──────────────────────────────
// 🪟 Bulk Archive Dialog Launcher (sends favorites from Taskpane context)
// ──────────────────────────────
// 🪟 Bulk Archive Dialog Launcher (sends favorites + selected message IDs)
async function openBulkArchiveDialog() {
  try {
    const dialogUrl = `${window.location.origin}/dialog.html`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 55, width: 40 },
      async (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("❌ Dialog launch failed:", asyncResult.error);
          return;
        }

        const dialog = asyncResult.value;
        console.log("✅ Bulk Archive dialog opened from Taskpane:", dialogUrl);

        // Listen for messages returned from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          console.log("📩 Message received from dialog:", msg);
        });

        const mbox: any = Office.context?.mailbox;
        let restIds: string[] = [];

        // ✅ Try multi-select (only works in list view)
        if (mbox?.getSelectedItemsAsync) {
          await new Promise<void>((resolve) => {
            mbox.getSelectedItemsAsync((res: any) => {
              if (
                res.status === Office.AsyncResultStatus.Succeeded &&
                Array.isArray(res.value) &&
                res.value.length > 0
              ) {
                restIds = res.value
                  .slice(0, 10)
                  .map((it: any) =>
                    mbox.convertToRestId(it.itemId, Office.MailboxEnums.RestVersion.v2_0)
                  );
              }
              resolve();
            });
          });
        }

        // ✅ Fallback: use the currently opened email if no multi-select
        if (restIds.length === 0 && mbox?.item?.itemId) {
          const restId = mbox.convertToRestId(
            mbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
          );
          restIds = [restId];
          console.log("📬 Using current email REST ID:", restId);
        }

        console.log("📦 Final REST IDs prepared for Bulk Archive:", restIds);

        // ✅ Combine favorites + REST IDs into one payload
        const payload = JSON.stringify({
          favorites: JSON.parse(localStorage.getItem("flowpoint:favorites") || "[]"),
          restIds,
        });

        // Send after short delay to ensure dialog listener is ready
        setTimeout(() => {
          try {
            dialog.messageChild(payload);
            console.log("✅ Sent payload to dialog:", payload);
          } catch (err) {
            console.error("❌ Failed to send payload to dialog:", err);
          }
        }, 1200);
      }
    );
  } catch (err) {
    console.error("❌ Error launching Bulk Archive dialog:", err);
  }
}
(window as any).openBulkArchiveDialog = openBulkArchiveDialog;

  // ──────────────────────────────
  // UI
  return (
    <FluentProvider theme={webDarkTheme} style={{ height: "100%", backgroundColor: "#2b2b2b" }}>
      {/* 🌐 Inline Slim Scrollbar Styles */}
            {/* 🌐 Inline Slim Scrollbar Styles */}
      <style>{`
        /* Universal scrollbar style */
        html, body {
          scrollbar-width: thin;
          scrollbar-color: #6b6b6b #2b2b2b;
        }
        ::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        ::-webkit-scrollbar-track {
          background: #2b2b2b;
        }
        ::-webkit-scrollbar-thumb {
          background-color: #6b6b6b;
          border-radius: 3px;
        }
        ::-webkit-scrollbar-thumb:hover {
          background-color: #a0a0a0;
        }

        /* 🔸 Slim scrollbars for inner lists */
        .scroll-section {
          overflow-y: auto;
          scrollbar-width: thin;
          scrollbar-color: #6b6b6b transparent;
        }
        .scroll-section::-webkit-scrollbar {
          width: 3px;
        }
        .scroll-section::-webkit-scrollbar-thumb {
          background-color: #6b6b6b;
          border-radius: 3px;
          transition: background-color 0.2s;
        }
        .scroll-section::-webkit-scrollbar-thumb:hover {
          background-color: #a0a0a0;
        }

        .favorites-scroll {
          max-height: 180px;
        }
        .library-scroll {
          max-height: 320px;
        }
      `}</style>


      <div style={{ display: "flex", flexDirection: "column", height: "100%", padding: 8, overflow: "hidden" }}>
        {/* 🔸 Header */}
       <Card style={{ padding: "0.35rem 0.5rem" }}>
  <div
    style={{
      display: "flex",
      justifyContent: "flex-start",
      alignItems: "center",
    }}
  >
    {!account ? (
      <Button
        size="small"
        appearance="primary"
        onClick={async () => {
          try {
            const result = await msalInstance.loginPopup(loginRequest);
            setAccount(result.account!);
            setToken(result.accessToken);
            setHeaderCollapsed(true);
          } catch (e) {
            console.error("Login failed:", e);
          }
        }}
      >
        Sign in
      </Button>
    ) : (
      <Button
        size="small"
        appearance="secondary"
        onClick={async () => {
          try {
            await msalInstance.logoutPopup();
          } catch (e) {
            console.error("Logout failed:", e);
          }
          setAccount(null);
          setToken("");
          setTechSiteId("");
          setActiveSiteId("");
          setDrives([]);
          setDriveItems([]);
          setFilteredItems([]);
          setFolderStack([]);
          setHeaderCollapsed(false);
        }}
      >
        Sign out
      </Button>
    )}
  </div>
</Card>

        {/* 📜 Scrollable Content Wrapper */}
        <div
          style={{
            overflowY: "auto",
            flexGrow: 1,
            display: "flex",
            flexDirection: "column",
            gap: 8,
            paddingTop: 8,
            paddingRight: 6,
            minHeight: 0,
          }}
        >
          {/* 📨 Email Info */}
          <Card style={{ padding: "0.3rem 0.5rem", lineHeight: 1.2, fontSize: "0.9rem" }}>
            <div style={{ marginBottom: "0.25rem" }}>
              <span style={{ color: darkOrange, fontSize: "0.8rem", display: "block" }}>From:</span>
              <span style={{ color: "white", wordBreak: "break-all" }}>{emailFrom || "—"}</span>
            </div>
            <div>
              <span style={{ color: darkOrange, fontSize: "0.8rem", display: "block" }}>Subject:</span>
              <span style={{ color: "white", wordBreak: "break-word" }}>{emailSubject || "—"}</span>
            </div>
          </Card>

          {/* ⭐ Favorites */}
          <Card style={{ padding: "0.5rem" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3 style={{ color: darkOrange, fontWeight: "bold", display: "flex", alignItems: "center", gap: 6 }}>
                <StarIcon style={{ width: 18, height: 18, color: darkOrange }} />
                Favorites
              </h3>
              <Button size="small" onClick={() => setFavoritesCollapsed(!favoritesCollapsed)}>
                {favoritesCollapsed ? "▼" : "▲"}
              </Button>
            </div>
{/* 🔶 Open Bulk Archive Button (compact version) */}
<div style={{ marginTop: "0.5rem", textAlign: "center" }}>
  <Button
    appearance="primary"
    size="small"
    onClick={() => openBulkArchiveDialog()}
    style={{
      backgroundColor: "#ff7a18",
      color: "#ffffff",
      border: "none",
      borderRadius: "2px",
      padding: "2px 14px", // tighter vertical + horizontal padding
      fontWeight: 600,
      fontSize: "0.85rem",
      lineHeight: "1rem",
      boxShadow: "0 0 4px rgba(255,122,24,0.4)",
      transition: "background-color 0.2s, box-shadow 0.2s",
    }}
    onMouseOver={(e) =>
      ((e.target as HTMLButtonElement).style.backgroundColor = "#ff8f3a")
    }
    onMouseOut={(e) =>
      ((e.target as HTMLButtonElement).style.backgroundColor = "#ff7a18")
    }
  >
    Open Bulk Archive
  </Button>
</div>


           {!favoritesCollapsed && (
  <div className="scroll-section favorites-scroll">
    <ul
      style={{
        color: "white",
        paddingLeft: "1rem",
        listStyle: "none",
        marginTop: "0.4rem",
        marginBottom: 0,
      }}
    >
      {favorites.length === 0 && <li>No favorites yet</li>}
      {favorites.map((fav) => (
        <li
          key={fav.id}
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 8,
            marginBottom: 4,
          }}
        >
          <span
            style={{
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              flex: 1,
            }}
          >
            {fav.name}
          </span>

          <div style={{ display: "flex", gap: 8 }}>
            <Tooltip content="Archive selected email(s) to this favorite" relationship="description">
              <CloudArrowUpIcon
                onClick={() => handleArchiveToPath(fav.driveId, fav.path)}
                style={{
                  width: 18,
                  height: 18,
                  cursor: "pointer",
                  color: darkOrange,
                }}
              />
            </Tooltip>

            <Tooltip content="Remove favorite" relationship="description">
              <StarIcon
                onClick={() => setFavorites(favorites.filter((f) => f.id !== fav.id))}
                style={{
                  width: 18,
                  height: 18,
                  cursor: "pointer",
                  color: yellowStar,
                }}
              />
            </Tooltip>
          </div>
        </li>
      ))}
    </ul>
  </div>
)}

          </Card>

          {/* 🏢 Library Selection (with collapse + auto-open) */}
          <Card style={{ padding: "0.5rem", position: "relative" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
              <label
                htmlFor="library-combobox"
                style={{ color: "white", fontWeight: 600, display: "block", marginBottom: 0 }}
              >
                Select Document Library:
              </label>
              <Button size="small" onClick={() => setLibCollapsed(!libCollapsed)}>
                {libCollapsed ? "▼" : "▲"}
              </Button>
            </div>

            {!libCollapsed && (
              <div style={{ position: "relative" }}>
                <Combobox
                  id="library-combobox"
                  appearance="outline"
                  value={selectedLibraryProfile?.Label ?? ""}
                  placeholder={mappingsLoading ? "Loading…" : "— Select a Library —"}
                  disabled={mappingsLoading || !libraryProfiles.length}
                  style={{ width: "100%" }}
                  onOptionSelect={(_, data) => {
                    const selectedProfile =
                      libraryProfiles.find((p) => p.Label === data.optionValue) || null;
                    setSelectedLibraryProfile(selectedProfile);
                    setLibraryAccessError("");
                    // ⚡ Auto-open immediately
                    if (selectedProfile) {
                      void handleLibraryOpen(selectedProfile);
                    } else {
                      setSelectedDriveId(null);
                      setSelectedDriveName(null);
                      setActiveSiteId("");
                      setDriveItems([]);
                      setFilteredItems([]);
                    }
                  }}
                >
                  {libraryProfiles.map((profile) => (
                    <Option key={profile.DriveId} value={profile.Label}>
                      {profile.Label}
                    </Option>
                  ))}
                </Combobox>

                {/* Tiny inline spinner to the right when loading/opening */}
                {loading && (
                  <div style={{ position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)" }}>
                    <Spinner size="tiny" />
                  </div>
                )}

                {libraryAccessError && (
                  <div style={{ color: "#ff6b6b", marginTop: "0.5rem", fontSize: "0.85rem" }}>
                    {libraryAccessError}
                  </div>
                )}
                {mappingsError && (
                  <div style={{ color: "#ff6b6b", marginTop: "0.5rem", fontSize: "0.85rem" }}>
                    {mappingsError}
                  </div>
                )}
              </div>
            )}
          </Card>

          {/* 📁 Folder Contents */}
          <Card style={{ padding: "0.5rem" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <h3 style={{ color: darkOrange, fontWeight: "bold" }}>
                {selectedDriveName
                  ? folderStack.length > 0
                    ? folderStack.map((f) => f.name).join(" / ")
                    : `${selectedDriveName} Contents`
                  : "Folder Contents"}
              </h3>
              {folderStack.length > 0 && (
                <Button
                  size="small"
                  appearance="secondary"
                  onClick={async () => {
                    const sid = getCurrentSiteId();
                    if (!token || !sid || !selectedDriveId || folderStack.length === 0) return;
                    const newStack = [...folderStack];
                    newStack.pop();
                    setFolderStack(newStack);
                    try {
                      setLoading(true);
                      if (newStack.length === 0) {
                        const items = await getDriveRootItems(token, sid, selectedDriveId);
                        setDriveItems(items);
                        setFilteredItems(items);
                      } else {
                        const parent = newStack[newStack.length - 1];
                        const items = await getDriveFolderItems(token, sid, selectedDriveId, parent.id);
                        setDriveItems(items);
                        setFilteredItems(items);
                      }
                      setSearchQuery("");
                    } catch (err) {
                      console.error("Error navigating back:", err);
                    } finally {
                      setLoading(false);
                    }
                  }}
                >
                  ⬅ Back
                </Button>
              )}
            </div>

            <Input
              placeholder="Search folder contents…"
              value={searchQuery}
              onChange={(e) => setSearchQuery((e.target as HTMLInputElement).value)}
              style={{ width: "100%", margin: "0.25rem 0" }}
              disabled={!selectedDriveId}
            />

            <div style={{ marginTop: 4, marginBottom: 6 }}>
              <Caption1 style={{ color: statusMsg ? "#fff" : "#999", whiteSpace: "normal" }}>
                {statusMsg || "Ready"}
              </Caption1>
            </div>
          {/* 🔁 Bulk retry section */}
          {bulkFailed.length > 0 && (
            <div style={{ marginBottom: "0.5rem" }}>
              <div style={{ color: "#ff6b6b", marginBottom: "0.25rem", fontSize: "0.85rem" }}>
                {bulkFailed.length} message{bulkFailed.length > 1 ? "s" : ""} failed to upload.
              </div>
              <Button
                size="small"
                appearance="primary"
                onClick={async () => {
                  const failedIds = bulkFailed.map((f) => f.restId);
                  setBulkFailed([]); // clear before retry
                  await bulkArchiveMessagesSequential(failedIds, selectedDriveId!, currentPath);
                }}
              >
                Retry Failed Uploads
              </Button>
            </div>
          )}

 <div className="scroll-section library-scroll">
  <ul
    style={{
      color: "white",
      paddingLeft: "1rem",
      listStyle: "none",
      marginBottom: 0,
    }}
  >
    {filteredItems.map((item) => {
      const isFav = favorites.some((f) => f.id === item.id);
      const pathForThisRow = buildPathWithLeaf(item.name);

      return (
        <li
          key={item.id}
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            cursor: item.folder ? "pointer" : "default",
            gap: 8,
          }}
        >
          <span
            onClick={() => {
              const sid = getCurrentSiteId();
              if (!item.folder || !token || !sid || !selectedDriveId) return;
              (async () => {
                try {
                  setLoading(true);
                  const items = await getDriveFolderItems(token, sid, selectedDriveId, item.id);
                  setFolderStack([...folderStack, { id: item.id, name: item.name }]);
                  setDriveItems(items);
                  setFilteredItems(items);
                  setSearchQuery("");
                } catch (err) {
                  console.error("Error loading folder items:", err);
                } finally {
                  setLoading(false);
                }
              })();
            }}
            style={{
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
              flex: 1,
            }}
          >
            {item.folder ? `📁 ${item.name}` : `📄 ${item.name}`}
          </span>

          <div style={{ display: "flex", gap: 8 }}>
            {item.folder && (
              <Tooltip content="Archive selected email(s) to this folder" relationship="description">
                <CloudArrowUpIcon
                  onClick={(e) => {
                    e.stopPropagation();
                    if (!selectedDriveId) {
                      setStatus("Please open a library first.");
                      return;
                    }
                    handleArchiveToPath(selectedDriveId, pathForThisRow);
                  }}
                  style={{
                    width: 18,
                    height: 18,
                    cursor: "pointer",
                    color: darkOrange,
                  }}
                />
              </Tooltip>
            )}
            {item.folder && (
              <Tooltip
                content={isFav ? "Remove favorite" : "Add favorite"}
                relationship="description"
              >
                <StarIcon
                  onClick={(e) => {
                    e.stopPropagation();
                    const exists = favorites.some((f) => f.id === item.id);
                    const path = pathForThisRow;
                    if (exists) {
                      setFavorites(favorites.filter((f) => f.id !== item.id));
                    } else {
                      setFavorites([
                        ...favorites,
                        { id: item.id, name: item.name, driveId: selectedDriveId!, path },
                      ]);
                    }
                  }}
                  style={{
                    width: 18,
                    height: 18,
                    cursor: "pointer",
                    color: isFav ? yellowStar : darkOrange,
                  }}
                />
              </Tooltip>
            )}
          </div>
        </li>
      );
    })}
    {selectedDriveId && !loading && filteredItems.length === 0 && (
      <li style={{ color: "#bbb" }}>No items.</li>
    )}
  </ul>
</div>
</Card>
</div>
</div>
</FluentProvider>
  );
}

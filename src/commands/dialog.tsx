// src/commands/dialog.tsx
import React, { useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  listDocumentLibraries,
  getDriveFolderItems,
  getAccessToken,
} from "../graphSharePoint";

/** Types */
type DriveRef = { id: string; name: string };
type FolderRef = { id: string; name: string };
type FavoriteRef = { id?: string; folderId?: string; name: string; driveId: string; path?: string };
type ParentPayload = { favorites?: FavoriteRef[]; restIds?: string[] };

/** SharePoint field names (same as taskpane) */
const FIELD_FROM = "From";
const FIELD_FROM_ADDRESS = "From_x002d_Address";
const FIELD_RECEIVED = "Received";
const FIELD_ATTACHMENT = "Attachment";
const FIELD_ORIGINAL_LINK = "OriginalMessageLink";
const FIELD_INTERNET_ID = "InternetMessageId";

/** Small helpers */
const encodeDrivePathForGraph = (path: string) =>
  path.split("/").filter(Boolean).map(encodeURIComponent).join("/");

const safeFileNameFromSubject = (subject: string) =>
  `${(subject || "Email").replace(/[^a-z0-9\\-_. ]/gi, "_")}.eml`;

async function withRetry<T>(fn: () => Promise<T>, max = 4, baseDelayMs = 500): Promise<T> {
  let attempt = 0, delay = baseDelayMs;
  while (true) {
    try { return await fn(); }
    catch (e: any) {
      const status = e?.status ?? e?.response?.status;
      const retriable = status === 429 || (status >= 500 && status <= 599);
      attempt++;
      if (!retriable || attempt >= max) throw e;
      await new Promise(r => setTimeout(r, delay));
      delay *= 2;
    }
  }
}

async function graphGET<T>(url: string, token: string): Promise<T> {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const err: any = new Error(`GET ${url} failed: ${r.status}`);
    err.status = r.status;
    err.body = await r.text().catch(() => "");
    throw err;
  }
  return r.json();
}

async function getSiteIdFromDrive(driveId: string, token: string): Promise<string> {
  const data = await graphGET<{ sharepointIds?: { siteId?: string } }>(
    `https://graph.microsoft.com/v1.0/drives/${driveId}?$select=sharepointIds`,
    token
  );
  const siteId = data?.sharepointIds?.siteId || "";
  if (!siteId) throw new Error("Could not resolve siteId for drive.");
  return siteId;
}

async function getDriveListFieldNames(token: string, siteId: string, driveId: string): Promise<Set<string>> {
  const cols = await graphGET<{ value: Array<{ name: string }> }>(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/list/columns?$select=name`,
    token
  );
  return new Set((cols.value || []).map(c => c.name));
}

async function waitForOfficeReady(): Promise<void> {
  return new Promise((resolve) => {
    if ((window as any).Office && (window as any).Office.context) resolve();
    else Office.onReady(() => resolve());
  });
}

/* ───────────────────────────────────────────── */

function BulkArchiveDialog() {
  const [token, setToken] = useState<string | null>(null);
  const [drives, setDrives] = useState<DriveRef[]>([]);
  const [folders, setFolders] = useState<FolderRef[]>([]);
  const [filteredFolders, setFilteredFolders] = useState<FolderRef[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedDrive, setSelectedDrive] = useState("");
  const [selectedFolder, setSelectedFolder] = useState("");
  const [favorites, setFavorites] = useState<FavoriteRef[]>([]);
  const [incomingRestIds, setIncomingRestIds] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [status, setStatus] = useState("Initializing…");

  /* ───────────────────────────────────────────── */
  // Recursive folder loader
  async function fetchAllFoldersRecursive(
    driveId: string,
    folderId: string = "root",
    prefix = ""
  ): Promise<FolderRef[]> {
    const children = await getDriveFolderItems(driveId, folderId);
    const all: FolderRef[] = [];
    for (const item of children) {
      if (item.folder) {
        const fullName = prefix ? `${prefix}/${item.name}` : item.name;
        all.push({ id: item.id, name: fullName });
        const sub = await fetchAllFoldersRecursive(driveId, item.id, fullName);
        all.push(...sub);
      }
    }
    return all;
  }

  /* ───────────────────────────────────────────── */
  // Init + receive parent payload
  useEffect(() => {
    async function initialize() {
      try {
        await waitForOfficeReady();
        console.log("✅ Office ready in dialog");

        setStatus("Signing in to Microsoft 365…");
        const accessToken = await getAccessToken();
        setToken(accessToken);

        setStatus("Loading document libraries…");
        const libs = await listDocumentLibraries();
        setDrives(libs || []);
        setStatus("Select a document library.");

        // Fallback favorites
        const stored = localStorage.getItem("flowpoint:favorites");
        if (stored) {
          const parsed = JSON.parse(stored);
          setFavorites(parsed);
          console.log("📦 Fallback favorites:", parsed);
        }

        // Listen for parent payload (favorites + restIds)
        const handler = (event: MessageEvent) => {
          try {
            const data: ParentPayload =
              typeof event.data === "string" ? JSON.parse(event.data) : event.data;
            if (Array.isArray(data?.favorites)) {
              const unique = Array.from(
                new Map(data.favorites.map(f => [f.folderId || f.id, f])).values()
              );
              setFavorites(unique);
              console.log("📥 Favorites from parent:", unique);
            }
            if (Array.isArray(data?.restIds)) {
              setIncomingRestIds(data.restIds);
              console.log("📥 REST IDs from parent:", data.restIds);
              if (data.restIds.length)
                setStatus(`Ready to archive ${Math.min(10, data.restIds.length)} email(s).`);
              else
                setStatus("No emails selected. Select up to 10 and reopen.");
            }
          } catch (e) {
            console.warn("⚠️ Non-JSON or irrelevant message:", e);
          }
        };
        window.addEventListener("message", handler);
        return () => window.removeEventListener("message", handler);
      } catch (err) {
        console.error("Init error:", err);
        setStatus("Failed to load document libraries.");
      } finally {
        setLoading(false);
      }
    }
    void initialize();
  }, []);

  /* ───────────────────────────────────────────── */
  // Folder load + search handlers
  async function loadFoldersForDrive(driveId: string, keepSearch = false) {
    setSelectedDrive(driveId);
    setFolders([]); setFilteredFolders([]); setSelectedFolder("");
    if (!driveId) return;
    try {
      setLoading(true);
      setStatus("Loading folders (including subfolders)…");
      const all = await fetchAllFoldersRecursive(driveId, "root");
      setFolders(all);
      const term = keepSearch ? searchTerm.trim().toLowerCase() : "";
      const visible = term ? all.filter(f => f.name.toLowerCase().includes(term)) : all;
      setFilteredFolders(visible);
      if (visible.length) setSelectedFolder(visible[0].id);
      setStatus(`Folders loaded (${all.length}).`);
    } catch (e) {
      console.error("Folder load error:", e);
      setStatus("Failed to load folders.");
    } finally { setLoading(false); }
  }

  const handleDriveChange = (e: React.ChangeEvent<HTMLSelectElement>) =>
    loadFoldersForDrive(e.target.value, true);

  function handleSearch(e: React.ChangeEvent<HTMLInputElement>) {
    const term = e.target.value;
    setSearchTerm(term);
    if (!term) { setFilteredFolders(folders); setSelectedFolder(""); return; }
    const lower = term.toLowerCase();
    const filtered = folders.filter(f => f.name.toLowerCase().includes(lower));
    setFilteredFolders(filtered);
    if (filtered.length) setSelectedFolder(filtered[0].id);
  }

  async function handleRefresh() {
    if (!selectedDrive) return alert("Please select a library first.");
    setLoading(true);
    const all = await fetchAllFoldersRecursive(selectedDrive, "root");
    setFolders(all);
    setFilteredFolders(all);
    setStatus(`Folder list updated (${all.length}).`);
    setLoading(false);
  }

  /* ───────────────────────────────────────────── */
  async function handleFavoriteClick(fav: FavoriteRef) {
    try {
      setStatus(`Loading favorite ${fav.name}…`);
      if (fav.driveId !== selectedDrive) await loadFoldersForDrive(fav.driveId);
      setSelectedFolder(fav.folderId || fav.id || "");
      setStatus(`Favorite selected: ${fav.name}`);
    } catch (e) {
      console.error("Favorite load error:", e);
      setStatus("Failed to load favorite.");
    }
  }

  /* ───────────────────────────────────────────── */
  // Upload + metadata patch
  async function uploadInChunksWithRetry(uploadUrl: string, blob: Blob) {
    const chunkSize = 5 * 1024 * 1024;
    const total = blob.size;
    for (let offset = 0; offset < total; ) {
      const slice = blob.slice(offset, Math.min(offset + chunkSize, total));
      const end = offset + slice.size - 1;
      await withRetry(async () => {
        const r = await fetch(uploadUrl, {
          method: "PUT",
          headers: {
            "Content-Length": String(slice.size),
            "Content-Range": `bytes ${offset}-${end}/${total}`,
          },
          body: slice,
        });
        if (!(r.ok || r.status === 202 || r.status === 201)) {
          const t = await r.text();
          const err: any = new Error(`Chunk upload failed ${r.status}: ${t}`);
          err.status = r.status;
          throw err;
        }
      });
      offset += slice.size;
    }
  }

  async function archiveMessageByRestId(
    restId: string,
    siteId: string,
    driveId: string,
    folderPath: string,
    token: string
  ) {
    const msg = await graphGET<{
      subject?: string;
      from?: { emailAddress?: { address?: string; name?: string } };
      hasAttachments?: boolean;
      receivedDateTime?: string;
      webLink?: string;
      internetMessageId?: string;
    }>(
      `https://graph.microsoft.com/v1.0/me/messages/${restId}?$select=subject,from,hasAttachments,receivedDateTime,webLink,internetMessageId`,
      token
    );

    const blob = await withRetry(async () => {
      const r = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages/${restId}/$value`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (!r.ok) throw new Error(`Download failed ${r.status}`);
      return r.blob();
    });

    const fileName = safeFileNameFromSubject(msg.subject || "Email");
    const encodedPath = encodeDrivePathForGraph(folderPath ? `${folderPath}/${fileName}` : fileName);

    const session = await withRetry(async () => {
      const r = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/createUploadSession`,
        {
          method: "POST",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ "@microsoft.graph.conflictBehavior": "rename" }),
        }
      );
      if (!r.ok) throw new Error(`Session failed ${r.status}`);
      return r.json();
    });

    await uploadInChunksWithRetry(session.uploadUrl, blob);

    const pathForFetch = folderPath
      ? `/${encodeDrivePathForGraph(folderPath)}/${encodeURIComponent(fileName)}`
      : `/${encodeURIComponent(fileName)}`;
    const uploaded = await graphGET<any>(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/root:${pathForFetch}`,
      token
    );

    const fields = await getDriveListFieldNames(token, siteId, driveId);
    const patch: Record<string, any> = {};
    if (fields.has(FIELD_FROM_ADDRESS)) patch[FIELD_FROM_ADDRESS] = msg.from?.emailAddress?.address || "";
    if (fields.has(FIELD_FROM)) patch[FIELD_FROM] = msg.from?.emailAddress?.name || "";
    if (fields.has(FIELD_RECEIVED)) patch[FIELD_RECEIVED] = msg.receivedDateTime ?? new Date().toISOString();
    if (fields.has(FIELD_ATTACHMENT)) patch[FIELD_ATTACHMENT] = !!msg.hasAttachments;
    if (fields.has(FIELD_ORIGINAL_LINK) && msg.webLink) patch[FIELD_ORIGINAL_LINK] = msg.webLink;
    if (fields.has(FIELD_INTERNET_ID) && msg.internetMessageId) patch[FIELD_INTERNET_ID] = msg.internetMessageId;

    if (Object.keys(patch).length)
      await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${uploaded.id}/listItem/fields`,
        {
          method: "PATCH",
          headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify(patch),
        }
      );
  }

  /* ───────────────────────────────────────────── */
  // Bulk archive main handler
  async function handleBulkArchive() {
  if (!selectedDrive) {
    setStatus("⚠️ Please select a document library first.");
    return;
  }
  if (!selectedFolder) {
    setStatus("⚠️ Please select a folder before archiving.");
    return;
  }

  const ids = incomingRestIds || [];
  if (ids.length === 0) {
    setStatus("⚠️ No emails selected. Select 1–10 emails, then reopen this dialog.");
    return;
  }

  // Enforce a cap of 10 emails
  const restIds = ids.slice(0, 10);
  if (ids.length > 10) {
    setStatus(`⚠️ ${ids.length} emails selected — only the first 10 will be processed.`);
  }


    try {
      setLoading(true);
      setStatus(`Archiving ${restIds.length} email(s)…`);
      const accessToken = token || (await getAccessToken());
      const siteId = await getSiteIdFromDrive(selectedDrive, accessToken);

      let success = 0, failed: string[] = [];
      for (let i = 0; i < restIds.length; i++) {
        const id = restIds[i];
        setStatus(`Uploading ${i + 1} of ${restIds.length}…`);
        try {
          await archiveMessageByRestId(id, siteId, selectedDrive, selectedFolder, accessToken);
          success++;
        } catch (e) {
          console.error("❌ Failed:", e);
          failed.push(id);
        }
      }

      if (!failed.length)
        setStatus(`✅ Archived ${success} email(s) successfully.`);
      else
        setStatus(`⚠️ ${success} succeeded, ${failed.length} failed (check console).`);
    } catch (e: any) {
      console.error("Bulk archive error:", e);
      setStatus(`❌ Bulk archive failed: ${e?.message || e}`);
    } finally {
      setLoading(false);
    }
  }

  /* ───────────────────────────────────────────── */
  // UI
  return (
    <div style={{ background: "#0f0f10", color: "#e5e7eb", fontFamily: "Segoe UI, sans-serif", padding: 20, minHeight: "100%" }}>
      <h3 style={{ color: "#ff7a18", fontWeight: 700, marginBottom: 6, textShadow: "0 0 8px rgba(255,122,24,0.5)" }}>
        Archive to SharePoint
      </h3>
      <p style={{ color: "#b5b8bf", fontSize: 12, marginTop: 0 }}>
        Choose a document library and target folder for bulk archiving.
      </p>

      <label style={{ display: "block", marginTop: 10, color: "#b5b8bf" }}>Document Library</label>
      <select
        style={{ width: "100%", padding: "6px", borderRadius: 6, border: "1px solid #2a2b2f", background: "#17181a", color: "#e5e7eb" }}
        value={selectedDrive}
        onChange={handleDriveChange}
        disabled={loading}
      >
        <option value="">-- Select a document library --</option>
        {drives.map((d) => (
          <option key={d.id} value={d.id}>{d.name}</option>
        ))}
      </select>

      <label style={{ display: "block", marginTop: 15, color: "#b5b8bf" }}>Folder / Subfolder</label>
      <input
        type="text"
        placeholder="Search folders..."
        value={searchTerm}
        onChange={handleSearch}
        style={{
          width: "100%", padding: "6px", borderRadius: 6,
          border: "1px solid #2a2b2f", background: "#17181a",
          color: "#e5e7eb", marginBottom: 6,
        }}
        disabled={!selectedDrive || loading || folders.length === 0}
      />
      <select
        style={{
          width: "100%", padding: "6px", borderRadius: 6, border: "1px solid #2a2b2f",
          background: "#17181a", color: "#e5e7eb", maxHeight: "200px", overflowY: "auto",
        }}
        value={selectedFolder}
        onChange={(e) => setSelectedFolder(e.target.value)}
        disabled={!filteredFolders.length || loading}
      >
        <option value="">-- Choose a folder --</option>
        {filteredFolders.map((f) => (
          <option key={f.id} value={f.id}>{f.name}</option>
        ))}
      </select>

      <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
        <button
          onClick={handleRefresh}
          style={{ background: "#2c2e33", color: "#fff", border: "1px solid #2a2b2f", borderRadius: 6, padding: "8px 14px", cursor: "pointer" }}
        >
          Update Folder List
        </button>
        <button
          onClick={handleBulkArchive}
          disabled={!selectedFolder}
          style={{
            background: selectedFolder ? "#ff7a18" : "#333",
            color: "#fff",
            border: "none",
            borderRadius: 6,
            padding: "8px 14px",
            fontWeight: 600,
            cursor: selectedFolder ? "pointer" : "not-allowed",
          }}
        >
          Bulk Archive
        </button>
      </div>

      <div style={{ marginTop: 20 }}>
        <h4 style={{ color: "#b5b8bf", marginBottom: 8 }}>⭐ Favorites</h4>
        {favorites.length ? (
          <div style={{ display: "grid", gap: 6 }}>
            {favorites.map((f) => (
              <button
                key={`${f.driveId}:${f.folderId || f.id}`}
                onClick={() => handleFavoriteClick(f)}
                style={{
                  textAlign: "left", background: "#141518", color: "#e5e7eb",
                  border: "1px solid #2a2b2f", borderRadius: 6, padding: "8px 10px", cursor: "pointer",
                }}
              >
                {f.name}
              </button>
            ))}
          </div>
        ) : <p style={{ fontSize: 12, color: "#666" }}>No favorites yet.</p>}
      </div>

      <p style={{
        marginTop: 15,
        fontSize: 12,
        color: status.includes("Failed")
          ? "#ff8f3a"
          : status.includes("✅") || status.includes("Ready") ? "#45ff82" : "#b5b8bf",
      }}>
        {loading ? "⏳ " : "✅ "} {status}
      </p>
    </div>
  );
}

/* ───────────────────────────────────────────── */
Office.onReady(() => {
  const container = document.getElementById("container");
  if (container) createRoot(container).render(<BulkArchiveDialog />);
});

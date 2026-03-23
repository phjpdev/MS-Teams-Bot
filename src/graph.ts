// src/graph.ts
import axios from "axios";

let cachedToken: string | null = null;
let tokenExpiry = 0;

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiry) return cachedToken;

  const res = await axios.post(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: process.env.GRAPH_CLIENT_ID!,
      client_secret: process.env.GRAPH_CLIENT_SECRET!,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    }),
    { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
  );

  cachedToken = res.data.access_token;
  tokenExpiry = Date.now() + (res.data.expires_in - 60) * 1000;
  return cachedToken;
}

function makeDriveRootUrl(relativePath: string) {
  return `https://graph.microsoft.com/v1.0/drives/${process.env.GRAPH_DRIVE_ID}/root:/${process.env.GRAPH_BASE_FOLDER}/${relativePath}`;
}

export async function uploadJsonToSharePoint(relativePath: string, json: any) {
  const token = await getAccessToken();

  const url = `${makeDriveRootUrl(relativePath)}:/content`;

  await axios.put(url, JSON.stringify(json, null, 2), {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });
}

export async function listFolder(path: string) {
  const token = await getAccessToken();

  const url = `${makeDriveRootUrl(path)}:/children`;

  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  return res.data.value;
}

export async function downloadFile(downloadUrl: string) {
  const res = await axios.get(downloadUrl);
  return res.data;
}

//Read JSON by known path
export async function downloadJsonFromSharePoint(
  relativePath: string
) {
  const token = await getAccessToken();

  const url = `https://graph.microsoft.com/v1.0/drives/${process.env.GRAPH_DRIVE_ID}/root:/${process.env.GRAPH_BASE_FOLDER}/${relativePath}:/content`;

  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  return res.data;
}

// read JSON by relative path (no need downloadUrl)
export async function getJsonFromSharePoint<T = any>(relativePath: string, fallback: T): Promise<T> {
  const token = await getAccessToken();
  const url = `${makeDriveRootUrl(relativePath)}:/content`;

  try {
    const res = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    // Graph may return string or object depending on file type.
    if (typeof res.data === "string") {
      return JSON.parse(res.data) as T;
    }
    return res.data as T;
  } catch {
    return fallback;
  }
}

/**
 * Find and download the most recently uploaded Excel/CSV file from a user's
 * OneDrive "Microsoft Teams Chat Files" folder. Teams stores files uploaded
 * in group chats there. Returns null if nothing found.
 */
export async function downloadRecentExcelFromUserDrive(
  userAadObjectId: string,
  maxAgeMinutes = 60
): Promise<{ fileName: string; buffer: Buffer } | null> {
  const token = await getAccessToken();
  const SUPPORTED = [".xlsx", ".xls", ".csv", ".tsv"];
  const cutoff = new Date(Date.now() - maxAgeMinutes * 60 * 1000);

  // Collect all candidate files from multiple sources, pick the most recent
  const candidates: { name: string; modified: Date; downloadUrl: string; source: string }[] = [];

  // Strategy 1: List known Teams chat files folders (most reliable, no indexing delay)
  const folderPaths = [
    "Microsoft Teams Chat Files",
    "Microsoft Teams-Chatdateien",  // German locale
  ];

  for (const folder of folderPaths) {
    try {
      const url = `https://graph.microsoft.com/v1.0/users/${userAadObjectId}/drive/root:/${folder}:/children?$orderby=lastModifiedDateTime desc&$top=10`;
      const res = await axios.get(url, {
        headers: { Authorization: `Bearer ${token}` },
      });

      for (const file of (res.data?.value ?? []) as any[]) {
        const name = (file.name ?? "").toLowerCase();
        if (!SUPPORTED.some((ext) => name.endsWith(ext))) continue;
        const downloadUrl = file["@microsoft.graph.downloadUrl"] ?? file["@content.downloadUrl"];
        if (!downloadUrl) continue;
        candidates.push({
          name: file.name,
          modified: new Date(file.lastModifiedDateTime ?? 0),
          downloadUrl,
          source: folder,
        });
      }
    } catch {
      // folder doesn't exist, skip
    }
  }

  // Strategy 2: Search user's OneDrive (covers files outside chat folders)
  try {
    const searchUrl = `https://graph.microsoft.com/v1.0/users/${userAadObjectId}/drive/root/search(q='.xlsx')?$top=20`;
    const res = await axios.get(searchUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    for (const file of (res.data?.value ?? []) as any[]) {
      const name = (file.name ?? "").toLowerCase();
      if (!SUPPORTED.some((ext) => name.endsWith(ext))) continue;
      const downloadUrl = file["@microsoft.graph.downloadUrl"] ?? file["@content.downloadUrl"];
      if (!downloadUrl) continue;
      // Avoid duplicates
      if (candidates.some((c) => c.name === file.name && c.modified.getTime() === new Date(file.lastModifiedDateTime ?? 0).getTime())) continue;
      candidates.push({
        name: file.name,
        modified: new Date(file.lastModifiedDateTime ?? 0),
        downloadUrl,
        source: "search",
      });
    }
  } catch (e: any) {
    console.error("Graph: OneDrive search error:", e?.response?.status, e?.message);
  }

  // Sort by most recent first
  candidates.sort((a, b) => b.modified.getTime() - a.modified.getTime());

  console.log(`Graph: found ${candidates.length} Excel candidates, cutoff=${cutoff.toISOString()}`);
  for (const c of candidates.slice(0, 5)) {
    console.log(`  "${c.name}" modified=${c.modified.toISOString()} source=${c.source}`);
  }

  // Pick the most recent file within the time window
  const match = candidates.find((c) => c.modified >= cutoff);
  if (!match) {
    console.log("Graph: no Excel file within time window");
    return null;
  }

  console.log(`Graph: downloading "${match.name}" (source=${match.source})`);
  const fileRes = await axios.get(match.downloadUrl, {
    responseType: "arraybuffer",
    timeout: 30000,
  });
  return { fileName: match.name, buffer: Buffer.from(fileRes.data) };
}

/** File metadata (eTag, lastModified) for change detection. */
export async function getFileMeta(relativePath: string): Promise<{
  eTag?: string;
  lastModifiedDateTime?: string;
  size?: number;
}> {
  const token = await getAccessToken();
  const url = makeDriveRootUrl(relativePath);

  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  return {
    eTag: res.data?.eTag,
    lastModifiedDateTime: res.data?.lastModifiedDateTime,
    size: res.data?.size,
  };
}

/** Text file content plus metadata for caching. */
export async function getTextFileContent(relativePath: string): Promise<{
  text: string;
  eTag?: string;
  lastModifiedDateTime?: string;
}> {
  const meta = await getFileMeta(relativePath);

  const token = await getAccessToken();
  const contentUrl = `${makeDriveRootUrl(relativePath)}:/content`;
  const res = await axios.get(contentUrl, {
    headers: { Authorization: `Bearer ${token}` },
    responseType: "text",
    transformResponse: (x) => x,
  });

  const text = typeof res.data === "string" ? res.data : JSON.stringify(res.data);

  return {
    text,
    eTag: meta.eTag,
    lastModifiedDateTime: meta.lastModifiedDateTime,
  };
}

/**
 * Append text to a file in SharePoint (read full file, append, write back).
 * If the file does not exist, creates it. For large files (e.g. 5–10 GB) consider
 * moving to Azure Blob Storage for true append/streaming.
 */
export async function appendTextToFile(relativePath: string, text: string): Promise<void> {
  const token = await getAccessToken();
  const fileUrl = makeDriveRootUrl(relativePath);

  try {
    const existing = await axios.get(`${fileUrl}:/content`, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: "text",
      transformResponse: (x) => x,
    });
    const combined = (existing.data ?? "") + text;
    await axios.put(`${fileUrl}:/content`, combined, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "text/plain",
      },
    });
  } catch {
    await axios.put(`${fileUrl}:/content`, text, {
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "text/plain",
      },
    });
  }
}



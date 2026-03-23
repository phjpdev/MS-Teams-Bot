// src/formatPipeline/fileParser.ts
// Downloads and parses Excel/CSV file attachments from Teams into TSV text.

import axios from "axios";
import * as XLSX from "xlsx";
import { TurnContext } from "botbuilder";
import { downloadRecentExcelFromUserDrive } from "../graph";

/** Supported file extensions for format pipeline. */
const SUPPORTED_EXTENSIONS = [".xlsx", ".xls", ".csv", ".tsv"];

/** Teams sends file uploads with this content type. */
const TEAMS_FILE_CONTENT_TYPE = "application/vnd.microsoft.teams.file.download.info";

export interface ParsedFileAttachment {
  fileName: string;
  tsvContent: string;
}

/** Check if a Teams attachment is a supported spreadsheet file. */
function isSupportedFile(attachment: any): boolean {
  // Check by file name extension
  const name = (attachment?.name ?? "").toLowerCase();
  if (SUPPORTED_EXTENSIONS.some((ext) => name.endsWith(ext))) return true;

  // Check by Teams file content type + fileType field
  if (attachment?.contentType === TEAMS_FILE_CONTENT_TYPE) {
    const fileType = (attachment.content?.fileType ?? "").toLowerCase();
    if (["xlsx", "xls", "csv", "tsv"].includes(fileType)) return true;
  }

  // Check by contentUrl extension (fallback for different attachment structures)
  const contentUrl = (attachment?.contentUrl ?? "").toLowerCase();
  if (contentUrl && SUPPORTED_EXTENSIONS.some((ext) => contentUrl.split("?")[0].endsWith(ext))) return true;

  // Check common MIME types for spreadsheets
  const ct = (attachment?.contentType ?? "").toLowerCase();
  if (
    ct === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || // .xlsx
    ct === "application/vnd.ms-excel" || // .xls
    ct === "text/csv"
  ) return true;

  return false;
}

/**
 * Download file content as a Buffer from a Teams file attachment.
 * Teams file attachments provide a pre-authenticated downloadUrl inside `attachment.content`.
 */
async function downloadAttachmentBuffer(attachment: any): Promise<Buffer> {
  // Teams file download info structure varies by context:
  // 1:1 chat: { contentType: "application/vnd.microsoft.teams.file.download.info",
  //             content: { downloadUrl: "https://...", uniqueId: "...", fileType: "xlsx" },
  //             name: "file.xlsx" }
  // Channel:  may also use contentUrl at top level
  const downloadUrl: string | undefined =
    attachment.content?.downloadUrl ??
    attachment.contentUrl ??
    attachment.content?.url;

  console.log(`Download URL for "${attachment.name}":`, downloadUrl ? downloadUrl.substring(0, 120) + "..." : "NONE");

  if (!downloadUrl) {
    throw new Error(
      `No download URL for "${attachment.name}". ` +
      `contentType=${attachment.contentType}, ` +
      `contentKeys=${attachment.content ? Object.keys(attachment.content).join(",") : "none"}, ` +
      `hasContentUrl=${!!attachment.contentUrl}`
    );
  }

  // Teams download URLs are pre-authenticated (token in query string)
  const res = await axios.get(downloadUrl, {
    responseType: "arraybuffer",
    timeout: 30000,
    headers: {
      "Accept": "application/octet-stream",
    },
  });
  console.log(`Download response for "${attachment.name}": status=${res.status}, size=${res.data?.byteLength ?? 0}`);
  return Buffer.from(res.data);
}

/** Parse an Excel buffer (.xlsx/.xls) into TSV text. Reads the first sheet. */
function excelBufferToTsv(buf: Buffer, fileName: string): string {
  const workbook = XLSX.read(buf, { type: "buffer" });

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error(`No sheets found in "${fileName}".`);

  const sheet = workbook.Sheets[sheetName];
  const rows: string[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
  }) as string[][];

  return rows
    .filter((row) => row.some((cell) => String(cell).trim().length > 0))
    .map((row) => row.map((cell) => String(cell).trim()).join("\t"))
    .join("\n");
}

/** Parse a CSV/TSV buffer into TSV text. */
function csvBufferToTsv(buf: Buffer, fileName: string): string {
  const text = buf.toString("utf-8");
  if (text.includes("\t")) return text;

  // Use xlsx's CSV parser for robust handling of quoted fields
  const workbook = XLSX.read(buf, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error(`No data found in "${fileName}".`);

  const sheet = workbook.Sheets[sheetName];
  const rows: string[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
  }) as string[][];

  return rows
    .filter((row) => row.some((cell) => String(cell).trim().length > 0))
    .map((row) => row.map((cell) => String(cell).trim()).join("\t"))
    .join("\n");
}

/**
 * Extract and parse all supported file attachments from a Teams message.
 * Returns parsed files (fileName + TSV) and any errors encountered.
 */
export async function parseFileAttachments(
  context: TurnContext
): Promise<{ files: ParsedFileAttachment[]; errors: string[] }> {
  const attachments = context.activity.attachments ?? [];
  const files: ParsedFileAttachment[] = [];
  const errors: string[] = [];

  // Log all attachments for debugging
  console.log(
    "All attachments:",
    JSON.stringify(
      attachments.map((a: any) => ({
        contentType: a.contentType,
        name: a.name,
        hasContentUrl: !!a.contentUrl,
        hasContent: !!a.content,
        contentKeys: a.content ? Object.keys(a.content) : [],
      }))
    )
  );

  for (const att of attachments) {
    if (!isSupportedFile(att)) continue;

    try {
      console.log(`Downloading file: ${att.name} (contentType: ${att.contentType})`);
      const buf = await downloadAttachmentBuffer(att);
      console.log(`Downloaded ${att.name}: ${buf.length} bytes`);

      const name = (att.name ?? "file").toLowerCase();
      let tsvContent: string;

      if (name.endsWith(".csv") || name.endsWith(".tsv")) {
        tsvContent = csvBufferToTsv(buf, att.name);
      } else {
        tsvContent = excelBufferToTsv(buf, att.name);
      }

      if (tsvContent.trim().length > 0) {
        console.log(`Parsed ${att.name}: ${tsvContent.split("\n").length} rows`);
        files.push({ fileName: att.name ?? "file", tsvContent });
      } else {
        errors.push(`File "${att.name}" was empty after parsing.`);
      }
    } catch (e: any) {
      const msg = `Failed to read "${att.name}": ${e?.message ?? String(e)}`;
      console.error(msg);
      errors.push(msg);
    }
  }

  return { files, errors };
}

/** True if the message has any supported file attachments. */
export function hasFileAttachments(context: TurnContext): boolean {
  return (context.activity.attachments ?? []).some(isSupportedFile);
}

/**
 * Fallback: download the most recently uploaded Excel/CSV from the user's
 * OneDrive "Microsoft Teams Chat Files" folder via Graph API.
 * Teams group chats/channels don't send file attachments to bots —
 * the files go to OneDrive/SharePoint instead.
 */
export async function parseFileFromUserDrive(
  userAadObjectId: string
): Promise<{ files: ParsedFileAttachment[]; errors: string[] }> {
  const files: ParsedFileAttachment[] = [];
  const errors: string[] = [];

  try {
    const result = await downloadRecentExcelFromUserDrive(userAadObjectId);
    if (!result) return { files, errors };

    const { fileName, buffer } = result;
    const name = fileName.toLowerCase();
    let tsvContent: string;

    if (name.endsWith(".csv") || name.endsWith(".tsv")) {
      tsvContent = csvBufferToTsv(buffer, fileName);
    } else {
      tsvContent = excelBufferToTsv(buffer, fileName);
    }

    if (tsvContent.trim().length > 0) {
      console.log(`Graph file parsed: ${fileName} → ${tsvContent.split("\n").length} rows`);
      files.push({ fileName, tsvContent });
    } else {
      errors.push(`File "${fileName}" was empty after parsing.`);
    }
  } catch (e: any) {
    const msg = `Failed to read file from OneDrive: ${e?.message ?? String(e)}`;
    console.error(msg);
    errors.push(msg);
  }

  return { files, errors };
}

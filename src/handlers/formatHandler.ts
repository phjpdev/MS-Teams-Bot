// src/handlers/formatHandler.ts
// Handles the complete Excel → OpenAI → preview → confirm → SharePoint pipeline.

import { TurnContext } from "botbuilder";
import {
  formatTableWithOpenAI,
  applyCorrectionsToFormattedData,
  FORMAT_DATA_TYPE_LABELS,
  FORMAT_SHAREPOINT_PATHS,
  formatPreviewSummary,
} from "../formatPipeline/formatData.js";
import {
  detectFormatCommand,
  isFormatConfirm,
  isFormatCancel,
  looksLikeCorrections,
  hasTabularData,
} from "../formatPipeline/detectFormat.js";
import {
  getPendingFormat,
  setPendingFormat,
  clearPendingFormat,
  getAwaitingTableData,
  setAwaitingTableData,
  clearAwaitingTableData,
} from "../formatPipeline/conversationState.js";
import { uploadJsonToSharePoint } from "../graph.js";
import { appendL0 } from "../l0.js";

interface FormatHandlerContext {
  context: TurnContext;
  text: string;
  embeddedTableTsv: string | null;
  rawText: string;
  conversationId: string;
  aadObjectId: string | undefined;
  fromName: string;
}

/** True when data is usable as table input — either TSV (tabs) or raw HTML table. */
function hasUsableTableData(data: string | null | undefined): data is string {
  if (!data) return false;
  return hasTabularData(data) || /<table/i.test(data);
}

async function safeSend(context: TurnContext, text: string) {
  const t = (text ?? "").trim();
  if (!t) return;
  await context.sendActivity(t);
}

function formatPreviewSnippet(json: any, maxLen = 500): string {
  try {
    const str = JSON.stringify(json, null, 2);
    if (str.length <= maxLen) return str;
    return str.slice(0, maxLen) + "\n... (full data will be saved on confirm).";
  } catch {
    return "(preview unavailable)";
  }
}

function stripWarnings(jsonToSave: any): object {
  return jsonToSave.warnings !== undefined
    ? { type: jsonToSave.type, columns: jsonToSave.columns, rows: jsonToSave.rows }
    : jsonToSave;
}

async function showPreviewAndAwaitConfirm(
  context: TurnContext,
  conversationId: string,
  dataType: string,
  cleanJson: any,
  warnings: string[]
) {
  const label = FORMAT_DATA_TYPE_LABELS[dataType as keyof typeof FORMAT_DATA_TYPE_LABELS];
  if (warnings.length > 0) {
    await safeSend(
      context,
      `Cleaned and formatted as **${label}**. The following issues were found (please fix or confirm to save anyway):`
    );
    for (const w of warnings) await safeSend(context, "• " + w);
    await safeSend(
      context,
      "Reply **Yes** or **Save** to save to SharePoint, **No** to cancel, or paste corrections."
    );
  } else {
    await safeSend(
      context,
      `Cleaned and formatted as **${label}**. Reply **Yes** or **Save** to save to SharePoint, or **No** to cancel.`
    );
  }
  await safeSend(context, formatPreviewSummary(cleanJson));
  await safeSend(context, "```json\n" + formatPreviewSnippet(cleanJson) + "\n```");
}

/**
 * Runs the full format pipeline for a single message.
 * Returns true if the message was handled (caller should not continue to AI).
 */
export async function handleFormatPipeline({
  context,
  text,
  embeddedTableTsv,
  rawText,
  conversationId,
  aadObjectId,
  fromName,
}: FormatHandlerContext): Promise<boolean> {
  const now = () => new Date().toISOString();
  const log = async (type: "user" | "bot", msg: string) => {
    try {
      await appendL0({ t: now(), type, cid: conversationId, uid: aadObjectId, name: fromName, msg });
    } catch (e) {
      console.error("L0 logging error:", e);
    }
  };

  // ── 1. Pending format: Yes / No / Corrections ────────────────────────────
  const pendingFormat = await getPendingFormat(conversationId);

  if (pendingFormat && isFormatConfirm(text)) {
    try {
      const path = FORMAT_SHAREPOINT_PATHS[pendingFormat.dataType];
      await uploadJsonToSharePoint(path, pendingFormat.json);
      await clearPendingFormat(conversationId);
      await safeSend(
        context,
        `Saved **${FORMAT_DATA_TYPE_LABELS[pendingFormat.dataType]}** to SharePoint (${path}).`
      );
      await log("user", text);
      await log("bot", `Saved ${pendingFormat.dataType} to ${path}`);
    } catch (e: any) {
      console.error("Format save error:", e);
      await safeSend(context, "Failed to save to SharePoint: " + (e?.message || String(e)));
    }
    return true;
  }

  if (pendingFormat && isFormatCancel(text)) {
    await clearPendingFormat(conversationId);
    await safeSend(context, "Cancelled. No data saved.");
    await log("user", text);
    await log("bot", "Cancelled format save.");
    return true;
  }

  if (pendingFormat && looksLikeCorrections(text)) {
    try {
      const result = await applyCorrectionsToFormattedData(
        pendingFormat.json,
        text,
        pendingFormat.dataType
      );
      if (result.error) {
        await safeSend(
          context,
          result.error +
            '\n\nReply **Yes** to save the current version, **No** to cancel, or paste the full table again with "here is the new timeline:".'
        );
        return true;
      }
      if (result.json) {
        const label = FORMAT_DATA_TYPE_LABELS[pendingFormat.dataType];
        const path = FORMAT_SHAREPOINT_PATHS[pendingFormat.dataType];
        const userWantsSave = /\b(save|speichern)\b/i.test(text);
        await setPendingFormat(conversationId, {
          dataType: pendingFormat.dataType,
          json: result.json,
          createdAt: now(),
        });
        if (userWantsSave) {
          try {
            await uploadJsonToSharePoint(path, result.json);
            await clearPendingFormat(conversationId);
            await safeSend(
              context,
              `Corrections applied and saved **${label}** to SharePoint (${path}).`
            );
            await safeSend(context, formatPreviewSummary(result.json));
            await log("user", "corrections+save");
            await log("bot", `Saved ${pendingFormat.dataType} to ${path}`);
            return true;
          } catch (e: any) {
            await safeSend(
              context,
              "Corrections applied but save failed: " +
                (e?.message || String(e)) +
                ". Reply **Yes** to retry save."
            );
            return true;
          }
        }
        await safeSend(
          context,
          `Corrections applied to **${label}**. Reply **Yes** or **Save** to save to SharePoint, or **No** to cancel.`
        );
        await safeSend(context, formatPreviewSummary(result.json));
        await safeSend(context, "```json\n" + formatPreviewSnippet(result.json) + "\n```");
        await log("user", "corrections");
        await log("bot", "Corrections applied; awaiting confirm.");
        return true;
      }
    } catch (e: any) {
      console.error("Apply corrections error:", e);
      await safeSend(
        context,
        "Could not apply corrections. Reply **Yes** to save the current version, **No** to cancel, or paste the full table again."
      );
      return true;
    }
  }

  if (pendingFormat) {
    const label = FORMAT_DATA_TYPE_LABELS[pendingFormat.dataType];
    await safeSend(
      context,
      `You have a pending **${label}**. Reply **Yes** or **Save** to save to SharePoint, **No** to cancel, or paste corrections (e.g. "Rohbau end 16.10.2026, Elektrik end 16.03.2026").`
    );
    return true;
  }

  // ── 2. Awaiting table paste (two-step flow) ──────────────────────────────
  const awaitingData = await getAwaitingTableData(conversationId);

  if (awaitingData && (isFormatCancel(text) || text.trim().toLowerCase() === "no")) {
    await clearAwaitingTableData(conversationId);
    await safeSend(
      context,
      'Cancelled. Send "format as timeline" (or matrix/budget) again when you are ready to paste the table.'
    );
    return true;
  }

  // Accept: explicit TSV/HTML, embeddedTableTsv, or multi-line rawText (space-delimited / Teams-stripped HTML)
  const rawLineCount = rawText.split("\n").filter((l: string) => l.trim().length > 0).length;
  const hasTableContent =
    hasUsableTableData(text) ||
    !!embeddedTableTsv ||
    hasUsableTableData(rawText) ||
    rawLineCount > 3;

  if (awaitingData && hasTableContent) {
    await clearAwaitingTableData(conversationId);
    const { dataType } = awaitingData;
    const tableData = embeddedTableTsv ?? (hasUsableTableData(rawText) ? rawText : text || rawText);
    try {
      await context.sendActivity({ type: "typing" });
      await safeSend(context, `Formatting as **${FORMAT_DATA_TYPE_LABELS[dataType]}**... this may take a moment.`);
      const result = await formatTableWithOpenAI(tableData, dataType);
      if (result.error) {
        await safeSend(
          context,
          "Formatting issue: " +
            result.error +
            '\n\nPaste the table again or try with "format as ' +
            dataType +
            '" followed by the table in one message.'
        );
        return true;
      }
      if (!result.json) {
        await safeSend(context, "Could not format the data. Please send a clear table (e.g. copy from Excel).");
        return true;
      }
      const cleanJson = stripWarnings(result.json as any);
      await setPendingFormat(conversationId, { dataType, json: cleanJson, createdAt: now() });
      await showPreviewAndAwaitConfirm(
        context,
        conversationId,
        dataType,
        cleanJson,
        result.warnings ?? []
      );
      await log("user", "table data (awaiting)");
      await log("bot", `Formatted ${dataType}; awaiting confirm.`);
    } catch (e: any) {
      console.error("Format from awaiting data error:", e);
      await safeSend(context, "Formatting failed: " + (e?.message || String(e)));
    }
    return true;
  }

  // catch-all: awaiting state set but no table found in this message
  if (awaitingData) {
    const label = FORMAT_DATA_TYPE_LABELS[awaitingData.dataType];
    await safeSend(
      context,
      `Still waiting for your **${label}** table. Paste the table data (e.g. copy from Excel) or reply **No** to cancel.`
    );
    return true;
  }

  // ── 3. Fresh format command / table detection ────────────────────────────
  const formatCmd = detectFormatCommand(text, embeddedTableTsv);
  if (!formatCmd) return false;

  const { dataType, tableData } = formatCmd;

  // If detectFormatCommand didn't find table data, try rawText as fallback
  // (Teams strips the HTML table from activity.text when typed text is present)
  const rawLineCount2 = rawText.split("\n").filter((l: string) => l.trim().length > 0).length;
  const effectiveTableData = hasUsableTableData(tableData)
    ? tableData
    : hasUsableTableData(rawText) || rawLineCount2 > 3
    ? rawText
    : null;

  if (!effectiveTableData) {
    await setAwaitingTableData(conversationId, dataType);
    await safeSend(
      context,
      `Paste the table data in your next message (e.g. copy from Excel). I will format it as **${FORMAT_DATA_TYPE_LABELS[dataType]}** and show a preview for you to confirm.`
    );
    return true;
  }

  try {
    await context.sendActivity({ type: "typing" });
    await safeSend(context, `Formatting as **${FORMAT_DATA_TYPE_LABELS[dataType]}**... this may take a moment.`);
    const result = await formatTableWithOpenAI(effectiveTableData, dataType);
    if (result.error) {
      await safeSend(
        context,
        "Formatting issue: " + result.error + "\n\nPlease correct the data and try again."
      );
      await log("user", text.slice(0, 200));
      await log("bot", "Format error: " + result.error);
      return true;
    }
    if (!result.json) {
      await safeSend(context, "Could not format the data. Please send a clear table (e.g. copy from Excel).");
      return true;
    }
    const cleanJson = stripWarnings(result.json as any);
    await setPendingFormat(conversationId, { dataType, json: cleanJson, createdAt: now() });
    await showPreviewAndAwaitConfirm(context, conversationId, dataType, cleanJson, result.warnings ?? []);
    await log("user", "format " + dataType);
    await log("bot", `Formatted ${dataType}; awaiting confirm.`);
  } catch (e: any) {
    console.error("Data format error:", e);
    await safeSend(context, "Formatting failed: " + (e?.message || String(e)));
  }

  return true;
}

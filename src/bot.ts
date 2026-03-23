// src/bot.ts — message router (thin orchestrator only)
import { ActivityHandler, TurnContext } from "botbuilder";
import { askAI } from "./ai/openai";
import { loadContext } from "./context";
import { appendL0, readRecentL0Events } from "./l0";
import { filterByTime, formatRecentMemory } from "./memory";
import { upsertUser } from "./userDirectory";
import { handleProfileCommand } from "./profileCommands";
import { saveConversationRef } from "./conversationRef";
import {
  getStandupSession,
  appendStandupResponse,
  parseStandupText,
  looksLikeStandupReply,
} from "./standup";
import { extractMessageParts } from "./formatPipeline/htmlParser";
import { hasFileAttachments, parseFileAttachments, parseFileFromUserDrive } from "./formatPipeline/fileParser";
import { handleFormatPipeline } from "./handlers/formatHandler";
import { handleReminders } from "./handlers/reminderHandler";
import { handleConfirmation } from "./handlers/confirmHandler";
import { handleActions } from "./handlers/actionHandler";

const MEMORY_MAX_ENTRIES = 100;
const MEMORY_MAX_DAYS = 20;

/** Track already-processed OneDrive files to avoid re-processing on every message. */
const processedFiles = new Set<string>();

function stripMentions(text: string) {
  return (text || "").replace(/<at>.*?<\/at>/g, "").trim();
}

function extractActions(aiText: string): any | null {
  const m = aiText.match(/<actions>\s*([\s\S]*?)\s*<\/actions>/i);
  if (!m) return null;
  try { return JSON.parse(m[1]); } catch {
    console.error("Invalid JSON inside <actions> block");
    return null;
  }
}

function removeActionsBlock(aiText: string) {
  return (aiText || "").replace(/<actions>[\s\S]*?<\/actions>/i, "").trim();
}

async function safeSend(context: TurnContext, text: string) {
  const t = (text ?? "").trim();
  if (!t) return;
  await context.sendActivity(t);
}

export class AIPMBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context: TurnContext, next) => {
      const rawText = context.activity.text ?? "";

      // Extract plain text + embedded table (Teams sends Excel paste as HTML)
      const { plainText, tableTsv } = extractMessageParts(rawText);
      const text = stripMentions(plainText);
      const conversationId = context.activity.conversation?.id ?? "unknown";
      const fromName = context.activity.from?.name ?? "User";
      const aadObjectId =
        (context.activity.from as any)?.aadObjectId ||
        (context.activity.from as any)?.objectId ||
        undefined;

      // ── Parse file attachments (Excel/CSV uploaded in chat) ────────────────
      // Must run BEFORE the empty-message check — user may send just a file with no text.
      let fileTsv: string | null = null;

      // 1) Try actual file attachments (Excel/CSV uploads in 1:1 chats)
      if (hasFileAttachments(context)) {
        try {
          await context.sendActivity({ type: "typing" });
          const { files, errors } = await parseFileAttachments(context);
          if (files.length > 0) {
            fileTsv = files.map((f) => f.tsvContent).join("\n");
          }
          if (errors.length > 0) {
            await safeSend(context, errors.join("\n"));
          }
        } catch (e: any) {
          console.error("File attachment parse error:", e);
          await safeSend(context, `Error reading file: ${e?.message ?? String(e)}`);
        }
      }

      // 2) Try Graph API fallback: download recent Excel from user's OneDrive
      //    Teams group chats/channels don't send file attachments to bots.
      //    Track processed files to avoid re-processing on every message.
      if (!fileTsv && aadObjectId) {
        try {
          const { files, errors } = await parseFileFromUserDrive(aadObjectId);
          if (files.length > 0) {
            const fileKey = files[0].fileName + "|" + files[0].tsvContent.length;
            if (!processedFiles.has(fileKey)) {
              processedFiles.add(fileKey);
              // Keep set from growing indefinitely
              if (processedFiles.size > 100) {
                const first = processedFiles.values().next().value;
                if (first) processedFiles.delete(first);
              }
              fileTsv = files.map((f) => f.tsvContent).join("\n");
              await safeSend(context, `Reading file **${files[0].fileName}**...`);
            }
          }
          if (errors.length > 0) {
            await safeSend(context, errors.join("\n"));
          }
        } catch (e: any) {
          console.error("Graph file fallback error:", e);
        }
      }

      // Merge: file attachment TSV takes priority, then embedded HTML table TSV
      const mergedTableTsv = fileTsv ?? tableTsv;

      if (!text && !mergedTableTsv) {
        await safeSend(context, "Please mention me with a message.");
        await next();
        return;
      }

      // ── Persist conversation ref (proactive standup messages) ────────────
      try {
        const ref = TurnContext.getConversationReference(context.activity);
        if (ref?.conversation?.id && ref?.serviceUrl) await saveConversationRef(ref as any);
      } catch (e) { console.error("Save conversation ref error:", e); }

      // ── Standup collection ───────────────────────────────────────────────
      try {
        const session = await getStandupSession();
        const today = new Date().toISOString().slice(0, 10);
        if (session?.date === today && aadObjectId && looksLikeStandupReply(text)) {
          const parsed = parseStandupText(text);
          await appendStandupResponse({
            userId: aadObjectId,
            userName: fromName,
            rawText: text,
            yesterday: parsed.yesterday,
            today: parsed.today,
            blockers: parsed.blockers,
            timestamp: new Date().toISOString(),
          });
        }
      } catch (e) { console.error("Standup response append error:", e); }

      // ── User directory sync ──────────────────────────────────────────────
      let speakerProfile: any = null;
      if (aadObjectId) {
        try { speakerProfile = await upsertUser({ aadObjectId, displayName: fromName }); }
        catch (e) { console.error("USER UPSERT ERROR:", e); }
      }

      // ── Profile commands (/profile, /role, /skill, /timezone) ────────────
      const profileResult = await handleProfileCommand({
        text,
        aadObjectId: aadObjectId ?? undefined,
        displayName: fromName,
      });
      if (profileResult.handled) {
        await safeSend(context, profileResult.reply);
        try {
          await appendL0({ t: new Date().toISOString(), type: "user", cid: conversationId, uid: aadObjectId, name: fromName, msg: text });
          await appendL0({ t: new Date().toISOString(), type: "bot", cid: conversationId, msg: profileResult.reply });
        } catch (e) { console.error("L0 LOGGING ERROR:", e); }
        await next();
        return;
      }

      // ── Format pipeline (Excel → OpenAI → preview → confirm → save) ──────
      const formatHandled = await handleFormatPipeline({
        context,
        text,
        embeddedTableTsv: mergedTableTsv,
        rawText,
        conversationId,
        aadObjectId,
        fromName,
      });
      if (formatHandled) { await next(); return; }

      // ── Reminders (due-soon / overdue / budget risk) ──────────────────────
      if (aadObjectId) await handleReminders(context, aadObjectId);

      // ── Task confirmation ("confirm" keyword) ─────────────────────────────
      if (aadObjectId) {
        const confirmed = await handleConfirmation(context, text, aadObjectId, conversationId);
        if (confirmed) { await next(); return; }
      }

      // ── Load context + memory ─────────────────────────────────────────────
      const ctx = await loadContext();
      const rawEvents = await readRecentL0Events(conversationId, MEMORY_MAX_ENTRIES * 2, MEMORY_MAX_DAYS);
      const recentMemory = formatRecentMemory(
        filterByTime(rawEvents, MEMORY_MAX_DAYS).slice(-MEMORY_MAX_ENTRIES)
      );

      // ── AI call ───────────────────────────────────────────────────────────
      const speakerRoles =
        speakerProfile?.roles?.length > 0 ? speakerProfile.roles.join(", ") : "(none)";
      const speakerCompetencies =
        speakerProfile?.competencies?.length > 0
          ? speakerProfile.competencies
              .map((c: { skill?: string; level?: number }) => [c.skill, c.level ?? 3].join(":"))
              .join(", ")
          : "(none)";

      const aiRaw = await askAI({
        userMessage: `SPEAKER:\n- name: ${fromName}\n- aadObjectId: ${aadObjectId ?? "unknown"}\n- timezone: ${speakerProfile?.timezone ?? "UTC"}\n- roles: ${speakerRoles}\n- competencies: ${speakerCompetencies}\n\nMESSAGE:\n${text}`,
        systemPromptTemplate: ctx.systemPromptTemplate,
        projectState: ctx.projectStateText,
        teamRoles: ctx.teamRolesText,
        rules: ctx.rulesText,
        recentMemory,
        userDirectory: ctx.userDirectoryText,
        taskRegistry: ctx.taskRegistryText,
        savedData: ctx.savedDataText,
      });

      const actionsObj = extractActions(aiRaw);
      const aiReply = removeActionsBlock(aiRaw);
      await safeSend(context, aiReply);

      // ── Execute AI actions ────────────────────────────────────────────────
      if (actionsObj?.actions && Array.isArray(actionsObj.actions)) {
        await handleActions({
          context,
          actions: actionsObj.actions,
          aadObjectId,
          fromName,
          conversationId,
          speakerTimezone: speakerProfile?.timezone ?? "UTC",
        });
      }

      // ── L0 audit log ──────────────────────────────────────────────────────
      try {
        await appendL0({ t: new Date().toISOString(), type: "user", cid: conversationId, uid: aadObjectId, name: fromName, msg: text });
        await appendL0({ t: new Date().toISOString(), type: "bot", cid: conversationId, msg: aiReply || "(no-text; actions-only reply)", model: process.env.AI_MODEL });
      } catch (e) { console.error("L0 LOGGING ERROR:", e); }

      await next();
    });
  }
}

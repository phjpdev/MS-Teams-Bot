import "dotenv/config";
import express from "express";
import cron from "node-cron";
import {
  CloudAdapter,
  TurnContext,
  ConfigurationBotFrameworkAuthentication,
} from "botbuilder";
import type { ConversationReference } from "botframework-schema";
import { AIPMBot } from "./src/bot";
import { syncTimelineToSharePoint } from "./src/timeline";
import { syncBudgetToSharePoint } from "./src/budget";
import { loadConversationRef } from "./src/conversationRef";
import {
  setStandupSession,
  getStandupSession,
  getStandupResponses,
  buildStandupSummaryAndRisks,
  sendStandupMessage,
  sendStandupSummaryAndChallenges,
} from "./src/standup";

const app = express();
app.use(express.json());

// Bot Framework Adapter (uses App Registration credentials)
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context: TurnContext, error: any) => {
  console.error("BOT ERROR:", error);
  await context.sendActivity("⚠️ The bot hit an error.");
};

const bot = new AIPMBot();

// On startup, ensure timeline and budget files reflect current TaskRegistry
syncTimelineToSharePoint().catch((e) => console.error("Startup timeline sync:", e));
syncBudgetToSharePoint().catch((e) => console.error("Startup budget sync:", e));

app.get("/", (req, res) => {
  res.send("Bot is running ✅");
});

app.get("/health", (req, res) => {
  res.status(200).send("OK");
});

// Scheduled standup trigger (called by cron at 09:00 and 11:00). action=ask | analyze
app.get("/api/standup/trigger", async (req, res) => {
  const secret = req.query.secret ?? (req.headers["x-standup-secret"] as string);
  if (process.env.STANDUP_TRIGGER_SECRET && secret !== process.env.STANDUP_TRIGGER_SECRET) {
    return res.status(401).send("Unauthorized");
  }
  const action = (req.query.action as string) || "ask";
  const refData = await loadConversationRef();
  if (!refData?.conversationReference) {
    return res
      .status(503)
      .send("No conversation reference. Send a message to the bot in the channel first.");
  }
  const appId = process.env.MicrosoftAppId;
  if (!appId) {
    return res.status(500).send("MicrosoftAppId not configured.");
  }
  try {
    await adapter.continueConversationAsync(
      appId,
      refData.conversationReference as Partial<ConversationReference>,
      async (context) => {
        if (action === "ask") {
          await setStandupSession();
          await sendStandupMessage(context);
        } else if (action === "analyze") {
          const session = await getStandupSession();
          const date = session?.date ?? new Date().toISOString().slice(0, 10);
          const responses = await getStandupResponses(date);
          const { summary, challenges } = await buildStandupSummaryAndRisks(date, responses);
          await sendStandupSummaryAndChallenges(context, summary, challenges);
        }
      }
    );
    res.status(200).send("OK");
  } catch (e: any) {
    console.error("Standup trigger error:", e);
    res.status(500).send(e?.message ?? String(e));
  }
});

// Messages endpoint (Teams/Azure Bot will POST here)
app.post("/api/messages", (req, res) => {
  adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

const port = Number(process.env.PORT || 3978);
app.listen(port, () => console.log(`Bot listening on port ${port}`));

// ── Scheduled standup cron jobs (runs inside the bot process) ──────────────
// No Azure WebJobs needed — these run in-process on the same App Service.

async function triggerStandup(action: "ask" | "analyze") {
  const refData = await loadConversationRef();
  if (!refData?.conversationReference) {
    console.log(`Standup ${action} skipped: no conversation reference saved yet.`);
    return;
  }
  const appId = process.env.MicrosoftAppId;
  if (!appId) {
    console.error(`Standup ${action} skipped: MicrosoftAppId not configured.`);
    return;
  }
  try {
    await adapter.continueConversationAsync(
      appId,
      refData.conversationReference as Partial<ConversationReference>,
      async (context) => {
        if (action === "ask") {
          await setStandupSession();
          await sendStandupMessage(context);
        } else {
          const session = await getStandupSession();
          const date = session?.date ?? new Date().toISOString().slice(0, 10);
          const responses = await getStandupResponses(date);
          const { summary, challenges } = await buildStandupSummaryAndRisks(date, responses);
          await sendStandupSummaryAndChallenges(context, summary, challenges);
        }
      }
    );
    console.log(`Standup ${action} sent at ${new Date().toISOString()}`);
  } catch (e) {
    console.error(`Standup ${action} error:`, e);
  }
}

// 9:00 AM weekdays — ask for status
cron.schedule("0 9 * * 1-5", () => {
  console.log("Cron: triggering standup ask...");
  triggerStandup("ask");
}, { timezone: "Europe/Berlin" });

// 11:00 AM weekdays — post summary + risks
cron.schedule("0 11 * * 1-5", () => {
  console.log("Cron: triggering standup analyze...");
  triggerStandup("analyze");
}, { timezone: "Europe/Berlin" });

console.log("Standup cron scheduled: ask=9:00 AM, analyze=11:00 AM (Mon-Fri, Europe/Berlin)");
// src/standup.ts — Scheduled standup: ask for updates, store responses, analyze risks, post summary
import { TurnContext } from "botbuilder";
import { getJsonFromSharePoint, uploadJsonToSharePoint } from "./graph";
import { loadTaskRegistry } from "./taskRegistry";
import { getOverdueTasks, getDueSoonTasks } from "./monitoring";

const RESPONSES_PATH = "bot-data/01_PROJECT_STATE/StandupResponses.json";
const SESSION_PATH = "bot-data/01_PROJECT_STATE/StandupSession.json";

export interface StandupResponseEntry {
  userId: string;
  userName: string;
  rawText: string;
  yesterday?: string;
  today?: string;
  blockers?: string;
  timestamp: string;
}

export interface StandupSession {
  date: string; // YYYY-MM-DD
  requestedAt: string;
}

export const STANDUP_QUESTION =
  "Good morning. Please provide your status in this format:\n\n" +
  "What was done yesterday?\n" +
  "What is planned for today?\n" +
  "Any blockers?\n\n" +
  "Replies will be compared with the project plan and deadlines. Inconsistencies or timeline risks may be flagged.";

function todayUtc(): string {
  return new Date().toISOString().slice(0, 10);
}

export async function getStandupSession(): Promise<StandupSession | null> {
  const data = await getJsonFromSharePoint(SESSION_PATH, null);
  if (!data || typeof data !== "object") return null;
  const d = data as any;
  if (!d.date || !d.requestedAt) return null;
  return { date: d.date, requestedAt: d.requestedAt };
}

export async function setStandupSession(): Promise<StandupSession> {
  const session: StandupSession = { date: todayUtc(), requestedAt: new Date().toISOString() };
  await uploadJsonToSharePoint(SESSION_PATH, session);
  return session;
}

export async function getStandupResponses(date: string): Promise<StandupResponseEntry[]> {
  const data = await getJsonFromSharePoint(RESPONSES_PATH, { byDate: {} });
  const byDate = (data as any).byDate;
  if (!byDate || typeof byDate !== "object") return [];
  const list = byDate[date];
  return Array.isArray(list) ? list : [];
}

export async function appendStandupResponse(entry: StandupResponseEntry): Promise<void> {
  const data = await getJsonFromSharePoint(RESPONSES_PATH, { byDate: {} });
  const byDate = (data as any).byDate ?? {};
  const date = entry.timestamp.slice(0, 10);
  const list: StandupResponseEntry[] = Array.isArray(byDate[date]) ? byDate[date] : [];
  list.push(entry);
  byDate[date] = list;
  await uploadJsonToSharePoint(RESPONSES_PATH, { byDate });
}

/** Simple extraction of Yesterday / Today / Blockers from user text */
export function parseStandupText(text: string): { yesterday?: string; today?: string; blockers?: string } {
  const t = (text || "").trim();
  const result: { yesterday?: string; today?: string; blockers?: string } = {};
  const sections = [
    { key: "yesterday" as const, patterns: [/yesterday\s*:?\s*/i, /what was done\s*:?\s*/i] },
    { key: "today" as const, patterns: [/today\s*:?\s*/i, /what is planned\s*:?\s*/i, /planned for today\s*:?\s*/i] },
    { key: "blockers" as const, patterns: [/blockers?\s*:?\s*/i, /blocker\s*:?\s*/i] },
  ];
  let remaining = t;
  for (const { key, patterns } of sections) {
    for (const re of patterns) {
      const m = remaining.match(re);
      if (m) {
        const start = m.index! + m[0].length;
        let end = remaining.length;
        for (const next of sections) {
          if (next.key === key) continue;
          for (const nextRe of next.patterns) {
            const nextM = remaining.slice(start).match(nextRe);
            if (nextM) end = Math.min(end, start + nextM.index!);
          }
        }
        result[key] = remaining.slice(start, end).trim() || undefined;
        break;
      }
    }
  }
  if (!result.yesterday && !result.today && !result.blockers && t.length > 0) {
    result.yesterday = t;
  }
  return result;
}

/** Detect if message looks like a standup reply (contains section headers or short status) */
export function looksLikeStandupReply(text: string): boolean {
  const lower = (text || "").toLowerCase();
  const hasSection =
    /yesterday\s*:/.test(lower) ||
    /today\s*:/.test(lower) ||
    /blockers?\s*:/.test(lower) ||
    /what was done\s*:/.test(lower) ||
    /what is planned\s*:/.test(lower);
  if (hasSection) return true;
  if (lower.length > 20 && (lower.includes("yesterday") || lower.includes("today") || lower.includes("blocker")))
    return true;
  return false;
}

/** Build standup summary and risk/challenge messages from responses and task registry */
export async function buildStandupSummaryAndRisks(
  date: string,
  responses: StandupResponseEntry[]
): Promise<{ summary: string; challenges: string[] }> {
  const [registry, overdue, dueSoon] = await Promise.all([
    loadTaskRegistry(),
    getOverdueTasks(),
    getDueSoonTasks(48),
  ]);

  const lines: string[] = [];
  lines.push(`Standup summary for ${date}`);
  lines.push("");

  for (const r of responses) {
    lines.push(`**${r.userName}**`);
    if (r.yesterday) lines.push(`- Yesterday: ${r.yesterday}`);
    if (r.today) lines.push(`- Today: ${r.today}`);
    if (r.blockers) lines.push(`- Blockers: ${r.blockers}`);
    lines.push("");
  }

  const challenges: string[] = [];

  for (const r of responses) {
    const userTasks = registry.tasks.filter((t) => t.ownerAadObjectId === r.userId);
    const myOverdue = overdue.filter((t) => t.ownerAadObjectId === r.userId);
    const myDueSoon = dueSoon.filter((t) => t.ownerAadObjectId === r.userId);

    if (myOverdue.length > 0) {
      const titles = myOverdue.map((t) => `"${t.title}"`).join(", ");
      challenges.push(
        `${r.userName}: You have overdue task(s): ${titles}. Please update status or confirm a new deadline.`
      );
    }

    const hasBlocker = (r.rawText || "").toLowerCase().includes("block");
    if (hasBlocker && r.blockers && r.blockers.toLowerCase() !== "none" && r.blockers.trim().length > 0) {
      challenges.push(`${r.userName} reported blocker(s): ${r.blockers}. Consider unblocking or escalating.`);
    }

    const todayLower = (r.today || "").toLowerCase();
    for (const t of myDueSoon) {
      const dueDate = t.dueAt ? t.dueAt.slice(0, 10) : "";
      if (dueDate === date && !todayLower.includes(t.title.toLowerCase().slice(0, 10))) {
        challenges.push(
          `${r.userName}: Task "${t.title}" is due today (${date}). Your "Today" section does not mention it. Please confirm if still on track.`
        );
      }
    }
  }

  if (overdue.length > 0 && responses.length > 0) {
    const names = [...new Set(overdue.map((t) => t.ownerDisplayName).filter(Boolean))];
    challenges.push(`Overall: Overdue tasks exist for: ${names.join(", ")}. Please update status.`);
  }

  return {
    summary: lines.join("\n").trim() || "No standup responses recorded for this date.",
    challenges,
  };
}

export async function sendStandupMessage(context: TurnContext): Promise<void> {
  await context.sendActivity(STANDUP_QUESTION);
}

export async function sendStandupSummaryAndChallenges(
  context: TurnContext,
  summary: string,
  challenges: string[]
): Promise<void> {
  await context.sendActivity(summary);
  if (challenges.length > 0) {
    const challengeBlock = "**Risks / follow-ups:**\n" + challenges.map((c) => "- " + c).join("\n");
    await context.sendActivity(challengeBlock);
  }
}

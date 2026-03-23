// src/l0.ts — Append-only L0 event ledger (NDJSON)
import { appendTextToFile, getTextFileContent } from "./graph";

const L0_PATH = "bot-data/02_L0_AUDIT_LOG/L0.ndjson";

export type L0Event =
  | {
      t: string;
      type: "user";
      cid: string;
      uid?: string;
      name?: string;
      msg: string;
    }
  | {
      t: string;
      type: "bot";
      cid: string;
      msg: string;
      model?: string;
    }
  | {
      t: string;
      type: "task";
      action: "upsert" | "confirm";
      taskId: string;
      owner?: string;
      dueAt?: string;
    };

/**
 * Append one event as a single NDJSON line. No redundant metadata.
 */
export async function appendL0(event: L0Event): Promise<void> {
  const line = JSON.stringify(event) + "\n";
  try {
    await appendTextToFile(L0_PATH, line);
  } catch (e) {
    console.error("L0 append error:", e);
  }
}

/**
 * Read recent conversation events (user + bot) for a given conversation.
 * For testing/medium size; for very large L0 consider streaming/tail.
 */
export async function readRecentL0Events(
  conversationId: string,
  limit = 20,
  maxDays?: number
): Promise<L0Event[]> {
  try {
    const { text } = await getTextFileContent(L0_PATH);
    const lines = text.split("\n").filter((s) => s.trim());
    const events: L0Event[] = [];
    const cutoffMs =
      maxDays != null ? Date.now() - maxDays * 24 * 60 * 60 * 1000 : 0;

    for (const line of lines) {
      try {
        const e = JSON.parse(line) as L0Event;
        if (e.type !== "user" && e.type !== "bot") continue;
        if (e.cid !== conversationId) continue;
        if (maxDays != null && new Date(e.t).getTime() < cutoffMs) continue;
        events.push(e);
      } catch {
        // skip malformed lines
      }
    }

    events.sort((a, b) => (a.t < b.t ? -1 : a.t > b.t ? 1 : 0));
    return events.slice(-limit);
  } catch {
    return [];
  }
}

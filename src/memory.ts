// src/memory.ts — Filter and format L0 events for AI context
import type { L0Event } from "./l0";

/**
 * Filter events by age (keep only last maxDays).
 */
export function filterByTime(events: L0Event[], maxDays: number): L0Event[] {
  const cutoff = Date.now() - maxDays * 24 * 60 * 60 * 1000;
  return events.filter((e) => new Date(e.t).getTime() >= cutoff);
}

/**
 * Format user/bot L0 events for GPT context (conversation history).
 * Task events are skipped.
 */
export function formatRecentMemory(events: L0Event[]): string {
  const conversation = events.filter(
    (e): e is L0Event & { msg: string } =>
      (e.type === "user" || e.type === "bot") && "msg" in e && e.msg != null
  );
  if (!conversation.length) return "No prior messages in scope.";

  return conversation
    .map((e) => {
      const who = e.type === "user" ? (e.name ?? "User") : "Bot";
      const text = String(e.msg).replace(/\s+/g, " ").trim();
      return `- [${e.t}] ${who}: ${text}`;
    })
    .join("\n");
}

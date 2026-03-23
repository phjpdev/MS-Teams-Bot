// src/handlers/confirmHandler.ts
// Handles task deadline confirmation ("confirm" keyword from task owner).

import { TurnContext } from "botbuilder";
import { findPendingConfirmationsForUser, confirmTaskDeadline } from "../taskRegistry.js";
import { syncTimelineToSharePoint } from "../timeline.js";
import { syncBudgetToSharePoint } from "../budget.js";
import { appendL0 } from "../l0.js";

async function safeSend(context: TurnContext, text: string) {
  const t = (text ?? "").trim();
  if (!t) return;
  await context.sendActivity(t);
}

function isConfirmText(text: string): boolean {
  const n = (text ?? "").toLowerCase().trim();
  return n === "confirm" || n === "confirmed" || n.startsWith("confirm ");
}

/**
 * Handles "confirm" messages from task owners.
 * Returns true if the message was a task confirmation and was handled.
 */
export async function handleConfirmation(
  context: TurnContext,
  text: string,
  aadObjectId: string,
  conversationId: string
): Promise<boolean> {
  if (!isConfirmText(text)) return false;

  const pending = await findPendingConfirmationsForUser(aadObjectId);
  if (!pending.length) return false;

  const latest = pending[pending.length - 1];
  const updated = await confirmTaskDeadline(latest.id, aadObjectId);

  if (!updated || updated.ownerAadObjectId !== aadObjectId) return false;

  try {
    await appendL0({
      t: new Date().toISOString(),
      type: "task",
      action: "confirm",
      taskId: updated.id,
      owner: updated.ownerDisplayName,
      dueAt: updated.dueAt ?? undefined,
    });
  } catch (e) {
    console.error("L0 task log error:", e);
  }

  await syncTimelineToSharePoint();
  await syncBudgetToSharePoint();
  await safeSend(
    context,
    `Confirmed. Deadline recorded for "${updated.title}" due ${updated.dueAt}${
      updated.dueTimezone ? ` (${updated.dueTimezone})` : ""
    }.`
  );
  return true;
}

// src/handlers/reminderHandler.ts
// Injects per-message reminders (due-soon, overdue, budget risk) with cooldown protection.

import { TurnContext } from "botbuilder";
import { getOverdueTasks, getDueSoonTasks, getBudgetRisks } from "../monitoring.js";
import { shouldSend, markSent } from "../notificationState.js";

const OVERDUE_COOLDOWN_MIN = 360;  // 6h
const DUE_SOON_COOLDOWN_MIN = 180; // 3h

async function safeSend(context: TurnContext, text: string) {
  const t = (text ?? "").trim();
  if (!t) return;
  await context.sendActivity(t);
}

/**
 * Sends any pending reminders for the current user.
 * Silently swallows errors so reminders never block the main flow.
 */
export async function handleReminders(
  context: TurnContext,
  aadObjectId: string
): Promise<void> {
  try {
    // Due soon (within 24h)
    const dueSoon = await getDueSoonTasks(24);
    for (const t of dueSoon.filter((t) => t.ownerAadObjectId === aadObjectId)) {
      const key = `${t.id}:dueSoon`;
      if (await shouldSend(key, DUE_SOON_COOLDOWN_MIN)) {
        await safeSend(
          context,
          `Reminder: "${t.title}" is due ${t.dueAt}${t.dueTimezone ? ` (${t.dueTimezone})` : ""}.`
        );
        await markSent(key);
      }
    }

    // Overdue
    const overdue = await getOverdueTasks();
    for (const t of overdue.filter((t) => t.ownerAadObjectId === aadObjectId)) {
      const key = `${t.id}:overdue`;
      if (await shouldSend(key, OVERDUE_COOLDOWN_MIN)) {
        await safeSend(
          context,
          `Overdue: "${t.title}" was due ${t.dueAt}. Please update status (in_progress/blocked/done) and note blockers.`
        );
        await markSent(key);
      }
    }

    // Budget overrun (variance > 20%)
    const budgetRisks = await getBudgetRisks(20);
    for (const t of budgetRisks.filter((t) => t.ownerAadObjectId === aadObjectId)) {
      const key = `${t.id}:budgetRisk`;
      if (await shouldSend(key, DUE_SOON_COOLDOWN_MIN)) {
        const est = t.estimatedCost ?? 0;
        const actual = t.actualCost ?? 0;
        const variance = actual - est;
        await safeSend(
          context,
          `Budget alert: "${t.title}" is over estimate (est. ${est}, actual ${actual}, variance +${variance}). Update actual cost or status if resolved.`
        );
        await markSent(key);
      }
    }
  } catch (e) {
    console.error("REMINDER ERROR:", e);
  }
}

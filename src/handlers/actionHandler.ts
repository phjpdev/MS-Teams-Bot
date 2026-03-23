// src/handlers/actionHandler.ts
// Executes structured actions returned by the AI in <actions> JSON blocks.

import { TurnContext } from "botbuilder";
import { upsertTask } from "../taskRegistry.js";
import { inferRequiredSkillsFromText, suggestBestOwner } from "../responsibilityEngine.js";
import { addUserRole, addUserSkill, loadUserDirectory } from "../userDirectory.js";
import { syncTimelineToSharePoint } from "../timeline.js";
import { syncBudgetToSharePoint } from "../budget.js";
import { appendL0 } from "../l0.js";

async function safeSend(context: TurnContext, text: string) {
  const t = (text ?? "").trim();
  if (!t) return;
  await context.sendActivity(t);
}

function isValidISO(dateStr?: string): boolean {
  if (!dateStr) return true;
  return !isNaN(new Date(dateStr).getTime());
}

export interface ActionContext {
  context: TurnContext;
  actions: any[];
  aadObjectId: string | undefined;
  fromName: string;
  conversationId: string;
  speakerTimezone: string;
}

/**
 * Processes all actions from the AI response.
 * Supports: update_speaker_profile, suggest_owner, upsert_task.
 */
export async function handleActions({
  context,
  actions,
  aadObjectId,
  fromName,
  conversationId,
  speakerTimezone,
}: ActionContext): Promise<void> {
  try {
    const userDir = await loadUserDirectory();

    for (const a of actions) {
      // ── update_speaker_profile ──────────────────────────────────────────
      if (a.type === "update_speaker_profile") {
        if (!aadObjectId) {
          await safeSend(
            context,
            "I can't save your profile here because I don't know your identity. Use a Teams channel or group chat where I can identify you."
          );
          continue;
        }
        const rolesAdded: string[] = [];
        const skillsAdded: string[] = [];
        if (Array.isArray(a.roles)) {
          for (const r of a.roles) {
            const role = String(r ?? "").trim();
            if (role) { await addUserRole(aadObjectId, role); rolesAdded.push(role); }
          }
        }
        if (Array.isArray(a.skills)) {
          for (const s of a.skills) {
            const skill = typeof s === "string" ? s : (s?.skill ?? "").trim();
            const level =
              typeof s === "object" && s !== null && typeof (s as any).level === "number"
                ? (s as any).level
                : undefined;
            if (skill) { await addUserSkill(aadObjectId, skill, level); skillsAdded.push(skill); }
          }
        }
        if (rolesAdded.length || skillsAdded.length) {
          const parts: string[] = [];
          if (rolesAdded.length) parts.push(`Roles: ${rolesAdded.join(", ")}`);
          if (skillsAdded.length) parts.push(`Skills: ${skillsAdded.join(", ")}`);
          await safeSend(context, `Profile updated in SharePoint. ${parts.join(". ")}`);
        }
        continue;
      }

      // ── suggest_owner ───────────────────────────────────────────────────
      if (a.type === "suggest_owner") {
        const title = String(a.title ?? "").trim() || "Unspecified task";
        const desc = String(a.description ?? "").trim();
        const requiredSkills = inferRequiredSkillsFromText(`${title} ${desc}`, userDir);
        const suggestion = await suggestBestOwner(userDir, requiredSkills);
        if (!suggestion) {
          await safeSend(
            context,
            "I can't suggest an owner yet because no matching competencies were found. Add relevant skills to UserDirectory.json (SharePoint)."
          );
          continue;
        }
        const matched = suggestion.matchedSkills.map((ms) => `${ms.skill}:${ms.level}`).join(", ");
        await safeSend(
          context,
          `Best suited owner (deterministic): ${suggestion.user.displayName}\n` +
            `- Matched skills: ${matched || "N/A"}\n` +
            `- Workload: ${suggestion.workload}\n\n` +
            `Reply with: "assign to ${suggestion.user.displayName}" to proceed, or specify another owner.`
        );
        continue;
      }

      // ── upsert_task ─────────────────────────────────────────────────────
      if (a.type !== "upsert_task" || typeof a.title !== "string") continue;

      // Resolve owner by display name if AAD ID missing
      let ownerAad: string | undefined = a.ownerAadObjectId;
      let ownerName: string | undefined = a.ownerDisplayName;
      if (!ownerAad && ownerName) {
        const match = userDir.users.find(
          (u) => u.displayName.toLowerCase() === String(ownerName).toLowerCase()
        );
        if (match) { ownerAad = match.aadObjectId; ownerName = match.displayName; }
      }

      // Suggest owner if still missing
      if (!ownerAad) {
        const requiredSkills = inferRequiredSkillsFromText(
          `${a.title} ${a.description ?? ""}`,
          userDir
        );
        const suggestion = await suggestBestOwner(userDir, requiredSkills);
        if (suggestion) {
          await safeSend(
            context,
            `Suggested owner for "${a.title}" is ${suggestion.user.displayName} (deterministic).\n` +
              `Please confirm by replying: "assign to ${suggestion.user.displayName}" (and restate deadline if needed).`
          );
          continue;
        }
        await safeSend(context, `A responsible owner must be defined for task "${a.title}".`);
        continue;
      }

      if (!isValidISO(a.dueAt)) {
        await safeSend(context, `I couldn't record "${a.title}" because the deadline format was invalid.`);
        continue;
      }

      // Enforce confirmation policy
      if (aadObjectId && ownerAad === aadObjectId) {
        a.status = "confirmed";
        a.dueNeedsConfirmation = false;
        a.dueConfirmedByAadObjectId = aadObjectId;
      } else if (a.dueAt) {
        a.status = "proposed";
        a.dueNeedsConfirmation = true;
      }

      const task = await upsertTask({
        title: a.title,
        description: a.description,
        ownerAadObjectId: ownerAad,
        ownerDisplayName: ownerName,
        dueAt: a.dueAt,
        dueTimezone: a.dueTimezone ?? speakerTimezone,
        status: a.status ?? "proposed",
        dueNeedsConfirmation: a.dueNeedsConfirmation ?? false,
        dueProposedByAadObjectId: a.dueProposedByAadObjectId ?? aadObjectId,
        dueProposedByName: a.dueProposedByName ?? fromName,
        dueConfirmedByAadObjectId: a.dueConfirmedByAadObjectId,
        startDate:
          typeof a.startDate === "string" && isValidISO(a.startDate) ? a.startDate : undefined,
        endDate:
          typeof a.endDate === "string" && isValidISO(a.endDate) ? a.endDate : undefined,
        estimatedCost: typeof a.estimatedCost === "number" ? a.estimatedCost : undefined,
        actualCost: typeof a.actualCost === "number" ? a.actualCost : undefined,
      });

      try {
        await appendL0({
          t: new Date().toISOString(),
          type: "task",
          action: "upsert",
          taskId: task.id,
          owner: task.ownerDisplayName,
          dueAt: task.dueAt ?? undefined,
        });
      } catch (e) {
        console.error("L0 task log error:", e);
      }

      await syncTimelineToSharePoint();
      await syncBudgetToSharePoint();

      if (task.dueNeedsConfirmation && task.ownerAadObjectId !== aadObjectId) {
        await safeSend(
          context,
          `${task.ownerDisplayName}: please confirm the deadline for "${task.title}" by replying "confirm".`
        );
      }
    }
  } catch (e) {
    console.error("ACTION APPLY ERROR:", e);
  }
}

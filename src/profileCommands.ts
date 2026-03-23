// src/profileCommands.ts
import {
  addUserRole,
  addUserSkill,
  loadUserDirectory,
  removeUserRole,
  removeUserSkill,
  setUserTimezone,
} from "./userDirectory";

type CmdResult =
  | { handled: true; reply: string }
  | { handled: false };

function parseIntSafe(s: string) {
  const n = Number(s);
  return Number.isFinite(n) ? n : undefined;
}

export async function handleProfileCommand(input: {
  text: string;
  aadObjectId?: string;
  displayName: string;
}) : Promise<CmdResult> {
  const t = input.text.trim();

  // Support both "/profile ..." and "profile ..."
  const lower = t.toLowerCase();
  const isProfile = lower.startsWith("/profile") || lower.startsWith("profile ");
  if (!isProfile) return { handled: false };

  if (!input.aadObjectId) {
    return { handled: true, reply: "I can’t identify your AAD identity in this chat. Try in a Teams group/channel (not anonymous)."};
  }

  const parts = t.replace(/^\/?profile\s*/i, "").trim().split(/\s+/);
  const sub = (parts[0] || "").toLowerCase();

  // HELP
  if (!sub || sub === "help") {
    return {
      handled: true,
      reply:
`Profile commands:
- /profile show
- /profile timezone Europe/Berlin
- /profile add-role "Backend Developer"
- /profile remove-role "Backend Developer"
- /profile add-skill azure 5
- /profile remove-skill azure`
    };
  }

  // SHOW
  if (sub === "show") {
    const dir = await loadUserDirectory();
    const me = dir.users.find(u => u.aadObjectId === input.aadObjectId);
    if (!me) return { handled: true, reply: "No profile found yet." };

    const roles = me.roles.length ? me.roles.join(", ") : "(none)";
    const skills = me.competencies.length
      ? me.competencies.map(c => `${c.skill}:${c.level ?? 3}`).join(", ")
      : "(none)";

    return {
      handled: true,
      reply:
`Profile for ${me.displayName}
- Timezone: ${me.timezone ?? "(not set)"}
- Roles: ${roles}
- Skills: ${skills}`
    };
  }

  // TIMEZONE
  if (sub === "timezone") {
    const tz = parts.slice(1).join(" ").trim();
    if (!tz) return { handled: true, reply: 'Usage: /profile timezone Europe/Berlin' };

    // Very light validation: must contain "/" for IANA-like tz
    if (!tz.includes("/")) {
      return { handled: true, reply: 'Timezone must be an IANA value like "Europe/Berlin", "America/Los_Angeles".' };
    }

    const updated = await setUserTimezone(input.aadObjectId, tz);
    return { handled: true, reply: `Saved timezone: ${updated?.timezone ?? tz}` };
  }

  // ADD ROLE
  if (sub === "add-role") {
    const role = t.match(/add-role\s+(.+)$/i)?.[1]?.trim() ?? "";
    if (!role) return { handled: true, reply: 'Usage: /profile add-role "Backend Developer"' };
    await addUserRole(input.aadObjectId, role);
    return { handled: true, reply: `Added role: ${role}` };
  }

  // REMOVE ROLE
  if (sub === "remove-role") {
    const role = t.match(/remove-role\s+(.+)$/i)?.[1]?.trim() ?? "";
    if (!role) return { handled: true, reply: 'Usage: /profile remove-role "Backend Developer"' };
    await removeUserRole(input.aadObjectId, role);
    return { handled: true, reply: `Removed role: ${role}` };
  }

  // ADD SKILL
  if (sub === "add-skill") {
    const skill = parts[1];
    const lvl = parseIntSafe(parts[2] ?? "");
    if (!skill) return { handled: true, reply: "Usage: /profile add-skill azure 5" };
    await addUserSkill(input.aadObjectId, skill, lvl);
    return { handled: true, reply: `Added skill: ${skill} (${lvl ?? 3})` };
  }

  // REMOVE SKILL
  if (sub === "remove-skill") {
    const skill = parts[1];
    if (!skill) return { handled: true, reply: "Usage: /profile remove-skill azure" };
    await removeUserSkill(input.aadObjectId, skill);
    return { handled: true, reply: `Removed skill: ${skill}` };
  }

  return { handled: true, reply: "Unknown profile command. Try: /profile help" };
}
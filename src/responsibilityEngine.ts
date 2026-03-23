// src/responsibilityEngine.ts
import { UserDirectory, UserProfile } from "./userDirectory";
import { countActiveTasksForUser } from "./taskRegistry";

export type OwnerSuggestion = {
  user: UserProfile;
  score: number;
  matchedSkills: Array<{ skill: string; level: number }>;
  workload: number;
};

/**
 * Build a normalized set of known skills from UserDirectory (dynamic, no hardcoding).
 */
function getAllKnownSkills(dir: UserDirectory): string[] {
  const s = new Set<string>();
  for (const u of dir.users) {
    for (const c of u.competencies ?? []) {
      const skill = (c.skill ?? "").trim().toLowerCase();
      if (skill) s.add(skill);
    }
  }
  return [...s];
}

/**
 * Infer required skills from text by matching against known skills.
 * - Exact substring match (case-insensitive).
 * - Also supports token/word match for single-word skills.
 */
export function inferRequiredSkillsFromText(text: string, dir: UserDirectory): string[] {
  const t = (text ?? "").toLowerCase();
  const skills = getAllKnownSkills(dir);

  const matched: string[] = [];

  for (const skill of skills) {
    if (!skill) continue;

    // direct substring match
    if (t.includes(skill)) {
      matched.push(skill);
      continue;
    }

    // token match (helps when skill is "react" and text has "React.")
    const tokens = t.split(/[^a-z0-9+.#/_-]+/i).filter(Boolean);
    if (tokens.includes(skill)) {
      matched.push(skill);
      continue;
    }
  }

  // de-dupe
  return [...new Set(matched)];
}

/**
 * Deterministic best owner selection:
 * score = sum(level * 10 for matched skills) - (workload * 5)
 */
export async function suggestBestOwner(dir: UserDirectory, requiredSkills: string[]): Promise<OwnerSuggestion | null> {
  const req = (requiredSkills ?? []).map(s => s.trim().toLowerCase()).filter(Boolean);
  if (!req.length) return null;

  const candidates = dir.users.filter(u => u.active);

  let best: OwnerSuggestion | null = null;

  for (const u of candidates) {
    const competencyMap = new Map<string, number>(
      (u.competencies ?? []).map(c => [(c.skill ?? "").trim().toLowerCase(), c.level ?? 3])
    );

    const matchedSkills: Array<{ skill: string; level: number }> = [];
    let skillScore = 0;

    for (const s of req) {
      const lvl = competencyMap.get(s);
      if (lvl) {
        matchedSkills.push({ skill: s, level: lvl });
        skillScore += lvl * 10;
      }
    }

    // must match at least one required skill
    if (skillScore === 0) continue;

    const workload = await countActiveTasksForUser(u.aadObjectId);
    const score = skillScore - workload * 5;

    const candidate: OwnerSuggestion = { user: u, score, matchedSkills, workload };

    if (!best || candidate.score > best.score) best = candidate;
  }

  return best;
}
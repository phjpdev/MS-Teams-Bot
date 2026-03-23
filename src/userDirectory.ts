// src/userDirectory.ts
import { getJsonFromSharePoint, uploadJsonToSharePoint } from "./graph";

export interface UserCompetency {
  skill: string;
  level?: number; // 1–5
}

export interface UserProfile {
  aadObjectId: string;
  displayName: string;
  email?: string;
  timezone?: string; // IANA e.g., "Europe/Berlin"
  roles: string[];
  competencies: UserCompetency[];
  active: boolean;
  updatedAt: string;
}

export interface UserDirectory {
  users: UserProfile[];
}

const USER_DIRECTORY_PATH = "bot-data/01_PROJECT_STATE/UserDirectory.json";

function nowIso() {
  return new Date().toISOString();
}

export async function loadUserDirectory(): Promise<UserDirectory> {
  const data = await getJsonFromSharePoint(USER_DIRECTORY_PATH, { users: [] });
  if (!data || typeof data !== "object" || !Array.isArray((data as any).users)) {
    return { users: [] };
  }
  return data as UserDirectory;
}

export async function saveUserDirectory(dir: UserDirectory) {
  await uploadJsonToSharePoint(USER_DIRECTORY_PATH, dir);
}

export async function upsertUser(params: {
  aadObjectId: string;
  displayName: string;
  email?: string;
  timezone?: string;
}) {
  const dir = await loadUserDirectory();
  const ts = nowIso();

  let user = dir.users.find(u => u.aadObjectId === params.aadObjectId);

  if (!user) {
    user = {
      aadObjectId: params.aadObjectId,
      displayName: params.displayName,
      email: params.email,
      timezone: params.timezone,
      roles: [],
      competencies: [],
      active: true,
      updatedAt: ts,
    };
    dir.users.push(user);
  } else {
    user.displayName = params.displayName || user.displayName;
    user.email = params.email || user.email;
    user.timezone = params.timezone || user.timezone;
    user.active = true;
    user.updatedAt = ts;
  }

  await saveUserDirectory(dir);
  return user;
}

export async function setUserTimezone(aadObjectId: string, timezone: string) {
  const dir = await loadUserDirectory();
  const u = dir.users.find(x => x.aadObjectId === aadObjectId);
  if (!u) return null;
  u.timezone = timezone;
  u.updatedAt = nowIso();
  await saveUserDirectory(dir);
  return u;
}

export async function addUserRole(aadObjectId: string, role: string) {
  const dir = await loadUserDirectory();
  const u = dir.users.find(x => x.aadObjectId === aadObjectId);
  if (!u) return null;

  const norm = role.trim();
  if (!norm) return u;

  const exists = u.roles.some(r => r.toLowerCase() === norm.toLowerCase());
  if (!exists) u.roles.push(norm);

  u.updatedAt = nowIso();
  await saveUserDirectory(dir);
  return u;
}

export async function removeUserRole(aadObjectId: string, role: string) {
  const dir = await loadUserDirectory();
  const u = dir.users.find(x => x.aadObjectId === aadObjectId);
  if (!u) return null;

  const target = role.trim().toLowerCase();
  u.roles = u.roles.filter(r => r.toLowerCase() !== target);

  u.updatedAt = nowIso();
  await saveUserDirectory(dir);
  return u;
}

export async function addUserSkill(aadObjectId: string, skill: string, level?: number) {
  const dir = await loadUserDirectory();
  const u = dir.users.find(x => x.aadObjectId === aadObjectId);
  if (!u) return null;

  const s = skill.trim();
  if (!s) return u;

  const lvl = typeof level === "number" ? Math.max(1, Math.min(5, level)) : 3;

  const existing = u.competencies.find(c => c.skill.toLowerCase() === s.toLowerCase());
  if (existing) existing.level = lvl;
  else u.competencies.push({ skill: s, level: lvl });

  u.updatedAt = nowIso();
  await saveUserDirectory(dir);
  return u;
}

export async function removeUserSkill(aadObjectId: string, skill: string) {
  const dir = await loadUserDirectory();
  const u = dir.users.find(x => x.aadObjectId === aadObjectId);
  if (!u) return null;

  const target = skill.trim().toLowerCase();
  u.competencies = u.competencies.filter(c => c.skill.toLowerCase() !== target);

  u.updatedAt = nowIso();
  await saveUserDirectory(dir);
  return u;
}

export function getPMs(dir: UserDirectory) {
  return dir.users.filter(u =>
    u.active && u.roles.some(r => r.toLowerCase().includes("pm") || r.toLowerCase().includes("project manager"))
  );
}
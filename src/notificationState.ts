// src/notificationState.ts
import { getJsonFromSharePoint, uploadJsonToSharePoint } from "./graph";

const PATH = "bot-data/01_PROJECT_STATE/NotificationState.json";

export type NotificationState = {
  lastSent: Record<string, string>;
};

function nowIso() {
  return new Date().toISOString();
}

export async function loadNotificationState(): Promise<NotificationState> {
  const data = await getJsonFromSharePoint(PATH, { lastSent: {} });
  if (!data || typeof data !== "object" || typeof (data as any).lastSent !== "object") {
    return { lastSent: {} };
  }
  return data as NotificationState;
}

export async function shouldSend(key: string, cooldownMinutes: number): Promise<boolean> {
  const state = await loadNotificationState();
  const last = state.lastSent[key];
  if (!last) return true;

  const lastMs = new Date(last).getTime();
  if (Number.isNaN(lastMs)) return true;

  const diffMin = (Date.now() - lastMs) / 60000;
  return diffMin >= cooldownMinutes;
}

export async function markSent(key: string): Promise<void> {
  const state = await loadNotificationState();
  state.lastSent[key] = nowIso();
  await uploadJsonToSharePoint(PATH, state);
}
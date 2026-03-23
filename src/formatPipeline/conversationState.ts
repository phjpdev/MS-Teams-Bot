// src/formatPipeline/conversationState.ts
// Unified per-conversation state for the format pipeline.
// Replaces pendingFormatConfirm.ts + awaitingTableData.ts (two separate SharePoint files → one).

import { getJsonFromSharePoint, uploadJsonToSharePoint } from "../graph.js";
import { type FormatDataType } from "./formatData.js";

const PATH = "bot-data/01_PROJECT_STATE/FormatConversationState.json";

// ── Types ────────────────────────────────────────────────────────────────────

export interface PendingFormatEntry {
  dataType: FormatDataType;
  json: object;
  createdAt: string;
}

export interface AwaitingTableEntry {
  dataType: FormatDataType;
  requestedAt: string;
}

export interface ConversationFormatState {
  pending?: PendingFormatEntry;   // formatted JSON awaiting Yes/No
  awaiting?: AwaitingTableEntry;  // set when bot asked user to paste table next
}

type StateMap = Record<string, ConversationFormatState>;

// ── Internal helpers ─────────────────────────────────────────────────────────

async function loadState(): Promise<StateMap> {
  const raw = await getJsonFromSharePoint(PATH, {});
  return (raw && typeof raw === "object" ? raw : {}) as StateMap;
}

async function saveState(state: StateMap): Promise<void> {
  await uploadJsonToSharePoint(PATH, state);
}

// ── Pending format (awaiting Yes/No confirmation) ────────────────────────────

export async function getPendingFormat(
  conversationId: string
): Promise<PendingFormatEntry | null> {
  const state = await loadState();
  return state[conversationId]?.pending ?? null;
}

export async function setPendingFormat(
  conversationId: string,
  entry: PendingFormatEntry
): Promise<void> {
  const state = await loadState();
  state[conversationId] = { ...state[conversationId], pending: entry, awaiting: undefined };
  await saveState(state);
}

export async function clearPendingFormat(conversationId: string): Promise<void> {
  const state = await loadState();
  if (state[conversationId]) {
    delete state[conversationId].pending;
    if (!state[conversationId].awaiting) delete state[conversationId];
  }
  await saveState(state);
}

// ── Awaiting table paste (two-step flow) ─────────────────────────────────────

export async function getAwaitingTableData(
  conversationId: string
): Promise<AwaitingTableEntry | null> {
  const state = await loadState();
  return state[conversationId]?.awaiting ?? null;
}

export async function setAwaitingTableData(
  conversationId: string,
  dataType: FormatDataType
): Promise<void> {
  const state = await loadState();
  state[conversationId] = {
    ...state[conversationId],
    awaiting: { dataType, requestedAt: new Date().toISOString() },
  };
  await saveState(state);
}

export async function clearAwaitingTableData(conversationId: string): Promise<void> {
  const state = await loadState();
  if (state[conversationId]) {
    delete state[conversationId].awaiting;
    if (!state[conversationId].pending) delete state[conversationId];
  }
  await saveState(state);
}

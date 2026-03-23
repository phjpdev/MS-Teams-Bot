// src/conversationRef.ts — Persist conversation reference for proactive (e.g. standup) messages
import { getJsonFromSharePoint, uploadJsonToSharePoint } from "./graph";

const PATH = "bot-data/01_PROJECT_STATE/StandupConversationRef.json";

export interface StoredConversationRef {
  conversationReference: {
    activityId?: string;
    user?: { id?: string; name?: string };
    bot?: { id?: string; name?: string };
    conversation?: { id?: string; conversationType?: string };
    channelId?: string;
    locale?: string;
    serviceUrl?: string;
  };
  updatedAt: string;
}

export async function loadConversationRef(): Promise<StoredConversationRef | null> {
  const data = await getJsonFromSharePoint(PATH, null);
  if (!data || typeof data !== "object") return null;
  const ref = (data as any).conversationReference;
  if (!ref || typeof ref !== "object" || !ref.conversation?.id || !ref.serviceUrl) return null;
  return data as StoredConversationRef;
}

export async function saveConversationRef(conversationReference: StoredConversationRef["conversationReference"]): Promise<void> {
  await uploadJsonToSharePoint(PATH, {
    conversationReference,
    updatedAt: new Date().toISOString(),
  });
}

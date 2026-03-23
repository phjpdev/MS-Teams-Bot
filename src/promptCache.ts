// src/promptCache.ts
import { getFileMeta, getTextFileContent } from "./graph";

type CacheEntry = {
  text: string;
  eTag?: string;
  lastModifiedDateTime?: string;
  fetchedAtMs: number;
};

const cache: Record<string, CacheEntry> = {};

/**
 * TTL is a safety check so we revalidate occasionally even if no one changes anything.
 * Change detection uses eTag/lastModified.
 */
const DEFAULT_TTL_MS = 60_000; // 60s (tunable)

export async function getCachedPrompt(params: {
  relativePath: string;
  ttlMs?: number;
  fallback: string;
}): Promise<string> {
  const { relativePath, ttlMs = DEFAULT_TTL_MS, fallback } = params;

  const existing = cache[relativePath];
  const now = Date.now();

  // fast path: cached and within TTL
  if (existing && now - existing.fetchedAtMs < ttlMs) {
    return existing.text;
  }

  // revalidate with metadata (cheap)
  try {
    const meta = await getFileMeta(relativePath);

    // if cached and unchanged, refresh fetchedAt and return cached
    if (
      existing &&
      meta.eTag &&
      existing.eTag === meta.eTag &&
      meta.lastModifiedDateTime &&
      existing.lastModifiedDateTime === meta.lastModifiedDateTime
    ) {
      existing.fetchedAtMs = now;
      return existing.text;
    }

    // changed or first load → download fresh
    const fresh = await getTextFileContent(relativePath);
    cache[relativePath] = {
      text: fresh.text,
      eTag: fresh.eTag,
      lastModifiedDateTime: fresh.lastModifiedDateTime,
      fetchedAtMs: now,
    };
    return fresh.text;
  } catch (e) {
    // if SharePoint fails, fall back to cached or fallback text
    if (existing?.text) return existing.text;
    return fallback;
  }
}

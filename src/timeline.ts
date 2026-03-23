// src/timeline.ts — Dedicated timeline file in SharePoint (03_TIMELINE)
import { loadTaskRegistry } from "./taskRegistry";
import { uploadJsonToSharePoint } from "./graph";

const TIMELINE_PATH = "bot-data/03_TIMELINE/Timeline.json";

export interface TimelineEntry {
  task: string;
  startDate: string;
  endDate: string;
  status: string;
  owner: string;
  deadline: string;
}

export interface TimelineFile {
  timeline: TimelineEntry[];
}

/**
 * Build timeline structure from TaskRegistry and upload to SharePoint.
 * Call after any task deadline or status change.
 */
export async function syncTimelineToSharePoint(): Promise<void> {
  try {
    const registry = await loadTaskRegistry();
    const timeline: TimelineEntry[] = registry.tasks.map((t) => ({
      task: t.title,
      startDate: t.startDate ?? t.createdAt ?? "",
      endDate: t.endDate ?? t.dueAt ?? "",
      status: t.status,
      owner: t.ownerDisplayName ?? t.ownerAadObjectId ?? "",
      deadline: t.dueAt ?? "",
    }));
    await uploadJsonToSharePoint(TIMELINE_PATH, { timeline });
  } catch (e) {
    console.error("Timeline sync error:", e);
  }
}

// src/taskRegistry.ts
import { getJsonFromSharePoint, uploadJsonToSharePoint } from "./graph";

export type TaskStatus = "proposed" | "confirmed" | "in_progress" | "blocked" | "done";

export interface TaskItem {
  id: string;
  title: string;
  description?: string;

  // Ownership
  ownerAadObjectId?: string;
  ownerDisplayName?: string;

  // Scheduling
  dueAt?: string; // ISO string
  dueTimezone?: string;
  startDate?: string; // ISO string, for timeline
  endDate?: string;   // ISO string, for timeline

  // Budget (for 04_BUDGET automation)
  estimatedCost?: number;
  actualCost?: number;

  // Confirmation workflow
  dueNeedsConfirmation?: boolean;
  dueProposedByAadObjectId?: string;
  dueProposedByName?: string;
  dueConfirmedByAadObjectId?: string;
  dueConfirmedAt?: string;

  // Metadata
  status: TaskStatus;
  createdAt: string;
  updatedAt: string;
}

export interface TaskRegistry {
  tasks: TaskItem[];
}

const TASK_REGISTRY_PATH = "bot-data/01_PROJECT_STATE/TaskRegistry.json";

function nowIso() {
  return new Date().toISOString();
}

function makeId() {
  // good enough unique id for SharePoint stored JSON
  return `task_${new Date().toISOString().replace(/[:.]/g, "-")}`;
}

export async function loadTaskRegistry(): Promise<TaskRegistry> {
  const data = await getJsonFromSharePoint(TASK_REGISTRY_PATH, { tasks: [] });
  if (!data || typeof data !== "object" || !Array.isArray((data as any).tasks)) {
    return { tasks: [] };
  }
  return data as TaskRegistry;
}

export async function saveTaskRegistry(registry: TaskRegistry) {
  await uploadJsonToSharePoint(TASK_REGISTRY_PATH, registry);
}

export async function upsertTask(partial: Partial<TaskItem> & { title: string }) {
  const registry = await loadTaskRegistry();
  const ts = nowIso();

  // Try match by id first
  let task: TaskItem | undefined;
  if (partial.id) {
    task = registry.tasks.find(t => t.id === partial.id);
  }

  // Otherwise match by title (case-insensitive)
  if (!task) {
    const titleKey = partial.title.trim().toLowerCase();
    task = registry.tasks.find(t => t.title.trim().toLowerCase() === titleKey);
  }

  if (!task) {
    task = {
      id: makeId(),
      title: partial.title.trim(),
      description: partial.description ?? "",
      status: partial.status ?? "proposed",
      createdAt: ts,
      updatedAt: ts,
      ownerAadObjectId: partial.ownerAadObjectId,
      ownerDisplayName: partial.ownerDisplayName,
      dueAt: partial.dueAt,
      dueTimezone: partial.dueTimezone,
      startDate: partial.startDate,
      endDate: partial.endDate,
      estimatedCost: partial.estimatedCost,
      actualCost: partial.actualCost,
      dueNeedsConfirmation: partial.dueNeedsConfirmation ?? false,
      dueProposedByAadObjectId: partial.dueProposedByAadObjectId,
      dueProposedByName: partial.dueProposedByName,
      dueConfirmedByAadObjectId: partial.dueConfirmedByAadObjectId,
      dueConfirmedAt: partial.dueConfirmedAt,
    };
    registry.tasks.push(task);
  } else {
    Object.assign(task, partial);
    task.updatedAt = ts;
  }

  await saveTaskRegistry(registry);
  return task;
}

export async function listTasksForUser(aadObjectId: string) {
  const registry = await loadTaskRegistry();
  return registry.tasks.filter(t => t.ownerAadObjectId === aadObjectId);
}

export async function findPendingConfirmationsForUser(aadObjectId: string) {
  const registry = await loadTaskRegistry();
  return registry.tasks.filter(
    t =>
      t.ownerAadObjectId === aadObjectId &&
      t.dueNeedsConfirmation === true &&
      !!t.dueAt
  );
}

export async function confirmTaskDeadline(taskId: string, confirmerAadObjectId: string) {
  const registry = await loadTaskRegistry();
  const task = registry.tasks.find(t => t.id === taskId);
  if (!task) return null;

  task.dueNeedsConfirmation = false;
  task.dueConfirmedByAadObjectId = confirmerAadObjectId;
  task.dueConfirmedAt = nowIso();
  task.status = task.status === "proposed" ? "confirmed" : task.status;
  task.updatedAt = nowIso();

  await saveTaskRegistry(registry);
  return task;
}

export async function countActiveTasksForUser(aadObjectId: string) {
  const registry = await loadTaskRegistry();
  return registry.tasks.filter(
    t =>
      t.ownerAadObjectId === aadObjectId &&
      t.status !== "done"
  ).length;
}
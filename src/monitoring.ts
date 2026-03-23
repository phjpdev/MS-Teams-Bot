// src/monitoring.ts
import { loadTaskRegistry, TaskItem } from "./taskRegistry";

export async function getOverdueTasks(): Promise<TaskItem[]> {
  const reg = await loadTaskRegistry();
  const now = Date.now();
  return reg.tasks.filter(t => {
    if (!t.dueAt) return false;
    const ms = new Date(t.dueAt).getTime();
    if (Number.isNaN(ms)) return false;
    return ms < now && t.status !== "done";
  });
}

export async function getDueSoonTasks(hours: number): Promise<TaskItem[]> {
  const reg = await loadTaskRegistry();
  const now = Date.now();
  const windowEnd = now + hours * 3600 * 1000;

  return reg.tasks.filter(t => {
    if (!t.dueAt) return false;
    const ms = new Date(t.dueAt).getTime();
    if (Number.isNaN(ms)) return false;
    return ms >= now && ms <= windowEnd && t.status !== "done";
  });
}

/** Tasks with budget overrun (variance exceeds threshold % of estimated cost). */
export async function getBudgetRisks(thresholdPercent = 20): Promise<TaskItem[]> {
  const reg = await loadTaskRegistry();
  const threshold = thresholdPercent / 100;
  return reg.tasks.filter(t => {
    const est = typeof t.estimatedCost === "number" ? t.estimatedCost : 0;
    const actual = typeof t.actualCost === "number" ? t.actualCost : 0;
    if (est <= 0) return false;
    const varianceRatio = (actual - est) / est;
    return varianceRatio > threshold && t.status !== "done";
  });
}

/** Tasks due within hours that are not done (milestone at risk). */
export async function getTimelineRisks(hours = 48): Promise<TaskItem[]> {
  return getDueSoonTasks(hours);
}
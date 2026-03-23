// src/budget.ts — Dedicated budget file in SharePoint (04_BUDGET)
import { loadTaskRegistry } from "./taskRegistry";
import { uploadJsonToSharePoint } from "./graph";

const BUDGET_PATH = "bot-data/04_BUDGET/Budget.json";

export interface BudgetEntry {
  taskName: string;
  estimatedCost: number;
  actualCost: number;
  variance: number;
  status: string;
}

export interface BudgetFile {
  budget: BudgetEntry[];
}

/**
 * Build budget structure from TaskRegistry and upload to SharePoint.
 * Variance = actualCost - estimatedCost. Call after any cost or task change.
 */
export async function syncBudgetToSharePoint(): Promise<void> {
  try {
    const registry = await loadTaskRegistry();
    const budget: BudgetEntry[] = registry.tasks.map((t) => {
      const estimated = typeof t.estimatedCost === "number" ? t.estimatedCost : 0;
      const actual = typeof t.actualCost === "number" ? t.actualCost : 0;
      return {
        taskName: t.title,
        estimatedCost: estimated,
        actualCost: actual,
        variance: actual - estimated,
        status: t.status,
      };
    });
    await uploadJsonToSharePoint(BUDGET_PATH, { budget });
  } catch (e) {
    console.error("Budget sync error:", e);
  }
}

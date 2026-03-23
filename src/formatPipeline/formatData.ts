// src/formatPipeline/formatData.ts — Format unstructured table data via OpenAI into Timeline / BudgetPlan / Qualifikationsmatrix
import OpenAI from "openai";
import "dotenv/config";

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY! });

export type FormatDataType = "timeline" | "budgetplan" | "qualifikationsmatrix";

const TIMELINE_SYSTEM = `You are a data-cleaning formatter. The user will send raw table data (e.g. from Excel). Your job is to CREATE A CLEAN structure from whatever they send. Handle as much as possible; do not reject for "ambiguous" or "missing" layout.

Behavior:
1. CLEAN: Remove columns that add no value (e.g. weekday names only with no dates, redundant labels). Keep task/activity names and any date or numeric columns.
2. NORMALIZE dates: Infer format (DD.MM.YY or MM.DD.YY) from the data. Prefer DD.MM.YY if values look like day-first (e.g. 26.11.25). Convert all to DD.MM.YYYY for output columns.
3. BUILD timeline grid: If the table has "start date", "end date", "duration", or date headers (numbers 1-17, week numbers, or actual dates), derive a list of date columns spanning the project (e.g. weekly from earliest to latest). For each task row, put "X" in every column whose date falls between that task's start and end (inclusive). If only one date is given, put "X" in that column. If no dates, leave values empty.
4. INVALID DATA: If a row has end date BEFORE start date, still include the row in "rows" but add to "warnings": "TaskName: end date (X) is before start date (Y). Please correct or confirm to save."
5. Only return {"error": "..."} when the input is completely empty or not table-like. Otherwise always return the schema below, with "warnings" optional.

Output schema (valid JSON only, no markdown):
{
  "type": "timeline",
  "columns": ["DD.MM.YYYY", ...],
  "rows": [
    { "task": "Task or milestone name", "values": ["X" or "", ...] }
  ],
  "warnings": ["Optional list of issues: end before start, missing dates, etc."]
}`;

const BUDGETPLAN_SYSTEM = `You are a data-cleaning formatter. The user will send raw table data (e.g. from Excel). CREATE A CLEAN budget structure. Handle as much as possible.

Behavior:
1. CLEAN: Remove non-numeric or non-date columns that are not posten names or amounts. Normalize decimal separators (e.g. 10.000 or 10,000 → 10000). Infer period columns (monthly MM/YYYY or similar).
2. NORMALIZE: Use numbers not strings. Ensure "Gesamt" / total column or compute it from row values.
3. INVALID DATA: If a row has negative amounts where they should be positive, or totals that don't match sum of values, still include the row and add to "warnings": "PostenName: brief issue. Please correct or confirm to save."
4. Only return {"error": "..."} when input is empty or not table-like. Otherwise always return the schema below.

Output schema (valid JSON only, no markdown):
{
  "type": "budgetplan",
  "columns": ["MM/YYYY", ..., "Gesamt"],
  "rows": [
    { "posten": "Item name", "values": [number, ...], "total": number }
  ],
  "warnings": ["Optional list of issues"]
}`;

const QUALIFIKATIONSMATRIX_SYSTEM = `You are a data-cleaning formatter. The user will send raw table data (e.g. from Excel). CREATE A CLEAN qualification matrix. Handle as much as possible.

Behavior:
1. CLEAN: Identify person columns and skill rows. Remove extra columns. Normalize percentages to 0-100 (if values are 0-1, multiply by 100).
2. Treat "x" or "X" in cells as 50 (default skill level). Do not ask the user how to interpret "x"; apply this default.
3. INVALID DATA: If a value is outside 0-100, clamp or add to "warnings". Still include the row.
4. Only return {"error": "..."} when input is empty or not table-like. Otherwise always return the schema below.

Output schema (valid JSON only, no markdown):
{
  "type": "qualifikationsmatrix",
  "columns": ["Person1", ...],
  "rows": [
    { "skill": "Skill name", "values": [0-100, ...] }
  ],
  "warnings": ["Optional list of issues"]
}`;

function getSystemPrompt(dataType: FormatDataType): string {
  switch (dataType) {
    case "timeline":
      return TIMELINE_SYSTEM;
    case "budgetplan":
      return BUDGETPLAN_SYSTEM;
    case "qualifikationsmatrix":
      return QUALIFIKATIONSMATRIX_SYSTEM;
    default:
      return TIMELINE_SYSTEM;
  }
}

function stripMarkdownCodeBlock(raw: string): string {
  const s = (raw || "").trim();
  const m = s.match(/^```(?:json)?\s*([\s\S]*?)```$/);
  if (m) return m[1].trim();
  return s;
}

function parseJsonResponse(raw: string): { json?: object; error?: string; warnings?: string[] } {
  const cleaned = stripMarkdownCodeBlock(raw);
  try {
    const obj = JSON.parse(cleaned);
    if (!obj || typeof obj !== "object") return { error: "Invalid format." };
    const hasValidStructure = Array.isArray((obj as any).columns) && Array.isArray((obj as any).rows);
    const errMsg = (obj as any).error;
    const warnings = Array.isArray((obj as any).warnings) ? (obj as any).warnings : undefined;
    if (errMsg && typeof errMsg === "string" && !hasValidStructure) {
      return { error: errMsg };
    }
    if (hasValidStructure) {
      return { json: obj, warnings };
    }
    return { error: errMsg || "Could not parse formatted data." };
  } catch {
    return { error: "Could not parse formatted data. Please send a clear table (e.g. copy from Excel)." };
  }
}

function validateFormattedJson(obj: any, dataType: FormatDataType): boolean {
  if (!obj || typeof obj !== "object") return false;
  if (obj.type !== dataType) return false;
  if (!Array.isArray(obj.columns) || !Array.isArray(obj.rows)) return false;
  return true;
}

export async function formatTableWithOpenAI(
  rawText: string,
  dataType: FormatDataType
): Promise<{ json?: object; error?: string; warnings?: string[] }> {
  const systemPrompt = getSystemPrompt(dataType);
  const userContent = (rawText || "").trim() || "No data provided.";
  const model = process.env.AI_MODEL || "gpt-4.1-mini";

  try {
    const response = await client.responses.create({
      model,
      input: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userContent },
      ],
    });
    const output = response.output_text || "";
    const result = parseJsonResponse(output);
    if (result.error) return result;
    if (result.json && !validateFormattedJson(result.json, dataType)) {
      return { error: "Format did not match expected schema. Please check your table and try again." };
    }
    return result;
  } catch (e: any) {
    console.error("formatTableWithOpenAI error:", e);
    return { error: e?.message || "Formatting failed. Please try again." };
  }
}

export const FORMAT_DATA_TYPE_LABELS: Record<FormatDataType, string> = {
  timeline: "Timeline",
  budgetplan: "Budget plan",
  qualifikationsmatrix: "Qualifikationsmatrix",
};

export const FORMAT_SHAREPOINT_PATHS: Record<FormatDataType, string> = {
  timeline: "bot-data/03_TIMELINE/TimelineTable.json",
  budgetplan: "bot-data/04_BUDGET/BudgetPlan.json",
  qualifikationsmatrix: "bot-data/01_PROJECT_STATE/Qualifikationsmatrix.json",
};

/** One-line summary of formatted JSON for preview (e.g. "Columns: 45 dates; Rows: Rohbau, Elektrik, concept, ..."). */
export function formatPreviewSummary(json: any): string {
  if (!json || typeof json !== "object") return "";
  const cols = Array.isArray(json.columns) ? json.columns : [];
  const rows = Array.isArray(json.rows) ? json.rows : [];
  const rowLabels = rows
    .slice(0, 5)
    .map((r: any) => (r.task ?? r.posten ?? r.skill ?? "?")).filter(Boolean);
  const more = rows.length > 5 ? ` +${rows.length - 5} more` : "";
  return `Columns: ${cols.length}${cols.length ? ` (${cols[0]} … ${cols[cols.length - 1]})` : ""}; Rows: ${rowLabels.join(", ")}${more}.`;
}

const APPLY_CORRECTIONS_SYSTEM = `You have a structured JSON document (timeline, budgetplan, or qualifikationsmatrix). The user will provide corrections in natural language (e.g. "Rohbau end date 16.10.2026", "Elektrik end 16.03.2026", "use original start", "update and save"). Your job is to apply those corrections to the JSON and return the complete updated document.

Rules:
- Output only valid JSON in the exact same schema (type, columns, rows). No markdown, no explanation.
- For timeline: If the user gives a new end date (or start date) for a task by name, update that row so the "X" marks span the correct range. Use DD.MM.YYYY. Use the EXACT year the user provides (e.g. 16.03.2026 means year 2026, not 2025). Extend the "columns" array if the new date is outside the current range. "Use original start" or "use origin one" means keep the existing start for that task.
- For budgetplan/qualifikationsmatrix: Update the numbers or values the user specifies.
- Do not add or require "owner" or "owners" — the schema has no owner field.
- If the user's message cannot be interpreted as corrections, return {"error": "Brief message: paste the full table again or reply Yes to save current version."}`;

/**
 * Apply user corrections to an existing formatted JSON (timeline/budgetplan/qualifikationsmatrix).
 * Used when the user has a pending format and sends e.g. "Rohbau end 16.10.2026, Elektrik 16.03.2026, update and save".
 */
export async function applyCorrectionsToFormattedData(
  pendingJson: object,
  userMessage: string,
  dataType: FormatDataType
): Promise<{ json?: object; error?: string }> {
  const model = process.env.AI_MODEL || "gpt-4.1-mini";
  const userContent =
    `Current JSON:\n${JSON.stringify(pendingJson, null, 2)}\n\nUser corrections:\n${(userMessage || "").trim()}`;

  try {
    const response = await client.responses.create({
      model,
      input: [
        { role: "system", content: APPLY_CORRECTIONS_SYSTEM },
        { role: "user", content: userContent },
      ],
    });
    const output = response.output_text || "";
    const result = parseJsonResponse(output);
    if (result.error) return result;
    if (result.json && !validateFormattedJson(result.json, dataType)) {
      return { error: "Corrected format did not match schema. Paste the full table again or reply Yes to save current version." };
    }
    const obj = result.json as any;
    const cleanJson =
      obj.warnings !== undefined ? { type: obj.type, columns: obj.columns, rows: obj.rows } : obj;
    return { json: cleanJson };
  } catch (e: any) {
    console.error("applyCorrectionsToFormattedData error:", e);
    return { error: e?.message || "Could not apply corrections. Paste the full table again or reply Yes to save current version." };
  }
}

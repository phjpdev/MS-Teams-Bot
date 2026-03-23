// src/context.ts
import { getJsonFromSharePoint } from "./graph";
import { getCachedPrompt } from "./promptCache";

const PROJECT_STATE_PATH = "bot-data/01_PROJECT_STATE/ProjectState.json";
const TEAM_ROLES_PATH = "bot-data/01_PROJECT_STATE/TeamRoles.json";
const USER_DIRECTORY_PATH = "bot-data/01_PROJECT_STATE/UserDirectory.json";
const TASK_REGISTRY_PATH = "bot-data/01_PROJECT_STATE/TaskRegistry.json";

// Saved format data (populated by the format pipeline)
const TIMELINE_PATH = "bot-data/03_TIMELINE/TimelineTable.json";
const BUDGET_PATH = "bot-data/04_BUDGET/BudgetPlan.json";
const QUALMATRIX_PATH = "bot-data/01_PROJECT_STATE/Qualifikationsmatrix.json";

// governance docs (optional; if you later move these too)
const RULESET_PATH = "bot-data/00_GOVERNANCE/Ruleset.md";
const LEADERSHIP_PATH = "bot-data/00_GOVERNANCE/Leadership_Guidelines.md";
const CHARTER_PATH = "bot-data/00_GOVERNANCE/Project_Charter.md";
const SYSTEM_PROMPT_PATH = "bot-data/00_GOVERNANCE/SystemPrompt.md";

import axios from "axios";

async function getTextFromSharePoint(relativePath: string, fallback = "N/A") {
  // Minimal read for .md via Graph content endpoint
  try {
    const tokenRes = await axios.post(
      `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: process.env.GRAPH_CLIENT_ID!,
        client_secret: process.env.GRAPH_CLIENT_SECRET!,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );
    const token = tokenRes.data.access_token;

    const url = `https://graph.microsoft.com/v1.0/drives/${process.env.GRAPH_DRIVE_ID}/root:/${process.env.GRAPH_BASE_FOLDER}/${relativePath}:/content`;
    const res = await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
    return typeof res.data === "string" ? res.data : JSON.stringify(res.data, null, 2);
  } catch {
    return fallback;
  }
}

export async function loadContext() {
  const projectStateObj = await getJsonFromSharePoint(PROJECT_STATE_PATH, {});
  const teamRolesObj = await getJsonFromSharePoint(TEAM_ROLES_PATH, {});
  const userDirectoryObj = await getJsonFromSharePoint(USER_DIRECTORY_PATH, { users: [] });
  const taskRegistryObj = await getJsonFromSharePoint(TASK_REGISTRY_PATH, { tasks: [] });

  // Load saved format data (timeline, budget, qualifikationsmatrix)
  const [timelineObj, budgetObj, qualmatrixObj] = await Promise.all([
    getJsonFromSharePoint(TIMELINE_PATH, null),
    getJsonFromSharePoint(BUDGET_PATH, null),
    getJsonFromSharePoint(QUALMATRIX_PATH, null),
  ]);

  const ruleset = await getTextFromSharePoint(RULESET_PATH, "N/A");
  const leadership = await getTextFromSharePoint(LEADERSHIP_PATH, "N/A");
  const charter = await getTextFromSharePoint(CHARTER_PATH, "N/A");

  const systemPromptTemplate = await getCachedPrompt({
    relativePath: SYSTEM_PROMPT_PATH,
    fallback: "N/A",
    ttlMs: 60_000,
  });

  // Build saved data text (only include non-empty data)
  const savedDataParts: string[] = [];
  if (timelineObj) savedDataParts.push(`### Timeline\n${JSON.stringify(timelineObj, null, 2)}`);
  if (budgetObj) savedDataParts.push(`### Budget Plan\n${JSON.stringify(budgetObj, null, 2)}`);
  if (qualmatrixObj) savedDataParts.push(`### Qualifikationsmatrix\n${JSON.stringify(qualmatrixObj, null, 2)}`);
  const savedDataText = savedDataParts.length > 0 ? savedDataParts.join("\n\n") : "";

  return {
    systemPromptTemplate,
    rulesText: [ruleset, leadership, charter].join("\n\n---\n\n"),
    projectStateText: JSON.stringify(projectStateObj, null, 2),
    teamRolesText: JSON.stringify(teamRolesObj, null, 2),
    userDirectoryText: JSON.stringify(userDirectoryObj, null, 2),
    taskRegistryText: JSON.stringify(taskRegistryObj, null, 2),
    savedDataText,
  };
}
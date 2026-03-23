// src/ai/prompt.ts

export function renderSystemPrompt(params: {
  template: string;
  nowUtc: string;
  projectState?: string;
  teamRoles?: string;
  rules?: string;
  recentMemory?: string;
  userDirectory?: string;
  taskRegistry?: string;
  savedData?: string;
}) {
  const map: Record<string, string> = {
    NOW_UTC: params.nowUtc,
    PROJECT_STATE: params.projectState ?? "N/A",
    TEAM_ROLES: params.teamRoles ?? "N/A",
    USER_DIRECTORY: params.userDirectory ?? "N/A",
    TASK_REGISTRY: params.taskRegistry ?? "N/A",
    RULES: params.rules ?? "N/A",
    RECENT_MEMORY: params.recentMemory ?? "N/A",
    SAVED_DATA: params.savedData ?? "N/A",

    // Optional: use these if the SharePoint template uses {{DEADLINE_RULES}} etc.
    DEADLINE_RULES: `If a user mentions ANY deadline:
1. Convert it to ISO 8601 in UTC.
2. Always return "dueAt" as ISO string.
3. Always return "dueTimezone".
4. If timezone not explicitly provided → use SPEAKER timezone.
5. NEVER return null for dueAt if a date/time was mentioned.
6. Use CURRENT UTC TIME to resolve relative expressions.
7. If date without time → default 17:00 UTC.`,
    USER_AWARENESS_RULES: `- Identify speaker using SPEAKER block.
- Use UserDirectory to map display names to AAD IDs.
- Use roles/competencies to suggest intelligent task owners.
- If someone assigns a deadline to another person → proposed + needs confirmation.
- If speaker commits to their own deadline → confirmed.`,
    ACTION_OUTPUT_RULES: `Return:
1) Human-readable response
2) <actions> JSON block
Support action types:
- update_speaker_profile
- upsert_task
- suggest_owner`,
  };

  const DATA_FORMAT_RULE = `
CRITICAL — DATA FORMATTING PIPELINE:
- When a user pastes table data (Excel, TSV, matrix, budget, timeline), the code pipeline handles it automatically BEFORE this AI is called.
- NEVER ask users to resend data with a trigger phrase.
- NEVER ask clarifying questions about table structure, column interpretation, or what "x" values mean.
- NEVER instruct users on how to format or prefix their data.
- If somehow raw tabular data reaches you, respond only with: "I'm processing your data." and nothing else.`;

  const savedDataSection = params.savedData
    ? `\n\nSAVED PROJECT DATA (Timeline, Budget, Qualifikationsmatrix — use this to answer questions about project data):\n${params.savedData}`
    : "";

  // Replace {{KEY}} placeholders
  return ((params.template || "") + DATA_FORMAT_RULE + savedDataSection)
    .replace(/\{\{([A-Z0-9_]+)\}\}/g, (_, key) => map[key] ?? `{{${key}}}`)
    .trim();
}

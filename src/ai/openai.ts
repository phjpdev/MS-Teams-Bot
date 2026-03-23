// src/ai/openai.ts 
import OpenAI from "openai";
import "dotenv/config";
import { renderSystemPrompt } from "./prompt.js";

const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY! });

export async function askAI(input: {
  userMessage: string;
  systemPromptTemplate: string;
  projectState?: string;
  teamRoles?: string;
  rules?: string;
  recentMemory?: string;
  userDirectory?: string;
  taskRegistry?: string;
  savedData?: string;
}) {
  const nowUtc = new Date().toISOString();

  const systemPrompt = renderSystemPrompt({
    template: input.systemPromptTemplate,
    nowUtc,
    projectState: input.projectState,
    teamRoles: input.teamRoles,
    rules: input.rules,
    recentMemory: input.recentMemory,
    userDirectory: input.userDirectory,
    taskRegistry: input.taskRegistry,
    savedData: input.savedData,
  });

  const model = process.env.AI_MODEL || "gpt-4.1-mini";

  const response = await client.responses.create({
    model,
    input: [
      { role: "system", content: systemPrompt },
      { role: "user", content: input.userMessage },
    ],
  });

  return response.output_text || "(empty response)";
}

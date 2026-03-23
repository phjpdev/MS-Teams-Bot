// src/formatPipeline/detectFormat.ts
// All data-type detection logic: trigger phrases, keyword scanning, structural analysis.

import { type FormatDataType } from "./formatData.js";

export { type FormatDataType };

export const FORMAT_TRIGGERS: { pattern: RegExp; dataType: FormatDataType }[] = [
  {
    pattern:
      /^\s*(\/format\s+timeline|format\s+as\s+timeline|here\s+is\s+the\s+new\s+timeline|here\s+is\s+the\s+timeline|this\s+is\s+the\s+timeline|new\s+timeline|save\s+as\s+timeline|timeline\s*:)\s*[:\s]*/i,
    dataType: "timeline",
  },
  {
    pattern:
      /^\s*(\/format\s+budget|format\s+as\s+budget|here\s+is\s+the\s+budget|new\s+budget|save\s+as\s+budget|budget\s*:)\s*[:\s]*/i,
    dataType: "budgetplan",
  },
  {
    pattern:
      /^\s*(\/format\s+(?:qualifikationsmatrix|matrix)|format\s+as\s+(?:qualifikationsmatrix|matrix)|(?:this\s+is\s+)?qualification\s+matrix|(?:here\s+is\s+the\s+)?qualifikationsmatrix|matrix\s*:)\s*[:\s]*/i,
    dataType: "qualifikationsmatrix",
  },
];

/** True when text contains tab-separated multi-row data (Excel paste). */
export function hasTabularData(text: string): boolean {
  return text.includes("\t") && text.split("\n").filter((l) => l.trim()).length > 1;
}

/** Detect data type from keywords anywhere in the message. */
export function detectDataTypeKeyword(text: string): FormatDataType | null {
  if (/qualif|matrix|metrix/i.test(text)) return "qualifikationsmatrix";
  if (/\btimeline\b|\bzeitplan\b|\bmeilenstein\b|\bgantt\b/i.test(text)) return "timeline";
  if (/\bbudget\b|\bkosten\b|\bbudgetplan\b|\bfinanzplan\b/i.test(text)) return "budgetplan";
  return null;
}

/** Detect data type purely from the structure of pasted tab-separated table content. */
export function detectDataTypeFromTableContent(text: string): FormatDataType | null {
  if (/start\s*date|end\s*date|startdatum|enddatum|start-datum|end-datum/i.test(text))
    return "timeline";
  if (/\bkw\b.*\d{1,2}|\d{1,2}\.\d{2}\.\d{4}.*\d{1,2}\.\d{2}\.\d{4}/i.test(text))
    return "timeline";

  if (
    /\bgesamt\b/i.test(text) &&
    /\bposten\b|\binvestition\b|\bkosten\b|\bausgaben\b/i.test(text)
  )
    return "budgetplan";
  if (
    /\b(januar|februar|m[äa]rz|april|mai|juni|juli|august|september|oktober|november|dezember)\b/i.test(
      text
    ) &&
    /\d{2,}/.test(text)
  )
    return "budgetplan";

  const numbers = text.match(/\b\d{1,3}\b/g) ?? [];
  if (numbers.length >= 10) {
    const inRange = numbers.filter((n) => parseInt(n) <= 100).length;
    if (inRange / numbers.length > 0.75) return "qualifikationsmatrix";
  }

  return null;
}

/**
 * Main entry point: given plain text (HTML already stripped) and optional embedded TSV,
 * returns the detected data type + the table data to format, or null if not a format command.
 */
export function detectFormatCommand(
  text: string,
  embeddedTableTsv: string | null = null
): { dataType: FormatDataType; tableData: string } | null {
  const t = (text || "").trim();

  // 1. Strict trigger phrase
  for (const { pattern, dataType } of FORMAT_TRIGGERS) {
    const m = t.match(pattern);
    if (m) {
      const rawData = t.slice(m[0].length).trim();
      const tableData = hasTabularData(rawData)
        ? rawData
        : embeddedTableTsv ?? "";
      return { dataType, tableData };
    }
  }

  // 2. Table present in message + keyword
  if (embeddedTableTsv) {
    const kwType = detectDataTypeKeyword(t) ?? detectDataTypeFromTableContent(embeddedTableTsv);
    if (kwType) return { dataType: kwType, tableData: embeddedTableTsv };
  }

  // 3. Plain TSV paste (no typed text) — detect type from table structure
  if (hasTabularData(t)) {
    const kwType = detectDataTypeKeyword(t);
    if (kwType) return { dataType: kwType, tableData: t };
    const structType = detectDataTypeFromTableContent(t);
    if (structType) return { dataType: structType, tableData: t };
  }

  // 4. Intent-only (user typed "here is the matrix" — Teams stripped the table)
  const hint = detectDataTypeKeyword(t);
  if (hint && /\b(here|this\s+is|save|send|new|update[d]?|paste|share)\b/i.test(t)) {
    const tableData = embeddedTableTsv ?? "";
    return { dataType: hint, tableData };
  }

  return null;
}

export function isFormatConfirm(text: string): boolean {
  return ["yes", "save", "ja", "speichern"].includes((text || "").toLowerCase().trim());
}

export function isFormatCancel(text: string): boolean {
  return ["no", "cancel", "discard", "abbrechen"].includes((text || "").toLowerCase().trim());
}

export function looksLikeCorrections(text: string): boolean {
  const t = (text || "").trim();
  const lower = t.toLowerCase();
  if (["yes", "save", "ja", "speichern", "no", "cancel", "discard", "abbrechen"].includes(lower))
    return false;
  if (t.length < 6) return false;
  const hasDate = /\d{1,2}\.\d{1,2}\.\d{2,4}/.test(t);
  const hasDateKeywords = /\b(end|start|date|update|due)\b/i.test(t);
  if (!hasDate && !hasDateKeywords) return false;
  return detectFormatCommand(text) === null;
}

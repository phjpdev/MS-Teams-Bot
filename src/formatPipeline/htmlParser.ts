// src/formatPipeline/htmlParser.ts
// Parses Teams HTML messages (Excel paste) into plain text + TSV table data.

/** Convert an HTML table (Teams rich paste) to tab-separated text. Returns null if no table found. */
export function htmlTableToTsv(html: string): string | null {
  if (!/<table/i.test(html)) return null;

  // Use string-splitting instead of regex matching — regex fails for tables with
  // row/col span attributes, nested elements, or multi-line cell content (common in Teams).
  const rows: string[] = [];
  const trParts = html.split(/<tr\b[^>]*>/gi);
  for (let i = 1; i < trParts.length; i++) {
    const trContent = trParts[i].split(/<\/tr>/gi)[0] ?? trParts[i];
    const cellParts = trContent.split(/<t[dh]\b[^>]*>/gi);
    const cells: string[] = [];
    for (let j = 1; j < cellParts.length; j++) {
      const cellHtml = cellParts[j].split(/<\/t[dh]>/gi)[0] ?? cellParts[j];
      const cellText = cellHtml
        .replace(/<br\s*\/?>/gi, " ")
        .replace(/<[^>]+>/g, "")
        .replace(/&nbsp;/g, " ")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&#\d+;/g, "")
        .replace(/\s+/g, " ")
        .trim();
      cells.push(cellText);
    }
    if (cells.length > 0 && cells.some((c) => c.length > 0)) {
      rows.push(cells.join("\t"));
    }
  }
  return rows.length > 0 ? rows.join("\n") : null;
}

/** Strip HTML tags and tables, returning only the plain typed text. */
export function htmlToPlainText(html: string): string {
  return html
    .replace(/<table[\s\S]*?<\/table>/gi, "")
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Given raw activity.text, returns:
 * - plainText: typed text without HTML / table markup
 * - tableTsv: tab-separated table content, or null if none
 */
export function extractMessageParts(rawText: string): {
  plainText: string;
  tableTsv: string | null;
} {
  if (/<table/i.test(rawText)) {
    return {
      plainText: htmlToPlainText(rawText),
      // If TSV extraction fails, pass raw HTML as fallback — OpenAI handles HTML tables fine
      tableTsv: htmlTableToTsv(rawText) ?? rawText,
    };
  }
  // Plain text — tabs mean Excel was pasted without typed text
  const tableTsv =
    rawText.includes("\t") && rawText.split("\n").filter((l) => l.trim()).length > 1
      ? rawText
      : null;
  return { plainText: rawText, tableTsv };
}

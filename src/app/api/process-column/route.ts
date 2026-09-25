import { NextRequest, NextResponse } from "next/server";
import * as cheerio from "cheerio";
import { GoogleGenAI } from "@google/genai";

export const maxDuration = 60;
export const dynamic = "force-dynamic";

const FALLBACK_MODEL_CHAIN = [
  "gemini-2.5-flash",
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-3.5-flash",
  "gemini-3.7-flash",
];

const STANDARD_HEADERS = {
  "User-Agent":
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
  Accept:
    "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
  "Accept-Language": "en-US,en;q=0.9",
};

function convertHtmlToCleanMarkdown(htmlText: string): string {
  const $ = cheerio.load(htmlText);
  $("script, style, noscript, nav, footer, header, svg, iframe, form").remove();

  $("table").each((_, tbl) => {
    const rows: string[] = [];
    $(tbl)
      .find("tr")
      .each((_, tr) => {
        const cells: string[] = [];
        $(tr)
          .find("th, td")
          .each((_, td) => {
            const txt = $(td).text().replace(/\s+/g, " ").trim();
            if (txt) cells.push(txt);
          });
        if (cells.length > 0) rows.push("| " + cells.join(" | ") + " |");
      });
    if (rows.length > 0) {
      $(tbl).replaceWith("\n\n" + rows.join("\n") + "\n\n");
    }
  });

  $("h1, h2, h3, h4, h5").each((_, h) => {
    const t = $(h).text().replace(/\s+/g, " ").trim();
    if (t) $(h).replaceWith(`\n### ${t}\n`);
  });

  return $("body").text().replace(/\n\s*\n\s*\n/g, "\n\n").replace(/\s+/g, " ").trim();
}

async function fetchPageWithTimeout(url: string, timeoutMs = 3500): Promise<{ text: string; isDiscontinued: boolean }> {
  if (!url || !url.startsWith("http")) return { text: "", isDiscontinued: false };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      headers: STANDARD_HEADERS,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (res.status === 404) {
      return { text: "HTTP 404 Not Found - Course link is dead or discontinued.", isDiscontinued: true };
    }
    if (!res.ok) return { text: "", isDiscontinued: false };
    const html = await res.text();
    const md = convertHtmlToCleanMarkdown(html);
    const hasDiscontinuedText = /no longer accepting applications|course discontinued|program discontinued|withdrawn for 2026|admissions closed permanently/i.test(md);
    return { text: md, isDiscontinued: hasDiscontinuedText };
  } catch {
    clearTimeout(timer);
    return { text: "", isDiscontinued: false };
  }
}

async function callGemini(
  apiKeys: string[],
  selectedModel: string,
  prompt: string,
  systemInstruction: string
): Promise<{ text: string; modelUsed: string; keyUsedIndex: number }> {
  let lastError: any = null;
  const modelsToTry = [selectedModel, ...FALLBACK_MODEL_CHAIN.filter((m) => m !== selectedModel)];

  for (let kIdx = 0; kIdx < apiKeys.length; kIdx++) {
    const key = apiKeys[kIdx].trim();
    if (!key) continue;
    const ai = new GoogleGenAI({ apiKey: key });

    for (let mIdx = 0; mIdx < modelsToTry.length; mIdx++) {
      const curModel = modelsToTry[mIdx];
      try {
        const response = await ai.models.generateContent({
          model: curModel,
          contents: prompt,
          config: {
            systemInstruction,
            temperature: 0.1,
            responseMimeType: "application/json",
          },
        });
        if (response?.text) {
          return { text: response.text, modelUsed: curModel, keyUsedIndex: kIdx + 1 };
        }
      } catch (err: any) {
        lastError = err;
      }
    }
  }
  throw new Error(`Gemini failover exhausted: ${lastError?.message || "Unknown error"}`);
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      columnName = "",
      columnIndex = 0,
      batchRows = [],
      apiKeys = [],
      model = "gemini-2.5-flash",
      customPrompt = "",
      fallbackDirectContent = "",
      cachedWebpages = {},
    } = body;

    if (!columnName) {
      return NextResponse.json({ error: "Missing required columnName." }, { status: 400 });
    }

    if (!Array.isArray(batchRows) || batchRows.length === 0) {
      return NextResponse.json({ error: "No rows provided for verification." }, { status: 400 });
    }

    if (!Array.isArray(apiKeys) || apiKeys.length === 0) {
      return NextResponse.json({ error: "At least one Gemini API key is required." }, { status: 400 });
    }

    // Scrape or load from cache for each row
    const scrapedContextPerCourse: { [rowId: number]: string } = {};
    const discontinuedFlags: { [rowId: number]: boolean } = {};
    const newlyScraped: { [url: string]: string } = {};

    const fetchTasks = batchRows.map(async (row: any) => {
      const rId = row.rowIndex;
      const url = row.url || "";
      if (url.startsWith("http")) {
        if (cachedWebpages && cachedWebpages[url]) {
          scrapedContextPerCourse[rId] = cachedWebpages[url];
        } else {
          const { text: md, isDiscontinued } = await fetchPageWithTimeout(url, 3500);
          if (md) {
            scrapedContextPerCourse[rId] = md;
            newlyScraped[url] = md;
          }
          if (isDiscontinued) {
            discontinuedFlags[rId] = true;
          }
        }
      }
    });

    await Promise.allSettled(fetchTasks);

    // Format rows for prompt
    const formattedBatch = batchRows.map((r: any) => {
      const rId = r.rowIndex;
      let webText = scrapedContextPerCourse[rId] || "";
      if (!webText && fallbackDirectContent) {
        webText = fallbackDirectContent.slice(0, 4000);
      }

      return {
        row_index: rId,
        entity_title: r.entityTitle || `Record #${rId + 1}`,
        campus: r.campus || "",
        current_box_value: r.currentValue !== undefined && r.currentValue !== null ? String(r.currentValue).trim() : "",
        is_page_marked_discontinued_or_404: Boolean(discontinuedFlags[rId]),
        scraped_webpage_content: webText
          ? webText.slice(0, 6000)
          : "Extract or synthesize verified academic value based on the program title.",
      };
    });

    const systemInstruction = `You are an expert academic database auditor, registrar, and quality assurance officer.
Your task is to perform an explicit BOX-BY-BOX VERIFICATION for the target column: "${columnName}".
Evaluate the current box value against the scraped 2026 webpage content.

DECISION PROTOCOL FOR "${columnName}":
1. DISCONTINUED COURSES:
   - If the webpage indicates the course is discontinued, not offered in 2026, or returns HTTP 404:
     - Set status: "DISCONTINUED"
     - Set verified_value: "Discontinued in 2026"
     - Set confidence: "HIGH"
     - Provide the reason and evidence quote.
2. MISSING / PLACEHOLDER VALUES:
   - If current box value is empty, null, or placeholder (e.g. 'tbc', 'abc', 'tbd', 'none', 'n/a', '-'):
     - Extract genuine 2026 data from the scraped webpage.
     - Never output placeholder or 'tbc'.
     - Set status: "AUTOFILLED".
3. OUTDATED VS CURRENT DATA:
   - If current box value has existing data:
     - Check if the scraped webpage contains updated 2026 figures or details.
     - If outdated: Extract the new 2026 published value. Status: "UPDATED".
     - If already matches: Keep the existing value. Status: "VERIFIED_CURRENT".
4. CONFIDENCE SCORING:
   - "HIGH": Exact figure or curriculum table found directly on official webpage.
   - "MEDIUM": Calculated (e.g. multiplied semester fee x 2 to get annual) or inferred from program structure.
   - "LOW": Webpage was sparse; synthesized based on title. Human review advised.
5. EVIDENCE SNIPPET:
   - Quote a 1-sentence excerpt from the scraped text supporting your finding.
6. DATA TYPE INTEGRITY:
   - If the target column is clearly a fee/number and the current value is numeric, output pure numbers without currency text (or standard numbers).
   - If the column is overview, output a clean 2-3 sentence description.
   - If structure, output comma-separated subject codes and credits.`;

    const prompt = `
TARGET COLUMN TO VERIFY: "${columnName}"
${customPrompt.trim() ? `ADDITIONAL USER INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}

Verify the following ${formattedBatch.length} records for column "${columnName}":
${JSON.stringify(formattedBatch, null, 2)}

REQUIRED JSON OUTPUT FORMAT:
{
  "results": [
    {
      "row_index": ${formattedBatch[0]?.row_index ?? 0},
      "is_up_to_date": false,
      "verified_value": "Genuine 2026 verified data...",
      "status": "UPDATED" | "AUTOFILLED" | "VERIFIED_CURRENT" | "DISCONTINUED",
      "confidence": "HIGH" | "MEDIUM" | "LOW",
      "reason": "Brief explanation of what was verified or changed",
      "evidence_snippet": "Exact quote or line from webpage where this fact was identified"
    }
  ]
}
`;

    const { text, modelUsed, keyUsedIndex } = await callGemini(apiKeys, model, prompt, systemInstruction);

    let results: any[] = [];
    const match = text.match(/\{[\s\S]*\}/);
    if (match) {
      try {
        const parsed = JSON.parse(match[0]);
        if (Array.isArray(parsed.results)) {
          results = parsed.results;
        }
      } catch (err) {
        console.warn("Parse error in process-column:", err);
      }
    }

    return NextResponse.json({
      success: true,
      columnName,
      columnIndex,
      results,
      modelUsed,
      keyUsedIndex,
      newlyScraped,
    });
  } catch (err: any) {
    console.error("Process Column Error:", err);
    return NextResponse.json(
      { error: err?.message || "Failed to process column verification." },
      { status: 500 }
    );
  }
}

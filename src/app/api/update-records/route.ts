import { NextRequest, NextResponse } from "next/server";
import * as cheerio from "cheerio";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { GoogleGenAI } from "@google/genai";

export const maxDuration = 60;
export const dynamic = "force-dynamic";

// Model auto-switch priority list
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

function normalizeKey(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function extractPageData(htmlText: string, sourceUrl: string): { tables: string[]; text: string; subLinks: string[] } {
  const $ = cheerio.load(htmlText);
  const subLinks: string[] = [];

  try {
    const parsedOrigin = new URL(sourceUrl).origin;
    $("a[href]").each((_, el) => {
      const href = $(el).attr("href");
      if (href) {
        try {
          const resolved = new URL(href, sourceUrl).toString();
          if (
            resolved.startsWith(parsedOrigin) &&
            !resolved.includes("#") &&
            !resolved.match(/\.(pdf|jpg|jpeg|png|zip|doc|docx|mp4|svg)$/i)
          ) {
            const isRelevant = /course|catalog|syllabus|schedule|curriculum|department|faculty|degree|program|result|timetable|grade/i.test(resolved);
            if (isRelevant && !subLinks.includes(resolved) && resolved !== sourceUrl) {
              subLinks.push(resolved);
            }
          }
        } catch {}
      }
    });
  } catch {}

  $("script, style, noscript, nav, footer, iframe, header, svg").remove();

  const tables: string[] = [];
  $("table").each((idx, el) => {
    const rows: string[] = [];
    $(el)
      .find("tr")
      .each((_, tr) => {
        const cols: string[] = [];
        $(tr)
          .find("th, td")
          .each((_, td) => {
            const txt = $(td).text().replace(/\s+/g, " ").trim();
            if (txt) cols.push(txt);
          });
        if (cols.length > 0) rows.push(cols.join(" | "));
      });
    if (rows.length > 0) {
      tables.push(`[TABLE ${idx + 1} from ${sourceUrl}]\n` + rows.join("\n"));
    }
  });

  const text = $("body").text().replace(/\s+/g, " ").trim();
  return { tables, text, subLinks };
}

/**
 * Deep crawler that tracks and lists all crawled pages
 */
async function deepCrawlUniversitySite(
  rootUrl: string,
  maxSubpages = 6
): Promise<{ content: string; warning?: string; pagesScraped: string[] }> {
  const visited = new Set<string>();
  const pagesScraped: string[] = [];
  const allTables: string[] = [];
  const textChunks: string[] = [];
  let antiBotDetected = false;

  try {
    const res = await fetch(rootUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
    if (!res.ok) {
      return {
        content: `Error fetching URL (HTTP ${res.status} ${res.statusText}).`,
        warning: `HTTP ${res.status}: The university site may block automated cloud scrapers or require login.`,
        pagesScraped: [rootUrl],
      };
    }
    const html = await res.text();
    visited.add(rootUrl);
    pagesScraped.push(rootUrl);

    if (
      html.includes("cf-browser-verification") ||
      html.includes("Access Denied") ||
      html.includes("Checking your browser") ||
      html.includes("Cloudflare")
    ) {
      antiBotDetected = true;
    }

    const { tables, text, subLinks } = extractPageData(html, rootUrl);
    allTables.push(...tables);
    textChunks.push(`=== ROOT PAGE (${rootUrl}) ===\n${text.slice(0, 15000)}`);

    const targetsToCrawl = subLinks.slice(0, maxSubpages);
    if (targetsToCrawl.length > 0) {
      const crawlPromises = targetsToCrawl.map(async (subUrl) => {
        if (visited.has(subUrl)) return;
        visited.add(subUrl);
        pagesScraped.push(subUrl);
        try {
          const subRes = await fetch(subUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
          if (subRes.ok) {
            const subHtml = await subRes.text();
            const extracted = extractPageData(subHtml, subUrl);
            allTables.push(...extracted.tables);
            textChunks.push(`=== SUBPAGE (${subUrl}) ===\n${extracted.text.slice(0, 10000)}`);
          }
        } catch {}
      });
      await Promise.allSettled(crawlPromises);
    }
  } catch (err: any) {
    return {
      content: `Failed to scrape university URL: ${err?.message}`,
      warning: `Scraper error: ${err?.message}`,
      pagesScraped: [rootUrl],
    };
  }

  const combined = `=== DEEP SCRAPED TABLES ===\n${allTables.join("\n\n")}\n\n=== MAIN CONTENT ===\n${textChunks.join("\n\n")}`;

  return {
    content: combined,
    warning: antiBotDetected ? "Cloudflare / Anti-bot challenge detected on target website." : undefined,
    pagesScraped,
  };
}

async function callGeminiWithFullFailover(
  apiKeys: string[],
  initialModel: string,
  prompt: string,
  systemInstruction: string
): Promise<{ text: string; modelUsed: string; keyUsedIndex: number }> {
  let lastError: Error | null = null;
  const modelsToTry = [initialModel, ...FALLBACK_MODEL_CHAIN.filter((m) => m !== initialModel)];

  for (let keyIdx = 0; keyIdx < apiKeys.length; keyIdx++) {
    const key = apiKeys[keyIdx].trim();
    if (!key) continue;

    const ai = new GoogleGenAI({ apiKey: key });

    for (let modelIdx = 0; modelIdx < modelsToTry.length; modelIdx++) {
      const currentModel = modelsToTry[modelIdx];

      try {
        const response = await ai.models.generateContent({
          model: currentModel,
          contents: prompt,
          config: {
            systemInstruction,
            temperature: 0.1,
            responseMimeType: "application/json",
          },
        });

        if (response && response.text) {
          return {
            text: response.text,
            modelUsed: currentModel,
            keyUsedIndex: keyIdx + 1,
          };
        }
      } catch (err: any) {
        lastError = err;
      }
    }
  }

  throw new Error(`All Gemini API keys & model fallbacks failed. Last error: ${lastError?.message || "Unknown error"}`);
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const url = (formData.get("url") as string | null) || "";
    const rawDirectContent = (formData.get("directContent") as string | null) || "";
    const apiKeysRaw = formData.get("apiKeys") as string | null;
    const selectedModel = (formData.get("model") as string | null) || "gemini-2.5-flash";
    const customPrompt = (formData.get("customPrompt") as string | null) || "";

    if (!file || !apiKeysRaw) {
      return NextResponse.json({ error: "Missing required Excel file or API key(s)." }, { status: 400 });
    }

    if (!url.trim() && !rawDirectContent.trim()) {
      return NextResponse.json({ error: "Please provide either a university URL or paste webpage text directly." }, { status: 400 });
    }

    const apiKeys = apiKeysRaw.split(",").map((k) => k.trim()).filter(Boolean);
    if (apiKeys.length === 0) {
      return NextResponse.json({ error: "At least one valid Google Gemini API key must be provided." }, { status: 400 });
    }

    // Step A: Parse uploaded Excel file
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rawExcelRows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (rawExcelRows.length === 0) {
      return NextResponse.json({ error: "The uploaded Excel sheet contains no records." }, { status: 400 });
    }

    const originalColumns = Object.keys(rawExcelRows[0] || {});

    // Step B: Resolve Web Content (Direct Content or Deep Web Scraper)
    let webText = "";
    let scraperWarning: string | undefined = undefined;
    let pagesScraped: string[] = [];

    if (rawDirectContent.trim().length > 30) {
      webText = `=== USER PROVIDED DIRECT WEBPAGE TEXT / HTML (AUTHENTIC SOURCE) ===\n${rawDirectContent.trim()}`;
      pagesScraped = ["Direct User Webpage Text / HTML (Bypassed Scraper)"];
    } else if (url.trim().startsWith("http")) {
      const crawlResult = await deepCrawlUniversitySite(url.trim(), 7);
      webText = crawlResult.content;
      scraperWarning = crawlResult.warning;
      pagesScraped = crawlResult.pagesScraped;
    } else {
      return NextResponse.json({ error: "Invalid URL provided and no direct content pasted." }, { status: 400 });
    }

    // Step C: Build Resilient Reconciliation Prompt
    const systemInstruction = `You are a world-class academic database administrator and university registrar AI.
Your mission is to rigorously reconcile, update, and autofill course, student, faculty, and marksheet records.
RULES:
1. FUZZY IDENTIFIER MATCHING: Match rows by course codes or titles flexibly (e.g. 'CS101' matches 'CS 101', 'CS-101', 'COMP101').
2. UPDATE OUTDATED INFORMATION: Overwrite any outdated fields with 2026 data found on the web content.
3. AUTOFILL BLANK / MISSING FIELDS: If an Excel cell is empty, blank (""), null, or missing, look up that record in the web content and populate it.
4. KEY PRESERVATION: Keep exact column keys as provided: ${JSON.stringify(originalColumns)}.
5. DO NOT RETURN AN UNMODIFIED SHEET IF DATA EXISTS: If the web content contains matching courses/grades, you MUST extract and update them.`;

    const prompt = `
${customPrompt.trim() ? `TEACHER'S SPECIAL INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
ORIGINAL EXCEL HEADERS: ${JSON.stringify(originalColumns)}
TOTAL ROWS IN EXCEL: ${rawExcelRows.length}

EXCEL ROWS TO RECONCILE (WITH ROW INDEX):
${JSON.stringify(
  rawExcelRows.slice(0, 450).map((r, idx) => ({ _row_index: idx, ...r })),
  null,
  1
)}

WEBPAGE DATA (COURSES, CATALOGS, RESULTS, TABLES):
${webText.slice(0, 80000)}

TASK:
- Compare every row with the webpage content.
- For each updated or auto-filled row, specify:
  * "row_index": The exact integer _row_index from the row object above.
  * "identifier": The course code or name matched.
  * "updated_data": An object containing the columns and their updated/populated values.
  * "changed_columns": Array of column names that were updated or autofilled.
  * "reason": Why it changed.
- If new courses exist on the web that are missing from Excel, include them in "new_rows".

OUTPUT STRICT JSON SCHEMA:
{
  "updated_rows": [
    {
      "row_index": 0,
      "identifier": "CS-101",
      "updated_data": { "Course Title": "Intro to AI", "Credits": "3" },
      "changed_columns": ["Course Title", "Credits"],
      "reason": "Updated outdated title and filled missing credits"
    }
  ],
  "new_rows": [
    {
      "data": { "Course Code": "CS-405", "Course Title": "Cloud Computing" },
      "reason": "New 2026 course catalog entry"
    }
  ],
  "summary": "Summary of updates made."
}
`;

    // Step D: AI Execution
    const { text: rawAiResponse, modelUsed, keyUsedIndex } = await callGeminiWithFullFailover(
      apiKeys,
      selectedModel,
      prompt,
      systemInstruction
    );

    let aiResult: any;
    try {
      aiResult = JSON.parse(rawAiResponse);
    } catch {
      const match = rawAiResponse.match(/\{[\s\S]*\}/);
      if (match) {
        aiResult = JSON.parse(match[0]);
      } else {
        throw new Error("Unable to parse AI JSON response.");
      }
    }

    // Step E: Resilient Column & Row Mapping
    const normalizedToOriginalCol: { [norm: string]: string } = {};
    originalColumns.forEach((c) => {
      normalizedToOriginalCol[normalizeKey(c)] = c;
    });

    const updatedExcelRows = [...rawExcelRows];
    const modifiedCellSet = new Set<string>();

    const updatedRowsList = aiResult.updated_rows || [];
    for (const item of updatedRowsList) {
      let rIdx = item.row_index;

      if (rIdx === undefined || rIdx < 0 || rIdx >= updatedExcelRows.length) {
        const id = String(item.identifier || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        if (id) {
          const foundIdx = updatedExcelRows.findIndex((row) =>
            Object.values(row).some((val) =>
              String(val).toLowerCase().replace(/[^a-z0-9]/g, "").includes(id)
            )
          );
          if (foundIdx !== -1) rIdx = foundIdx;
        }
      }

      if (rIdx !== undefined && rIdx >= 0 && rIdx < updatedExcelRows.length) {
        const uData = item.updated_data || {};
        const cCols = item.changed_columns || [];

        for (const [key, val] of Object.entries(uData)) {
          const normKey = normalizeKey(key);
          const targetCol = normalizedToOriginalCol[normKey];
          if (targetCol) {
            const oldVal = updatedExcelRows[rIdx][targetCol];
            if (String(oldVal) !== String(val)) {
              updatedExcelRows[rIdx][targetCol] = val;
              modifiedCellSet.add(`${rIdx}_${targetCol}`);
            }
          }
        }

        for (const col of cCols) {
          const targetCol = normalizedToOriginalCol[normalizeKey(col)];
          if (targetCol) {
            modifiedCellSet.add(`${rIdx}_${targetCol}`);
          }
        }
      }
    }

    const startNewRowIdx = updatedExcelRows.length;
    const newRowsList = aiResult.new_rows || [];
    for (let i = 0; i < newRowsList.length; i++) {
      const rawNewData = newRowsList[i].data || {};
      const alignedRow: any = {};

      originalColumns.forEach((c) => {
        const norm = normalizeKey(c);
        let foundVal = "";
        for (const [k, v] of Object.entries(rawNewData)) {
          if (normalizeKey(k) === norm) {
            foundVal = String(v);
            break;
          }
        }
        alignedRow[c] = foundVal;
      });

      const targetIdx = startNewRowIdx + i;
      updatedExcelRows.push(alignedRow);
      originalColumns.forEach((c) => {
        modifiedCellSet.add(`${targetIdx}_${c}`);
      });
    }

    // Step F: Build Styled Excel Workbook
    const outputWb = new ExcelJS.Workbook();
    const ws = outputWb.addWorksheet("2026 Updated Records");

    ws.columns = originalColumns.map((col) => ({ header: col, key: col, width: 22 }));

    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF1E3C72" },
      };
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    const yellowFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };

    updatedExcelRows.forEach((rowObj, rIdx) => {
      const addedRow = ws.addRow(rowObj);
      originalColumns.forEach((colKey, cIdx) => {
        const cell = addedRow.getCell(cIdx + 1);
        cell.border = {
          top: { style: "thin", color: { argb: "FFD9D9D9" } },
          left: { style: "thin", color: { argb: "FFD9D9D9" } },
          bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
          right: { style: "thin", color: { argb: "FFD9D9D9" } },
        };

        if (modifiedCellSet.has(`${rIdx}_${colKey}`)) {
          cell.fill = yellowFill;
          cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF7F6000" } };
        }
      });
    });

    ws.columns.forEach((col) => {
      let maxLen = 14;
      col.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 4, 45);
    });

    // Summary tab
    const summaryWs = outputWb.addWorksheet("Update Summary");
    summaryWs.addRow(["RecordSync Execution Summary"]);
    summaryWs.addRow(["Total Courses/Rows Processed", rawExcelRows.length]);
    summaryWs.addRow(["Autofilled & Updated Rows", updatedRowsList.length]);
    summaryWs.addRow(["Newly Discovered Courses/Rows", newRowsList.length]);
    summaryWs.addRow(["AI Model Used", modelUsed]);
    summaryWs.addRow(["API Key Account Used", keyUsedIndex]);
    if (scraperWarning) summaryWs.addRow(["Scraper Diagnostics", scraperWarning]);
    if (customPrompt.trim()) summaryWs.addRow(["Teacher Custom Instructions", customPrompt.trim()]);
    summaryWs.addRow(["Pages Scraped List:", pagesScraped.join(" | ")]);
    summaryWs.addRow(["AI Summary Notes", aiResult.summary || "Reconciliation completed."]);

    const excelBuffer = await outputWb.xlsx.writeBuffer();

    return new NextResponse(excelBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="Updated_2026_${file.name}"`,
        "X-Updated-Count": String(updatedRowsList.length),
        "X-New-Count": String(newRowsList.length),
        "X-Model-Used": modelUsed,
        "X-Key-Used": String(keyUsedIndex),
        "X-Summary-Notes": encodeURIComponent(aiResult.summary || ""),
        "X-Scraper-Warning": encodeURIComponent(scraperWarning || ""),
        "X-Pages-Scraped": encodeURIComponent(JSON.stringify(pagesScraped)),
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err?.message || "Internal server error occurred." }, { status: 500 });
  }
}

import { NextRequest, NextResponse } from "next/server";
import * as cheerio from "cheerio";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
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
            const isRelevant = /course|program|catalog|syllabus|schedule|curriculum|department|faculty|degree|admission|fee/i.test(resolved);
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
 * Intelligent Crawler:
 * 1. Crawls root URL
 * 2. Reads direct in-row course links (e.g. course_detail_web_link, course_fee_web_link) directly from the Excel sheet!
 * 3. Crawls priority sublinks up to maxPages limit.
 */
async function crawlUniversityPages(
  rootUrl: string,
  inSheetUrls: string[],
  maxPages = 20
): Promise<{ content: string; warning?: string; pagesScraped: string[] }> {
  const visited = new Set<string>();
  const pagesScraped: string[] = [];
  const allTables: string[] = [];
  const textChunks: string[] = [];
  let antiBotDetected = false;

  // 1. Fetch root page if provided
  if (rootUrl && rootUrl.startsWith("http")) {
    try {
      const res = await fetch(rootUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
      if (res.ok) {
        const html = await res.text();
        visited.add(rootUrl);
        pagesScraped.push(rootUrl);

        if (html.includes("cf-browser-verification") || html.includes("Access Denied") || html.includes("Cloudflare")) {
          antiBotDetected = true;
        }

        const { tables, text, subLinks } = extractPageData(html, rootUrl);
        allTables.push(...tables);
        textChunks.push(`=== ROOT (${rootUrl}) ===\n${text.slice(0, 10000)}`);

        // Enqueue high-priority sublinks from root
        for (const link of subLinks) {
          if (!inSheetUrls.includes(link) && pagesScraped.length < maxPages) {
            inSheetUrls.push(link);
          }
        }
      }
    } catch (e: any) {
      console.warn("Root fetch error:", e?.message);
    }
  }

  // 2. Fetch specific course detail & fee pages (direct in-sheet links)
  const linksToCrawl = inSheetUrls.filter((u) => u && u.startsWith("http") && !visited.has(u)).slice(0, maxPages);

  const crawlTasks = linksToCrawl.map(async (linkUrl) => {
    if (visited.has(linkUrl)) return;
    visited.add(linkUrl);
    pagesScraped.push(linkUrl);

    try {
      const res = await fetch(linkUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
      if (res.ok) {
        const html = await res.text();
        const extracted = extractPageData(html, linkUrl);
        allTables.push(...extracted.tables);
        textChunks.push(`=== PAGE: ${linkUrl} ===\n${extracted.text.slice(0, 8000)}`);
      }
    } catch (err: any) {
      console.warn(`Error crawling ${linkUrl}:`, err?.message);
    }
  });

  await Promise.allSettled(crawlTasks);

  const combined = `=== SCRAPED TABLES ===\n${allTables.join("\n\n")}\n\n=== SCRAPED WEBPAGES ===\n${textChunks.join("\n\n")}`;

  return {
    content: combined,
    warning: antiBotDetected ? "Cloudflare / Anti-bot verification detected on site." : undefined,
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

    const apiKeys = apiKeysRaw.split(",").map((k) => k.trim()).filter(Boolean);
    if (apiKeys.length === 0) {
      return NextResponse.json({ error: "At least one valid Google Gemini API key must be provided." }, { status: 400 });
    }

    // Step A: Parse uploaded Excel file buffer
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: "buffer" });

    // FIX 1: Target the actual DATA sheet (e.g. 'Data', 'Records', or longest data sheet) instead of sheet 0 ('What_Was_Done')
    let targetSheetName = workbook.SheetNames[0];
    let maxRowCount = 0;
    for (const name of workbook.SheetNames) {
      const s = workbook.Sheets[name];
      const rows = XLSX.utils.sheet_to_json(s, { defval: "" });
      if (rows.length > maxRowCount) {
        maxRowCount = rows.length;
        targetSheetName = name;
      }
    }

    const worksheet = workbook.Sheets[targetSheetName];
    const rawExcelRows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (rawExcelRows.length === 0) {
      return NextResponse.json({ error: "The uploaded Excel sheet contains no records." }, { status: 400 });
    }

    const originalColumns = Object.keys(rawExcelRows[0] || {});

    // FIX 2: Harvest in-sheet web links (like course_detail_web_link, course_fee_web_link)
    const inSheetUrls: string[] = [];
    rawExcelRows.forEach((row) => {
      Object.entries(row).forEach(([colKey, val]) => {
        const strVal = String(val).trim();
        if (strVal.startsWith("http://") || strVal.startsWith("https://")) {
          if (!inSheetUrls.includes(strVal)) {
            inSheetUrls.push(strVal);
          }
        }
      });
    });

    // Step B: Crawl targeted university pages
    let webText = "";
    let scraperWarning: string | undefined = undefined;
    let pagesScraped: string[] = [];

    if (rawDirectContent.trim().length > 30) {
      webText = `=== USER DIRECT PASTED CONTENT ===\n${rawDirectContent.trim()}`;
      pagesScraped = ["Direct User Webpage Text / HTML (Bypassed Scraper)"];
    } else {
      const crawlResult = await crawlUniversityPages(url.trim(), inSheetUrls, 25);
      webText = crawlResult.content;
      scraperWarning = crawlResult.warning;
      pagesScraped = crawlResult.pagesScraped;
    }

    // Step C: Build AI Reconciliation Prompt
    const systemInstruction = `You are a world-class academic database administrator and university registrar AI.
You are reconciling the main course catalog sheet ('${targetSheetName}') against live university webpage data.
CRITICAL RULES:
1. AUTO-FILL: Fill in all empty fields, 'tbc' placeholders, or missing details (e.g., overview, structure, career_prospects, fee_per_year, entry_requirements) using the scraped course/program details.
2. UPDATE: If course names, fees, duration, or titles have updated for 2026, overwrite them with current values.
3. FUZZY MATCH: Match rows by course_name, specialization, campus, or course_detail_web_link.
4. EXACT COLUMNS: Maintain exact column names: ${JSON.stringify(originalColumns)}.`;

    const prompt = `
${customPrompt.trim() ? `TEACHER'S INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
SHEET NAME BEING UPDATED: "${targetSheetName}"
ORIGINAL COLUMNS: ${JSON.stringify(originalColumns)}
TOTAL RECORDS IN SHEET: ${rawExcelRows.length}

EXCEL RECORDS TO RECONCILE (WITH ROW INDEX):
${JSON.stringify(
  rawExcelRows.slice(0, 450).map((r, idx) => ({ _row_index: idx, ...r })),
  null,
  1
)}

LIVE UNIVERSITY WEBPAGE DATA (INCLUDING SPECIFIC PROGRAM PAGES & TABLES):
${webText.slice(0, 85000)}

TASK:
- For every row where fields were empty, 'tbc', or outdated, extract the correct information from the scraped pages and populate/update it.
- Output JSON specifying row_index, identifier, updated_data, and changed_columns.

JSON SCHEMA:
{
  "updated_rows": [
    {
      "row_index": 0,
      "identifier": "Advanced Diploma in English",
      "updated_data": { "overview ": "Comprehensive overview text...", "fee_per_year": 45000 },
      "changed_columns": ["overview ", "fee_per_year"],
      "reason": "Populated overview and updated fee from official program page"
    }
  ],
  "new_rows": [],
  "summary": "Summary of populated and updated fields."
}
`;

    // Step D: Call Gemini with failover
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
        throw new Error("Unable to parse AI response into JSON.");
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

    // Process new rows
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

    // Step F: Build Styled Excel Workbook preserving all original sheets
    const outputWb = new ExcelJS.Workbook();

    // Preserve any existing notes sheet (e.g. 'What_Was_Done')
    for (const sName of workbook.SheetNames) {
      if (sName !== targetSheetName) {
        const origOtherSheet = workbook.Sheets[sName];
        const otherRows: any[] = XLSX.utils.sheet_to_json(origOtherSheet, { header: 1 });
        const newOtherWs = outputWb.addWorksheet(sName);
        otherRows.forEach((r) => newOtherWs.addRow(r));
      }
    }

    // Add main updated Data worksheet
    const ws = outputWb.addWorksheet(targetSheetName);
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

    // Add Summary Sheet
    const summaryWs = outputWb.addWorksheet("Update Summary");
    summaryWs.addRow(["RecordSync Execution Summary"]);
    summaryWs.addRow(["Sheet Updated", targetSheetName]);
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

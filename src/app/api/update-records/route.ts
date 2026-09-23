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

/**
 * Clean HTML and extract text + structured tables
 */
function extractPageData(htmlText: string, sourceUrl: string): { tables: string[]; text: string; subLinks: string[] } {
  const $ = cheerio.load(htmlText);

  // Extract internal sublinks before removing navigation elements
  const subLinks: string[] = [];
  try {
    const parsedOrigin = new URL(sourceUrl).origin;
    $("a[href]").each((_, el) => {
      const href = $(el).attr("href");
      if (href) {
        try {
          const resolved = new URL(href, sourceUrl).toString();
          // Filter to same domain and relevant course/academic subpages
          if (
            resolved.startsWith(parsedOrigin) &&
            !resolved.includes("#") &&
            !resolved.match(/\.(pdf|jpg|jpeg|png|zip|doc|docx|mp4|svg)$/i)
          ) {
            // Priority keywords in URL
            const isRelevant = /course|catalog|syllabus|schedule|curriculum|department|faculty|degree|program|result|timetable|grade/i.test(resolved);
            if (isRelevant && !subLinks.includes(resolved) && resolved !== sourceUrl) {
              subLinks.push(resolved);
            }
          }
        } catch {
          // ignore malformed URLs
        }
      }
    });
  } catch {
    // ignore
  }

  // Remove clutter
  $("script, style, noscript, nav, footer, iframe, header, svg").remove();

  // Extract structured tables
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
 * DEEP CRAWLER: Crawls root page + internal subpages/tabs where course data is hidden
 */
async function deepCrawlUniversitySite(rootUrl: string, maxSubpages = 6): Promise<string> {
  const visited = new Set<string>();
  const allTables: string[] = [];
  const textChunks: string[] = [];

  // Step 1: Fetch root page
  try {
    const res = await fetch(rootUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status} ${res.statusText}`);
    }
    const html = await res.text();
    visited.add(rootUrl);

    const { tables, text, subLinks } = extractPageData(html, rootUrl);
    allTables.push(...tables);
    textChunks.push(`=== ROOT PAGE (${rootUrl}) ===\n${text.slice(0, 15000)}`);

    // Step 2: Concurrently crawl inside pages (up to maxSubpages)
    const targetsToCrawl = subLinks.slice(0, maxSubpages);
    if (targetsToCrawl.length > 0) {
      const crawlPromises = targetsToCrawl.map(async (subUrl) => {
        if (visited.has(subUrl)) return;
        visited.add(subUrl);
        try {
          const subRes = await fetch(subUrl, { headers: STANDARD_HEADERS, cache: "no-store" });
          if (subRes.ok) {
            const subHtml = await subRes.text();
            const extracted = extractPageData(subHtml, subUrl);
            allTables.push(...extracted.tables);
            textChunks.push(`=== SUBPAGE (${subUrl}) ===\n${extracted.text.slice(0, 10000)}`);
          }
        } catch (e: any) {
          console.warn(`Failed crawling subpage ${subUrl}:`, e?.message);
        }
      });

      await Promise.allSettled(crawlPromises);
    }
  } catch (err: any) {
    throw new Error(`Failed to crawl university URL ${rootUrl}: ${err?.message}`);
  }

  const combinedTables = allTables.join("\n\n");
  const combinedText = textChunks.join("\n\n");

  return `=== DEEP SCRAPED TABLES ACROSS ALL PAGES ===\n${combinedTables}\n\n=== DEEP SCRAPED CONTENT (ROOT + SUBPAGES) ===\n${combinedText}`;
}

/**
 * Multi-API Key & Model Fallback Call
 */
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
        console.warn(
          `Key #${keyIdx + 1} Model ${currentModel} error: ${err?.message || err}. Trying next fallback...`
        );
        lastError = err;
      }
    }
  }

  throw new Error(
    `All provided Gemini API keys & model fallbacks failed. Last error: ${
      lastError?.message || "Unknown error"
    }`
  );
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const url = formData.get("url") as string | null;
    const apiKeysRaw = formData.get("apiKeys") as string | null;
    const selectedModel = (formData.get("model") as string | null) || "gemini-2.5-flash";
    const customPrompt = (formData.get("customPrompt") as string | null) || "";

    if (!file || !url || !apiKeysRaw) {
      return NextResponse.json(
        { error: "Missing required fields: file, url, or apiKeys." },
        { status: 400 }
      );
    }

    const apiKeys = apiKeysRaw.split(",").map((k) => k.trim()).filter(Boolean);
    if (apiKeys.length === 0) {
      return NextResponse.json(
        { error: "At least one valid Google Gemini API key must be provided." },
        { status: 400 }
      );
    }

    // Step A: Parse uploaded Excel file
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const excelRows: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (excelRows.length === 0) {
      return NextResponse.json(
        { error: "The uploaded Excel sheet is empty." },
        { status: 400 }
      );
    }

    const columns = Object.keys(excelRows[0] || {});

    // Step B: Deep Crawl Root + Subpages (extracts data hidden in tabs/sublinks)
    const webText = await deepCrawlUniversitySite(url, 7);

    // Step C: AI Reconciliation Prompt
    const systemInstruction = `You are an expert university registrar AI assistant and academic database updater.
Your goal is to thoroughly reconcile, update, and autofill course/student/faculty records from legacy Excel sheets with live university webpage data gathered from the root portal and all linked subpages.
CRITICAL MANDATES:
1. UPDATE OUTDATED INFORMATION: If a course title, instructor, credit hour, code, grade, or status has changed or has a 2026 update anywhere across the scraped pages, UPDATE IT.
2. AUTOFILL MISSING / BLANK DATA: For every cell that is blank (""), null, or missing, find the matching course/record in the scraped text or tables and fill it in.
3. FUZZY & FLEXIBLE MATCHING: University websites often format course codes with or without hyphens (e.g., 'CS101', 'CS 101', 'CS-101', 'COMP 101'). Match them intelligently!
4. NEW RECORDS: Add new 2026 courses or student records under 'new_rows'.
5. TEACHER'S CUSTOM INSTRUCTIONS: Obey any specific user prompt strictly.
6. PRESERVE COLUMNS: Retain exact column headers: ${JSON.stringify(columns)}.`;

    const teacherInstructionSection = customPrompt.trim()
      ? `TEACHER'S UNIQUE INSTRUCTION (HIGHEST PRIORITY):\n"${customPrompt.trim()}"\n`
      : "";

    const prompt = `
${teacherInstructionSection}
ORIGINAL EXCEL HEADERS: ${JSON.stringify(columns)}
TOTAL ROWS IN EXCEL SHEET: ${excelRows.length}

EXCEL ROWS TO RECONCILE & UPDATE (JSON):
${JSON.stringify(excelRows.slice(0, 450), null, 1)}

DEEP-SCRAPED UNIVERSITY WEBPAGE DATA (INCLUDING SUBPAGES & TABLES):
${webText.slice(0, 75000)}

INSTRUCTIONS:
1. Search all scraped tables and subpage texts to find corresponding entries for each row.
2. Update any outdated values (e.g. 2026 credits, codes, instructors, room numbers, results).
3. Fill any empty or missing fields with data found on the portal.
4. If a row is modified or populated, list the row index in 'updated_rows', include the updated row data, and list which column names changed.

REQUIRED JSON FORMAT:
{
  "updated_rows": [
    {
      "row_index": 0,
      "updated_data": { "CourseCode": "CS101", "CourseTitle": "Introduction to AI", "Credits": "3" },
      "changed_columns": ["CourseTitle", "Credits"],
      "reason": "Autofilled title and updated credits from subpage catalog table"
    }
  ],
  "new_rows": [
    {
      "data": { "CourseCode": "CS499", "CourseTitle": "Senior Capstone Project" },
      "reason": "Newly introduced 2026 course found in department curriculum subpage"
    }
  ],
  "summary": "Detailed summary of updated and autofilled courses found across university pages."
}
`;

    // Step D: Execute AI reasoning
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
      const jsonMatch = rawAiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        aiResult = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error("Failed to parse AI response into JSON format.");
      }
    }

    // Step E: Create Workbook with Openpyxl / ExcelJS Styling
    const outputWb = new ExcelJS.Workbook();
    const ws = outputWb.addWorksheet("2026 Updated Records");

    ws.columns = columns.map((col) => ({ header: col, key: col, width: 22 }));

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

    const updatedExcelRows = [...excelRows];
    const modifiedCellSet = new Set<string>();

    const updatedRowsList = aiResult.updated_rows || [];
    for (const item of updatedRowsList) {
      const rIdx = item.row_index;
      const uData = item.updated_data || {};
      const cCols = item.changed_columns || [];

      if (rIdx !== undefined && rIdx >= 0 && rIdx < updatedExcelRows.length) {
        for (const colKey of columns) {
          if (uData[colKey] !== undefined && String(uData[colKey]) !== String(updatedExcelRows[rIdx][colKey])) {
            updatedExcelRows[rIdx][colKey] = uData[colKey];
            modifiedCellSet.add(`${rIdx}_${colKey}`);
          }
        }
        for (const colKey of cCols) {
          modifiedCellSet.add(`${rIdx}_${colKey}`);
        }
      }
    }

    const startNewRowIdx = updatedExcelRows.length;
    const newRowsList = aiResult.new_rows || [];
    for (let i = 0; i < newRowsList.length; i++) {
      const newDataObj = newRowsList[i].data || {};
      const targetIdx = startNewRowIdx + i;
      updatedExcelRows.push(newDataObj);
      for (const colKey of columns) {
        modifiedCellSet.add(`${targetIdx}_${colKey}`);
      }
    }

    const yellowFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };

    updatedExcelRows.forEach((rowObj, rIdx) => {
      const addedRow = ws.addRow(rowObj);
      columns.forEach((colKey, cIdx) => {
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

    const summaryWs = outputWb.addWorksheet("Update Summary");
    summaryWs.addRow(["RecordSync Execution Summary"]);
    summaryWs.addRow(["Total Courses/Rows Processed", excelRows.length]);
    summaryWs.addRow(["Autofilled & Updated Rows", updatedRowsList.length]);
    summaryWs.addRow(["Newly Discovered Courses/Rows", newRowsList.length]);
    summaryWs.addRow(["AI Model Used", modelUsed]);
    summaryWs.addRow(["API Key Account Used", keyUsedIndex]);
    if (customPrompt.trim()) {
      summaryWs.addRow(["Teacher Custom Instructions", customPrompt.trim()]);
    }
    summaryWs.addRow(["AI Summary Notes", aiResult.summary || "All records matched and updated."]);

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
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json(
      { error: err?.message || "Internal server error occurred." },
      { status: 500 }
    );
  }
}

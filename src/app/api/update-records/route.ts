import { NextRequest, NextResponse } from "next/server";
import * as cheerio from "cheerio";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { GoogleGenAI } from "@google/genai";

// Vercel Serverless maximum execution time configuration (60 seconds)
export const maxDuration = 60;
export const dynamic = "force-dynamic";

// Fallback Model Chain: If model hits rate limit/quota, automatically switch to backup model
const FALLBACK_MODEL_CHAIN = [
  "gemini-2.5-flash",
  "gemini-3.5-flash-lite",
  "gemini-3.1-flash-lite",
  "gemini-3.5-flash",
  "gemini-3.7-flash",
];

/**
 * 1. Web Scraping Layer: Extracts HTML text and tables from University links
 */
async function scrapeUniversityWebpage(url: string): Promise<string> {
  const headers = {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    Accept:
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
  };

  const response = await fetch(url, { headers, cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Failed to fetch university URL. Status: ${response.status} ${response.statusText}`);
  }

  const htmlText = await response.text();
  const $ = cheerio.load(htmlText);

  // Remove scripts, inline styles, metadata to clean context
  $("script, style, noscript, nav, footer, iframe, header, svg").remove();

  // Extract structured tables dynamically
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
        if (cols.length > 0) {
          rows.push(cols.join(" | "));
        }
      });
    if (rows.length > 0) {
      tables.push(`--- UNIVERSITY WEBPAGE TABLE ${idx + 1} ---\n` + rows.join("\n"));
    }
  });

  const bodyText = $("body").text().replace(/\s+/g, " ").trim();

  return `=== EXTRACTED WEBPAGE TABLES ===\n${tables.join(
    "\n\n"
  )}\n\n=== MAIN WEBPAGE TEXT ===\n${bodyText.slice(0, 35000)}`;
}

/**
 * 2. Multi-API Key & Dual Model Fallback Gemini Call
 * Loops through API Keys AND Model Chain to ensure zero quota failures!
 */
async function callGeminiWithFullFailover(
  apiKeys: string[],
  initialModel: string,
  prompt: string,
  systemInstruction: string
): Promise<{ text: string; modelUsed: string; keyUsedIndex: number }> {
  let lastError: Error | null = null;

  // Build model try order starting with user chosen model
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
          `Key #${keyIdx + 1} with Model '${currentModel}' failed: ${err?.message || err}. Trying next model/key...`
        );
        lastError = err;
        // Continue loop to try next model or next API key
      }
    }
  }

  throw new Error(
    `All provided Gemini API keys and model fallbacks failed or reached quota limits. Last error: ${
      lastError?.message || "Unknown API error"
    }`
  );
}

/**
 * Main API Handler
 */
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const url = formData.get("url") as string | null;
    const apiKeysRaw = formData.get("apiKeys") as string | null;
    const selectedModel = (formData.get("model") as string | null) || "gemini-2.5-flash";

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

    // Step A: Parse uploaded Excel file buffer with XLSX
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

    // Step B: Scrape University Webpage
    const webText = await scrapeUniversityWebpage(url);

    // Step C: AI Reconciliation Prompt
    const systemInstruction = `You are an expert university registrar AI assistant. 
Your task is to update legacy Excel student/course/department records with fresh 2026 webpage data.
Keep original column headers exactly as provided. Identify changed grades, updated statuses, modified course codes, or new student additions.
Return strictly formatted JSON matching the required schema.`;

    const prompt = `
ORIGINAL EXCEL COLUMNS: ${JSON.stringify(columns)}
LEGACY EXCEL DATA (JSON - First 150 rows):
${JSON.stringify(excelRows.slice(0, 150), null, 2)}

FRESH 2026 SCRAPED UNIVERSITY WEBPAGE CONTENT:
${webText}

INSTRUCTIONS:
1. Match entity records (students, courses, faculty, IDs) from the Excel sheet with the scraped 2026 web data.
2. For any row with 2026 updates, provide the modified row and specify which column names changed.
3. If new 2026 records exist on the webpage that are not in the original Excel sheet, add them under 'new_rows'.
4. Do NOT change column header names. Use exact column keys: ${JSON.stringify(columns)}.

REQUIRED JSON RESPONSE STRUCTURE:
{
  "updated_rows": [
    {
      "row_index": 0,
      "updated_data": { "Col1": "Val", "Col2": "UpdatedVal" },
      "changed_columns": ["Col2"],
      "reason": "Updated 2026 grade based on university result table"
    }
  ],
  "new_rows": [
    {
      "data": { "Col1": "Val", "Col2": "Val" },
      "reason": "Newly enrolled 2026 record"
    }
  ],
  "summary": "2-sentence executive summary of reconciled records."
}
`;

    // Step D: Call Gemini with Dual Failover (API Keys + Auto Model Switch)
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
      // Regex extraction fallback
      const jsonMatch = rawAiResponse.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        aiResult = JSON.parse(jsonMatch[0]);
      } else {
        throw new Error("Failed to parse AI output into valid JSON.");
      }
    }

    // Step E: Create Styled Excel Workbook with ExcelJS
    const outputWb = new ExcelJS.Workbook();
    const ws = outputWb.addWorksheet("2026 Updated Records");

    // Add Headers
    ws.columns = columns.map((col) => ({ header: col, key: col, width: 20 }));

    // Apply header styling
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

    // Populate data & track cell modifications
    const updatedExcelRows = [...excelRows];
    const modifiedCellSet = new Set<string>();

    // Process AI updates
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

    // Process newly added rows
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

    // Soft Yellow Fill (#FFF2CC) for changed cells
    const yellowFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };

    // Add rows to worksheet & format
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

    // Auto-fit column widths
    ws.columns.forEach((col) => {
      let maxLen = 12;
      col.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 4, 40);
    });

    // Add Summary Worksheet
    const summaryWs = outputWb.addWorksheet("Update Summary");
    summaryWs.addRow(["RecordSync Execution Summary"]);
    summaryWs.addRow(["Original Rows Count", excelRows.length]);
    summaryWs.addRow(["Updated Existing Rows", updatedRowsList.length]);
    summaryWs.addRow(["Newly Added Rows", newRowsList.length]);
    summaryWs.addRow(["Model Used (With Auto-Switch)", modelUsed]);
    summaryWs.addRow(["API Key Account # Used", keyUsedIndex]);
    summaryWs.addRow(["AI Reconciliation Notes", aiResult.summary || "No notes."]);

    // Export binary buffer
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

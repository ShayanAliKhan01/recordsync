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

async function scrapeUniversityWebpage(url: string): Promise<string> {
  const headers = {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    Accept:
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
  };

  const response = await fetch(url, { headers, cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Failed to fetch university URL. Status: ${response.status} ${response.statusText}`);
  }

  const htmlText = await response.text();
  const $ = cheerio.load(htmlText);

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
        if (cols.length > 0) {
          rows.push(cols.join(" | "));
        }
      });
    if (rows.length > 0) {
      tables.push(`--- UNIVERSITY TABLE ${idx + 1} ---\n` + rows.join("\n"));
    }
  });

  const bodyText = $("body").text().replace(/\s+/g, " ").trim();
  return `=== EXTRACTED WEBPAGE TABLES ===\n${tables.join(
    "\n\n"
  )}\n\n=== MAIN WEBPAGE TEXT ===\n${bodyText.slice(0, 45000)}`;
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
        console.warn(
          `Key #${keyIdx + 1} Model ${currentModel} failed: ${err?.message || err}. Trying next fallback...`
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

    // Step B: Scrape University Webpage
    const webText = await scrapeUniversityWebpage(url);

    // Step C: Build Intelligent Prompt (Handling Batching, Auto-fill, and Teacher's Custom Instructions)
    const systemInstruction = `You are an expert university registrar AI assistant and academic database administrator.
Your task is to reconcile, update, and autofill course/student/faculty records from legacy Excel sheets with live university webpage data.
CRITICAL RULES:
1. AUTO-FILL: If an Excel row has empty, missing, or 'N/A' fields (e.g. course title, credit hours, instructor, department, room, prerequisite, 2026 grade), inspect the scraped web page data and AUTO-FILL those blank fields with correct information.
2. UPDATE: If existing fields have changed in 2026 (e.g. course code changed from CS-101 to CS-1101, or updated grade/status), update the field.
3. TEACHER'S CUSTOM INSTRUCTION: If the teacher gave custom instructions, obey them strictly.
4. NEW RECORDS: If the webpage contains new courses or student records not present in the Excel sheet, append them under new_rows.
5. PRESERVE COLUMNS: Keep exact column keys: ${JSON.stringify(columns)}.`;

    const teacherInstructionSection = customPrompt.trim()
      ? `TEACHER'S UNIQUE CUSTOM INSTRUCTIONS (PRIORITY):\n"${customPrompt.trim()}"\n`
      : "";

    // Support processing large sheets (up to 400 course rows) by passing compact row data
    const prompt = `
${teacherInstructionSection}
ORIGINAL EXCEL COLUMNS: ${JSON.stringify(columns)}
TOTAL ROWS IN EXCEL: ${excelRows.length}

EXCEL ROWS TO RECONCILE (JSON):
${JSON.stringify(excelRows.slice(0, 450), null, 1)}

LIVE SCRAPED UNIVERSITY WEBPAGE CONTENT:
${webText}

INSTRUCTIONS:
1. Compare each row (course, student, faculty, or code) against the live university webpage.
2. AUTO-FILL BLANKS: For any row where one or more columns are blank, empty (""), null, or missing, look up that course/entity on the webpage and fill in the missing data.
3. Mark updated/autofilled column names in 'changed_columns'.
4. If there are new courses or records on the webpage not in the Excel, include them in 'new_rows'.

REQUIRED JSON RESPONSE STRUCTURE:
{
  "updated_rows": [
    {
      "row_index": 0,
      "updated_data": { "CourseCode": "CS101", "CourseTitle": "Intro to AI", "Credits": "3" },
      "changed_columns": ["CourseTitle", "Credits"],
      "reason": "Autofilled missing Course Title and updated credits from university portal"
    }
  ],
  "new_rows": [
    {
      "data": { "CourseCode": "CS401", "CourseTitle": "Advanced Machine Learning" },
      "reason": "New 2026 course catalog entry"
    }
  ],
  "summary": "2-sentence executive summary of autofilled and updated courses."
}
`;

    // Step D: Execute AI reasoning with dual failover
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

    // Apply updates and autofills
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

    // Process new 2026 courses / rows
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

    // Soft yellow fill (#FFF2CC)
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

    // Summary sheet
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

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

function extractCleanPageText(htmlText: string): string {
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
        if (cols.length > 0) rows.push(cols.join(" | "));
      });
    if (rows.length > 0) tables.push(rows.join("\n"));
  });

  const bodyText = $("body").text().replace(/\s+/g, " ").trim();
  return (tables.length > 0 ? `[TABLES]\n${tables.join("\n\n")}\n\n` : "") + bodyText;
}

async function fetchUrlContent(url: string): Promise<string> {
  try {
    const res = await fetch(url, { headers: STANDARD_HEADERS, cache: "no-store" });
    if (!res.ok) return "";
    const html = await res.text();
    return extractCleanPageText(html);
  } catch {
    return "";
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

    // Find main Data worksheet
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
      return NextResponse.json({ error: "No records found in the Excel sheet." }, { status: 400 });
    }

    const originalColumns = Object.keys(rawExcelRows[0] || {});

    // Step B: Crawl specific course detail URLs
    const pagesScraped: string[] = [];
    const scrapedContextPerUrl = new Map<string, string>();

    // 1. Root URL
    if (url.trim().startsWith("http")) {
      pagesScraped.push(url.trim());
      const rootText = await fetchUrlContent(url.trim());
      scrapedContextPerUrl.set(url.trim(), rootText.slice(0, 8000));
    }

    // 2. Direct in-sheet course links
    const targetInSheetUrls: string[] = [];
    rawExcelRows.forEach((r) => {
      for (const val of Object.values(r)) {
        const str = String(val).trim();
        if (str.startsWith("http://") || str.startsWith("https://")) {
          if (!targetInSheetUrls.includes(str)) targetInSheetUrls.push(str);
        }
      }
    });

    const linksToFetch = targetInSheetUrls.slice(0, 30);
    const fetchPromises = linksToFetch.map(async (link) => {
      pagesScraped.push(link);
      const content = await fetchUrlContent(link);
      if (content.length > 50) {
        scrapedContextPerUrl.set(link, content.slice(0, 6000));
      }
    });

    await Promise.allSettled(fetchPromises);

    // If user pasted text directly
    if (rawDirectContent.trim().length > 30) {
      pagesScraped.push("Direct User Webpage Text / HTML");
      scrapedContextPerUrl.set("direct_paste", rawDirectContent.slice(0, 30000));
    }

    // Step C: Build Concise Micro-Rows for Gemini
    // Filter down only to rows that have missing fields or correspond to fetched URLs
    const candidateRows = rawExcelRows.map((r, idx) => {
      // Find course name and link
      let courseName = "";
      let detailLink = "";
      for (const [k, v] of Object.entries(r)) {
        const norm = normalizeKey(k);
        if (norm.includes("coursename") || norm.includes("program") || norm.includes("coursetitle")) {
          if (!courseName && v) courseName = String(v);
        }
        if (norm.includes("detail") && norm.includes("link")) {
          if (!detailLink && v) detailLink = String(v);
        }
      }

      // Check for empty or placeholder columns
      const blankCols: string[] = [];
      originalColumns.forEach((c) => {
        const val = String(r[c] || "").trim().toLowerCase();
        if (!val || val === "tbc" || val === "n/a" || val === "null" || val === "none") {
          blankCols.push(c);
        }
      });

      return {
        _row_index: idx,
        course_name: courseName || String(r[originalColumns[0]] || ""),
        course_link: detailLink,
        missing_fields: blankCols,
      };
    });

    // Take candidates that have missing fields
    const candidatesToUpdate = candidateRows.filter((c) => c.missing_fields.length > 0).slice(0, 50);

    // Aggregate relevant scraped pages
    let aggregatedWebContext = "";
    for (const [linkUrl, textSnippet] of scrapedContextPerUrl.entries()) {
      aggregatedWebContext += `\n--- SOURCE: ${linkUrl} ---\n${textSnippet}\n`;
    }

    const systemInstruction = `You are an automated academic data extraction AI.
Your job is to populate missing or outdated university course details (overview, structure, fee, requirements, credit hours) from live web text into the course rows.
RULES:
1. For each course row in the input list, search the scraped web text for its course name or course link.
2. If found, extract the exact description/overview, eligibility criteria, credit hours, or fee.
3. You MUST populate the missing fields requested for each matched course.
4. Output valid JSON strictly following the schema.`;

    const prompt = `
${customPrompt.trim() ? `SPECIAL INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
TARGET COURSES THAT NEED FIELDS POPULATED (${candidatesToUpdate.length} Courses):
${JSON.stringify(candidatesToUpdate, null, 2)}

LIVE SCRAPED UNIVERSITY WEBPAGES & TABLES:
${aggregatedWebContext.slice(0, 75000)}

INSTRUCTIONS:
- For every course that you can match in the scraped text, provide its _row_index and an 'updated_data' dictionary containing the populated column values.
- Do NOT return an empty list if data exists in the scraped pages!

REQUIRED JSON FORMAT:
{
  "updated_rows": [
    {
      "row_index": 0,
      "updated_data": {
        "overview ": "Advanced Diploma in English provides foundational knowledge in linguistics and literature...",
        "fee_per_year": 43262
      },
      "changed_columns": ["overview ", "fee_per_year"],
      "reason": "Extracted from program course page"
    }
  ],
  "summary": "Populated missing course overviews, structure, and details from NUML program pages."
}
`;

    // Step D: Execute Gemini
    let aiResult: any = { updated_rows: [], summary: "" };
    let modelUsed = selectedModel;
    let keyUsedIndex = 1;

    try {
      const response = await callGemini(apiKeys, selectedModel, prompt, systemInstruction);
      modelUsed = response.modelUsed;
      keyUsedIndex = response.keyUsedIndex;
      const jsonMatch = response.text.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        aiResult = JSON.parse(jsonMatch[0]);
      }
    } catch (e: any) {
      console.warn("Gemini call error:", e?.message);
    }

    // Step E: Apply updates to Excel
    const normalizedToOriginalCol: { [norm: string]: string } = {};
    originalColumns.forEach((c) => {
      normalizedToOriginalCol[normalizeKey(c)] = c;
    });

    const updatedExcelRows = [...rawExcelRows];
    const modifiedCellSet = new Set<string>();

    const updatedRowsList = aiResult.updated_rows || [];
    for (const item of updatedRowsList) {
      const rIdx = item.row_index;
      if (rIdx !== undefined && rIdx >= 0 && rIdx < updatedExcelRows.length) {
        const uData = item.updated_data || {};
        const cCols = item.changed_columns || [];

        for (const [key, val] of Object.entries(uData)) {
          const normKey = normalizeKey(key);
          const targetCol = normalizedToOriginalCol[normKey];
          if (targetCol && val !== undefined && val !== null) {
            updatedExcelRows[rIdx][targetCol] = val;
            modifiedCellSet.add(`${rIdx}_${targetCol}`);
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

    // Step F: Build Styled Excel Workbook
    const outputWb = new ExcelJS.Workbook();

    // Preserve non-target sheets
    for (const sName of workbook.SheetNames) {
      if (sName !== targetSheetName) {
        const origOtherSheet = workbook.Sheets[sName];
        const otherRows: any[] = XLSX.utils.sheet_to_json(origOtherSheet, { header: 1 });
        const newOtherWs = outputWb.addWorksheet(sName);
        otherRows.forEach((r) => newOtherWs.addRow(r));
      }
    }

    // Main updated Data worksheet
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

    // Summary Sheet
    const summaryWs = outputWb.addWorksheet("Update Summary");
    summaryWs.addRow(["RecordSync Execution Summary"]);
    summaryWs.addRow(["Sheet Updated", targetSheetName]);
    summaryWs.addRow(["Total Courses in Sheet", rawExcelRows.length]);
    summaryWs.addRow(["Autofilled & Updated Rows", updatedRowsList.length]);
    summaryWs.addRow(["AI Model Used", modelUsed]);
    summaryWs.addRow(["API Key Account Used", keyUsedIndex]);
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
        "X-New-Count": "0",
        "X-Model-Used": modelUsed,
        "X-Key-Used": String(keyUsedIndex),
        "X-Summary-Notes": encodeURIComponent(aiResult.summary || "Courses updated."),
        "X-Pages-Scraped": encodeURIComponent(JSON.stringify(pagesScraped)),
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err?.message || "Internal server error occurred." }, { status: 500 });
  }
}

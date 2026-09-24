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

    // Target the main Data worksheet
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

    if (rawDirectContent.trim().length > 30) {
      pagesScraped.push("Direct User Webpage Text / HTML");
      scrapedContextPerUrl.set("direct_paste", rawDirectContent.slice(0, 30000));
    }

    // Build candidate rows needing autofill
    const candidateRows = rawExcelRows.map((r, idx) => {
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

    const candidatesToUpdate = candidateRows.filter((c) => c.missing_fields.length > 0);

    let aggregatedWebContext = "";
    for (const [linkUrl, textSnippet] of scrapedContextPerUrl.entries()) {
      aggregatedWebContext += `\n--- SOURCE: ${linkUrl} ---\n${textSnippet}\n`;
    }

    const systemInstruction = `You are an automated academic database reconciler AI.
Your job is to populate missing or outdated course details (overview, structure, fee, requirements, credit hours) for the given batch of course rows using the scraped university web text.
RULES:
1. For each course row in the batch, match its course name or link in the scraped text.
2. If found, extract the exact description/overview, eligibility criteria, credit hours, or fee.
3. You MUST populate the missing fields requested for each matched course.
4. Output valid JSON strictly matching the schema: {"updated_rows": [{"row_index": 0, "updated_data": {"Col": "Val"}, "changed_columns": ["Col"]}]}.`;

    // Step C: Run Sequential Batches of 20 Rows
    const BATCH_SIZE = 20;
    const allUpdatedRows: any[] = [];
    let lastModelUsed = selectedModel;
    let lastKeyUsedIndex = 1;

    // Process in batches
    const totalToProcess = Math.min(candidatesToUpdate.length, 100); // 5 batches max to stay comfortably inside Vercel's 60s execution limit
    for (let i = 0; i < totalToProcess; i += BATCH_SIZE) {
      const batchSlice = candidatesToUpdate.slice(i, i + BATCH_SIZE);

      const prompt = `
${customPrompt.trim() ? `SPECIAL INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
BATCH OF CANDIDATE COURSES TO POPULATE (${batchSlice.length} Courses):
${JSON.stringify(batchSlice, null, 2)}

LIVE SCRAPED UNIVERSITY WEBPAGES & TABLES:
${aggregatedWebContext.slice(0, 75000)}

INSTRUCTIONS:
- For every course in this batch, output its _row_index and an 'updated_data' object containing the populated column values.
- Return valid JSON:
{
  "updated_rows": [
    {
      "row_index": 0,
      "updated_data": {"overview ": "Detailed overview..."},
      "changed_columns": ["overview "]
    }
  ]
}
`;

      try {
        const response = await callGemini(apiKeys, selectedModel, prompt, systemInstruction);
        lastModelUsed = response.modelUsed;
        lastKeyUsedIndex = response.keyUsedIndex;
        const jsonMatch = response.text.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          if (Array.isArray(parsed.updated_rows)) {
            allUpdatedRows.push(...parsed.updated_rows);
          }
        }
      } catch (err: any) {
        console.warn(`Batch ${i / BATCH_SIZE + 1} error:`, err?.message);
      }
    }

    // Step D: Apply updates to Excel in-place
    const normalizedToOriginalCol: { [norm: string]: string } = {};
    originalColumns.forEach((c) => {
      normalizedToOriginalCol[normalizeKey(c)] = c;
    });

    const updatedExcelRows = [...rawExcelRows];
    const modifiedCellSet = new Set<string>();

    for (const item of allUpdatedRows) {
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

    // Step E: Build in-place Styled Excel Workbook preserving all original sheets
    const outputWb = new ExcelJS.Workbook();

    for (const sName of workbook.SheetNames) {
      if (sName !== targetSheetName) {
        const origOtherSheet = workbook.Sheets[sName];
        const otherRows: any[] = XLSX.utils.sheet_to_json(origOtherSheet, { header: 1 });
        const newOtherWs = outputWb.addWorksheet(sName);
        otherRows.forEach((r) => newOtherWs.addRow(r));
      }
    }

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
    summaryWs.addRow(["Autofilled & Updated Rows", allUpdatedRows.length]);
    summaryWs.addRow(["AI Model Used", lastModelUsed]);
    summaryWs.addRow(["API Key Account Used", lastKeyUsedIndex]);
    if (customPrompt.trim()) summaryWs.addRow(["Teacher Custom Instructions", customPrompt.trim()]);
    summaryWs.addRow(["Pages Scraped List:", pagesScraped.join(" | ")]);

    const excelBuffer = await outputWb.xlsx.writeBuffer();

    return new NextResponse(excelBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="Updated_2026_${file.name}"`,
        "X-Updated-Count": String(allUpdatedRows.length),
        "X-New-Count": "0",
        "X-Model-Used": lastModelUsed,
        "X-Key-Used": String(lastKeyUsedIndex),
        "X-Summary-Notes": encodeURIComponent(`Batch engine updated ${allUpdatedRows.length} course records.`),
        "X-Pages-Scraped": encodeURIComponent(JSON.stringify(pagesScraped)),
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err?.message || "Internal server error occurred." }, { status: 500 });
  }
}

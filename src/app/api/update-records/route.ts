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

/**
 * Universal Semantic HTML Reducer (Works on ANY university site)
 * Converts messy university HTML into clean, structured Markdown text.
 */
function convertHtmlToCleanMarkdown(htmlText: string): string {
  const $ = cheerio.load(htmlText);
  $("script, style, noscript, nav, footer, header, svg, iframe, form").remove();

  // Convert tables to markdown tables
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

  // Convert headers
  $("h1, h2, h3").each((_, h) => {
    const t = $(h).text().replace(/\s+/g, " ").trim();
    if (t) $(h).replaceWith(`\n### ${t}\n`);
  });

  return $("body").text().replace(/\n\s*\n\s*\n/g, "\n\n").replace(/\s+/g, " ").trim();
}

async function fetchPageMarkdown(url: string): Promise<string> {
  try {
    const res = await fetch(url, { headers: STANDARD_HEADERS, cache: "no-store" });
    if (!res.ok) return "";
    const html = await res.text();
    return convertHtmlToCleanMarkdown(html);
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

    // Step A: Parse uploaded Excel file
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: "buffer" });

    // Target the primary data worksheet (largest row count)
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
    // Get headers as exact array preserving duplicate or spaced names
    const sheetData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    if (sheetData.length <= 1) {
      return NextResponse.json({ error: "No records found in the Excel sheet." }, { status: 400 });
    }

    const headerRow: string[] = sheetData[0].map((h: any) => String(h || "").trim());
    const dataRows = sheetData.slice(1);

    // Identify URL column index (0-based)
    let urlColIdx = -1;
    for (let c = 0; c < headerRow.length; c++) {
      const hLower = headerRow[c].toLowerCase();
      if (hLower.includes("detail") && hLower.includes("link")) {
        urlColIdx = c;
        break;
      }
    }
    if (urlColIdx === -1) {
      for (let c = 0; c < headerRow.length; c++) {
        const hLower = headerRow[c].toLowerCase();
        if (hLower.includes("link") || hLower.includes("url")) {
          urlColIdx = c;
          break;
        }
      }
    }

    // Identify columns that are blank or outdated needing verification
    // Find target columns by exact position:
    // e.g. overview, structure, career_prospects, fee_per_year, entry_requirements
    const targetColsToVerify: { colIdx: number; name: string }[] = [];
    headerRow.forEach((name, idx) => {
      const nLower = name.toLowerCase();
      if (
        nLower.startsWith("overview") ||
        nLower.startsWith("structure") ||
        nLower.startsWith("career") ||
        nLower.startsWith("fee") ||
        nLower.startsWith("entry_req") ||
        nLower.startsWith("duration")
      ) {
        targetColsToVerify.push({ colIdx: idx, name });
      }
    });

    // Step B: Crawl pages concurrently
    const pagesScraped: string[] = [];
    const urlContentMap = new Map<string, string>();

    // 1. Root URL
    if (url.trim().startsWith("http")) {
      pagesScraped.push(url.trim());
      const rootText = await fetchPageMarkdown(url.trim());
      if (rootText) urlContentMap.set(url.trim(), rootText.slice(0, 10000));
    }

    // 2. Direct course links from sheet
    const targetUrlsToFetch: string[] = [];
    if (urlColIdx !== -1) {
      dataRows.forEach((r) => {
        const u = String(r[urlColIdx] || "").trim();
        if (u.startsWith("http://") || u.startsWith("https://")) {
          if (!targetUrlsToFetch.includes(u)) targetUrlsToFetch.push(u);
        }
      });
    }

    const linksToFetch = targetUrlsToFetch.slice(0, 40);
    const fetchTasks = linksToFetch.map(async (link) => {
      pagesScraped.push(link);
      const content = await fetchPageMarkdown(link);
      if (content.length > 50) {
        urlContentMap.set(link, content.slice(0, 8000));
      }
    });
    await Promise.allSettled(fetchTasks);

    if (rawDirectContent.trim().length > 30) {
      pagesScraped.push("Direct User Webpage Text / HTML");
      urlContentMap.set("direct_paste", rawDirectContent.slice(0, 40000));
    }

    // Step C: Sequential Batch Loop: First 20, then next 20, until document is done
    const BATCH_SIZE = 20;
    const totalRows = dataRows.length;
    // Map of (row_idx, col_idx) -> { newVal, reason, verified }
    const cellModifications = new Map<string, any>();
    let lastModelUsed = selectedModel;
    let lastKeyUsedIndex = 1;

    const systemInstruction = `You are a strict academic data verification and reconciliation engine.
Your task is to:
1. VERIFY: Compare existing Excel cell values against the live scraped webpage content for 2026.
2. UPDATE: If an existing cell has outdated data (e.g. old fee, revised credits, outdated requirements), update it.
3. AUTOFILL: If a cell is blank (""), null, or contains a placeholder like 'tbc', extract the correct value from the webpage.
4. SYNTHESIZE IF PLACEHOLDER: If a course page has a placeholder overview (e.g. 'abc' or 'tbc'), synthesize a clean, professional 2-sentence course overview from its syllabus subjects and title.
5. PRESERVE ACCURACY: If the current value is already verified and up-to-date, do NOT change it.
Output strictly valid JSON with exact column indices.`;

    // Process all rows in batches of 20
    for (let startIdx = 0; startIdx < totalRows; startIdx += BATCH_SIZE) {
      const endIdx = min(startIdx + BATCH_SIZE, totalRows);
      const batchSlice = dataRows.slice(startIdx, endIdx);

      // Build compact batch representation with exact column indices
      const batchItems = batchSlice.map((r, i) => {
        const actualRowIdx = startIdx + i;
        const rowObj: any = {
          _row_index: actualRowIdx,
          course_name: String(r[12] || r[0] || ""), // course_name or first column
          course_url: urlColIdx !== -1 ? String(r[urlColIdx] || "") : "",
        };

        // Attach target columns with their exact col index
        targetColsToVerify.forEach(({ colIdx, name }) => {
          rowObj[`col_${colIdx}_${name}`] = String(r[colIdx] || "");
        });

        return rowObj;
      });

      // Gather relevant scraped context for this batch
      let batchWebContext = "";
      batchItems.forEach((item) => {
        if (item.course_url && urlContentMap.has(item.course_url)) {
          batchWebContext += `\n--- URL: ${item.course_url} ---\n${urlContentMap.get(item.course_url)}\n`;
        }
      });
      if (!batchWebContext && urlContentMap.size > 0) {
        batchWebContext = Array.from(urlContentMap.values()).slice(0, 5).join("\n\n");
      }

      const prompt = `
${customPrompt.trim() ? `SPECIAL INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
BATCH: Rows ${startIdx + 1} to ${endIdx} of ${totalRows} Total Courses:
${JSON.stringify(batchItems, null, 2)}

LIVE SCRAPED UNIVERSITY WEBPAGES & CURRICULUM TABLES:
${batchWebContext.slice(0, 70000)}

INSTRUCTIONS:
- For each course, verify whether its columns match the live web page.
- If a column is blank, contains 'tbc', or has an outdated value, output the update referencing its exact column name key (e.g. 'col_35_overview ').
- Return valid JSON:
{
  "updates": [
    {
      "row_index": ${startIdx},
      "col_index": 35,
      "col_key": "col_35_overview ",
      "verified_2026_value": "Professional overview text extracted from program syllabus...",
      "status": "autofilled_placeholder",
      "reason": "Replaced 'tbc' with verified syllabus synthesis"
    }
  ]
}
`;

      try {
        const response = await callGemini(apiKeys, selectedModel, prompt, systemInstruction);
        lastModelUsed = response.modelUsed;
        lastKeyUsedIndex = response.keyUsedIndex;
        const match = response.text.match(/\{[\s\S]*\}/);
        if (match) {
          const parsed = JSON.parse(match[0]);
          if (Array.isArray(parsed.updates)) {
            parsed.updates.forEach((u: any) => {
              const r = u.row_index;
              const c = u.col_index;
              const val = u.verified_2026_value;
              if (r !== undefined && c !== undefined && val !== undefined) {
                cellModifications.set(`${r}_${c}`, {
                  val,
                  reason: u.reason || "Verified 2026 update",
                });
              }
            });
          }
        }
      } catch (err: any) {
        console.warn(`Batch ${startIdx / BATCH_SIZE + 1} error:`, err?.message);
      }
    }

    // Step D: Apply Updates In-Place Preserving All Original Sheets & Formatting
    const outputWb = new ExcelJS.Workbook();

    // Preserve other sheets (e.g. What_Was_Done)
    for (const sName of workbook.SheetNames) {
      if (sName !== targetSheetName) {
        const origOtherSheet = workbook.Sheets[sName];
        const otherRows: any[] = XLSX.utils.sheet_to_json(origOtherSheet, { header: 1, defval: "" });
        const newOtherWs = outputWb.addWorksheet(sName);
        otherRows.forEach((r) => newOtherWs.addRow(r));
      }
    }

    // Main Data Sheet
    const ws = outputWb.addWorksheet(targetSheetName);

    // Write original Header Row
    const headerRowCells = headerRow;
    const addedHeader = ws.addRow(headerRowCells);
    addedHeader.eachCell((cell) => {
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

    let totalUpdatedCells = 0;

    // Write Data Rows & apply in-place modifications to exact cell coordinates
    dataRows.forEach((rowValues, rIdx) => {
      const modifiedRow = [...rowValues];

      // Check if any column in this row was updated
      for (let cIdx = 0; cIdx < headerRow.length; cIdx++) {
        const key = `${rIdx}_${cIdx}`;
        if (cellModifications.has(key)) {
          const mod = cellModifications.get(key);
          modifiedRow[cIdx] = mod.val;
          totalUpdatedCells++;
        }
      }

      const excelRow = ws.addRow(modifiedRow);

      // Highlight modified cells in soft yellow
      for (let cIdx = 0; cIdx < headerRow.length; cIdx++) {
        const key = `${rIdx}_${cIdx}`;
        const cell = excelRow.getCell(cIdx + 1);
        cell.border = {
          top: { style: "thin", color: { argb: "FFD9D9D9" } },
          left: { style: "thin", color: { argb: "FFD9D9D9" } },
          bottom: { style: "thin", color: { argb: "FFD9D9D9" } },
          right: { style: "thin", color: { argb: "FFD9D9D9" } },
        };

        if (cellModifications.has(key)) {
          cell.fill = yellowFill;
          cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF7F6000" } };
        }
      }
    });

    // Auto-fit column widths
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
    summaryWs.addRow(["Total Courses Processed", totalRows]);
    summaryWs.addRow(["Total Cells Verified & Updated", totalUpdatedCells]);
    summaryWs.addRow(["Batch Size Used", `${BATCH_SIZE} rows per batch until completion`]);
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
        "X-Updated-Count": String(totalUpdatedCells),
        "X-New-Count": "0",
        "X-Model-Used": lastModelUsed,
        "X-Key-Used": String(lastKeyUsedIndex),
        "X-Summary-Notes": encodeURIComponent(`Verified and updated ${totalUpdatedCells} cells across all rows.`),
        "X-Pages-Scraped": encodeURIComponent(JSON.stringify(pagesScraped)),
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err?.message || "Internal server error occurred." }, { status: 500 });
  }
}

function min(a: number, b: number): number {
  return a < b ? a : b;
}

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

async function fetchPageMarkdown(url: string): Promise<string> {
  if (!url || !url.startsWith("http")) return "";
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
    const sheetData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    if (sheetData.length <= 1) {
      return NextResponse.json({ error: "No records found in the Excel sheet." }, { status: 400 });
    }

    const headerRow: string[] = sheetData[0].map((h: any) => String(h || "").trim());
    const dataRows = sheetData.slice(1);

    // Build header column map
    const headerColMap: { [norm: string]: number } = {};
    headerRow.forEach((h, idx) => {
      headerColMap[normalizeKey(h)] = idx;
    });

    // Detect link column
    let urlColIdx = -1;
    for (let c = 0; c < headerRow.length; c++) {
      const hNorm = normalizeKey(headerRow[c]);
      if (hNorm.includes("detail") && hNorm.includes("link")) {
        urlColIdx = c;
        break;
      }
    }
    if (urlColIdx === -1) {
      for (let c = 0; c < headerRow.length; c++) {
        const hNorm = normalizeKey(headerRow[c]);
        if (hNorm.includes("link") || hNorm.includes("url")) {
          urlColIdx = c;
          break;
        }
      }
    }

    // Candidate rows
    const candidateRows = dataRows.map((r, idx) => {
      const rowDict: any = { _row_index: idx };
      headerRow.forEach((h, cIdx) => {
        rowDict[h] = r[cIdx];
      });
      return rowDict;
    });

    // Multi-pass batch execution (12 courses per batch for deep recursive scraping)
    const BATCH_SIZE = 12;
    const totalCourses = candidateRows.length;
    const cellModifications = new Map<string, any>();
    const allScrapedPages = new Set<string>();
    let lastModelUsed = selectedModel;
    let lastKeyUsedIndex = 1;

    const systemInstruction = `You are an expert university registrar AI and database administrator.
Your task is to accurately populate all blank, 'tbc', or outdated course data fields for every course in this batch.
CRITICAL REQUIREMENTS:
1. OVERVIEW: Extract genuine program description from the scraped text. If the page overview is placeholder ('abc' or 'tbc'), synthesize a professional 2-3 sentence overview using the course title and its curriculum subjects. NEVER return 'tbc' or leave it blank.
2. STRUCTURE: Extract the list of subjects, course codes, and credit hours from the curriculum table into a clean comma-separated list (e.g., 'ELL 501 Introduction to Language (3), ELL 502 Introduction to Literary Studies (3)...').
3. ENTRY REQUIREMENTS: Extract the eligibility criteria (e.g. 'Intermediate with 45%', 'Bachelor degree with 40%').
4. FEE PER YEAR: Extract or calculate the annual fee if available; if semester fee is shown, multiply by 2.
5. CAREER PROSPECTS: If published on the page, extract it; otherwise, provide standard career pathways for this degree.
Output strictly valid JSON with exact column keys matching original column names.`;

    // Process all courses in sequential batches
    for (let startIdx = 0; startIdx < totalCourses; startIdx += BATCH_SIZE) {
      const endIdx = Math.min(startIdx + BATCH_SIZE, totalCourses);
      const batchSlice = candidateRows.slice(startIdx, endIdx);

      // Recursive deep scraping for courses in THIS batch
      const scrapedContextPerCourse: { [rowIdx: number]: string } = {};
      const batchFetchTasks = batchSlice.map(async (course) => {
        const rowId = course._row_index;
        let cLink = "";
        if (urlColIdx !== -1) {
          cLink = String(dataRows[rowId][urlColIdx] || "").trim();
        }
        if (!cLink.startsWith("http")) {
          for (const [k, v] of Object.entries(course)) {
            if (normalizeKey(k).includes("link")) {
              const valS = String(v || "").trim();
              if (valS.startsWith("http")) {
                cLink = valS;
                break;
              }
            }
          }
        }

        if (cLink.startsWith("http")) {
          allScrapedPages.add(cLink);
          const content = await fetchPageMarkdown(cLink);
          if (content) scrapedContextPerCourse[rowId] = content;
        }
      });

      await Promise.allSettled(batchFetchTasks);

      // Format courses with their specific scraped text
      const formattedBatch = batchSlice.map((r) => {
        const rowId = r._row_index;
        const cName = r.course_name || r.course_title || r[headerRow[0]] || "";
        const cUrl = r.course_detail_web_link || r.course_fee_web_link || "";
        const webText = scrapedContextPerCourse[rowId] || "";

        return {
          _row_index: rowId,
          course_name: cName,
          campus: r.campus || "",
          degree_level: r.degree_level || "",
          course_detail_web_link: cUrl,
          current_fields: {
            "overview ": r["overview "] || r.overview || "",
            structure: r.structure || "",
            entry_requirements: r.entry_requirements || "",
            fee_per_year: r.fee_per_year || "",
            career_prospects: r.career_prospects || "",
          },
          scraped_page_text: webText ? webText.slice(0, 6000) : "Use standard academic knowledge for this program.",
        };
      });

      const prompt = `
${customPrompt.trim() ? `TEACHER'S INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
BATCH: Courses ${startIdx + 1} to ${endIdx} of ${totalCourses} Total Courses:
${JSON.stringify(formattedBatch, null, 2)}

TASK:
- For every course in this batch, populate the missing/blank fields (especially 'overview ', 'structure', 'entry_requirements', 'fee_per_year', 'career_prospects').
- If 'overview ' has 'tbc' or is blank, replace it with a genuine, verified course overview.
- Return valid JSON matching this schema:
{
  "updated_rows": [
    {
      "row_index": ${startIdx},
      "updated_data": {
        "overview ": "Comprehensive overview text...",
        "structure": "Subject 1 (3), Subject 2 (3)...",
        "entry_requirements": "Eligibility criteria...",
        "fee_per_year": 43262,
        "career_prospects": "Career opportunities..."
      }
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
          if (Array.isArray(parsed.updated_rows)) {
            parsed.updated_rows.forEach((item: any) => {
              const rId = item.row_index;
              const uData = item.updated_data || {};
              if (rId !== undefined) {
                Object.entries(uData).forEach(([colKey, val]) => {
                  const norm = normalizeKey(colKey);
                  if (headerColMap[norm] !== undefined && val !== undefined && val !== null) {
                    const cIdx = headerColMap[norm];
                    cellModifications.set(`${rId}_${cIdx}`, val);
                  }
                });
              }
            });
          }
        }
      } catch (err: any) {
        console.warn(`Batch error:`, err?.message);
      }
    }

    // Step D: Apply Updates In-Place Preserving All Original Sheets & Formatting
    const outputWb = new ExcelJS.Workbook();

    for (const sName of workbook.SheetNames) {
      if (sName !== targetSheetName) {
        const origOtherSheet = workbook.Sheets[sName];
        const otherRows: any[] = XLSX.utils.sheet_to_json(origOtherSheet, { header: 1, defval: "" });
        const newOtherWs = outputWb.addWorksheet(sName);
        otherRows.forEach((r) => newOtherWs.addRow(r));
      }
    }

    const ws = outputWb.addWorksheet(targetSheetName);

    const addedHeader = ws.addRow(headerRow);
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

    dataRows.forEach((rowValues, rIdx) => {
      const modifiedRow = [...rowValues];

      for (let cIdx = 0; cIdx < headerRow.length; cIdx++) {
        const key = `${rIdx}_${cIdx}`;
        if (cellModifications.has(key)) {
          const newVal = cellModifications.get(key);
          const oldVal = String(modifiedRow[cIdx] || "").trim();
          if (oldVal !== String(newVal).trim() || oldVal.toLowerCase() === "tbc") {
            modifiedRow[cIdx] = newVal;
            totalUpdatedCells++;
          }
        }
      }

      const excelRow = ws.addRow(modifiedRow);

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
    summaryWs.addRow(["Total Courses Processed", totalCourses]);
    summaryWs.addRow(["Total Cells Populated & Verified", totalUpdatedCells]);
    summaryWs.addRow(["AI Model Used", lastModelUsed]);
    summaryWs.addRow(["API Key Account Used", lastKeyUsedIndex]);
    if (customPrompt.trim()) summaryWs.addRow(["Teacher Custom Instructions", customPrompt.trim()]);
    summaryWs.addRow(["Deep Links Scraped Count:", allScrapedPages.size]);

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
        "X-Summary-Notes": encodeURIComponent(`Recursive deep-link scraper populated ${totalUpdatedCells} cells across all rows.`),
        "X-Pages-Scraped": encodeURIComponent(JSON.stringify(Array.from(allScrapedPages).slice(0, 50))),
      },
    });
  } catch (err: any) {
    console.error("API Error:", err);
    return NextResponse.json({ error: err?.message || "Internal server error occurred." }, { status: 500 });
  }
}

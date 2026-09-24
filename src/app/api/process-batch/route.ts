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

async function fetchPageWithTimeout(url: string, timeoutMs = 3500): Promise<string> {
  if (!url || !url.startsWith("http")) return "";
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      headers: STANDARD_HEADERS,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (!res.ok) return "";
    const html = await res.text();
    return convertHtmlToCleanMarkdown(html);
  } catch {
    clearTimeout(timer);
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
    const body = await req.json();
    const {
      batchCourses = [],
      apiKeys = [],
      model = "gemini-2.5-flash",
      customPrompt = "",
      fallbackDirectContent = "",
      batchIndex = 0,
      totalBatches = 1,
    } = body;

    if (!Array.isArray(batchCourses) || batchCourses.length === 0) {
      return NextResponse.json({ error: "No courses provided in this batch." }, { status: 400 });
    }

    if (!Array.isArray(apiKeys) || apiKeys.length === 0) {
      return NextResponse.json({ error: "At least one Gemini API key is required." }, { status: 400 });
    }

    // Parallel deep scraping for links in THIS batch only (with fast 3.5s timeout)
    const scrapedContextPerCourse: { [rowId: number]: string } = {};
    const scrapedUrls: string[] = [];

    const fetchTasks = batchCourses.map(async (course: any) => {
      const rowId = course.rowIndex;
      const url = course.url || "";
      if (url.startsWith("http")) {
        scrapedUrls.push(url);
        const md = await fetchPageWithTimeout(url, 3500);
        if (md) scrapedContextPerCourse[rowId] = md;
      }
    });

    await Promise.allSettled(fetchTasks);

    // Format courses with their scraped context
    const formattedBatch = batchCourses.map((c: any) => {
      const rowId = c.rowIndex;
      let webText = scrapedContextPerCourse[rowId] || "";
      if (!webText && fallbackDirectContent) {
        webText = fallbackDirectContent.slice(0, 4000);
      }

      return {
        row_index: rowId,
        course_name: c.courseName || "",
        campus: c.campus || "",
        degree_level: c.degreeLevel || "",
        course_link: c.url || "",
        current_fields: c.currentFields || {},
        scraped_page_text: webText ? webText.slice(0, 5000) : "Extract details based on standard university catalog standards for this title.",
      };
    });

    const systemInstruction = `You are an expert university registrar AI and database administrator.
Your task is to accurately populate all blank, 'tbc', or outdated course data fields for every course in this batch.
CRITICAL REQUIREMENTS:
1. OVERVIEW: Extract genuine program description from the scraped text. If the page overview is placeholder ('abc' or 'tbc'), synthesize a professional 2-3 sentence overview using the course title and its curriculum subjects. NEVER return 'tbc' or leave it blank.
2. STRUCTURE: Extract the list of subjects, course codes, and credit hours from the curriculum table into a clean comma-separated list (e.g., 'ELL 501 Introduction to Language (3), ELL 502 Introduction to Literary Studies (3)...').
3. ENTRY REQUIREMENTS: Extract the eligibility criteria (e.g. 'Intermediate with 45%', 'Bachelor degree with 40%').
4. FEE PER YEAR: Extract or calculate the annual fee if available; if semester fee is shown, multiply by 2.
5. CAREER PROSPECTS: If published on the page, extract it; otherwise, provide standard career pathways for this degree.
Output strictly valid JSON with exact column keys matching original column names.`;

    const prompt = `
${customPrompt.trim() ? `TEACHER'S INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}
BATCH: Processing Batch ${batchIndex + 1} of ${totalBatches} (${batchCourses.length} courses):
${JSON.stringify(formattedBatch, null, 2)}

TASK:
- For every course in this batch, populate the missing/blank fields (especially 'overview', 'overview ', 'structure', 'entry_requirements', 'fee_per_year', 'career_prospects').
- If 'overview' has 'tbc' or is blank, replace it with a genuine, verified course overview.
- Return valid JSON matching this schema:
{
  "updated_rows": [
    {
      "row_index": ${batchCourses[0]?.rowIndex ?? 0},
      "updated_data": {
        "overview": "Comprehensive overview text...",
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

    const { text, modelUsed, keyUsedIndex } = await callGemini(apiKeys, model, prompt, systemInstruction);

    let updatedRows: any[] = [];
    const match = text.match(/\{[\s\S]*\}/);
    if (match) {
      try {
        const parsed = JSON.parse(match[0]);
        if (Array.isArray(parsed.updated_rows)) {
          updatedRows = parsed.updated_rows;
        }
      } catch (parseErr) {
        console.warn("JSON parse error:", parseErr);
      }
    }

    return NextResponse.json({
      success: true,
      updatedRows,
      modelUsed,
      keyUsedIndex,
      scrapedUrls,
    });
  } catch (err: any) {
    console.error("Batch processing error:", err);
    return NextResponse.json(
      { error: err?.message || "Failed to process batch." },
      { status: 500 }
    );
  }
}

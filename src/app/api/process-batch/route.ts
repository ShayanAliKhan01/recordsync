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

async function fetchPageWithTimeout(url: string, timeoutMs = 3500): Promise<{ text: string; isDiscontinued: boolean }> {
  if (!url || !url.startsWith("http")) return { text: "", isDiscontinued: false };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      headers: STANDARD_HEADERS,
      cache: "no-store",
      signal: controller.signal,
    });
    clearTimeout(timer);
    if (res.status === 404) {
      return { text: "HTTP 404 Not Found - Course link is dead or discontinued.", isDiscontinued: true };
    }
    if (!res.ok) return { text: "", isDiscontinued: false };
    const html = await res.text();
    const md = convertHtmlToCleanMarkdown(html);
    const hasDiscontinuedText = /no longer accepting applications|course discontinued|program discontinued|withdrawn for 2026|admissions closed permanently/i.test(md);
    return { text: md, isDiscontinued: hasDiscontinuedText };
  } catch {
    clearTimeout(timer);
    return { text: "", isDiscontinued: false };
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
      targetColumns = [], // Array of { index: number, name: string }
      batchCourses = [],  // Array of { rowIndex, entityTitle, campus, url, rowUrls: Record<string, string>, currentFields: Record<string, any> }
      apiKeys = [],
      model = "gemini-2.5-flash",
      customPrompt = "",
      fallbackDirectContent = "",
      cachedWebpages = {},
    } = body;

    if (!Array.isArray(batchCourses) || batchCourses.length === 0) {
      return NextResponse.json({ error: "No courses provided in this batch." }, { status: 400 });
    }

    if (!Array.isArray(apiKeys) || apiKeys.length === 0) {
      return NextResponse.json({ error: "At least one Gemini API key is required." }, { status: 400 });
    }

    // Scrape or load from cache for all dedicated link columns associated with each course
    const scrapedContextPerCourse: { [rowId: number]: Record<string, string> } = {};
    const discontinuedFlags: { [rowId: number]: boolean } = {};
    const newlyScraped: { [url: string]: string } = {};

    const fetchTasks: Promise<any>[] = [];

    batchCourses.forEach((course: any) => {
      const rId = course.rowIndex;
      scrapedContextPerCourse[rId] = {};

      const urlsToFetch: { label: string; url: string }[] = [];

      // Check if specific columns contain links where specific data is available
      if (course.rowUrls && typeof course.rowUrls === "object") {
        for (const [colName, urlVal] of Object.entries(course.rowUrls)) {
          if (typeof urlVal === "string" && urlVal.startsWith("http")) {
            urlsToFetch.push({ label: colName, url: urlVal.trim() });
          }
        }
      }

      // Fallback single URL
      if (urlsToFetch.length === 0 && course.url && String(course.url).startsWith("http")) {
        urlsToFetch.push({ label: "general_link", url: String(course.url).trim() });
      }

      urlsToFetch.forEach(({ label, url }) => {
        fetchTasks.push(
          (async () => {
            if (cachedWebpages && cachedWebpages[url]) {
              scrapedContextPerCourse[rId][label] = cachedWebpages[url];
            } else {
              const { text: md, isDiscontinued } = await fetchPageWithTimeout(url, 3500);
              if (md) {
                scrapedContextPerCourse[rId][label] = md;
                newlyScraped[url] = md;
              }
              if (isDiscontinued) {
                discontinuedFlags[rId] = true;
              }
            }
          })()
        );
      });
    });

    await Promise.allSettled(fetchTasks);

    // Target column names for prompt
    const colNames = targetColumns.map((c: any) => c.name);

    // Format courses with their source-attributed scraped texts
    const formattedBatch = batchCourses.map((c: any) => {
      const rId = c.rowIndex;
      const sources = scrapedContextPerCourse[rId] || {};

      const scrapedSources: Record<string, string> = {};
      for (const [colSource, md] of Object.entries(sources)) {
        scrapedSources[`scraped_from_${colSource}`] = md.slice(0, 5000);
      }

      if (Object.keys(scrapedSources).length === 0 && fallbackDirectContent) {
        scrapedSources["fallback_direct_content"] = fallbackDirectContent.slice(0, 4000);
      }

      return {
        row_index: rId,
        entity_title: c.entityTitle || `Record #${rId + 1}`,
        campus: c.campus || "",
        current_target_boxes: c.currentFields || {},
        is_page_marked_discontinued_or_404: Boolean(discontinuedFlags[rId]),
        available_scraped_sites: scrapedSources,
      };
    });

    const systemInstruction = `You are an expert academic database auditor, registrar, and quality assurance officer.
Your task is to accurately verify and populate the TARGET COLUMNS for every course in this batch.
TARGET COLUMNS: ${JSON.stringify(colNames)}

CRITICAL DEDICATED SITE MATCHING:
In this university database, specific columns have their official data available on specific dedicated websites:
1. FEE / TUITION COLUMNS (e.g., 'fee_per_year', 'tuition_fee', 'per_credit_fee'):
   - Inspect the website scraped from the FEE LINK (e.g., 'scraped_from_course_fee_web_link' or 'scraped_from_fee_url').
   - Calculate or extract the 2026 annual tuition. If semester fee is listed, multiply by 2. Return clean numeric/currency text.
2. OVERVIEW / DESCRIPTION COLUMNS (e.g., 'overview', 'overview ', 'course_description'):
   - Inspect the website scraped from the COURSE DETAIL LINK (e.g., 'scraped_from_course_detail_web_link').
   - Extract genuine 2-3 sentence overview. If blank or 'tbc', synthesize a verified overview. NEVER leave 'tbc'.
3. STRUCTURE / CURRICULUM COLUMNS (e.g., 'structure', 'syllabus'):
   - Inspect the course detail page syllabus/curriculum table.
   - Extract all subjects, course codes, and credit hours into a clean comma-separated list.
4. ENTRY REQUIREMENTS / ELIGIBILITY:
   - Check the admission portal link ('scraped_from_course_apply_web_link') or the course detail page.
   - Extract minimum degree, percentage, or GPA criteria.
5. If the current box already accurately matches the 2026 data, keep it and mark "VERIFIED_CURRENT". If blank or 'tbc', mark "AUTOFILLED". If outdated, mark "UPDATED".
6. DISCONTINUED: If any official source says the course is discontinued or returns 404, set status: "DISCONTINUED" and value: "Discontinued in 2026".

Output strictly valid JSON with exact column keys matching the target column names.`;

    const prompt = `
TARGET COLUMNS TO VERIFY: ${JSON.stringify(colNames)}
${customPrompt.trim() ? `ADDITIONAL USER INSTRUCTION:\n"${customPrompt.trim()}"\n` : ""}

Process the following ${formattedBatch.length} courses:
${JSON.stringify(formattedBatch, null, 2)}

REQUIRED JSON OUTPUT FORMAT:
{
  "updated_rows": [
    {
      "row_index": ${formattedBatch[0]?.row_index ?? 0},
      "status": "UPDATED" | "AUTOFILLED" | "VERIFIED_CURRENT" | "DISCONTINUED",
      "confidence": "HIGH" | "MEDIUM" | "LOW",
      "reason": "Brief summary of changes made",
      "evidence_snippet": "Quote from webpage supporting the data",
      "fields": {
        ${colNames.map((name: string) => `"${name}": "Verified 2026 value"`).join(",\n        ")}
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
      } catch (err) {
        console.warn("Parse error in process-batch:", err);
      }
    }

    return NextResponse.json({
      success: true,
      updatedRows,
      modelUsed,
      keyUsedIndex,
      newlyScraped,
    });
  } catch (err: any) {
    console.error("Process Batch Error:", err);
    return NextResponse.json(
      { error: err?.message || "Failed to process batch." },
      { status: 500 }
    );
  }
}

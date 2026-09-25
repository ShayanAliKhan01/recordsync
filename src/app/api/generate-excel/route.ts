import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 45;
export const dynamic = "force-dynamic";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const targetSheetName = (formData.get("targetSheetName") as string | null) || "";
    const modificationsRaw = (formData.get("modifications") as string | null) || "{}";
    const auditLogsRaw = (formData.get("auditLogs") as string | null) || "[]";
    const scrapedCacheRaw = (formData.get("scrapedCache") as string | null) || "{}";
    const metadataRaw = (formData.get("metadata") as string | null) || "{}";

    if (!file) {
      return NextResponse.json({ error: "Missing uploaded Excel file." }, { status: 400 });
    }

    const modifications: Record<string, any> = JSON.parse(modificationsRaw);
    const auditLogs: any[] = JSON.parse(auditLogsRaw);
    const scrapedCache: Record<string, string> = JSON.parse(scrapedCacheRaw);
    const metadata = JSON.parse(metadataRaw);

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer as any);

    // Get target worksheet or fallback to first sheet
    let ws = wb.getWorksheet(targetSheetName);
    if (!ws) {
      ws = wb.worksheets[0];
    }

    const yellowFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFF2CC" },
    };

    const redDiscontinuedFill: ExcelJS.Fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFCE4D6" },
    };

    let updatedCount = 0;

    // Apply cell modifications directly in-place using exact coordinates
    for (const [key, modInfo] of Object.entries(modifications)) {
      const parts = key.split("_");
      if (parts.length !== 2) continue;
      const rIdx = parseInt(parts[0], 10);
      const cIdx = parseInt(parts[1], 10);

      // Data row index rIdx maps to Excel row (rIdx + 2) because row 1 is header
      const row = ws.getRow(rIdx + 2);
      const cell = row.getCell(cIdx + 1);

      let val = modInfo;
      let status = "UPDATED";
      if (typeof modInfo === "object" && modInfo !== null && "value" in modInfo) {
        val = modInfo.value;
        status = modInfo.status || "UPDATED";
      }

      cell.value = val;

      if (status === "DISCONTINUED") {
        cell.fill = redDiscontinuedFill;
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFC00000" } };
      } else {
        cell.fill = yellowFill;
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF7F6000" } };
      }
      updatedCount++;
    }

    // 1. Add or replace "Verification_Audit" sheet
    if (auditLogs.length > 0) {
      const existingAudit = wb.getWorksheet("Verification_Audit");
      if (existingAudit) wb.removeWorksheet(existingAudit.id);

      const auditWs = wb.addWorksheet("Verification_Audit");
      const headerRow = auditWs.addRow([
        "Row #",
        "Course / Entity Name",
        "Column Name",
        "Previous Box Value",
        "Verified 2026 Value",
        "Status",
        "Confidence",
        "Evidence & Verification Reason",
        "Source Web Link",
      ]);

      headerRow.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF1E3C72" },
        };
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });

      auditLogs.forEach((log) => {
        const r = auditWs.addRow([
          log.rowIndex + 1,
          log.entityTitle || "",
          log.columnName || "",
          log.previousValue !== undefined ? String(log.previousValue) : "",
          log.verifiedValue !== undefined ? String(log.verifiedValue) : "",
          log.status || "VERIFIED",
          log.confidence || "HIGH",
          log.reason || log.evidenceSnippet || "",
          log.sourceUrl || "",
        ]);

        const statusCell = r.getCell(6);
        const confCell = r.getCell(7);

        if (log.status === "DISCONTINUED") {
          statusCell.fill = redDiscontinuedFill;
          statusCell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFC00000" } };
        } else if (log.status === "UPDATED" || log.status === "AUTOFILLED") {
          statusCell.fill = yellowFill;
          statusCell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FF7F6000" } };
        }

        if (log.confidence === "LOW") {
          confCell.fill = redDiscontinuedFill;
          confCell.font = { name: "Calibri", size: 10, bold: true, color: { argb: "FFC00000" } };
        }
      });

      auditWs.columns = [
        { width: 10 },
        { width: 35 },
        { width: 25 },
        { width: 30 },
        { width: 45 },
        { width: 20 },
        { width: 15 },
        { width: 50 },
        { width: 40 },
      ];
    }

    // 2. Add or replace "Scraped_Audit_Cache" sheet (Persistent Webpage Data Storage)
    const scrapedUrls = Object.keys(scrapedCache);
    if (scrapedUrls.length > 0) {
      const existingCache = wb.getWorksheet("Scraped_Audit_Cache");
      if (existingCache) wb.removeWorksheet(existingCache.id);

      const cacheWs = wb.addWorksheet("Scraped_Audit_Cache");
      const cHeader = cacheWs.addRow([
        "#",
        "Source Webpage URL",
        "Scraped Text Sample (First 2000 chars)",
        "Cached Timestamp",
      ]);

      cHeader.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0F2027" },
        };
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
      });

      scrapedUrls.forEach((url, i) => {
        const textSample = (scrapedCache[url] || "").slice(0, 2000);
        cacheWs.addRow([i + 1, url, textSample, new Date().toISOString()]);
      });

      cacheWs.columns = [
        { width: 8 },
        { width: 45 },
        { width: 80 },
        { width: 25 },
      ];
    }

    // 3. Add or replace "Update Summary" sheet
    let summaryWs = wb.getWorksheet("Update Summary");
    if (summaryWs) wb.removeWorksheet(summaryWs.id);
    summaryWs = wb.addWorksheet("Update Summary");

    summaryWs.addRow(["RecordSync Dynamic Execution Summary"]);
    summaryWs.addRow(["Sheet Updated", ws.name]);
    summaryWs.addRow(["Total Boxes Verified & Populated", updatedCount]);
    summaryWs.addRow(["Columns Verified", metadata.columnsVerified || "Dynamic Target Columns"]);
    summaryWs.addRow(["AI Model Used", metadata.modelUsed || "gemini-2.5-flash"]);
    summaryWs.addRow(["API Key Account Used", metadata.keyUsed || "1"]);
    if (metadata.customPrompt) {
      summaryWs.addRow(["User Custom Instructions", metadata.customPrompt]);
    }
    summaryWs.addRow(["Webpages Scraped & Cached", scrapedUrls.length]);
    summaryWs.addRow(["Timestamp", new Date().toISOString()]);

    summaryWs.columns.forEach((col) => {
      col.width = 35;
    });

    const outputBuffer = await wb.xlsx.writeBuffer();

    return new NextResponse(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="Updated_2026_${file.name}"`,
        "X-Updated-Count": String(updatedCount),
        "X-Model-Used": metadata.modelUsed || "gemini-2.5-flash",
        "X-Key-Used": String(metadata.keyUsed || "1"),
      },
    });
  } catch (err: any) {
    console.error("Generate Excel Error:", err);
    return NextResponse.json(
      { error: err?.message || "Failed to generate Excel file." },
      { status: 500 }
    );
  }
}

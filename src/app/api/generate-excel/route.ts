import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export const maxDuration = 30;
export const dynamic = "force-dynamic";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;
    const targetSheetName = (formData.get("targetSheetName") as string | null) || "";
    const modificationsRaw = (formData.get("modifications") as string | null) || "{}";
    const metadataRaw = (formData.get("metadata") as string | null) || "{}";

    if (!file) {
      return NextResponse.json({ error: "Missing uploaded Excel file." }, { status: 400 });
    }

    const modifications: Record<string, any> = JSON.parse(modificationsRaw);
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

    let updatedCount = 0;

    // Apply cell modifications directly in-place
    for (const [key, newVal] of Object.entries(modifications)) {
      const parts = key.split("_");
      if (parts.length !== 2) continue;
      const rIdx = parseInt(parts[0], 10);
      const cIdx = parseInt(parts[1], 10);

      // Data row index rIdx maps to Excel row (rIdx + 2) because row 1 is the header
      const row = ws.getRow(rIdx + 2);
      const cell = row.getCell(cIdx + 1);

      cell.value = newVal;
      cell.fill = yellowFill;
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF7F6000" } };
      updatedCount++;
    }

    // Add or replace "Update Summary" sheet
    let summaryWs = wb.getWorksheet("Update Summary");
    if (summaryWs) {
      wb.removeWorksheet(summaryWs.id);
    }
    summaryWs = wb.addWorksheet("Update Summary");

    summaryWs.addRow(["RecordSync 2026 Batch Execution Summary"]);
    summaryWs.addRow(["Sheet Updated", ws.name]);
    summaryWs.addRow(["Total Cells Populated & Verified", updatedCount]);
    summaryWs.addRow(["AI Model Used", metadata.modelUsed || "gemini-2.5-flash"]);
    summaryWs.addRow(["API Key Account Used", metadata.keyUsed || "1"]);
    if (metadata.customPrompt) {
      summaryWs.addRow(["Teacher Custom Instructions", metadata.customPrompt]);
    }
    summaryWs.addRow(["Deep Links Scraped Count", metadata.scrapedPagesCount || 0]);
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

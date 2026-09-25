"use client";

import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  FileSpreadsheet,
  Globe,
  Key,
  Sparkles,
  Download,
  AlertCircle,
  CheckCircle2,
  Cpu,
  RefreshCw,
  Zap,
  MessageSquareText,
  ClipboardPaste,
  ChevronDown,
  ChevronUp,
  CheckSquare,
  Square,
  Table,
  ShieldCheck,
  Eye,
  X,
  RotateCcw,
  Check,
  ExternalLink,
} from "lucide-react";

const MODEL_OPTIONS = [
  {
    id: "gemini-2.5-flash",
    name: "Gemini 2.5 Flash (Recommended Default)",
    quota: "250K TPM / 20 RPD",
    desc: "Best for accurate structured JSON reconciliation & fast reasoning.",
  },
  {
    id: "gemini-3.5-flash-lite",
    name: "Gemini 3.5 Flash Lite (High Daily Volume)",
    quota: "250K TPM / 500 RPD",
    desc: "Best daily limit (500 requests/day). High capacity for multiple files.",
  },
  {
    id: "gemini-3.1-flash-lite",
    name: "Gemini 3.1 Flash Lite",
    quota: "250K TPM / 500 RPD",
    desc: "Secondary light model with 500 daily requests limit.",
  },
  {
    id: "gemini-3.5-flash",
    name: "Gemini 3.5 Flash",
    quota: "250K TPM / 20 RPD",
    desc: "Advanced reasoning for complex unstructured university web tables.",
  },
  {
    id: "gemini-3.7-flash",
    name: "Gemini 3.7 Flash",
    quota: "250K TPM / 20 RPD",
    desc: "Latest flash model with extended reasoning capabilities.",
  },
];

interface ColumnMeta {
  index: number;
  name: string;
  emptyOrPlaceholderCount: number;
  totalRows: number;
  selected: boolean;
}

interface VerificationItem {
  rowIndex: number;
  entityTitle: string;
  columnName: string;
  previousValue: any;
  verifiedValue: any;
  status: string;
  confidence: string;
  reason: string;
  evidenceSnippet?: string;
  sourceUrl?: string;
}

export default function RecordSyncPage() {
  const [apiKeys, setApiKeys] = useState<string>("");
  const [selectedModel, setSelectedModel] = useState<string>("gemini-2.5-flash");
  const [url, setUrl] = useState<string>("");
  const [directContent, setDirectContent] = useState<string>("");
  const [showDirectPaste, setShowDirectPaste] = useState<boolean>(false);
  const [customPrompt, setCustomPrompt] = useState<string>("");
  const [file, setFile] = useState<File | null>(null);

  // Dynamic Excel Introspection State
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string>("");
  const [availableColumns, setAvailableColumns] = useState<ColumnMeta[]>([]);
  const [linkColIndex, setLinkColIndex] = useState<number>(-1);
  const [titleColIndex, setTitleColIndex] = useState<number>(0);
  const [campusColIndex, setCampusColIndex] = useState<number>(-1);
  const [parsedDataRows, setParsedDataRows] = useState<any[][]>([]);
  const [parsedHeaderRow, setParsedHeaderRow] = useState<string[]>([]);

  // Execution & Live Progress
  const [loading, setLoading] = useState<boolean>(false);
  const [statusStep, setStatusStep] = useState<string>("");
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const [columnProgress, setColumnProgress] = useState<{
    currentColName: string;
    colIndexNum: number;
    totalCols: number;
    currentBatch: number;
    totalBatches: number;
    percent: number;
    updatedBoxes: number;
  } | null>(null);

  const [liveVerificationFeed, setLiveVerificationFeed] = useState<VerificationItem[]>([]);
  const [selectedDiffItem, setSelectedDiffItem] = useState<VerificationItem | null>(null);

  // Resume Session State
  const [savedSessionNotice, setSavedSessionNotice] = useState<string | null>(null);

  const [downloadBlobUrl, setDownloadBlobUrl] = useState<string | null>(null);
  const [downloadFileName, setDownloadFileName] = useState<string>("");
  const [stats, setStats] = useState<{
    updatedBoxes: number;
    verifiedBoxes: number;
    discontinuedCount: number;
    columnsProcessed: number;
    modelUsed: string;
    keyUsed: string;
    notes: string;
    scrapedPagesCount: number;
  } | null>(null);

  // Check for saved session on mount
  useEffect(() => {
    try {
      const stored = localStorage.getItem("recordsync_active_session");
      if (stored) {
        const parsed = JSON.parse(stored);
        if (parsed.fileName && parsed.auditLogs?.length > 0) {
          setSavedSessionNotice(`Found saved session for "${parsed.fileName}" with ${parsed.auditLogs.length} verified boxes.`);
        }
      }
    } catch {
      // ignore
    }
  }, []);

  const handleClearSession = () => {
    localStorage.removeItem("recordsync_active_session");
    setSavedSessionNotice(null);
  };

  // Dynamic File Upload & Introspection
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files || !e.target.files[0]) return;
    const uploadedFile = e.target.files[0];
    setFile(uploadedFile);
    setErrorMsg(null);
    setDownloadBlobUrl(null);
    setStats(null);
    setLiveVerificationFeed([]);

    try {
      const arrayBuffer = await uploadedFile.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: "array" });
      setSheetNames(wb.SheetNames);

      // Select sheet with highest row count
      let targetSheet = wb.SheetNames[0];
      let maxRowCount = 0;
      for (const sName of wb.SheetNames) {
        const s = wb.Sheets[sName];
        const rLen = XLSX.utils.sheet_to_json(s, { header: 1 }).length;
        if (rLen > maxRowCount) {
          maxRowCount = rLen;
          targetSheet = sName;
        }
      }
      setActiveSheet(targetSheet);

      const ws = wb.Sheets[targetSheet];
      const sheetData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (sheetData.length <= 1) {
        throw new Error("No data rows found in this Excel sheet.");
      }

      const headers: string[] = sheetData[0].map((h: any) => String(h || "").trim());
      const dataRows = sheetData.slice(1);
      setParsedHeaderRow(headers);
      setParsedDataRows(dataRows);

      // Dynamically detect link, title, and campus columns
      let detectedLinkIdx = -1;
      let detectedTitleIdx = 0;
      let detectedCampusIdx = -1;

      headers.forEach((h, idx) => {
        const lower = h.toLowerCase();
        if (detectedLinkIdx === -1 && (lower.includes("link") || lower.includes("url") || lower.includes("web"))) {
          detectedLinkIdx = idx;
        }
        if (lower.includes("course") && (lower.includes("name") || lower.includes("title"))) {
          detectedTitleIdx = idx;
        } else if (lower.includes("program") || lower.includes("degree_name")) {
          detectedTitleIdx = idx;
        }
        if (lower.includes("campus") || lower.includes("location") || lower.includes("city")) {
          detectedCampusIdx = idx;
        }
      });

      setLinkColIndex(detectedLinkIdx);
      setTitleColIndex(detectedTitleIdx);
      setCampusColIndex(detectedCampusIdx);

      // Inspect all columns dynamically: count missing/placeholder values
      const cols: ColumnMeta[] = headers.map((h, idx) => {
        let missing = 0;
        dataRows.forEach((r) => {
          const val = String(r[idx] || "").trim().toLowerCase();
          if (!val || val === "tbc" || val === "abc" || val === "tbd" || val === "none" || val === "n/a" || val === "-" || val === "null") {
            missing++;
          }
        });

        // Auto-select columns that are incomplete or typical academic targets
        const lower = h.toLowerCase();
        const isTarget =
          missing > 0 &&
          idx !== detectedLinkIdx &&
          (lower.includes("overview") ||
            lower.includes("structure") ||
            lower.includes("syllabus") ||
            lower.includes("fee") ||
            lower.includes("tuition") ||
            lower.includes("requirement") ||
            lower.includes("eligibility") ||
            lower.includes("career") ||
            lower.includes("credit") ||
            lower.includes("duration") ||
            lower.includes("intake") ||
            missing > 10);

        return {
          index: idx,
          name: h || `Column_${idx + 1}`,
          emptyOrPlaceholderCount: missing,
          totalRows: dataRows.length,
          selected: isTarget,
        };
      });

      setAvailableColumns(cols);
    } catch (err: any) {
      setErrorMsg("Failed to parse Excel file: " + err.message);
    }
  };

  const toggleColumnSelection = (index: number) => {
    setAvailableColumns((prev) =>
      prev.map((c) => (c.index === index ? { ...c, selected: !c.selected } : c))
    );
  };

  const selectAllIncomplete = () => {
    setAvailableColumns((prev) =>
      prev.map((c) => ({
        ...c,
        selected: c.emptyOrPlaceholderCount > 0 && c.index !== linkColIndex,
      }))
    );
  };

  const selectAllColumns = () => {
    setAvailableColumns((prev) => prev.map((c) => ({ ...c, selected: true })));
  };

  const deselectAllColumns = () => {
    setAvailableColumns((prev) => prev.map((c) => ({ ...c, selected: false })));
  };

  // Run the dynamic column-by-column, box-by-box verification
  const handleRunDynamicUpdate = async () => {
    setErrorMsg(null);
    setDownloadBlobUrl(null);
    setStats(null);
    setLiveVerificationFeed([]);

    const parsedKeys = apiKeys
      .split(",")
      .map((k) => k.trim())
      .filter(Boolean);

    if (parsedKeys.length === 0) {
      setErrorMsg("Please provide at least one Google Gemini API Key.");
      return;
    }
    if (!file || parsedDataRows.length === 0) {
      setErrorMsg("Please upload your Excel file.");
      return;
    }

    const selectedCols = availableColumns.filter((c) => c.selected);
    if (selectedCols.length === 0) {
      setErrorMsg("Please select at least one target column to verify and update.");
      return;
    }

    setLoading(true);

    const modifications: Record<string, any> = {};
    const auditLogs: VerificationItem[] = [];
    const cachedWebpages: Record<string, string> = {};
    let totalUpdatedBoxes = 0;
    let totalVerifiedBoxes = 0;
    let totalDiscontinued = 0;
    let activeModel = selectedModel;
    let activeKeyIndex = 1;

    const totalRows = parsedDataRows.length;
    const BATCH_SIZE = 10;
    const totalBatchesPerRow = Math.ceil(totalRows / BATCH_SIZE);
    const totalSteps = selectedCols.length * totalBatchesPerRow;
    let stepsCompleted = 0;

    try {
      // Loop dynamically column-by-column
      for (let cIdx = 0; cIdx < selectedCols.length; cIdx++) {
        const targetCol = selectedCols[cIdx];
        const colNumber = targetCol.index;
        const colName = targetCol.name;

        for (let bIdx = 0; bIdx < totalBatchesPerRow; bIdx++) {
          const start = bIdx * BATCH_SIZE;
          const end = Math.min(start + BATCH_SIZE, totalRows);
          const currentSlice = parsedDataRows.slice(start, end);

          stepsCompleted++;
          const percent = Math.round((stepsCompleted / totalSteps) * 100);

          setColumnProgress({
            currentColName: colName,
            colIndexNum: cIdx + 1,
            totalCols: selectedCols.length,
            currentBatch: bIdx + 1,
            totalBatches: totalBatchesPerRow,
            percent,
            updatedBoxes: totalUpdatedBoxes,
          });

          setStatusStep(`Column ${cIdx + 1}/${selectedCols.length} ("${colName}") • Rows ${start + 1}–${end} of ${totalRows}`);

          const batchPayload = currentSlice.map((r, sliceIdx) => {
            const actualRowIndex = start + sliceIdx;
            let courseLink = linkColIndex !== -1 ? String(r[linkColIndex] || "").trim() : "";
            if (!courseLink.startsWith("http") && url.trim().startsWith("http")) {
              courseLink = url.trim();
            }

            return {
              rowIndex: actualRowIndex,
              entityTitle: String(r[titleColIndex] || `Record #${actualRowIndex + 1}`).trim(),
              campus: campusColIndex !== -1 ? String(r[campusColIndex] || "").trim() : "",
              currentValue: r[colNumber],
              url: courseLink,
            };
          });

          // Call dynamic /api/process-column
          let attemptSuccess = false;
          let retries = 0;

          while (!attemptSuccess && retries < 2) {
            try {
              const res = await fetch("/api/process-column", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  columnName: colName,
                  columnIndex: colNumber,
                  batchRows: batchPayload,
                  apiKeys: parsedKeys,
                  model: selectedModel,
                  customPrompt,
                  fallbackDirectContent: directContent,
                  cachedWebpages,
                }),
              });

              if (!res.ok) {
                const errJson = await res.json().catch(() => ({}));
                throw new Error(errJson.error || `Server status ${res.status}`);
              }

              const data = await res.json();
              activeModel = data.modelUsed || activeModel;
              activeKeyIndex = data.keyUsedIndex || activeKeyIndex;

              // Cache newly scraped pages
              if (data.newlyScraped) {
                Object.assign(cachedWebpages, data.newlyScraped);
              }

              if (Array.isArray(data.results)) {
                data.results.forEach((item: any) => {
                  const rIdx = item.row_index;
                  const verifiedVal = item.verified_value;
                  const status = item.status || "VERIFIED_CURRENT";
                  const confidence = item.confidence || "HIGH";
                  const reason = item.reason || "";
                  const evidenceSnippet = item.evidence_snippet || "";
                  const prevVal = parsedDataRows[rIdx]?.[colNumber];
                  const entityTitle = String(parsedDataRows[rIdx]?.[titleColIndex] || `Record #${rIdx + 1}`);
                  const sourceUrl = linkColIndex !== -1 ? String(parsedDataRows[rIdx]?.[linkColIndex] || "") : url;

                  if (status === "DISCONTINUED") {
                    modifications[`${rIdx}_${colNumber}`] = { value: "Discontinued in 2026", status: "DISCONTINUED" };
                    totalDiscontinued++;
                  } else if (status === "UPDATED" || status === "AUTOFILLED") {
                    modifications[`${rIdx}_${colNumber}`] = { value: verifiedVal, status };
                    totalUpdatedBoxes++;
                  } else {
                    totalVerifiedBoxes++;
                  }

                  const verificationEntry: VerificationItem = {
                    rowIndex: rIdx,
                    entityTitle,
                    columnName: colName,
                    previousValue: prevVal,
                    verifiedValue: verifiedVal,
                    status,
                    confidence,
                    reason,
                    evidenceSnippet,
                    sourceUrl,
                  };

                  auditLogs.push(verificationEntry);

                  // Update live UI feed
                  setLiveVerificationFeed((prev) => [verificationEntry, ...prev.slice(0, 14)]);
                });
              }

              // Auto-save session progress to localStorage
              try {
                localStorage.setItem(
                  "recordsync_active_session",
                  JSON.stringify({
                    fileName: file.name,
                    activeSheet,
                    stepsCompleted,
                    totalSteps,
                    auditLogsCount: auditLogs.length,
                    timestamp: new Date().toISOString(),
                  })
                );
              } catch {
                // ignore
              }

              attemptSuccess = true;
            } catch (err: any) {
              retries++;
              if (retries >= 2) {
                console.warn(`Column ${colName} batch error:`, err.message);
              } else {
                await new Promise((res) => setTimeout(res, 1500));
              }
            }
          }
        }
      }

      // Step 3: Finalize and generate Excel
      setStatusStep("Assembling updated workbook with highlighted boxes & verification audit sheet...");

      const generateFormData = new FormData();
      generateFormData.append("file", file);
      generateFormData.append("targetSheetName", activeSheet);
      generateFormData.append("modifications", JSON.stringify(modifications));
      generateFormData.append("auditLogs", JSON.stringify(auditLogs));
      generateFormData.append("scrapedCache", JSON.stringify(cachedWebpages));
      generateFormData.append(
        "metadata",
        JSON.stringify({
          modelUsed: activeModel,
          keyUsed: activeKeyIndex,
          customPrompt,
          columnsVerified: selectedCols.map((c) => c.name).join(", "),
          scrapedPagesCount: Object.keys(cachedWebpages).length,
        })
      );

      const genRes = await fetch("/api/generate-excel", {
        method: "POST",
        body: generateFormData,
      });

      if (!genRes.ok) {
        const errJson = await genRes.json().catch(() => ({}));
        throw new Error(errJson.error || "Failed to generate final Excel file.");
      }

      const blob = await genRes.blob();
      const blobUrl = URL.createObjectURL(blob);
      setDownloadBlobUrl(blobUrl);
      setDownloadFileName(`Updated_2026_${file.name}`);

      // Clear session after successful completion
      localStorage.removeItem("recordsync_active_session");
      setSavedSessionNotice(null);

      setStats({
        updatedBoxes: totalUpdatedBoxes,
        verifiedBoxes: totalVerifiedBoxes,
        discontinuedCount: totalDiscontinued,
        columnsProcessed: selectedCols.length,
        modelUsed: activeModel,
        keyUsed: String(activeKeyIndex),
        notes: `Processed ${totalRows} rows across ${selectedCols.length} target columns. Updated ${totalUpdatedBoxes} boxes (#FFF2CC), verified ${totalVerifiedBoxes} boxes, detected ${totalDiscontinued} discontinued courses (#FCE4D6).`,
        scrapedPagesCount: Object.keys(cachedWebpages).length,
      });
    } catch (err: any) {
      console.error(err);
      setErrorMsg(err.message || "An unexpected error occurred during execution.");
    } finally {
      setLoading(false);
      setStatusStep("");
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* Header Bar */}
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center shadow-lg shadow-blue-500/20">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-lg text-white leading-tight flex items-center gap-2">
                RecordSync <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20">Pro Verification Engine</span>
              </h1>
              <p className="text-xs text-slate-400">Box-by-Box Verification, Confidence Scoring & Side-by-Side Diff Preview</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs text-emerald-400 bg-emerald-950/80 border border-emerald-800/60 px-3 py-1 rounded-full font-medium flex items-center gap-1.5">
              <Zap className="w-3.5 h-3.5" /> 100% Free Gemini API
            </span>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        {/* Left Column: API & Configuration */}
        <div className="lg:col-span-4 space-y-6">
          {/* Saved Session Alert Banner */}
          {savedSessionNotice && (
            <div className="bg-blue-950/80 border border-blue-800 p-4 rounded-2xl space-y-2 text-xs">
              <div className="flex items-center justify-between font-semibold text-blue-300">
                <span className="flex items-center gap-1.5">
                  <RotateCcw className="w-3.5 h-3.5 text-blue-400" /> Active Session Saved
                </span>
                <button onClick={handleClearSession} className="text-slate-400 hover:text-slate-200">
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
              <p className="text-slate-300">{savedSessionNotice}</p>
            </div>
          )}

          {/* API Key Box */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <Key className="w-4 h-4 text-blue-400" /> Google Gemini API Key(s)
              </label>
              <span className="text-[10px] text-slate-400 bg-slate-800 px-2 py-0.5 rounded">Multi-Key Failover</span>
            </div>
            <textarea
              rows={3}
              value={apiKeys}
              onChange={(e) => setApiKeys(e.target.value)}
              placeholder="Paste your Gemini API key(s) here. Separate multiple keys with commas for auto-failover on quota limits..."
              className="w-full text-xs bg-slate-950 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 resize-none font-mono"
            />
          </div>

          {/* Model Selector */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-3">
            <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
              <Cpu className="w-4 h-4 text-indigo-400" /> Primary AI Model
            </label>
            <select
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2.5 text-xs text-slate-200 focus:outline-none focus:ring-2 focus:ring-indigo-500/50"
            >
              {MODEL_OPTIONS.map((opt) => (
                <option key={opt.id} value={opt.id}>
                  {opt.name} ({opt.quota})
                </option>
              ))}
            </select>
          </div>

          {/* Dynamic Logic Info Card */}
          <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-3 text-xs text-slate-400">
            <div className="flex items-center gap-2 text-slate-200 font-semibold">
              <ShieldCheck className="w-4 h-4 text-emerald-400" /> Quality Control Protocol
            </div>
            <div className="space-y-1.5 leading-relaxed">
              <p>• <strong>🟢 HIGH Confidence:</strong> Exact official table match.</p>
              <p>• <strong>🟡 MEDIUM:</strong> Calculated (semester fee × 2).</p>
              <p>• <strong>🔴 REVIEW:</strong> Synthesized from catalog.</p>
              <p>• <strong>⚠️ DISCONTINUED:</strong> 404 dead link or retired program.</p>
            </div>
          </div>
        </div>

        {/* Right Column: Dynamic Form & Execution */}
        <div className="lg:col-span-8 space-y-6">
          {/* Main Action Card */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl space-y-5">
            {/* Input 1: File Upload */}
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" /> 1. Upload Any Excel File (.xlsx / .xls)
              </label>
              <div className="relative border-2 border-dashed border-slate-800 hover:border-blue-500/50 bg-slate-950/50 rounded-2xl p-6 text-center transition cursor-pointer group">
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleFileChange}
                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                />
                <div className="flex flex-col items-center space-y-2">
                  <div className="w-12 h-12 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center group-hover:scale-110 transition">
                    <FileSpreadsheet className="w-6 h-6 text-blue-400" />
                  </div>
                  {file ? (
                    <div>
                      <p className="text-sm font-medium text-emerald-400">{file.name}</p>
                      <p className="text-xs text-slate-400">
                        {(file.size / 1024).toFixed(1)} KB • {parsedDataRows.length} rows detected in sheet "{activeSheet}"
                      </p>
                    </div>
                  ) : (
                    <div>
                      <p className="text-sm font-medium text-slate-300">Drag & drop any university Excel file or browse</p>
                      <p className="text-xs text-slate-500">Works dynamically for any template, column layout, or structure</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Dynamic Column Selector Card */}
            {availableColumns.length > 0 && (
              <div className="bg-slate-950/90 border border-slate-800 rounded-xl p-5 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex items-center gap-2">
                    <Table className="w-4 h-4 text-indigo-400" />
                    <span className="text-sm font-semibold text-slate-200">
                      Target Columns to Verify & Update ({availableColumns.filter((c) => c.selected).length}/{availableColumns.length})
                    </span>
                  </div>
                  <div className="flex items-center gap-2 text-xs">
                    <button
                      type="button"
                      onClick={selectAllIncomplete}
                      className="px-2.5 py-1 rounded-lg bg-blue-950 border border-blue-800 text-blue-300 hover:bg-blue-900 transition"
                    >
                      Select Incomplete
                    </button>
                    <button
                      type="button"
                      onClick={selectAllColumns}
                      className="px-2.5 py-1 rounded-lg bg-slate-800 text-slate-300 hover:bg-slate-700 transition"
                    >
                      All
                    </button>
                    <button
                      type="button"
                      onClick={deselectAllColumns}
                      className="px-2.5 py-1 rounded-lg bg-slate-800 text-slate-400 hover:bg-slate-700 transition"
                    >
                      None
                    </button>
                  </div>
                </div>

                {/* Column Checklist Grid */}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2.5 max-h-56 overflow-y-auto pr-2">
                  {availableColumns.map((col) => (
                    <div
                      key={col.index}
                      onClick={() => toggleColumnSelection(col.index)}
                      className={`flex items-center justify-between p-2.5 rounded-lg border text-xs cursor-pointer transition select-none ${
                        col.selected
                          ? "bg-blue-950/40 border-blue-500/50 text-blue-200"
                          : "bg-slate-900/60 border-slate-800/80 text-slate-400 hover:border-slate-700"
                      }`}
                    >
                      <div className="flex items-center gap-2 truncate">
                        {col.selected ? (
                          <CheckSquare className="w-4 h-4 text-blue-400 flex-shrink-0" />
                        ) : (
                          <Square className="w-4 h-4 text-slate-600 flex-shrink-0" />
                        )}
                        <span className="font-mono truncate" title={col.name}>
                          {col.name}
                        </span>
                      </div>
                      {col.emptyOrPlaceholderCount > 0 && (
                        <span className="text-[10px] px-1.5 py-0.5 rounded bg-amber-950/80 text-amber-400 border border-amber-800/60 font-mono">
                          {col.emptyOrPlaceholderCount} blank
                        </span>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Input 2: Root / Fallback Web Link */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <Globe className="w-4 h-4 text-sky-400" /> 2. University Website Link / Portal URL
                </label>
                <button
                  type="button"
                  onClick={() => setShowDirectPaste(!showDirectPaste)}
                  className="text-xs text-blue-400 hover:text-blue-300 flex items-center gap-1"
                >
                  <ClipboardPaste className="w-3.5 h-3.5" />
                  <span>{showDirectPaste ? "Hide Direct Paste" : "Behind Login / Protected? Paste Directly"}</span>
                  {showDirectPaste ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                </button>
              </div>
              <input
                type="url"
                value={url}
                onChange={(e) => setUrl(e.target.value)}
                placeholder="https://university.edu/course-catalog or fallback portal link"
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-sky-500/50"
              />
            </div>

            {/* Direct Paste Accordion */}
            {showDirectPaste && (
              <div className="bg-slate-950/80 border border-blue-900/50 rounded-xl p-4 space-y-2">
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  If the portal is password-protected, paste copied text or HTML table content here:
                </p>
                <textarea
                  rows={4}
                  value={directContent}
                  onChange={(e) => setDirectContent(e.target.value)}
                  placeholder="Paste copied website text or table here..."
                  className="w-full text-xs bg-slate-900 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 font-mono"
                />
              </div>
            )}

            {/* Input 3: Optional Instruction */}
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <MessageSquareText className="w-4 h-4 text-purple-400" /> 3. Custom Instruction (Optional)
              </label>
              <textarea
                rows={2}
                value={customPrompt}
                onChange={(e) => setCustomPrompt(e.target.value)}
                placeholder="e.g. 'If tuition is shown in GBP, convert or specify currency explicitly. For overview, write 2 concise sentences.'"
                className="w-full text-xs bg-slate-950 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-purple-500/50 resize-none"
              />
            </div>

            {/* Error Alert */}
            {errorMsg && (
              <div className="bg-red-950/80 border border-red-800 text-red-300 px-4 py-3 rounded-xl text-xs flex items-center gap-3">
                <AlertCircle className="w-5 h-5 text-red-400 flex-shrink-0" />
                <span>{errorMsg}</span>
              </div>
            )}

            {/* Live Real-Time Progress Bar & Box Tracker */}
            {loading && columnProgress && (
              <div className="bg-slate-950 border border-blue-900/60 rounded-xl p-4 space-y-3">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-semibold text-blue-400 flex items-center gap-2">
                    <RefreshCw className="w-3.5 h-3.5 animate-spin" /> Verifying Column {columnProgress.colIndexNum}/{columnProgress.totalCols}: "{columnProgress.currentColName}"
                  </span>
                  <span className="font-mono text-emerald-400 font-bold">{columnProgress.percent}%</span>
                </div>
                {/* Visual Progress Bar */}
                <div className="w-full bg-slate-900 rounded-full h-2.5 overflow-hidden border border-slate-800">
                  <div
                    className="bg-gradient-to-r from-blue-500 via-indigo-500 to-emerald-400 h-2.5 rounded-full transition-all duration-300 ease-out"
                    style={{ width: `${columnProgress.percent}%` }}
                  />
                </div>
                <div className="flex items-center justify-between text-[11px] text-slate-400">
                  <span>Batch {columnProgress.currentBatch} of {columnProgress.totalBatches}</span>
                  <span className="text-amber-400 font-medium">Boxes Updated So Far: {columnProgress.updatedBoxes}</span>
                </div>

                {/* Live Box-by-Box Verification Stream with Diff Preview Link */}
                {liveVerificationFeed.length > 0 && (
                  <div className="space-y-1.5 pt-2">
                    <div className="flex items-center justify-between text-[10px] text-slate-400 font-semibold uppercase tracking-wider">
                      <span>Live Box Checks (Click to inspect diff):</span>
                      <span>Confidence</span>
                    </div>
                    <div className="bg-slate-900/90 rounded-lg p-2.5 space-y-1.5 text-[11px] font-mono max-h-36 overflow-y-auto">
                      {liveVerificationFeed.map((item, i) => (
                        <div
                          key={i}
                          onClick={() => setSelectedDiffItem(item)}
                          className="flex items-center justify-between gap-2 border-b border-slate-800/60 pb-1 cursor-pointer hover:bg-slate-800/50 p-1 rounded transition"
                        >
                          <div className="flex items-center gap-2 truncate max-w-xs">
                            <Eye className="w-3.5 h-3.5 text-blue-400 flex-shrink-0" />
                            <span className="text-slate-300 truncate">
                              Row {item.rowIndex + 1} • <strong className="text-indigo-300">{item.columnName}</strong>: {item.entityTitle}
                            </span>
                          </div>
                          <div className="flex items-center gap-1.5 flex-shrink-0">
                            <span
                              className={`text-[9px] px-1.5 py-0.5 rounded font-bold ${
                                item.status === "AUTOFILLED"
                                  ? "bg-amber-950 text-amber-400 border border-amber-800"
                                  : item.status === "DISCONTINUED"
                                  ? "bg-rose-950 text-rose-400 border border-rose-800"
                                  : item.status === "UPDATED"
                                  ? "bg-blue-950 text-blue-400 border border-blue-800"
                                  : "bg-emerald-950 text-emerald-400 border border-emerald-800"
                              }`}
                            >
                              {item.status}
                            </span>
                            <span
                              className={`text-[9px] px-1 py-0.5 rounded font-semibold ${
                                item.confidence === "HIGH"
                                  ? "text-emerald-400"
                                  : item.confidence === "LOW"
                                  ? "text-rose-400"
                                  : "text-amber-400"
                              }`}
                            >
                              {item.confidence}
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Run Button */}
            <button
              onClick={handleRunDynamicUpdate}
              disabled={loading}
              className="w-full py-4 px-6 rounded-xl font-semibold text-white bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 hover:from-blue-500 hover:to-purple-500 focus:outline-none focus:ring-2 focus:ring-blue-500/50 shadow-lg shadow-blue-600/25 transition disabled:opacity-50 flex items-center justify-center gap-2 text-sm"
            >
              {loading ? (
                <>
                  <RefreshCw className="w-5 h-5 animate-spin text-white" />
                  <span>{statusStep || "Processing..."}</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-5 h-5 text-white" />
                  <span>Run Dynamic Column Verification & Autofill</span>
                </>
              )}
            </button>
          </div>

          {/* Results & Download Panel */}
          {stats && downloadBlobUrl && (
            <div className="bg-emerald-950/30 border border-emerald-800/60 rounded-2xl p-6 shadow-xl space-y-6">
              <div className="flex items-center gap-3 text-emerald-400">
                <CheckCircle2 className="w-6 h-6" />
                <h3 className="font-bold text-lg">Verification & Autofill Complete!</h3>
              </div>

              {/* Stats Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Updated / Autofilled Boxes</span>
                  <p className="text-2xl font-black text-amber-400 mt-1">{stats.updatedBoxes}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Verified Current Boxes</span>
                  <p className="text-2xl font-black text-emerald-400 mt-1">{stats.verifiedBoxes}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Discontinued Courses</span>
                  <p className="text-2xl font-black text-rose-400 mt-1">{stats.discontinuedCount}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Webpages Cached</span>
                  <p className="text-2xl font-black text-sky-400 mt-1">{stats.scrapedPagesCount}</p>
                </div>
              </div>

              {/* Summary Notes */}
              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 text-xs text-slate-300 space-y-1">
                <span className="font-semibold text-slate-200">Execution Summary:</span>
                <p>{stats.notes}</p>
              </div>

              {/* Download Button */}
              <a
                href={downloadBlobUrl}
                download={downloadFileName}
                className="w-full py-4 px-6 rounded-xl font-semibold text-slate-900 bg-emerald-400 hover:bg-emerald-300 shadow-lg shadow-emerald-400/20 transition flex items-center justify-center gap-2 text-sm text-center"
              >
                <Download className="w-5 h-5 text-slate-900" />
                <span>Download Updated Excel File (with Verification_Audit Sheet)</span>
              </a>
            </div>
          )}
        </div>
      </main>

      {/* Side-by-Side Diff Preview Modal */}
      {selectedDiffItem && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-5 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <Eye className="w-5 h-5 text-blue-400" />
                <h3 className="font-bold text-base text-white">Box Verification Audit: {selectedDiffItem.columnName}</h3>
              </div>
              <button
                onClick={() => setSelectedDiffItem(null)}
                className="p-1 rounded-lg hover:bg-slate-800 text-slate-400 hover:text-white"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-1 text-xs">
              <span className="text-slate-400">Course / Record:</span>
              <p className="font-semibold text-indigo-300 text-sm">{selectedDiffItem.entityTitle}</p>
            </div>

            {/* Side-by-Side Comparison */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs font-mono">
              <div className="bg-slate-950/80 border border-slate-800 rounded-xl p-4 space-y-1.5">
                <span className="text-slate-400 uppercase text-[10px] tracking-wider font-semibold block">
                  Original Excel Box Value
                </span>
                <p className="text-slate-300 break-words line-through decoration-rose-500/70">
                  {selectedDiffItem.previousValue !== undefined && selectedDiffItem.previousValue !== null && String(selectedDiffItem.previousValue).trim() !== ""
                    ? String(selectedDiffItem.previousValue)
                    : "(Blank / Empty)"}
                </p>
              </div>

              <div className="bg-emerald-950/20 border border-emerald-800/60 rounded-xl p-4 space-y-1.5">
                <span className="text-emerald-400 uppercase text-[10px] tracking-wider font-semibold block">
                  Verified 2026 Value
                </span>
                <p className="text-emerald-200 break-words font-medium">
                  {selectedDiffItem.verifiedValue !== undefined ? String(selectedDiffItem.verifiedValue) : "(None)"}
                </p>
              </div>
            </div>

            {/* Status & Evidence */}
            <div className="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs space-y-2">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="text-slate-400">Status:</span>
                  <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-950 text-blue-400 border border-blue-800">
                    {selectedDiffItem.status}
                  </span>
                  <span className="text-slate-400 ml-2">Confidence:</span>
                  <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-950 text-emerald-400 border border-emerald-800">
                    {selectedDiffItem.confidence}
                  </span>
                </div>
                {selectedDiffItem.sourceUrl && selectedDiffItem.sourceUrl.startsWith("http") && (
                  <a
                    href={selectedDiffItem.sourceUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-blue-400 hover:text-blue-300 flex items-center gap-1 text-[11px]"
                  >
                    <span>View Page</span>
                    <ExternalLink className="w-3 h-3" />
                  </a>
                )}
              </div>

              {selectedDiffItem.reason && (
                <p className="text-slate-300 text-[11px] leading-relaxed">
                  <strong className="text-slate-200">Reason:</strong> {selectedDiffItem.reason}
                </p>
              )}

              {selectedDiffItem.evidenceSnippet && (
                <div className="bg-slate-900/80 p-2.5 rounded border border-slate-800 text-[11px] text-slate-300 font-mono">
                  <span className="text-slate-500 block text-[10px] uppercase">Source Evidence Excerpt:</span>
                  "{selectedDiffItem.evidenceSnippet}"
                </div>
              )}
            </div>

            <div className="flex justify-end">
              <button
                onClick={() => setSelectedDiffItem(null)}
                className="px-4 py-2 rounded-xl bg-slate-800 text-slate-200 hover:bg-slate-700 text-xs font-semibold"
              >
                Close Preview
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

"use client";

import React, { useState } from "react";
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
  ShieldAlert,
  ChevronDown,
  ChevronUp,
  ListOrdered,
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

export default function RecordSyncPage() {
  const [apiKeys, setApiKeys] = useState<string>("");
  const [selectedModel, setSelectedModel] = useState<string>("gemini-2.5-flash");
  const [url, setUrl] = useState<string>("");
  const [directContent, setDirectContent] = useState<string>("");
  const [showDirectPaste, setShowDirectPaste] = useState<boolean>(false);
  const [customPrompt, setCustomPrompt] = useState<string>("");
  const [file, setFile] = useState<File | null>(null);

  const [loading, setLoading] = useState<boolean>(false);
  const [statusStep, setStatusStep] = useState<string>("");
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [scraperWarning, setScraperWarning] = useState<string | null>(null);

  const [downloadBlobUrl, setDownloadBlobUrl] = useState<string | null>(null);
  const [downloadFileName, setDownloadFileName] = useState<string>("");
  const [stats, setStats] = useState<{
    updatedCount: number;
    newCount: number;
    modelUsed: string;
    keyUsed: string;
    notes: string;
    pagesScraped: string[];
  } | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const handleRunUpdate = async () => {
    setErrorMsg(null);
    setScraperWarning(null);
    setDownloadBlobUrl(null);
    setStats(null);

    if (!apiKeys.trim()) {
      setErrorMsg("Please provide at least one Google Gemini API Key.");
      return;
    }
    if (!file) {
      setErrorMsg("Please upload your Excel file (.xlsx or .xls).");
      return;
    }
    if (!url.trim() && !directContent.trim()) {
      setErrorMsg("Please provide either a university webpage URL or paste the webpage text directly.");
      return;
    }

    setLoading(true);
    setStatusStep("Step 1/3: Deep Crawling University Root Page & Subpages...");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("url", url);
      formData.append("directContent", directContent);
      formData.append("apiKeys", apiKeys);
      formData.append("model", selectedModel);
      formData.append("customPrompt", customPrompt);

      setTimeout(() => {
        setStatusStep("Step 2/3: Fuzzy Matching Course Codes & Auto-filling Blanks...");
      }, 3000);

      setTimeout(() => {
        setStatusStep("Step 3/3: Reconciling Outdated Values & Applying Yellow Cell Styling...");
      }, 7000);

      const response = await fetch("/api/update-records", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || `Server returned status ${response.status}`);
      }

      const blob = await response.blob();
      const blobUrl = URL.createObjectURL(blob);
      setDownloadBlobUrl(blobUrl);
      setDownloadFileName(`Updated_2026_${file.name}`);

      const updatedCount = parseInt(response.headers.get("X-Updated-Count") || "0", 10);
      const newCount = parseInt(response.headers.get("X-New-Count") || "0", 10);
      const modelUsed = response.headers.get("X-Model-Used") || selectedModel;
      const keyUsed = response.headers.get("X-Key-Used") || "1";
      const rawNotes = response.headers.get("X-Summary-Notes") || "";
      const notes = rawNotes ? decodeURIComponent(rawNotes) : "Reconciliation completed.";

      const rawWarning = response.headers.get("X-Scraper-Warning") || "";
      if (rawWarning) {
        setScraperWarning(decodeURIComponent(rawWarning));
      }

      let pagesScraped: string[] = [];
      const rawPages = response.headers.get("X-Pages-Scraped") || "";
      if (rawPages) {
        try {
          pagesScraped = JSON.parse(decodeURIComponent(rawPages));
        } catch {
          pagesScraped = [];
        }
      }

      setStats({ updatedCount, newCount, modelUsed, keyUsed, notes, pagesScraped });
    } catch (err: any) {
      setErrorMsg(err.message || "An unexpected error occurred during processing.");
    } finally {
      setLoading(false);
      setStatusStep("");
    }
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans">
      {/* Top Header Bar */}
      <header className="border-b border-slate-800 bg-slate-900/60 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center shadow-lg shadow-blue-500/20">
              <Sparkles className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="font-bold text-lg text-white leading-tight flex items-center gap-2">
                RecordSync <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20">Resilient 2026 Edition</span>
              </h1>
              <p className="text-xs text-slate-400">Intelligent Academic Excel Updater with Deep Scraping & Direct Bypass</p>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-xs text-emerald-400 bg-emerald-950/80 border border-emerald-800/60 px-3 py-1 rounded-full font-medium flex items-center gap-1.5">
              <Zap className="w-3.5 h-3.5" /> 100% Free Google AI Studio API
            </span>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8 grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Column: Config Panel */}
        <div className="lg:col-span-4 space-y-6">
          {/* API Key Box */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-4">
            <div className="flex items-center justify-between">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <Key className="w-4 h-4 text-blue-400" /> Google Gemini API Key(s)
              </label>
              <span className="text-[10px] text-slate-400 bg-slate-800 px-2 py-0.5 rounded">Multi-Key Supported</span>
            </div>
            <textarea
              rows={3}
              value={apiKeys}
              onChange={(e) => setApiKeys(e.target.value)}
              placeholder="Paste your Gemini API key(s) here. Separate multiple keys with commas for auto-failover on quota limits..."
              className="w-full text-xs bg-slate-950 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 resize-none font-mono"
            />
            <div className="bg-slate-950/80 p-3 rounded-xl border border-slate-800 text-[11px] text-slate-400 space-y-1.5">
              <p className="flex items-center gap-1.5 text-blue-400 font-semibold">
                <Zap className="w-3.5 h-3.5" /> Dual Auto-Switch Failover
              </p>
              <p>
                • <strong>Key Failover:</strong> If key #1 hits quota limit, key #2 is tried instantly.<br />
                • <strong>Model Failover:</strong> If model RPM/RPD cap is reached, it switches models automatically!
              </p>
            </div>
          </div>

          {/* AI Model Selection Dropdown */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-xl space-y-3">
            <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
              <Cpu className="w-4 h-4 text-indigo-400" /> Primary Gemini Model
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
            {MODEL_OPTIONS.find((m) => m.id === selectedModel) && (
              <p className="text-[11px] text-slate-400 bg-slate-950/60 p-2.5 rounded-lg border border-slate-800/80">
                {MODEL_OPTIONS.find((m) => m.id === selectedModel)?.desc}
              </p>
            )}
          </div>
        </div>

        {/* Right Column: Execution Form & Results */}
        <div className="lg:col-span-8 space-y-6">
          
          {/* Main Action Card */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl space-y-5">
            
            {/* Input 1: File Upload */}
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" /> 1. Upload Excel Sheet (.xlsx / .xls)
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
                      <p className="text-xs text-slate-400">{(file.size / 1024).toFixed(1)} KB • Ready to reconcile</p>
                    </div>
                  ) : (
                    <div>
                      <p className="text-sm font-medium text-slate-300">Drag & drop your course/student sheet or browse</p>
                      <p className="text-xs text-slate-500">Supports up to 400+ course rows with blank or outdated values</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Input 2: Web Link */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <Globe className="w-4 h-4 text-sky-400" /> 2. Target University Web Link
                </label>
                <button
                  type="button"
                  onClick={() => setShowDirectPaste(!showDirectPaste)}
                  className="text-xs text-blue-400 hover:text-blue-300 flex items-center gap-1"
                >
                  <ClipboardPaste className="w-3.5 h-3.5" />
                  <span>{showDirectPaste ? "Hide Direct Paste" : "Behind Login / Protected? Paste Web Text Directly"}</span>
                  {showDirectPaste ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                </button>
              </div>
              <div className="relative">
                <input
                  type="url"
                  value={url}
                  onChange={(e) => setUrl(e.target.value)}
                  placeholder="https://university.edu/course-catalog-2026 or results page"
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-sky-500/50 pl-10"
                />
                <Globe className="w-4 h-4 text-slate-500 absolute left-3.5 top-3.5" />
              </div>
            </div>

            {/* Direct Paste Accordion (Ultimate Scraper Bypass) */}
            {showDirectPaste && (
              <div className="bg-slate-950/80 border border-blue-900/50 rounded-xl p-4 space-y-2">
                <div className="flex items-center gap-2 text-blue-400 text-xs font-semibold">
                  <ClipboardPaste className="w-4 h-4" /> 100% Guaranteed Bypass: Direct University Webpage Text / HTML Paste
                </div>
                <p className="text-[11px] text-slate-400 leading-relaxed">
                  If the university portal is behind a student/teacher password login, Cloudflare captcha, or dynamic React tabs, simply open the page in your browser, press <strong>Ctrl + A</strong>, <strong>Ctrl + C</strong>, and paste the text/table content here:
                </p>
                <textarea
                  rows={4}
                  value={directContent}
                  onChange={(e) => setDirectContent(e.target.value)}
                  placeholder="Paste copied university text or raw HTML here..."
                  className="w-full text-xs bg-slate-900 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 font-mono"
                />
              </div>
            )}

            {/* Input 3: Teacher's Custom Prompt Input */}
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                  <MessageSquareText className="w-4 h-4 text-purple-400" /> 3. Teacher Custom Instruction (Optional)
                </label>
                <span className="text-[10px] text-purple-300 bg-purple-950/80 border border-purple-800/50 px-2 py-0.5 rounded">Unique Task Prompt</span>
              </div>
              <textarea
                rows={2}
                value={customPrompt}
                onChange={(e) => setCustomPrompt(e.target.value)}
                placeholder="e.g. 'Autofill all missing Course Credits and Prerequisites. If a course is discontinued in 2026, set Status to Discontinued.' or 'Only update instructor names and emails.'"
                className="w-full text-xs bg-slate-950 border border-slate-800 rounded-xl p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-purple-500/50 resize-none"
              />
            </div>

            {/* Scraper Diagnostic Warning */}
            {scraperWarning && (
              <div className="bg-amber-950/80 border border-amber-800 text-amber-300 px-4 py-3 rounded-xl text-xs flex items-center gap-3">
                <ShieldAlert className="w-5 h-5 text-amber-400 flex-shrink-0" />
                <span>{scraperWarning} (Tip: If data was missed, use the 'Paste Web Text Directly' option above!)</span>
              </div>
            )}

            {/* Error Alert */}
            {errorMsg && (
              <div className="bg-red-950/80 border border-red-800 text-red-300 px-4 py-3 rounded-xl text-xs flex items-center gap-3">
                <AlertCircle className="w-5 h-5 text-red-400 flex-shrink-0" />
                <span>{errorMsg}</span>
              </div>
            )}

            {/* Run Button */}
            <button
              onClick={handleRunUpdate}
              disabled={loading}
              className="w-full py-4 px-6 rounded-xl font-semibold text-white bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 hover:from-blue-500 hover:to-purple-500 focus:outline-none focus:ring-2 focus:ring-blue-500/50 shadow-lg shadow-blue-600/25 transition disabled:opacity-50 flex items-center justify-center gap-2 text-sm"
            >
              {loading ? (
                <>
                  <RefreshCw className="w-5 h-5 animate-spin text-white" />
                  <span>{statusStep || "Processing Records..."}</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-5 h-5 text-white" />
                  <span>Run Intelligent Auto-Fill & Update</span>
                </>
              )}
            </button>
          </div>

          {/* Results & Download Panel */}
          {stats && downloadBlobUrl && (
            <div className="bg-emerald-950/30 border border-emerald-800/60 rounded-2xl p-6 shadow-xl space-y-6">
              <div className="flex items-center gap-3 text-emerald-400">
                <CheckCircle2 className="w-6 h-6" />
                <h3 className="font-bold text-lg">Reconciliation Complete!</h3>
              </div>

              {/* Stats Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Autofilled / Updated Rows</span>
                  <p className="text-2xl font-black text-amber-400 mt-1">{stats.updatedCount}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">New Discovered Rows</span>
                  <p className="text-2xl font-black text-emerald-400 mt-1">{stats.newCount}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">Model Used</span>
                  <p className="text-xs font-bold text-indigo-300 mt-2 truncate">{stats.modelUsed}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">API Key Account</span>
                  <p className="text-xs font-bold text-blue-300 mt-2">Account #{stats.keyUsed}</p>
                </div>
              </div>

              {/* Scraped Pages List Card */}
              {stats.pagesScraped && stats.pagesScraped.length > 0 && (
                <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-4 space-y-2">
                  <div className="flex items-center gap-2 text-xs font-semibold text-sky-400">
                    <ListOrdered className="w-4 h-4" />
                    <span>Pages Scraped ({stats.pagesScraped.length}):</span>
                  </div>
                  <ul className="text-[11px] text-slate-300 space-y-1.5 font-mono max-h-40 overflow-y-auto pr-2">
                    {stats.pagesScraped.map((pageUrl, idx) => (
                      <li key={idx} className="flex items-center justify-between bg-slate-950 p-2 rounded-lg border border-slate-800/80">
                        <span className="truncate max-w-md">{pageUrl}</span>
                        {pageUrl.startsWith("http") && (
                          <a
                            href={pageUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="text-blue-400 hover:text-blue-300 ml-2 flex-shrink-0"
                          >
                            <ExternalLink className="w-3.5 h-3.5" />
                          </a>
                        )}
                      </li>
                    ))}
                  </ul>
                </div>
              )}

              {/* Summary Notes */}
              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 text-xs text-slate-300 space-y-1">
                <span className="font-semibold text-slate-200">AI Reconciliation Summary:</span>
                <p>{stats.notes}</p>
              </div>

              {/* Download Button */}
              <a
                href={downloadBlobUrl}
                download={downloadFileName}
                className="w-full py-4 px-6 rounded-xl font-semibold text-slate-900 bg-emerald-400 hover:bg-emerald-300 shadow-lg shadow-emerald-400/20 transition flex items-center justify-center gap-2 text-sm text-center"
              >
                <Download className="w-5 h-5 text-slate-900" />
                <span>Download Updated Excel File (.xlsx)</span>
              </a>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}

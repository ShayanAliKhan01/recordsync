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
  const [file, setFile] = useState<File | null>(null);

  const [loading, setLoading] = useState<boolean>(false);
  const [statusStep, setStatusStep] = useState<string>("");
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const [downloadBlobUrl, setDownloadBlobUrl] = useState<string | null>(null);
  const [downloadFileName, setDownloadFileName] = useState<string>("");
  const [stats, setStats] = useState<{
    updatedCount: number;
    newCount: number;
    modelUsed: string;
    keyUsed: string;
    notes: string;
  } | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const handleRunUpdate = async () => {
    setErrorMsg(null);
    setDownloadBlobUrl(null);
    setStats(null);

    if (!apiKeys.trim()) {
      setErrorMsg("Please provide at least one Google Gemini API Key.");
      return;
    }
    if (!file) {
      setErrorMsg("Please upload your legacy Excel file (.xlsx or .xls).");
      return;
    }
    if (!url.trim() || !url.startsWith("http")) {
      setErrorMsg("Please enter a valid university URL starting with http:// or https://");
      return;
    }

    setLoading(true);
    setStatusStep("Step 1/3: Extracting University Web Content & Table Data...");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("url", url);
      formData.append("apiKeys", apiKeys);
      formData.append("model", selectedModel);

      setTimeout(() => {
        setStatusStep("Step 2/3: AI Entity Matching (With Key & Model Auto-Switch Failover)...");
      }, 3000);

      setTimeout(() => {
        setStatusStep("Step 3/3: Styling Updated Cells with Soft Yellow Highlight...");
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

      setStats({ updatedCount, newCount, modelUsed, keyUsed, notes });
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
                RecordSync <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20">Vercel Ready 2026</span>
              </h1>
              <p className="text-xs text-slate-400">Intelligent Academic Excel & Web Record Reconciliation</p>
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
                1. <strong>Key Failover:</strong> Automatically switches from Key #1 to Key #2 if quota exhausts.<br />
                2. <strong>Model Auto-Switch:</strong> Automatically switches models (e.g. Flash → Flash-Lite) if model-specific RPM/RPD cap is reached!
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
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl space-y-6">
            
            {/* Input 1: File Upload */}
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-emerald-400" /> 1. Upload Legacy Excel File (.xlsx)
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
                      <p className="text-xs text-slate-400">{(file.size / 1024).toFixed(1)} KB • Click or drag to replace</p>
                    </div>
                  ) : (
                    <div>
                      <p className="text-sm font-medium text-slate-300">Drag & drop old mark sheet or click to browse</p>
                      <p className="text-xs text-slate-500">Supports .xlsx and .xls formats</p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Input 2: Web Link */}
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-200 flex items-center gap-2">
                <Globe className="w-4 h-4 text-sky-400" /> 2. Target 2026 University Web Link
              </label>
              <div className="relative">
                <input
                  type="url"
                  value={url}
                  onChange={(e) => setUrl(e.target.value)}
                  placeholder="https://university.edu/course-results/2026"
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-sm text-slate-200 focus:outline-none focus:ring-2 focus:ring-sky-500/50 pl-10"
                />
                <Globe className="w-4 h-4 text-slate-500 absolute left-3.5 top-3.5" />
              </div>
            </div>

            {/* Error Message Alert */}
            {errorMsg && (
              <div className="bg-red-950/80 border border-red-800 text-red-300 px-4 py-3 rounded-xl text-xs flex items-center gap-3">
                <AlertCircle className="w-5 h-5 text-red-400 flex-shrink-0" />
                <span>{errorMsg}</span>
              </div>
            )}

            {/* Action Button */}
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
                  <span>Run Intelligent 2026 Update</span>
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
                  <span className="text-xs text-slate-400 font-medium">Updated Cells</span>
                  <p className="text-2xl font-black text-amber-400 mt-1">{stats.updatedCount}</p>
                </div>
                <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4">
                  <span className="text-xs text-slate-400 font-medium">New 2026 Rows</span>
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

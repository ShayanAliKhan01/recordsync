# RecordSync 2026 - Academic Excel & Web Record Updater

RecordSync is a Next.js (App Router) + Tailwind CSS web application built for university teachers to reconcile legacy Excel records (e.g. 2023/2024 data) with fresh 2026 university web data.

---

## 🌟 Key Features

1. **Vercel Native Deployment**
   - Built on Next.js 14+ (App Router) for zero-configuration, 100% free hosting on Vercel.

2. **Dual Failover Strategy (Zero Quota Exhaustion)**
   - **Multi-API Key Failover**: Input multiple free Google AI Studio API keys (comma-separated). If Key #1 hits rate limits (HTTP 429), it automatically switches to Key #2!
   - **Automatic Model Auto-Switch**: If a specific model reaches its daily or minute limit (e.g., `gemini-2.5-flash` at 20 RPD), the system automatically switches models in real-time (`gemini-3.5-flash-lite` -> `gemini-3.1-flash-lite` -> `gemini-3.5-flash` -> `gemini-3.7-flash`).

3. **Dynamic University Web Scraping API Route**
   - Cleans HTML text, removes navigation metadata, and extracts structured tables dynamically.

4. **Excel Visual Cell Highlighting (`#FFF2CC`)**
   - Powered by `exceljs`. Applies soft yellow background highlights to changed or newly added cells.

---

## 🚀 How to Run Locally

```bash
# 1. Navigate to project directory
cd recordsync

# 2. Install dependencies
npm install

# 3. Start local development server
npm run dev
```
Open `http://localhost:3000` in your browser.

---

## 🌐 How to Deploy to Vercel (100% Free)

1. Push the `recordsync` project folder to your GitHub repository.
2. Sign in to [Vercel](https://vercel.com/) and click **Add New Project**.
3. Import your repository. Vercel automatically detects Next.js.
4. Click **Deploy**. Done!

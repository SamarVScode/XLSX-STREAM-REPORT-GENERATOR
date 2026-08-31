# ⚡ EI Stream Report Server (Render Edition)

High-performance, ultra-low-memory asynchronous report generation service for processing large Excel spreadsheets (`.xlsx`, `.xlsb`, `.csv`, `.ods`) from Google Drive / Google Sheets on **Render.com Free Tier (512 MB RAM)**.

---

## 🌟 Key Highlights

* **In-Flight Stream Processing**: Uses Rust-based `python-calamine` and XML `iterparse` to read data row-by-row with **flat ~30 MB – 35 MB RAM** consumption.
* **350 MB+ File Support on Free Tiers**: Processes multi-million row files on standard 512 MB instances without Out-of-Memory (OOM) crashes.
* **Dual Format Engine**: Native support for both standard `.xlsx` (OpenXML) and compact binary `.xlsb` (BIFF12).
* **Automatic Disk Reclamation**: Downloaded raw input files are **immediately deleted** from disk as soon as report summaries are generated.
* **Google Apps Script (GAS) Compatible**: Non-blocking async queue with time-triggered polling (`/convert-async` $\rightarrow$ `/job/{id}` $\rightarrow$ `/job/{id}/result`).
* **Built-in Web Test Bench**: Visit `/test` in any browser to submit links and view real-time status.

---

## 📊 Supported Report Generators

1. **EI Summary Report** (`ei`)
2. **Forward Pendency Report** (`forward_pendency`)
3. **Reverse Pendency Report** (`reverse_pendency`)
4. **Conversion Summary Report** (`conversion` — Sameday / D-1)
5. **NPS Report** (`nps`)
6. **SCM TAT 24Hrs Performance Report** (`tat` / `scm_tat`)
7. **VMS Adherence Report** (`vms_adherence`)
8. **2nd Attempt Adherence Report** (`second_attempt_adherence`)
9. **EOB Priority Report** (`eob`)
10. **Untraceable Report** (`untraceable`)

---

## 🚀 Quick Start (Local Run)

```bash
# Install dependencies
pip install -r requirements.txt

# Run test suite
pytest test_server.py -v

# Start local server on port 8000
python main.py
```

---

## ☁️ Deploy to Render.com Free Tier

1. Connect your repository to **[Render.com](https://dashboard.render.com/)**.
2. Create a **New Web Service**:
   * **Root Directory**: `ei_stream_server` (or `.` if deploying root repository)
   * **Build Command**: `pip install -r requirements.txt`
   * **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
   * **Instance Type**: **Free** (512 MB RAM)
3. Set Environment Variables:
   * `API_KEY`: `OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ` (or your custom secret)
   * `MAX_CONCURRENT_JOBS`: `1`

---

## 📡 API Reference

* `GET /health` — Health check & active jobs metrics (requires `X-API-KEY` header).
* `GET /test` — Interactive web test bench UI.
* `GET /convert-async?drive_url=...&report_type=...` — Submit async processing job.
* `GET /job/{job_id}` — Poll job status (`processing` \| `done` \| `error`).
* `GET /job/{job_id}/result` — Download generated `.xlsx` report file.

# Render.com Free Tier Hosting Guide — EI Stream Service

This guide explains how to deploy **`ei_stream_server`** on **Render.com Free Tier (512 MB RAM)**.

---

## 🌟 Why This Server Runs on 512 MB RAM Free Tier
* **In-Flight Stream Engine**: Uses Rust-based `python-calamine` and XML `iterparse` to read `.xlsx` and `.xlsb` row-by-row.
* **Flat Memory**: Constant **~30 MB – 35 MB RAM** consumption even when processing 350 MB+ files.
* **Auto-Cleanup**: The downloaded 350 MB input file is **immediately deleted** from disk as soon as the summary is written.

---

## 🚀 Step-by-Step Deployment on Render.com

### Step 1: Push to GitHub
Commit and push the `ei_stream_server` directory to your GitHub repository.

### Step 2: Create Web Service on Render
1. Log in to [dashboard.render.com](https://dashboard.render.com/).
2. Click **New +** $\rightarrow$ **Web Service**.
3. Select your GitHub repository.

### Step 3: Configure Settings
* **Name**: `ei-stream-service` (or your preferred name)
* **Region**: `Singapore` (or nearest to your users)
* **Branch**: `main`
* **Root Directory**: `ei_stream_server` (or `.` if deploying from root of repository)
* **Runtime**: `Python 3`
* **Build Command**: `pip install -r requirements.txt`
* **Start Command**: `uvicorn main:app --host 0.0.0.0 --port $PORT`
* **Instance Type**: **Free** (512 MB RAM, 0.5 CPU)

### Step 4: Set Environment Variables
Under **Environment Variables**, add:
* `API_KEY`: `OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ` (or your custom secret key)
* `MAX_CONCURRENT_JOBS`: `1`
* `PYTHON_VERSION`: `3.11.0`

### Step 5: Click "Deploy Web Service"
Render will build and launch your service. In ~2 minutes, your service will be live at:
`https://<your-service-name>.onrender.com`

---

## 🔗 Connecting to Google Apps Script (GAS)
In [`ei_report_trigger/Code.js`](file:///C:/Users/User/Desktop/server/ei_report_trigger/Code.js), update:
```javascript
var SERVER_URL = "https://your-service-name.onrender.com";
var API_KEY    = "OoV81VZ6ugIQ5qu_JNKfDM0jEp0SQyhpuZMaPTv5BbQ";
```

---

## 🧪 Interactive Test Bench
Visit `https://<your-service-name>.onrender.com/test` in your browser to submit Drive URLs and monitor live job progress interactively.

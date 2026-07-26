# 🚀 Cold Email Automation Engine (Google Workspace)

This repository contains a lightweight, high-performance cold email execution engine built entirely within **Google Apps Script** and the **Gemini 3.1 Flash Lite API**. It converts static rows of potential leads into deeply contextualized, human-written emails drafted directly into your Gmail Drafts folder—with your resume PDF already natively attached.

The pipeline executes the **Cold Mail Mastery Framework**—combining extreme structural brevity with hyper-targeted company observation anchors and strict anti-AI lexical filters.

---

## ✨ Key Features & Upgrades

* **Native Google Drive Injection:** Automatically fetches your physical Resume PDF from Drive and attaches it securely to every generated Gmail draft.
* **Run-Level Memory (Anti-Monotony):** The engine tracks every opening hook style, quantified stat, and highlighted project used during a batch run. It explicitly bans reusing them in subsequent emails to ensure 100% variance.
* **Granular Project-Matching Matrix:** Replaces broad domain buckets with a highly specific 7-tier matrix (Media, SaaS, E-commerce, AI/ML, Hardware/CV, FinTech, Enterprise/ERP) to map your background flawlessly without forcing unrelated metrics.
* **Personal Context Prioritization:** Supports an optional column for recent LinkedIn posts or shared backgrounds, instructing the AI to anchor the hook on *human* events rather than static company descriptions.
* **Deterministic Layout & Rotation:** Enforces a rigid two-block output structure for perfect mobile readability and mathematically rotates subject line formulas (A/B/C) to prevent repetitive phrasing.
* **Execution Deduplication:** Built-in safeguards actively track processed email/company pairs in memory to prevent duplicate drafts from being generated if a row appears twice.
* **UI-Driven Control System:** Tick checkboxes next to specific companies in your Google Sheet and click the custom top-menu button to process only those rows.

---

## 🗂️ Project Modular Architecture

The repository environment structure inside the Apps Script workspace follows a decoupled architecture pattern:

```text
├── Code.gs             # Runtime engine logic, framework rules, run-memory arrays, and regex parsing
└── Resume.json.html    # Profile database layer containing structured JSON object properties

```

To modify your background experiences or projects over time, update the JSON object inside the `Resume.json.html` file. The execution engine updates its contextual mapping automatically upon the next trigger.

---

## 🛠️ Google Sheets Data Structure

The script reads headers dynamically. Ensure your Google Sheet contains these exact columns (spelling matters) in Row 1.

| Column Header | Description |
| --- | --- |
| `Name` | Recruiter or Executive's First/Last Name. |
| `Email` | Direct professional email address. |
| `Title` | Recruiter / Executive Job Title. |
| `Company` | Registered name of the company. |
| `Industry` | Vertical market categorization (e.g., SaaS, FinTech). |
| `Company Description` | Raw text detailing what the company builds or services. |
| `Personal Context` | **(Optional but Recommended)** Paste a recent LinkedIn post, company news, or shared ground here. If present, the AI prioritizes this for a highly authentic hook. |
| `Status` | **(Crucial Setup)** Used by the engine to track and skip processed rows (marks as "Drafted"). |
| `Draft Now` | **(Crucial Setup)** Insert **Checkboxes** here (`Insert > Checkbox`). Tick these to select rows for processing. |

*(Note: You can include other columns like SNo or Website anywhere in the sheet; the script maps the required headers dynamically by name).*

---

## 🚀 Execution & Setup Guide

### Step 1: Obtain your Keys and IDs

1. **Gemini API Key:** Navigate to [Google AI Studio](https://aistudio.google.com/), log in, click **"Get API Key"**, and generate a new free-tier key.
2. **Google Drive File ID:** Open your Google Drive, right-click your Resume PDF, and click **"Copy Link"**. Extract the long string of random characters between `/d/` and `/view`. This is your File ID.

### Step 2: Inject the Code into Google Sheets

1. Open your master outreach Google Sheet.
2. Go to **Extensions** > **Apps Script**.
3. Clear out any placeholder code inside the `Code.gs` file and paste the new `Code.gs` provided in this repository.
4. Replace `YOUR_GEMINI_API_KEY_HERE` with your AI Studio key.
5. Replace `YOUR_GOOGLE_DRIVE_FILE_ID_HERE` with your PDF's File ID.
6. In the left panel, click the **`+` (Add a file)** icon and select **HTML**. Name it exactly **`Resume.json`**.
7. Paste your JSON profile context into `Resume.json.html` and click **Save**.

### Step 3: Run the Engine & Authorize

1. Refresh your Google Sheet in your web browser.
2. A new custom menu called **`🚀 Cold Email Engine`** will appear in the top toolbar.
3. The first time you run it, Google will prompt an **"Authorization Required"** window to allow access to Sheets, Gmail, and Drive. Click **Review Permissions** > select your account > **Advanced** > **Go to Untitled project (unsafe)** > **Allow**.

---

## 🚦 Daily Operation Workflow

```text
[ Tick Checkboxes ] ──> [ Click Custom Menu ] ──> [ Gmail Drafts Folder ]
                                                              │
                                                  (Review & Click Send!)
                                            *Resume is auto-attached natively*

```

1. **Select Targets:** Scroll through your spreadsheet and tick the checkboxes in the **Draft Now** column for the companies you want to contact today. If you found a great LinkedIn post for a lead, paste it into the **Personal Context** column first.
2. **Execute:** Click **`🚀 Cold Email Engine`** > **`Generate Drafts for Ticked Rows`**.
3. **Automated State Tracking:** The script processes the targeted rows applying run-level memory, drafts them into Gmail with your PDF attached, updates the sheet to `Drafted`, and resets the checkboxes.
4. **Review & Dispatch:** Open your Gmail Drafts folder. Review the layout, ensure the PDF is present, and hit **Send**.
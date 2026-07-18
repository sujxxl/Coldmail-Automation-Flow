Here is the complete, fully updated `README.md` for your repository. It includes the latest architectural changes, the UI-driven checkbox system, and the decoupled JSON context configuration.

---

# 🚀 Cold Email Automation Engine (Google Workspace)

This repository contains a lightweight, high-performance cold email execution engine built entirely within **Google Apps Script** and the **Gemini API** (free tier). It converts static rows of potential leads into deeply contextualized, human-written emails drafted directly into your Gmail Drafts folder.

The pipeline executes the **Cold Mail Mastery Framework**—combining extreme structural brevity with hyper-targeted company observation anchors and strict anti-AI lexical filters.

---

## ✨ Features

* **Zero-Cost Operation:** Runs entirely on Google Apps Script (free) and the Gemini 2.5 Flash API (free tier).
* **Decoupled Profile Context:** Your resume data is isolated in a separate `Resume.json` file. Update your projects once, and the engine automatically parses the newest data.
* **UI-Driven Control System:** No need to run bulk scripts blindly. Tick checkboxes next to specific companies in your Google Sheet and click the custom top-menu button to process only those rows.
* **Anti-AI Detection Filter:** Hardcoded lexical constraints strip out robotic phrases (e.g., *"hope this email finds you well"*, *"esteemed organization"*).
* **Dynamic Hook Integration:** Evaluates the company's industry and operational description to draft a unique opening observation.
* **Human-in-the-Loop Safeguard:** Pushes outputs to your **Gmail Drafts folder**, allowing you to maintain final editorial control and attach your resume manually before sending.

---

## 🗂️ Project Modular Architecture

The repository environment structure inside the Apps Script workspace follows a decoupled architecture pattern:

```text
├── Code.gs             # Runtime engine logic, framework rules, API calls, and UI menus
└── Resume.json.html    # Profile database layer containing structured JSON object properties

```

To modify your background experiences or projects over time, update the `technical_pillars` values inside the `Resume.json` file. The core execution engine naturally updates its content parsing automatically upon the next trigger.

---

## 🛠️ Google Sheets Data Structure

To ensure the execution script matches your data indices flawlessly, verify your Google Sheet contains these exact columns starting at **Column A**:

| Column | Header Name | Description |
| --- | --- | --- |
| **A** | `SNo` | Serial Number. |
| **B** | `Name` | Recruiter or Executive's First/Last Name. |
| **C** | `Email` | Direct professional email address. |
| **D** | `Title` | Recruiter / Executive Job Title. |
| **E** | `Company` | Registered name of the company. |
| **F** | `Website` | Link to the corporate domain. |
| **G** | `Industry` | Vertical market categorization (e.g., SaaS, FinTech). |
| **H** | `Company Description` | Raw text detailing what the company builds or services. |
| **I** | `Status` | **(Crucial Setup)** Used by the engine to track and skip processed rows (marks as "Drafted"). |
| **J** | `Draft Now` | **(Crucial Setup)** Insert **Checkboxes** here (`Insert > Checkbox`). Tick these to select rows for processing. |

---

## 🚀 Execution & Setup Guide

### Step 1: Obtain a Free Gemini API Key

1. Navigate to [Google AI Studio](https://aistudio.google.com/).
2. Log in with your standard Google Account.
3. Click **"Get API Key"** in the sidebar, create a new key, and save it securely.

### Step 2: Inject the Code into Google Sheets

1. Open your master outreach Google Sheet.
2. Go to **Extensions** > **Apps Script**.
3. Clear out any placeholder code inside the `Code.gs` file and paste the `Code.gs` provided in this repository.
4. Replace the variable `YOUR_GEMINI_API_KEY_HERE` at the top of the script with your actual key.
5. In the left panel, click the **`+` (Add a file)** icon and select **HTML**. Name it exactly **`Resume.json`**.
6. Paste your JSON profile context into `Resume.json.html` and click **Save**.

### Step 3: Run the Engine & Authorize

1. Refresh your Google Sheet in your web browser.
2. A new custom menu called **`🚀 Cold Email Engine`** will appear in the top toolbar.
3. The first time you run it, Google will prompt an **"Authorization Required"** window. Click **Review Permissions** > select your account > **Advanced** > **Go to Untitled project (unsafe)** > **Allow**.

---

## 🚦 Daily Operation Workflow

```
[ Tick Checkboxes ] ──> [ Click Custom Menu ] ──> [ Gmail Drafts Folder ]
                                                           │
                                                (Review, Attach PDF, Send)

```

1. **Select Targets:** Scroll through your spreadsheet and tick the checkboxes in **Column J** for the companies you want to contact today.
2. **Execute:** Click **`🚀 Cold Email Engine`** > **`Generate Drafts for Ticked Rows`**.
3. **Automated State Tracking:** The script will process the targeted rows, instantly draft them into your Gmail, set `Column I` to `Drafted`, and reset the checkboxes to `FALSE`.
4. **Review & Dispatch:** Open your Gmail Drafts folder. Quickly review the copy, attach your latest PDF resume, and hit **Send**.
/**
 * Automatically creates a custom action menu in the Google Sheets toolbar upon load.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Cold Email Engine')
    .addItem('Generate Drafts for Ticked Rows', 'processSelectedCheckboxes')
    .addToUi();
}

/**
 * Iterates through rows, maps headers dynamically, validates checkbox states, 
 * parses external JSON resume context, processes payloads via Gemini 3.1 Flash Lite API, and updates states.
 */
function processSelectedCheckboxes() {
  const API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // <--- Insert your free Google AI Studio API Key
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Dynamic header position lookup map
  const colIndex = {
    name: headers.findIndex(h => String(h).trim().toLowerCase() === 'name'),
    email: headers.findIndex(h => String(h).trim().toLowerCase() === 'email'),
    title: headers.findIndex(h => String(h).trim().toLowerCase() === 'title'),
    company: headers.findIndex(h => String(h).trim().toLowerCase() === 'company'),
    industry: headers.findIndex(h => String(h).trim().toLowerCase() === 'industry'),
    description: headers.findIndex(h => String(h).trim().toLowerCase() === 'company description'),
    status: headers.findIndex(h => String(h).trim().toLowerCase() === 'status'),
    checkbox: headers.findIndex(h => String(h).trim().toLowerCase() === 'draft now')
  };
  
  // Guard clause checking for structural spreadsheet configurations
  if (colIndex.email === -1 || colIndex.company === -1 || colIndex.status === -1 || colIndex.checkbox === -1) {
    SpreadsheetApp.getUi().alert("Error: Critical column headers missing! Ensure your sheet has 'Email', 'Company', 'Status', and 'Draft Now' spelled exactly correctly in Row 1.");
    return;
  }
  
  // DYNAMIC CONTEXT FETCH: Load and parse the separate Resume.json file
  let myBackground;
  try {
    let rawJsonText = HtmlService.createHtmlOutputFromFile('Resume.json.html').getContent();
    rawJsonText = rawJsonText.replace(/\u00a0/g, ' ').replace(/&nbsp;/g, ' ').trim();
    const parsedResume = JSON.parse(rawJsonText);
    myBackground = JSON.stringify(parsedResume, null, 2);
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: Could not read or parse Resume.json file. Detail: " + err.toString());
    return;
  }
  
  // THE COLD EMAIL MASTERY STRICT RULES BLOCK
  const emailRules = `
  You are an elite copywriter executing the Cold Mail Mastery Framework. Draft an explicit, ultra-short cold email that mirrors real human outreach—NOT a job application dump.
  
  Enforce the 4-Part Framework rules strictly:
  1. THE HOOK: Open IMMEDIATELY with a sharp, specific observation about what their company does based on the Industry/Company Description. 
     * CRITICAL FILTER AGAINST BIO-DUMPING: Do NOT copy-paste or regurgitate their generic company summary description. Instead, pick ONE specific action, engineering product, or core solution they build and speak casually about how they approach it.
     * NEVER use conversational filler like "I hope this email finds you well", "My name is...", or "I am writing to express my interest". Start directly with the hook.
  2. THE CONNECTION: Link their company's core focus/engineering domain to an engineering challenge or niche you respect.
  3. THE CONGRUENCE (CONTEXT MATCHING RULE): Review the provided JSON background context. Look at the target company's description and dynamically pick EXACTLY ONE technical pillar based on this smart matching matrix:
     * Web/Dev/SaaS: Highlight "Elgrace Talents Platform & Systems Full-Stack Web App" or "CampusLink/Aviksoft".
     * AI/ML/Data: Highlight "AntiDOPE.ai" or the "Autonomous Waypoint Guidance System (YOLOv11s)".
     * Hardware/Robotics/Systems: Highlight "Rank 7 ISDC 2025" with Team Inferno DTU.
  4. THE ASK: A single, low-friction call to action asking for a brief 15-minute sync or a reply to look at a 60-second video demo.

  STRICT LAYOUT, VISUAL DENSITY & PHRASING RULES:
  - Do NOT put too many line breaks or empty spaces in the body. Combine the statements cleanly into two compact, highly skimmable text blocks:
    * Block 1: Your hook and connection sentence combined cleanly.
    * Block 2: Your matched background sentence and action statement combined.
  - MANDATORY INSERTION STATEMENT: You MUST organically include this exact phrase right before or as part of your final call to action: "I've added my resume for you to review."
  - Total body length MUST be under 80-100 words max (3 sentences total).
  - ZERO AI Buzzwords: Ban words like 'delighted', 'synergy', 'testament', 'foster', 'keen interest', 'esteemed organization', or 'pioneering'.
  `;

  let processedCount = 0;

  // Row looping starts at Index 1 to skip column headers
  for (let i = 1; i < data.length; i++) {
    let isChecked = data[i][colIndex.checkbox];
    let status = data[i][colIndex.status];
    let recruiterEmail = data[i][colIndex.email];
    
    let recruiterName = colIndex.name !== -1 ? data[i][colIndex.name] : "Team";
    let recruiterTitle = colIndex.title !== -1 ? data[i][colIndex.title] : "Executive";
    let companyName = data[i][colIndex.company];
    let industry = colIndex.industry !== -1 ? data[i][colIndex.industry] : "Tech Sector";
    let companyDescription = colIndex.description !== -1 ? data[i][colIndex.description] : "";
    
    // Clean up name fields safely
    let cleanedFirstName = String(recruiterName).split(' ')[0].trim();
    let greetingPrefix = "";
    if (!cleanedFirstName || cleanedFirstName.toLowerCase() === "team" || cleanedFirstName.toLowerCase() === "hiring") {
      greetingPrefix = "Hi there";
    } else {
      greetingPrefix = "Hi " + cleanedFirstName;
    }
    
    // Process checked matching targets only
    if (isChecked === true && recruiterEmail && status !== "Drafted") {
      
      const prompt = `
      Generate a raw cold outreach email draft following this exact structural instruction set.
      
      Sender Background (Parsed via JSON Context file):
      ${myBackground}
      
      Framework Constraints:
      ${emailRules}
      
      Recipient Details:
      - Title: ${recruiterTitle}
      - Company Name: ${companyName}
      - Industry: ${industry}
      - Company Profile: ${companyDescription}
      
      SUBJECT LINE FORMULA ENGINE RULES:
      Dynamically generate a compelling, high-converting subject line using one of these three Cold Mail Mastery patterns based on what fits best. Do NOT write standard application lines:
      - Pattern A (The Role Reframe): Template: [A bold, unique angle on the engineering problem or tech stack] | [Target Team Area] at ${companyName} (e.g., 'Application for the Jack of All, Master of Some Role | Founder's Office at ${companyName}' or 'Building scalable web infra at scale | Core Platform Team')
      - Pattern B (The Shared Journey): Template: [A specific product or system angle you loved about their recent growth/niche] (e.g., 'absolutely loved your journey building out the new core protocol stack')
      - Pattern C (The Function + Credibility Signal): Template: [Niche Dev/Engineering Focus Area] Space | DTU Engineering Network, [1 quantified engineering achievement signal asset]
      
      STRICT OUTPUT FORMAT PATTERN REQUIRED:
      You must wrap the components inside absolute explicit structural tokens. Do not use any markdown bolding, backticks, or asterisks. Return your text formatted precisely like this:

      [SUBJECT_START]
      [Your selected interesting, high-converting subject line formula text here]
      [SUBJECT_END]
      [BODY_START]
      [Start your body directly with the hook sentence here. Do NOT include greetings like "Hi" or sign-offs like "Best" inside this token block.]
      [BODY_END]
      `;
      
      // UPGRADED TO THE EXPLICIT GEMINI 3.1 FLASH LITE ROUTE STRINGS
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${API_KEY}`;
      const payload = {
        "contents": [{ "parts": [{ "text": prompt }] }]
      };
      
      const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      };
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        const json = JSON.parse(responseText);
        
        if (responseCode !== 200 || !json.candidates || json.candidates.length === 0 || !json.candidates[0].content) {
          throw new Error("Gemini API Error Response: " + responseText);
        }
        
        const aiResponse = json.candidates[0].content.parts[0].text.trim();
        
        // SECURE TEXT ISOLATION EXTRACTION VIA REGEX TOKENS
        let subjectMatch = aiResponse.match(/\[SUBJECT_START\]([\s\S]*?)\[SUBJECT_END\]/);
        let bodyMatch = aiResponse.match(/\[BODY_START\]([\s\S]*?)\[BODY_END\]/);
        
        let subject = `Quick question regarding ${companyName}`;
        let emailBody = "";
        
        if (subjectMatch && subjectMatch[1]) {
          subject = subjectMatch[1].trim();
        }
        if (bodyMatch && bodyMatch[1]) {
          emailBody = bodyMatch[1].trim();
        } else {
          emailBody = aiResponse.replace(/\[.*?\]/g, "").trim();
        }
        
        // SYSTEM CONCATENATION PROCESS
        let structuredBody = greetingPrefix + ",\n\n" + emailBody;
        structuredBody += `\n\nBest,\nSujal Bhati\nDelhi Technological University (DTU)\nlinkedin.com/sujalbhati`;
        
        // Instantiate message inside Gmail drafts folder
        GmailApp.createDraft(recruiterEmail, subject, structuredBody);
        
        // Update states dynamically by tracked indexes
        sheet.getRange(i + 1, colIndex.status + 1).setValue("Drafted");
        sheet.getRange(i + 1, colIndex.checkbox + 1).setValue(false);
        SpreadsheetApp.flush(); 
        processedCount++;
        
      } catch (e) {
        Logger.log("Failed execution processing index row " + (i + 1) + ": " + e.toString());
        sheet.getRange(i + 1, colIndex.status + 1).setValue("Error");
        sheet.getRange(i + 1, colIndex.checkbox + 1).setValue(false);
        SpreadsheetApp.flush();
      }
      
      // Speed optimized: 2 seconds delay keeps us completely safe and moving fast
      Utilities.sleep(2000); 
    }
  }

  // Final notification alert
  if (processedCount > 0) {
    SpreadsheetApp.getUi().alert(`Success! Generated ${processedCount} personalized draft(s) inside your Gmail.`);
  } else {
    SpreadsheetApp.getUi().alert("No rows were processed. Check that checkboxes are active and Status cells are blank.");
  }
}
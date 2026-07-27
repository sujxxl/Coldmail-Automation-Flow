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
 * parses external JSON resume context, attaches a native PDF from Google Drive,
 * processes payloads via Gemini API, and updates states.
 */
function processSelectedCheckboxes() {
  const API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // <--- Insert your free Google AI Studio API Key

  // ==========================================
  // 💾 GOOGLE DRIVE RESUME CONFIGURATION
  // ==========================================
  // Open Google Drive, right-click your Resume PDF -> Get Link.
  // The long string of random characters between '/d/' and '/view' is your File ID.
  const RESUME_FILE_ID = "YOUR_GOOGLE_DRIVE_FILE_ID_HERE"; // <--- Insert your Resume File ID here

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
    personalContext: headers.findIndex(h => String(h).trim().toLowerCase() === 'personal context'),
    status: headers.findIndex(h => String(h).trim().toLowerCase() === 'status'),
    checkbox: headers.findIndex(h => String(h).trim().toLowerCase() === 'draft now')
  };

  if (colIndex.email === -1 || colIndex.company === -1 || colIndex.status === -1 || colIndex.checkbox === -1) {
    SpreadsheetApp.getUi().alert("Error: Critical column headers missing! Ensure your sheet has 'Email', 'Company', 'Status', and 'Draft Now' spelled exactly correctly in Row 1.");
    return;
  }

  let resumeBlob;
  try {
    resumeBlob = DriveApp.getFileById(RESUME_FILE_ID).getBlob();
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: Could not access your Resume PDF from Google Drive. Verify your RESUME_FILE_ID is correct.");
    return;
  }

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

  // ==========================================
  // 🧠 RUN-LEVEL MEMORY (this is the core fix for monotony)
  // ==========================================
  // Everything used in a previous draft THIS RUN gets banned from the next
  // one, so five emails in a row can't all open the same way or cite the
  // same stat. Without this, the model has no way of knowing what it just
  // wrote two rows ago and defaults to its single highest-probability phrasing.
  const runMemory = {
    usedOpenerStyles: [],      // e.g. "direct-critique", "curious-question"
    usedProjectsHighlighted: [], // dynamically populated by whatever the LLM extracted
    usedStatPhrases: [],       // exact quantified phrases already used
    usedSubjectPatterns: []    // rotated deterministically
  };

  const OPENER_STYLES = [
    "direct-technical-observation", // casually point out how they solve a specific problem
    "curious-question",             // open with a genuine question about a technical tradeoff they made
    "specific-detail-callout",      // name one exact feature/product/post and react to it
    "casual-stack-recognition"      // casually mention a part of their tech stack or scaling journey
  ];

  const SUBJECT_PATTERNS = ["A", "B", "C"];

  let processedCount = 0;
  const draftedThisRun = new Set(); // dedupe guard — fixes duplicate-draft bug

  for (let i = 1; i < data.length; i++) {
    let isChecked = data[i][colIndex.checkbox];
    let status = data[i][colIndex.status];
    let recruiterEmail = data[i][colIndex.email];
    let companyName = data[i][colIndex.company];

    if (isChecked !== true || !recruiterEmail || status === "Drafted") continue;

    // Dedupe guard: skip if this exact recipient was already drafted in this run
    const dedupeKey = String(recruiterEmail).trim().toLowerCase() + "|" + String(companyName).trim().toLowerCase();
    if (draftedThisRun.has(dedupeKey)) {
      sheet.getRange(i + 1, colIndex.status + 1).setValue("Duplicate - Skipped");
      sheet.getRange(i + 1, colIndex.checkbox + 1).setValue(false);
      continue;
    }

    let recruiterName = colIndex.name !== -1 ? data[i][colIndex.name] : "Team";
    let recruiterTitle = colIndex.title !== -1 ? data[i][colIndex.title] : "Executive";
    let industry = colIndex.industry !== -1 ? data[i][colIndex.industry] : "Tech Sector";
    let companyDescription = colIndex.description !== -1 ? data[i][colIndex.description] : "";
    let personalContext = colIndex.personalContext !== -1 ? data[i][colIndex.personalContext] : "";

    let cleanedFirstName = String(recruiterName).split(' ')[0].trim();
    let greetingPrefix = (!cleanedFirstName || cleanedFirstName.toLowerCase() === "team" || cleanedFirstName.toLowerCase() === "hiring") 
      ? "Hi there" 
      : "Hi " + cleanedFirstName;

    const subjectPattern = SUBJECT_PATTERNS[processedCount % SUBJECT_PATTERNS.length];
    runMemory.usedSubjectPatterns.push(subjectPattern);

    const avoidBlock = `
    DO NOT reuse any of the following in this email — they've already been used in earlier emails this batch:
    - Opener styles already used: ${runMemory.usedOpenerStyles.length ? runMemory.usedOpenerStyles.join(", ") : "none yet"}
    - Projects already extracted/highlighted in the last 2 emails: ${runMemory.usedProjectsHighlighted.slice(-2).join(", ") || "none yet"}
    - Exact stat phrases already used: ${runMemory.usedStatPhrases.length ? runMemory.usedStatPhrases.join(" | ") : "none yet"}
    Pick an opener style from this list that has NOT been used yet: ${OPENER_STYLES.join(", ")}.
    `;

    // ==========================================
    // 🧠 DYNAMIC PROMPT ENGINE (Website Snoop + JSON Extraction)
    // ==========================================
    const emailRules = `
    You are a developer sending a quick, casual email to an engineering lead. Write exactly as if you just closed their website tab and immediately typed a quick thought on your phone.

    DYNAMIC CONTEXT MATCHING (CRITICAL):
    You must actively read the provided "Sender Background (JSON)". Analyze the recipient's Industry and Company Profile, then extract EXACTLY ONE project, past role, or technical skill cluster from the JSON that best matches their specific domain.
    Do NOT invent projects. Only use facts explicitly found in the JSON. If nothing perfectly aligns, find the closest systems-engineering or backend overlap.

    4-PART FRAMEWORK (strict):
    1. THE HOOK (The "Website Snoop" Open): Open by casually mentioning something specific you noticed about what they are building or scaling.
       ${personalContext ? "PRIORITIZE the Personal Context provided below over the generic company description. React to their actual post or news." : "Anchor the hook in ONE specific product or engineering decision from their company description."}
       - USE CONVERSATIONAL OPENERS LIKE: "Was just looking at how you guys handle...", "Noticed your team is building out...", "Saw you are tackling...", or "Really dig how you are approaching..."
       - BANNED STRUCTURES (CRITICAL): You are STRICTLY FORBIDDEN from using textbook statements or premise-conclusion setups. NEVER write phrases like "Managing X requires Y", "Your approach relies heavily on Z", "Automotive retail demands...", or "The shift toward X suggests Y".
    2. THE CONNECTION: Casually link their work to your own build.
       - USE TRANSITIONS LIKE: "Reminds me of a similar headache I had when...", "We had to solve a similar edge-case with...", or "I was just navigating similar technical hurdles while..."
       - BANNED TRANSITIONS: Do NOT use formal phrasing like "I recently navigated these same technical hurdles" or "I enjoyed working on..."
    3. THE CONGRUENCE: Seamlessly integrate the ONE specific project/role you dynamically extracted from the JSON. State ONE quantified result naturally. Do not reuse a stat phrase already flagged as used.
    4. THE ASK: One low-friction ask. Get on a call, talk it through, or connect. VARY the exact wording each time.

    ${avoidBlock}

    LAYOUT — THIS IS MANDATORY, NOT A SUGGESTION:
    The email body must be exactly TWO parts, generated and returned SEPARATELY (see output format below) — never merge them into one paragraph:
    - MAIN_BLOCK: The Hook + Connection + Congruence combined into one tight paragraph. 2–3 sentences, under 65 words. No greeting or sign-off.
    - ACTION_BLOCK: The Ask, ALONE, as its own separate short line. 1 sentence, under 20 words. 
    Mention the attached resume organically in either block.
    No markdown, no asterisks, no bold, no buzzwords ('delighted', 'synergy', 'pioneering', etc.). 
    `;

    const subjectFormulaText = subjectPattern === "A"
      ? `Pattern A (The Role Reframe): [a bold, specific angle on the actual engineering problem, lowercase] | [Target Team] at ${companyName}`
      : subjectPattern === "B"
      ? `Pattern B (The Shared Journey): [one specific, genuine thing about their recent work/growth/niche — not generic praise, lowercase]`
      : `Pattern C (The Function + Credibility Signal): [Specific niche/domain] | DTU Engineering Network, [one specific quantified signal]`;

    const prompt = `
    Generate a raw cold outreach email draft following this exact structural instruction set.

    Sender Background (JSON):
    ${myBackground}

    Framework Constraints:
    ${emailRules}

    Recipient Details:
    - Title: ${recruiterTitle}
    - Company Name: ${companyName}
    - Industry: ${industry}
    - Company Profile: ${companyDescription}
    ${personalContext ? `- Personal Context (their LinkedIn post / recent activity / shared ground): ${personalContext}` : ""}

    SUBJECT LINE (Formula assigned by rotation):
    ${subjectFormulaText}

    STRICT OUTPUT FORMAT — wrap output in these exact tokens, no markdown:
    [SUBJECT_START]
    [subject line text]
    [SUBJECT_END]
    [MAIN_BLOCK_START]
    [Hook + Connection + Congruence combined, 2-3 sentences, no greeting, no sign-off, no ask/CTA in here]
    [MAIN_BLOCK_END]
    [ACTION_BLOCK_START]
    [Just the ask/CTA, 1 short sentence, nothing else]
    [ACTION_BLOCK_END]
    [PROJECT_USED_START]
    [name of the ONE project/role you extracted from the JSON, e.g. "Elgrace Talents" — just the name]
    [PROJECT_USED_END]
    [OPENER_STYLE_START]
    [which opener style you used, from this list: ${OPENER_STYLES.join(", ")}]
    [OPENER_STYLE_END]
    [STAT_PHRASE_START]
    [the exact quantified stat phrase you used]
    [STAT_PHRASE_END]
    `;

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${API_KEY}`;
    const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
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

      let subjectMatch = aiResponse.match(/\[SUBJECT_START\]([\s\S]*?)\[SUBJECT_END\]/);
      let mainBlockMatch = aiResponse.match(/\[MAIN_BLOCK_START\]([\s\S]*?)\[MAIN_BLOCK_END\]/);
      let actionBlockMatch = aiResponse.match(/\[ACTION_BLOCK_START\]([\s\S]*?)\[ACTION_BLOCK_END\]/);
      let projectMatch = aiResponse.match(/\[PROJECT_USED_START\]([\s\S]*?)\[PROJECT_USED_END\]/);
      let openerMatch = aiResponse.match(/\[OPENER_STYLE_START\]([\s\S]*?)\[OPENER_STYLE_END\]/);
      let statMatch = aiResponse.match(/\[STAT_PHRASE_START\]([\s\S]*?)\[STAT_PHRASE_END\]/);

      let subject = `Quick question regarding ${companyName}`;
      let emailBody = "";

      if (subjectMatch && subjectMatch[1]) subject = subjectMatch[1].trim();

      if (mainBlockMatch && mainBlockMatch[1] && actionBlockMatch && actionBlockMatch[1]) {
        emailBody = mainBlockMatch[1].trim() + "\n\n" + actionBlockMatch[1].trim();
      } else if (mainBlockMatch && mainBlockMatch[1]) {
        emailBody = mainBlockMatch[1].trim();
      } else {
        emailBody = aiResponse.replace(/\[.*?\]/g, "").trim();
      }

      if (projectMatch && projectMatch[1]) runMemory.usedProjectsHighlighted.push(projectMatch[1].trim());
      if (openerMatch && openerMatch[1]) runMemory.usedOpenerStyles.push(openerMatch[1].trim());
      if (statMatch && statMatch[1]) runMemory.usedStatPhrases.push(statMatch[1].trim());

      let structuredBody = greetingPrefix + ",\n\n" + emailBody;
      structuredBody += `\n\nBest,\nSujal Bhati\nDelhi Technological University (DTU)\nlinkedin.com/sujalbhati`;

      GmailApp.createDraft(recruiterEmail, subject, structuredBody, {
        attachments: [resumeBlob]
      });

      draftedThisRun.add(dedupeKey);
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

    Utilities.sleep(2000);
  }

  if (processedCount > 0) {
    SpreadsheetApp.getUi().alert(`Success! Generated ${processedCount} personalized draft(s) with your Resume attached inside your Gmail.`);
  } else {
    SpreadsheetApp.getUi().alert("No rows were processed. Check that checkboxes are active and Status cells are blank.");
  }
}
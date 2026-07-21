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
 *
 * FIXES IN THIS VERSION (vs original):
 *  - Run-level memory: tracks every opener style, quantified stat, and subject
 *    formula already used in this batch and explicitly bans reusing them —
 *    this is what was causing every email to read like the same template.
 *  - Subject-line formula is now assigned deterministically by rotation
 *    (A/B/C/A/B/C...) instead of "let the model pick" — that's why every
 *    subject line came out as "Scaling X | Y" before.
 *  - Granular project-matching matrix (was 3 broad buckets that all resolved
 *    to the same "Elgrace Talents / 1000+ users / 85%" story regardless of
 *    company). Now matches on industry sub-type and forces a different
 *    highlighted project if the same one was used in the last 2 drafts.
 *  - Dedupe guard: skips a recipient (email+company) already drafted in this
 *    execution, fixing the duplicate-draft bug seen for Utthunga.
 *  - Optional "Personal Context" column: paste a recruiter's LinkedIn post,
 *    recent activity, or shared background here. When present, the hook is
 *    required to reference THIS instead of only the company description —
 *    this is the single biggest lever for making hooks feel non-generic.
 *  - Banned generic-opener list enforced directly in the prompt — and
 *    broadened to ban the underlying RHETORICAL SHAPE ("Company's shift
 *    toward X suggests Y"), not just specific phrases, since the model was
 *    paraphrasing around exact-phrase bans while keeping the same structure.
 *  - CTA phrasing must vary — no more identical "would you be open to a
 *    15-minute sync" every time.
 *  - HARD two-block layout: the model must return the hook/connection/
 *    congruence as one block and the ask as a separate block via distinct
 *    output tokens (MAIN_BLOCK / ACTION_BLOCK). The script then joins them
 *    with a guaranteed blank line — this is enforced in code, not left to
 *    the model's own paragraph-break judgment, which is what was producing
 *    unreadable single-paragraph walls of text before.
 *  - Enterprise/ERP bucket added to the project matrix so unrelated niche
 *    projects (e.g. a computer-vision stat) stop getting force-fit onto
 *    companies where they make no sense.
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
    // NEW — optional. Paste a recruiter's recent LinkedIn post, a company
    // news trigger, or a shared-ground note (same college, same city, etc).
    // If this column doesn't exist, the script still runs fine — it just
    // falls back to company-only context (the old, weaker behavior).
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
    usedProjectsHighlighted: [], // e.g. "Elgrace Talents", "AntiDOPE.ai"
    usedStatPhrases: [],       // exact quantified phrases already used
    usedSubjectPatterns: []    // rotated deterministically but logged anyway
  };

  const OPENER_STYLES = [
    "direct-technical-critique",   // lead with a sharp opinion on how they solve a specific problem
    "curious-question",            // open with a genuine question about a specific technical tradeoff they made
    "specific-detail-callout",     // name one exact feature/product/post and react to it
    "contrarian-observation"       // point out something non-obvious or slightly surprising about their approach
  ];

  const SUBJECT_PATTERNS = ["A", "B", "C"]; // rotated by row index, not left to the model to pick

  // Granular project matching — replaces the old 3-bucket system that kept
  // resolving to the same project regardless of company type.
  const PROJECT_MATRIX = `
  Match the company to the MOST SPECIFIC bucket below — do not default to the
  broadest one. If two projects could fit, prefer the more specific match:
  - Media / video / content delivery / streaming infra: FFmpeg/Sharp media-pipeline work
  - General web SaaS / platforms / internal tools: Elgrace Talents full-stack platform
  - E-commerce / supply chain / logistics / inventory / retail systems: CampusLink/Aviksoft
  - AI / ML / data / analytics / NLP: AntiDOPE.ai
  - Computer vision / autonomous systems / robotics / hardware / firmware / IIoT: Rank 7 ISDC 2025, Team Inferno DTU, YOLOv11s waypoint guidance
  - Fintech / payments / transaction infra / security-critical backend: pick whichever project involved the most rigorous data-integrity or security work — do not force-fit Elgrace Talents just because it's the strongest metric.
  - Enterprise software / ERP / IT staffing / custom business-app development / systems integration: Elgrace Talents full-stack platform or CampusLink/Aviksoft — whichever involved more backend/integration work. Do NOT use the YOLOv11s/computer-vision project here just because it has a strong percentage stat — ERP and CV are unrelated domains and forcing that connection reads as a non-sequitur.
  If the company doesn't cleanly match any bucket, pick the closest one and
  say so honestly through specific technical language — do not stretch a
  media-optimization story to cover a supply-chain company just because the
  user-count metric sounds impressive. A weaker but genuinely relevant project
  beats a stronger but unrelated one every time.
  `;

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
    let greetingPrefix = "";
    if (!cleanedFirstName || cleanedFirstName.toLowerCase() === "team" || cleanedFirstName.toLowerCase() === "hiring") {
      greetingPrefix = "Hi there";
    } else {
      greetingPrefix = "Hi " + cleanedFirstName;
    }

    // Deterministic subject pattern rotation — this is what fixes every
    // subject line collapsing into "Scaling X | Y". No AI discretion here.
    const subjectPattern = SUBJECT_PATTERNS[processedCount % SUBJECT_PATTERNS.length];
    runMemory.usedSubjectPatterns.push(subjectPattern);

    // Build the "avoid repeating yourself" block from run memory so far
    const avoidBlock = `
    DO NOT reuse any of the following in this email — they've already been
    used in earlier emails this batch and reusing them is the #1 failure mode:
    - Opener styles already used: ${runMemory.usedOpenerStyles.length ? runMemory.usedOpenerStyles.join(", ") : "none yet"}
    - Projects already highlighted in the last 2 emails: ${runMemory.usedProjectsHighlighted.slice(-2).join(", ") || "none yet"}
    - Exact stat phrases already used: ${runMemory.usedStatPhrases.length ? runMemory.usedStatPhrases.join(" | ") : "none yet"}
    Pick an opener style from this list that has NOT been used yet if possible: ${OPENER_STYLES.join(", ")}.
    `;

    const emailRules = `
    You are an elite copywriter executing the Cold Mail Mastery Framework. Draft an explicit, ultra-short cold email that reads like a real person wrote it about THIS specific person/company — not a template with the company name swapped in.

    ${PROJECT_MATRIX}

    4-PART FRAMEWORK (strict):
    1. THE HOOK: Open with a sharp, specific observation.
       ${personalContext ? "PRIORITIZE the Personal Context provided below over the generic company description — reference their actual post, recent move, or specific detail. This is what separates a real email from a template." : "No personal context was provided for this recipient, so anchor the hook in ONE specific product/engineering decision from the company description — never restate the description generically."}
       - BANNED: not just these exact phrases but this entire RHETORICAL SHAPE — "[Company]'s [transition/shift/focus/move] toward/into/on [X] suggests/requires/demands [derivative technical need]." Any sentence that follows that logical structure is banned even if the wording is different. Examples of the banned shape (do not paraphrase around these, avoid the underlying structure entirely): "Your transition into/from...", "[Company]'s ability to... is an impressive...", "[Company] shifting toward... suggests...", "Handling X at that scale requires a rigorous approach to Y."
       - Instead, open with a concrete, specific claim, question, or reaction — something a person would actually say out loud, not an inference chain.
       - Never use filler like "I hope this email finds you well" or "I am writing to express my interest."
    2. THE CONNECTION: Link their specific technical focus to an engineering challenge or niche you genuinely respect — in your own words, not a repeated formula.
    3. THE CONGRUENCE: Pick EXACTLY ONE project from the matrix above that most specifically fits this company's actual domain (not the broadest possible match, and not just the strongest metric). State ONE quantified result. Do not reuse a stat phrase already flagged as used. If no project genuinely fits, use plain language about relevant engineering experience rather than forcing an unrelated project in for the sake of a number.
    4. THE ASK: One low-friction ask — get on a call, talk it through, connect, or have them revert with thoughts. VARY the exact wording each time; do not default to "would you be open to a 15-minute sync."

    ${avoidBlock}

    LAYOUT — THIS IS MANDATORY, NOT A SUGGESTION:
    The email body must be exactly TWO parts, generated and returned SEPARATELY (see output format below) — never merge them into one paragraph:
    - MAIN_BLOCK: the Hook + Connection + Congruence, combined into one tight paragraph. 2–3 sentences, under 65 words.
    - ACTION_BLOCK: the Ask, ALONE, as its own separate short line. 1 sentence, under 20 words. This must NOT contain the hook or congruence content — it is only the call to action (e.g. asking to get on a call, talk, connect, or have them revert).
    Mention the attached resume inside MAIN_BLOCK or ACTION_BLOCK (vary which one and the phrasing) — do not make it a third block.
    No markdown, no asterisks, no bold, no buzzwords ('delighted', 'synergy', 'testament', 'foster', 'keen interest', 'esteemed organization', 'pioneering', 'robust', 'innovative', 'cutting-edge').
    `;

    const subjectFormulaText = subjectPattern === "A"
      ? `Pattern A (The Role Reframe): [a bold, specific angle on the actual engineering problem] | [Target Team] at ${companyName}`
      : subjectPattern === "B"
      ? `Pattern B (The Shared Journey): [one specific, genuine thing about their recent work/growth/niche — not generic praise]`
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
    ${personalContext ? `- Personal Context (their LinkedIn post / recent activity / shared ground — PRIORITIZE this for the hook): ${personalContext}` : ""}

    SUBJECT LINE — you MUST use this exact formula for this email (assigned by rotation, do not switch to a different pattern):
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
    [name of the ONE project you highlighted, e.g. "Elgrace Talents" — just the name, nothing else]
    [PROJECT_USED_END]
    [OPENER_STYLE_START]
    [which opener style you used, from this list: ${OPENER_STYLES.join(", ")}]
    [OPENER_STYLE_END]
    [STAT_PHRASE_START]
    [the exact quantified stat phrase you used, e.g. "improved handling performance by 85%"]
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

      // Join MAIN_BLOCK and ACTION_BLOCK with a guaranteed blank line between
      // them — this is what actually enforces the two-block readable layout.
      // We don't rely on the model to insert its own paragraph break here;
      // the script forces it every time regardless of what the model outputs.
      if (mainBlockMatch && mainBlockMatch[1] && actionBlockMatch && actionBlockMatch[1]) {
        emailBody = mainBlockMatch[1].trim() + "\n\n" + actionBlockMatch[1].trim();
      } else if (mainBlockMatch && mainBlockMatch[1]) {
        // Fallback if the model forgot to separate out the action block
        emailBody = mainBlockMatch[1].trim();
      } else {
        emailBody = aiResponse.replace(/\[.*?\]/g, "").trim();
      }

      // Update run memory so the NEXT row's prompt knows what to avoid
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
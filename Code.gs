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
 * Iterates through rows, validates checkbox states, parses external JSON resume context,
 * processes the payload via Gemini API, and moves results into Gmail Drafts.
 */
function processSelectedCheckboxes() {
    const API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // <--- Insert your free Google AI Studio API Key
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // DYNAMIC CONTEXT FETCH: Load and parse the separate Resume.json file
    let myBackground;
    try {
        let rawJsonText = HtmlService.createHtmlOutputFromFile('Resume.json').getContent();

        // FIX: Clean invisible web formatting and non-breaking spaces that break JSON engines
        rawJsonText = rawJsonText.replace(/\u00a0/g, ' ').replace(/&nbsp;/g, ' ').trim();

        const parsedResume = JSON.parse(rawJsonText);
        myBackground = JSON.stringify(parsedResume, null, 2);
    } catch (err) {
        SpreadsheetApp.getUi().alert("Error: Could not read or parse Resume.json file. Make sure it contains valid JSON formatting. Detail: " + err.toString());
        return;
    }

    // THE COLD EMAIL MASTERY STRICT RULES BLOCK
    const emailRules = `
  You are an elite executive copywriter executing the Cold Mail Mastery Framework. Draft an explicit, ultra-short cold email that mirrors real human output.
  
  Enforce the 4-Part Framework rules strictly:
  1. THE HOOK: Stop them from closing the email. Open IMMEDIATELY with a sharp, specific observation about what their company does based on the Industry/Company Description. 
     * NEVER use conversational filler like "I hope this email finds you well", "My name is...", or "I am writing to express my interest". Start directly with the hook.
  2. THE CONNECTION: Link their company's core focus/engineering domain to an engineering challenge or niche you respect.
  3. THE CONGRUENCE: Review the 'experience' and 'projects' blocks in the provided JSON context. Pick EXACTLY ONE specific, quantified project result or past internship highlight that maps cleanly to their tech space. Do NOT dump multiple elements or list everything.
  4. THE ASK: A single, low-friction call to action asking for a brief 15-minute sync or a reply to look at a 60-second video demo.
  
  Tone, Style, & Formatting Constraints:
  - Total length MUST be under 120-150 words (maximum 3-4 sentences total).
  - Keep it crisp, conversational, and direct. 
  - ZERO AI Buzzwords: Ban words like 'delighted', 'synergy', 'testament', 'foster', 'keen interest', 'esteemed organization', or 'pioneering'.
  - Stop text generation immediately at the final call to action sentence. Do not include signature blocks.
  `;

    let processedCount = 0;

    // Row looping starts at Index 1 to skip column headers
    for (let i = 1; i < data.length; i++) {
        let recruiterName = data[i][1];      // Column B: Name
        let recruiterEmail = data[i][2];     // Column C: Email
        let recruiterTitle = data[i][3];     // Column D: Title
        let companyName = data[i][4];        // Column E: Company
        let industry = data[i][6];           // Column G: Industry
        let companyDescription = data[i][7]; // Column H: Company Description
        let status = data[i][8];             // Column I: Status
        let isChecked = data[i][9];          // Column J: Checkbox (True/False)

        // Process checked matching targets only
        if (isChecked === true && recruiterEmail && status !== "Drafted") {

            const prompt = `
      Execute the Cold Mail Mastery framework to generate a raw cold outreach email draft.
      
      Sender Background (Parsed via JSON Context file):
      ${myBackground}
      
      Framework Constraints:
      ${emailRules}
      
      Recipient Details:
      - Recruiter/Executive Name: ${recruiterName}
      - Title: ${recruiterTitle}
      - Company Name: ${companyName}
      - Industry: ${industry}
      - Company Profile: ${companyDescription}
      
      Output Requirements:
      Line 1 MUST be exactly in this format following the Function + Credibility Signal structure:
      Subject: Application for [Specific Dev/Engineering] Role | DTU Engineering, [Mention the 1 specific asset you matched like 'ISDC Rank 7' or 'Full-Stack Developer']
      
      Line 2 onwards must be the email body text. End exactly at the final call to action sentence. No markdown formatting or bold markers.
      `;

            const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${API_KEY}`;
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
                const json = JSON.parse(response.getContentText());
                const aiResponse = json.candidates[0].content.parts[0].text;

                let lines = aiResponse.split('\n');
                let subject = `Software Role | DTU Electrical Engineering`;
                let body = aiResponse;

                if (lines[0].toLowerCase().startsWith("subject:")) {
                    subject = lines[0].replace(/subject:/i, "").trim();
                    body = lines.slice(1).join('\n').trim();
                }

                // Formulate clear human sign-off block
                body += `\n\nBest,\nSujal Bhati\nDelhi Technological University (DTU)\nlinkedin.com/sujalbhati`;

                // Instantiate message inside Gmail drafts folder
                GmailApp.createDraft(recruiterEmail, subject, body);

                // Update states
                sheet.getRange(i + 1, 9).setValue("Drafted");
                sheet.getRange(i + 1, 10).setValue(false);
                SpreadsheetApp.flush();
                processedCount++;

            } catch (e) {
                Logger.log("Failed execution processing index row " + (i + 1) + ": " + e.toString());
                sheet.getRange(i + 1, 9).setValue("Error");
                sheet.getRange(i + 1, 10).setValue(false);
                SpreadsheetApp.flush();
            }

            // Delay block to protect your free-tier token allocation
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
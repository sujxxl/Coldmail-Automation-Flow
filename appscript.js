function generateColdEmailDrafts() {
    const API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // <--- Put your free Gemini API key here
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // 1. SENDER BACKGROUND & CREDIBILITY SIGNALS (Function + Credibility Formula)
    const myBackground = `
  Sender Name: Sujal Bhati
  College/Identity Anchor: DTU (Delhi Technological University) / DU Alum Engineering Network
  Target Roles: Software Engineering, AI/ML, Full-Stack, Hardware-Software Interdisciplinary Roles
  Core Quantified Technical Pillars (Use 1 relevant pillar per email max):
  - Pillar A (Full-Stack/VPS): Built custom React frontends hosted on Vercel with Supabase backends; self-hosted and deployed complex enterprise ERPNext/Frappe CRM instances natively on VPS.
  - Pillar B (Robotics/AI/ML): Junior Software Lead for Team Inferno DTU. Ran electrical & software subsystems using ROS frameworks to secure Rank 7 at the International Space Drone Challenge (ISDC).
  - Pillar C (Cybersecurity/Ethical Hacking): Team Member at EHAX DTU (Ethical Hacking Society). Architected multi-stage CTF vulnerability challenges, OSINT analysis, and server security exploits.
  `;

    // 2. THE COLD EMAIL MASTERY STRICT RULES BLOCK
    const emailRules = `
  You are an elite ghostwriter executing the Cold Mail Mastery Framework. Write a cold email that feels like it was individually crafted by a human over 20 minutes.
  
  Follow the 4-Part Framework strictly:
  1. THE HOOK: Stop them from closing the email. Open IMMEDIATELY with a sharp, specific observation about what their company does based on the Industry/Company Description. 
     * NEVER use conversational filler like "I hope this email finds you well", "My name is...", or "I am writing to express my interest". Start directly with the hook.
  2. THE CONNECTION: Show you understand their world. Link their company's core focus/engineering domain to an engineering challenge or niche you respect.
  3. THE CONGRUENCE: Match yourself to them. Bring in EXACTLY ONE specific, quantified project result from the Sender Background (e.g., Rank 7 International Space Drone Challenge using ROS, or deploying full-stack systems on a VPS). Do NOT dump the whole resume.
  4. THE ASK: A single, low-friction call to action. Ask for a brief 15-minute sync or a quick reply to look at a 60-second video clip of your technical stacks.
  
  Tone & Style Rules:
  - Total length MUST be under 120-150 words (maximum 3-4 sentences total).
  - Keep it crisp, conversational, and direct. 
  - ZERO AI Buzzwords: Ban words like 'delighted', 'synergy', 'testament', 'foster', 'keen interest', 'esteemed organization', or 'pioneering'.
  - Do not include signature placeholders like "[Your Name]". Stop generating text right at the call to action sentence.
  `;

    // Loop through rows starting at index 1 (skipping headers)
    for (let i = 1; i < data.length; i++) {
        let recruiterName = data[i][1];      // Column B: Name
        let recruiterEmail = data[i][2];     // Column C: Email
        let recruiterTitle = data[i][3];     // Column D: Title
        let companyName = data[i][4];        // Column E: Company
        let industry = data[i][6];           // Column G: Industry
        let companyDescription = data[i][7]; // Column H: Company Description
        let status = data[i][8];             // Column I: Status (Tracking column)

        // Process only if Email exists and hasn't been drafted yet
        if (recruiterEmail && status !== "Drafted" && status !== "Skipped") {

            const prompt = `
      Execute the Cold Mail Mastery guide to generate a raw cold outreach email draft.
      
      Sender Background:
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
      Line 1 MUST be exactly in this format following the Function + Credibility Signal structure from the guide:
      Subject: Application for [Specific Dev/Engineering] Role | DTU Engineering, [Mention the 1 specific asset you used like 'ISDC Rank 7' or 'Full-Stack Developer']
      
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

                // Clean signature addition
                body += `\n\nBest,\nSujal Bhati\nDelhi Technological University (DTU)\nlinkedin.com/sujalbhati`;

                // Push straight to your Gmail Drafts
                GmailApp.createDraft(recruiterEmail, subject, body);

                // Mark row as complete
                sheet.getRange(i + 1, 9).setValue("Drafted");
                SpreadsheetApp.flush();

            } catch (e) {
                Logger.log("Failed on row " + (i + 1) + ": " + e.toString());
                sheet.getRange(i + 1, 9).setValue("Error");
                SpreadsheetApp.flush();
            }

            // Safety pause for free tier rate compliance
            Utilities.sleep(2000);
        }
    }
}
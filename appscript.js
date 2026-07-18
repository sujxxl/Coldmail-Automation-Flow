function generateColdEmailDrafts() {
    const API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // <--- Put your API key here
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // 1. INPUT YOUR FIXED RESUME BACKGROUND CONTEXT HERE
    const myBackground = `
  Name: Sujal Bhati
  Education: B.Tech in Electrical Engineering at Delhi Technological University (DTU).
  Key Skills/Exp: Full-stack dev (React, Supabase, Frappe/ERPNext), Cybersecurity (CTFs, ethical hacking), Robotics (ROS, Junior Software Lead for Team Inferno DTU space drone team).
  `;

    // 2. INPUT THE COLD EMAIL MASTERY RULES HERE
    const emailRules = `
  - Keep the total email length under 4 sentences.
  - Never start with conversational filler like "I hope this email finds you well" or "My name is...".
  - Open directly with the personalized company hook/research provided.
  - Make it sound brief, human, casual yet respectful, and completely organic. No AI buzzwords like 'delighted', 'synergy', 'testament', or 'foster'.
  - Conclude with a low-friction call to action asking if they have 2 minutes next week or if you can send over a 60-second video of your work.
  `;

    // Loop through rows (skip header row 0)
    for (let i = 1; i < data.length; i++) {
        let recruiterName = data[i][0];
        let recruiterEmail = data[i][1];
        let companyName = data[i][2];
        let companyHook = data[i][3];
        let status = data[i][4];

        // Only process if email exists and status isn't already "Drafted"
        if (recruiterEmail && status !== "Drafted") {

            const prompt = `
      You are an expert ghostwriter. Write a highly personalized, short cold email to a recruiter based on these details.
      
      Sender Background:
      ${myBackground}
      
      Strict Writing Rules to Follow:
      ${emailRules}
      
      Recipient Details:
      - Recruiter Name: ${recruiterName}
      - Company Name: ${companyName}
      - Specific Company Context/Research Hook: ${companyHook}
      
      Output ONLY the draft in raw text. Separate the Subject Line and the Body clearly. No markdown formatting.
      `;

            // Call Gemini API (using flash model for speed and free tier compatibility)
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

                // Parse Subject and Body from AI response
                let lines = aiResponse.split('\n');
                let subject = "Quick question regarding " + companyName;
                let body = aiResponse;

                if (lines[0].toLowerCase().includes("subject:")) {
                    subject = lines[0].replace(/subject:/i, "").trim();
                    body = lines.slice(1).join('\n').trim();
                }

                // Create the Draft in your Gmail account
                GmailApp.createDraft(recruiterEmail, subject, body);

                // Update the spreadsheet status
                sheet.getRange(i + 1, 5).setValue("Drafted");
                SpreadsheetApp.flush(); // Update sheet in real-time

            } catch (e) {
                Logger.log("Error processing row " + (i + 1) + ": " + e.toString());
                sheet.getRange(i + 1, 5).setValue("Error");
            }

            // Small pause to stay safely inside free tier rate limits
            Utilities.sleep(2000);
        }
    }
}
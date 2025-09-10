/**
 * @license
 * Copyright 2025 Peter Horner
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @author Peter Horner
 * @youtube https://www.youtube.com/@PeterHornerGoogleTech
 *
 * @fileoverview
 * This Google Apps Script uses the Gemini AI to automatically process new emails in Gmail.
 * It identifies emails that require a direct response, ignores notifications and marketing,
 * and applies Gmail labels ('To Respond' and 'Processed') to help you manage your inbox.
 *
 * --- INSTRUCTIONS FOR USE ---
 *
 * **Part 1: Initial Setup**
 *
 * 1.  **Create the Script:**
 * - Go to script.google.com and create a new project.
 * - Name your project (e.g., "AI Email Assistant").
 * - Copy and paste the ENTIRE content of this file into the `Code.gs` editor.
 *
 * 2.  **Configure the Manifest File:**
 * - In the editor, click the 'Project Settings' ⚙️ icon on the left.
 * - Check the box "Show 'appsscript.json' manifest file in editor".
 * - Return to the 'Editor' 📝 view. You will now see the `appsscript.json` file.
 * - Delete the contents of that file and replace it with the provided `appsscript.json` code.
 *
 * 3.  **Get Your Gemini API Key:**
 * - Go to Google AI Studio: https://aistudio.google.com/
 * - Click "Get API key" and then "Create API key in new project".
 * - Copy the generated API key. KEEP THIS KEY PRIVATE LIKE A PASSWORD.
 *
 * 4.  **Store Your API Key Securely:**
 * - In the script (`Code.gs`), find the `setupScriptProperties()` function.
 * - Replace "PASTE_YOUR_API_KEY_HERE" with the API key you just copied.
 * - Save the script (Ctrl+S or Cmd+S).
 * - From the function dropdown menu at the top, select `setupScriptProperties` and click "Run".
 * - You will be asked to authorize the script. Follow the prompts to allow it.
 * - After it runs successfully, it's a good security practice to remove your key from the code.
 *
 * 5.  **Create Gmail Labels:**
 * - In your Gmail account, manually create two new labels:
 * - `Processed`
 * - `To Respond`
 * - (Note: The script will create these for you if they don't exist, but it's good practice to set them up first).
 *
 *
 * **Part 2: Running the Script**
 *
 * 1.  **Manual Run:**
 * - To test the script, select the `processInbox` function from the dropdown menu and click "Run".
 * - Check the Execution Log at the bottom of the screen to see its progress.
 *
 * 2.  **Automated Run (Trigger):**
 * - To have the script run automatically, click the 'Triggers' ⏰ icon on the left.
 * - Click "+ Add Trigger" in the bottom right.
 * - Set up the trigger with the following settings:
 * - Choose which function to run: `processInbox`
 * - Select event source: `Time-driven`
 * - Select type of time-based trigger: `Hour timer` (or your preferred frequency)
 * - Select hour interval: `Every hour`
 * - Click "Save". The script will now run automatically in the background.
 *
 * --- DISCLAIMER ---
 *
 * This script is provided as-is, without any warranty. The author, Peter Horner, is not
 * responsible for any issues, data loss, or unexpected costs that may arise from its use.
 * By using this script, you agree to do so at your own risk.
 *
 * This script uses the Google AI (Gemini) API. Depending on your usage, this may incur
 * costs. Please review Google's API pricing. You are responsible for all API costs
 * associated with your use of this script.
 *
 * Always keep your API key secure and private. Do not share it or commit it to public
 * code repositories.
 *
 */

/**
 * Main function to process a batch of unprocessed emails.
 * It iterates through each email thread, analyzes its content, and applies appropriate labels.
 */
function processInbox() {
  // Retrieve a batch of email threads that are unread, have the 'Test' label,
  // and do not yet have the 'Processed' label.
  const threads = getUnprocessedEmails();

  // Loop through each retrieved email thread.
  for (const thread of threads) {
    // Use a try-catch block to handle potential errors for a single thread,
    // allowing the script to continue with the next threads if one fails.
    try {
      // Get the first message from the current email thread.
      const message = thread.getMessages()[0];
      // Send the email message to the AI for analysis.
      const analysis = getEmailAnalysis(message);

      // If the analysis returns a result, apply the corresponding Gmail labels.
      if (analysis) {
        applyGmailLabels(thread, analysis);
      }
    } catch (e) {
      // Log an error message to the console if processing a thread fails.
      console.error(`Failed to process thread ${thread.getId()}: ${e.toString()}`);
    }
  }
}

/**
 * Searches Gmail for unprocessed emails and returns them.
 * @returns {GmailApp.GmailThread[]} An array of Gmail threads matching the search criteria.
 */
function getUnprocessedEmails() {
  // Searches for threads that are unread, have the 'Test' label,
  // but do NOT have the 'Processed' label. This prevents re-processing emails.
  // The '0, 5' parameters limit the search to the first 5 results to avoid exceeding execution limits.
  return GmailApp.search('is:unread label:Test -label:Processed', 0, 5);
}

/**
 * A utility function for testing. It retrieves unprocessed emails
 * and logs their subjects to the console.
 */
function testReadingEmails() {
  // Get the same batch of emails as the main processing function.
  const threads = getUnprocessedEmails();
  // Loop through the threads and print the subject of the first message in each.
  for (const thread of threads) {
    console.log(thread.getFirstMessageSubject());
  }
}

/**
 * Analyzes an email message using the Gemini AI API to determine if a response is required.
 * @param {GmailApp.GmailMessage} message The email message object to analyze.
 * @returns {Object|null} A JSON object with the analysis result, e.g., { "requiresResponse": true }.
 */
function getEmailAnalysis(message) {
  // Retrieve the Gemini API key stored in the script's properties.
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  // Specify the AI model to use. 'flash' is a fast and cost-effective model.
  const model = 'gemini-1.5-flash-latest';
  // Construct the full API endpoint URL.
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  // Define the prompt with detailed instructions for the AI.
  // This prompt tells the AI exactly how to analyze the email and what JSON format to return.
  const prompt = `
    Analyze the following email and determine if it requires a direct response from me. Your output must be a single JSON object with one key:
- "requiresResponse": a boolean (true or false).

Set "requiresResponse" to true ONLY if the email meets one of the following criteria:
1. It is a direct, personal message specifically addressed to me that asks a question.
2. It is a direct request for a specific deliverable or a task that I am personally responsible for.
3. It explicitly needs my review, approval, or a decision from me.
4. It is a message I initiated and is now pending my next response.

Set "requiresResponse" to false in all other cases, including but not limited to:
1. Automated notifications, system-generated emails, or transactional messages (e.g., file access requests, comment threads, purchase receipts).
2. All Google Calendar updates, meeting invitations, and confirmations.
3. Broad announcements, mass newsletters, marketing, promotional emails, or general invitations to events.
4. Informational notes that do not ask for a reply, such as "thank you" messages or "for your information" updates.
5. Emails with a general call to action that is not a direct request to me (e.g., "register now," "sign up here," "view our latest blog post").
6. Messages where I am CC'd for informational purposes and my action is not required.

    Email Subject: "${message.getSubject()}"
    Email Body:
    ---
    ${message.getPlainBody()}
    ---
  `;

  // Log the generated prompt for debugging purposes.
  console.log(prompt);

  // Prepare the data payload in the format required by the Gemini API.
  const payload = { "contents": [{ "parts": [{ "text": prompt }] }] };
  // Configure the HTTP request options.
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    // The payload must be converted to a JSON string.
    'payload': JSON.stringify(payload)
  };

  // Make the API call to the Gemini service.
  const httpResponse = UrlFetchApp.fetch(url, options);
  // Parse the JSON response received from the API.
  const responseData = JSON.parse(httpResponse.getContentText());
  // Extract the text content from the response, clean up any markdown code fences (```), and trim whitespace.
  const jsonString = responseData.candidates[0].content.parts[0].text.replace(/```json|```/g, '').trim();
  // Parse the cleaned text into a final JSON object.
  return JSON.parse(jsonString);
}

/**
 * Applies Gmail labels to a thread based on the AI analysis.
 * @param {GmailApp.GmailThread} thread The email thread to label.
 * @param {Object} analysis The analysis object containing the "requiresResponse" boolean.
 */
function applyGmailLabels(thread, analysis) {
  const subject = thread.getFirstMessageSubject();
  // Log the AI's decision to the console for monitoring.
  console.log(analysis.requiresResponse);

  // Check if the analysis determined that a response is required.
  if (analysis.requiresResponse) {
    // Get the user-defined label 'To Respond'.
    const toRespondLabel = GmailApp.getUserLabelByName('To Respond');
    // Apply the 'To Respond' label to the thread.
    thread.addLabel(toRespondLabel);
    console.log(`Email "${subject}" labeled as 'To Respond'.`);
  }

  // Always apply the 'Processed' label to prevent the script from analyzing this email again.
  const processedLabel = GmailApp.getUserLabelByName('Processed');
  thread.addLabel(processedLabel);
  console.log(`Email "${subject}" labeled as 'Processed'.`);
}

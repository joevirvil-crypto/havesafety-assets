/**
 * ==========================================
 * Haven Knowledge Base Web App
 * Code.gs
 * ==========================================
 */

/**
 * Serves the main HTML page.
 */
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Knowledge Base")
    .setFaviconUrl(
      "https://cdn.jsdelivr.net/gh/joevirvil-crypto/havesafety-assets@main/haven-favicon.png"
    )
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Includes HTML/CSS/JS files.
 */
function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* =========================================================
   SEARCH KNOWLEDGE
========================================================= */

/**
 * Searches Topic, Content, and Article ID.
 *
 * @param {string} keyword
 * @return {Array}
 */
function searchKnowledge(keyword) {

  if (!keyword) return [];

  const searchTerm = keyword.toString().trim().toLowerCase();

  if (searchTerm === "") return [];

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("KnowledgeBase");

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) return [];

  const results = [];

  for (let i = 1; i < data.length; i++) {

    const topic = (data[i][0] || "").toString();
    const content = (data[i][1] || "").toString();
    const articleId = (data[i][2] || "").toString();

    if (
      topic.toLowerCase().includes(searchTerm) ||
      content.toLowerCase().includes(searchTerm) ||
      articleId.toLowerCase().includes(searchTerm)
    ) {

      results.push({
        id: i,
        topic: topic,
        articleId: articleId
      });

    }

  }

  return results;

}

/* =========================================================
   GET ARTICLE
========================================================= */

/**
 * Returns one article.
 *
 * Column A = Topic
 * Column B = Content
 * Column C = Article ID
 * Column D = Updated
 * Column E = PDF File IDs
 *
 * Multiple PDFs supported.
 *
 * Example:
 * 1abc123,2def456,3ghi789
 */
function getArticle(id) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("KnowledgeBase");

  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();

  const rowIndex = Number(id);

  if (
    isNaN(rowIndex) ||
    rowIndex < 1 ||
    rowIndex >= data.length
  ) {
    return null;
  }

  const updated = data[rowIndex][3];

// Column E contains comma-separated PDF IDs
const pdfIds = (data[rowIndex][4] || "").toString();

const pdfs = pdfIds
  .split(",")
  .map(id => id.trim())
  .filter(id => id !== "")
  .map(id => ({
    url: "https://drive.google.com/file/d/" + id + "/view"
  }));

return {
  topic: data[rowIndex][0] || "",
  content: data[rowIndex][1] || "",
  articleId: data[rowIndex][2] || "",
  updated: updated
    ? Utilities.formatDate(
        new Date(updated),
        Session.getScriptTimeZone(),
        "MMMM d, yyyy 'at' h:mm a"
      )
    : "",
  pdfs: pdfs
};

}

/* =========================================================
   ON EDIT
========================================================= */

/**
 * Automatically:
 * - Updates timestamp when Column B changes
 * - Auto-resizes edited row
 * - Auto-resizes edited column
 */
function onEdit(e) {

  const sheet = e.range.getSheet();

  if (sheet.getName() !== "KnowledgeBase") return;

  const row = e.range.getRow();
  const column = e.range.getColumn();

  // Skip header row
  if (row < 2) return;

  // Timestamp when Column B changes
  if (column === 2) {

    const timestampCell = sheet.getRange(row, 4);

    if (e.range.getValue() !== "") {
      timestampCell.setValue(new Date());
    } else {
      timestampCell.clearContent();
    }

  }

  // Auto resize row
  sheet.autoResizeRows(row, 1);

  // Auto resize edited column
  sheet.autoResizeColumns(column, 1);

}
/**
 * Knowledge Base AI Assistant backend module for Google Apps Script.
 * Retrieves data from the active Google Sheet and queries Gemini 3 Flash.
 */

// Global configuration constant
var CONFIG = {
  // Set your Gemini API key here or store it in Script Properties
  GEMINI_API_KEY: "AQ.Ab8RN6K09sVlQUoWe1J7B45CHx6MOg9monuff52dDk_MePd1Cg", // Insert your Gemini API key here or set script property 'GEMINI_API_KEY'
  SHEET_NAME: "KnowledgeBase", // Optional: Target sheet name, falls back to first sheet if not found
  MODEL_NAME: "gemini-3-flash-preview"
};


/**
 * Reads data from the Google Sheet and formats it into structured text context.
 * @return {string} Formatted context string from sheet rows.
 */
function getSheetKnowledgeContext() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return "Error: No active Google Spreadsheet bound to this script.";
    }

    // Try finding the specific sheet, otherwise default to the first tab
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    if (!data || data.length < 2) {
      return "The knowledge base sheet is currently empty or has no data rows.";
    }

    var headers = data[0];
    var formattedText = "KNOWLEDGE BASE RECORDS:\n\n";

    // Loop through rows starting from row index 1 (skipping header row)
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      // Skip empty rows
      if (row.join("").trim() === "") continue;

      formattedText += "--- Record #" + r + " ---\n";
      for (var c = 0; c < headers.length; c++) {
        var header = headers[c] ? headers[c].toString().trim() : "Field " + (c + 1);
        var value = row[c] !== undefined && row[c] !== null ? row[c].toString().trim() : "";
        if (value !== "") {
          formattedText += header + ": " + value + "\n";
        }
      }
      formattedText += "\n";
    }

    return formattedText;

  } catch (error) {
    Logger.log("Error reading spreadsheet: " + error.toString());
    return "Error reading sheet data: " + error.toString();
  }
}


/**
 * Main entry point called from index.html via google.script.run
 * Sends user query and sheet context to Gemini AI and returns response.
 * @param {string} userQuestion - Question asked by user in the chat UI.
 * @return {string} AI answer generated strictly from sheet data.
 */
function askKnowledgeBaseAI(userQuestion) {
  try {
    if (!userQuestion || userQuestion.trim() === "") {
      return "Please enter a valid question.";
    }

    // Retrieve API key from CONFIG or Script Properties
    var apiKey = CONFIG.GEMINI_API_KEY || PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    
    if (!apiKey) {
      return "Configuration Error: Gemini API key is missing. Please set GEMINI_API_KEY in Code.gs or Script Properties.";
    }

    // Step 1: Get latest context from Google Sheet
    var sheetContext = getSheetKnowledgeContext();

    // Step 2: Prepare Gemini API Request
    var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + CONFIG.MODEL_NAME + ":generateContent?key=" + apiKey;

    var systemInstructionText = 
      "You are a helpful, professional AI Customer Support & Knowledge Base Assistant. " +
      "Your sole task is to answer user questions using ONLY the provided Knowledge Base Records. " +
      "If the user asks something that cannot be answered or inferred from the provided records, politely reply: " +
      "'I apologize, but I don't have that information in my current knowledge base database.' " +
      "Keep answers concise, clear, polite, and directly based on the data provided.";

    var fullPromptText = "CONTEXT FROM DATA SHEET:\n" + sheetContext + "\n\nUSER QUESTION:\n" + userQuestion;

    var payload = {
      "contents": [
        {
          "role": "user",
          "parts": [{ "text": fullPromptText }]
        }
      ],
      "systemInstruction": {
        "parts": [{ "text": systemInstructionText }]
      }
    };

    var fetchOptions = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };


    var response = UrlFetchApp.fetch(apiUrl, fetchOptions);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    var jsonResult = JSON.parse(responseText);

    if (responseCode === 200 && jsonResult.candidates && jsonResult.candidates.length > 0) {
      var candidate = jsonResult.candidates[0];
      if (candidate.content && candidate.content.parts && candidate.content.parts.length > 0) {
        return candidate.content.parts[0].text;
      }
    }

    if (jsonResult.error) {
      Logger.log("Gemini API Error: " + JSON.stringify(jsonResult.error));
      return "AI Service Error: " + (jsonResult.error.message || "Failed to retrieve response from AI service.");
    }

    return "Sorry, I was unable to find an answer to your question in the knowledge base.";

  } catch (err) {
    Logger.log("Execution Error: " + err.toString());
    return "An internal server error occurred while processing your query: " + err.toString();
  }
}
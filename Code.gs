const SHEET_NAME = "Sheet1";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    // If headers don't exist yet, add them
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Score"]);
    }

    // Append the player's name and score
    sheet.appendRow([new Date(), data.name, data.stoppedAt]);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();

  let scores = [];
  // Skip header row (start at i = 1)
  for (let i = 1; i < rows.length; i++) {
    scores.push({
      timestamp: rows[i][0],
      name: rows[i][1],
      stoppedAt: rows[i][2] // This holds the final score
    });
  }

  // Sort descending by score (highest score first)
  scores.sort((a, b) => b.stoppedAt - a.stoppedAt);

  return ContentService.createTextOutput(JSON.stringify(scores))
    .setMimeType(ContentService.MimeType.JSON);
}

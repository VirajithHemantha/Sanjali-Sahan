var SPREADSHEET_ID = "1UNWOffY0xDk3rdgBGlUGqYLKv08w9FEMeqmeW7xCTQU";

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    // If accessed directly via browser without parameters, return a friendly message
    if (!e || !e.parameter || !e.parameter.sheet) {
      return ContentService.createTextOutput("Success: Web App is deployed and working correctly!").setMimeType(ContentService.MimeType.TEXT);
    }

    var sheetName = e.parameter.sheet;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    
    // Auto-create sheets and headers if they don't exist
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (sheetName === "rsvp") {
        sheet.appendRow(["Date", "Name", "Guests", "Notes"]);
        sheet.getRange(1, 1, 1, 4).setFontWeight("bold"); // Make headers bold
      } else if (sheetName === "wish") {
        sheet.appendRow(["Date", "Name", "Message"]);
        sheet.getRange(1, 1, 1, 3).setFontWeight("bold"); // Make headers bold
      }
    }
    
    // Append the incoming data
    if (sheetName === "rsvp") {
      sheet.appendRow([
        e.parameter.date || new Date().toLocaleString(),
        e.parameter.name || "Unknown",
        e.parameter.guests || "0",
        e.parameter.notes || ""
      ]);
    } else if (sheetName === "wish") {
      sheet.appendRow([
        e.parameter.date || new Date().toLocaleString(),
        e.parameter.name || "Unknown",
        e.parameter.message || ""
      ]);
    }
    
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.toString()).setMimeType(ContentService.MimeType.TEXT);
  }
}

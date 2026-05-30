// Google Apps Script — paste this into script.google.com
// Linked to: Chou Collection SMS Opt-In sheet
// Sheet ID: 1Eg3aWDXip-7UCgPKJQiL9gC4p4RHuautfZsEQeIg5k8

const SHEET_ID = "1Eg3aWDXip-7UCgPKJQiL9gC4p4RHuautfZsEQeIg5k8";
const SHEET_NAME = "Opt-Ins";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME)
                  || SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Full Name", "Mobile Phone", "Email",
        "Property Address", "Consent Given", "User Agent"
      ]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold")
           .setBackground("#8b6f47").setFontColor("#ffffff");
    }

    sheet.appendRow([
      new Date().toLocaleString("en-US", {timeZone: "America/Los_Angeles"}),
      data.name || "",
      data.phone || "",
      data.email || "",
      data.property || "",
      data.consent === true ? "YES — Express consent given via web form" : "NO",
      data.userAgent || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: "error", message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: "ok", message: "Chou Collection SMS Opt-In endpoint"}))
    .setMimeType(ContentService.MimeType.JSON);
}

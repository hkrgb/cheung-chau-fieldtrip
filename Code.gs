const SHEET_ID = "1imqkcaee0bSMYuWf5YenNS5gxWOGvjl3bELJ1v8Nq2A";
const SHEET_NAME = "提交紀錄";

function doGet() {
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    const groupName = data.groupName || "";
    const answers = data.answers || {};

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    const questionIds = Object.keys(answers);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 2 + questionIds.length)
        .setValues([["提交時間", "組別"].concat(questionIds)]);
    } else {
      questionIds.forEach(function(id) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        if (headers.indexOf(id) === -1) {
          sheet.getRange(1, headers.length + 1).setValue(id);
        }
      });
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(function(h) {
      if (h === "提交時間") return new Date();
      if (h === "組別") return groupName;
      return answers[h] || data[h] || "";
    });
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

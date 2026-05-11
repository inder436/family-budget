const SHEET_NAME = "Budget";

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const params   = e.parameter;
  const action   = params.action;
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  let sheet      = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["year","month","category","amount","description","updated"]);
  }

  let result;
  if      (action === "get") result = getData(sheet);
  else if (action === "set") result = setData(sheet, params);
  else                        result = { error: "Unknown action" };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getData(sheet) {
  const rows = sheet.getDataRange().getValues();
  const data = {};
  for (let i = 1; i < rows.length; i++) {
    const [year, month, category, amount, description] = rows[i];
    const k = `${year}-${String(month).padStart(2,"0")}-${category}`;
    data[k] = amount;
    // also store description under key with -desc suffix
    if (description !== undefined && description !== "") {
      data[k + "-desc"] = description;
    }
  }
  return { ok: true, data };
}

function setData(sheet, params) {
  const year     = params.year;
  const month    = params.month;
  const category = decodeURIComponent(params.category || "");
  const amount   = parseFloat(params.amount) || 0;
  const desc     = decodeURIComponent(params.desc || "");

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(year) &&
        String(rows[i][1]) === String(month) &&
        String(rows[i][2]) === String(category)) {
      sheet.getRange(i+1, 4).setValue(amount);
      if (desc !== "") sheet.getRange(i+1, 5).setValue(desc);
      sheet.getRange(i+1, 6).setValue(new Date().toISOString());
      return { ok: true, action: "updated" };
    }
  }
  sheet.appendRow([year, month, category, amount, desc, new Date().toISOString()]);
  return { ok: true, action: "inserted" };
}

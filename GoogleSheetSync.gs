const doc = SpreadsheetApp.openById("1nFUoOWcA7hQ6B_724h9rwMiE0fzUto9IU4nxHzdxXhI");

function doPost(e) {
  const action = ((e && e.parameter && e.parameter.action) || "").toLowerCase();

  if (action === "tabs") {
    const tabs = doc.getSheets().map((sheet) => sheet.getName());
    return jsonResponse({ status: "success", tabs });
  }

  if (action === "sync_assignments") {
    const rawRecords = (e && e.parameter && e.parameter.records) || "[]";
    const records = JSON.parse(rawRecords);

    if (!Array.isArray(records)) {
      return jsonResponse({ status: "error", message: "records must be a JSON array" });
    }

    const grouped = {};
    for (const record of records) {
      const className = String((record && record.className) || "").trim();
      if (!className) continue;
      if (!grouped[className]) grouped[className] = [];

      grouped[className].push({
        assignmentName: String((record && record.assignmentName) || ""),
        dueDate: String((record && record.dueDate) || ""),
        className,
      });
    }

    const updatedClasses = [];
    let rowsWritten = 0;

    for (const className in grouped) {
      const sheet = doc.getSheetByName(className);
      if (!sheet) {
        continue;
      }

      const classRecords = grouped[className].slice();
      classRecords.sort((a, b) => {
        const aDate = parseMmDdYyyy(a.dueDate);
        const bDate = parseMmDdYyyy(b.dueDate);
        return aDate.getTime() - bDate.getTime();
      });

      clearAssignmentColumns(sheet);

      if (classRecords.length > 0) {
        const values = classRecords.map((item) => [item.assignmentName, item.dueDate, item.className]);
        sheet.getRange(2, 1, values.length, 1).setValues(values.map((row) => [row[0]]));
        sheet.getRange(2, 2, values.length, 1).setValues(values.map((row) => [row[1]]));
        sheet.getRange(2, 4, values.length, 1).setValues(values.map((row) => [row[2]]));
      }

      updatedClasses.push(className);
      rowsWritten += classRecords.length;
    }

    return jsonResponse({
      status: "success",
      updatedClasses,
      rowsWritten,
      requestedClasses: Object.keys(grouped),
    });
  }

  const textValue = (e && e.parameter && e.parameter.text) || "SUCCESS";
  return jsonResponse({ status: "success", text: textValue });
}

function clearAssignmentColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const rowCount = lastRow - 1;

  sheet.getRange(2, 1, rowCount, 1).clearContent();
  sheet.getRange(2, 2, rowCount, 1).clearContent();
  sheet.getRange(2, 4, rowCount, 1).clearContent();
}

function parseMmDdYyyy(value) {
  const match = String(value || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return new Date(8640000000000000);

  const month = Number(match[1]) - 1;
  const day = Number(match[2]);
  const year = Number(match[3]);
  return new Date(year, month, day);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

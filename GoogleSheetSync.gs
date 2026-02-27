const doc = SpreadsheetApp.openById("1nFUoOWcA7hQ6B_724h9rwMiE0fzUto9IU4nxHzdxXhI");
const EXCLUDED_TAB_NAMES = new Set(["dashboard", "class[template]"]);

function doPost(e) {
  const action = ((e && e.parameter && e.parameter.action) || "").toLowerCase();

  if (action === "tabs") {
    const tabs = doc.getSheets().map((sheet) => sheet.getName());
    return jsonResponse({
      status: "success",
      spreadsheetId: doc.getId(),
      spreadsheetName: doc.getName(),
      tabs,
    });
  }

  if (action === "dump_tabs") {
    const maxRows = Number((e && e.parameter && e.parameter.maxRows) || 250);
    const safeMaxRows = isNaN(maxRows) ? 250 : Math.max(1, Math.min(maxRows, 2000));
    const dump = doc.getSheets().map((sheet) => {
      const lastRow = sheet.getLastRow();
      const rowCount = Math.min(Math.max(0, lastRow - 1), safeMaxRows);
      const values = rowCount > 0 ? sheet.getRange(2, 1, rowCount, 4).getValues() : [];

      const rows = values
        .map((row, index) => ({
          rowNumber: index + 2,
          assignmentName: String(row[0] || "").trim(),
          dueDate: formatDueDate(row[1]),
          className: String(row[3] || "").trim(),
        }))
        .filter((row) => row.assignmentName || row.dueDate || row.className);

      return {
        sheetName: sheet.getName(),
        lastRow,
        nonEmptyRowsInDump: rows.length,
        rows,
      };
    });

    return jsonResponse({
      status: "success",
      spreadsheetId: doc.getId(),
      spreadsheetName: doc.getName(),
      maxRows: safeMaxRows,
      sheets: dump,
    });
  }

  if (action === "clear_all_class_tabs") {
    const clearedTabs = [];
    let clearedRows = 0;

    const sheets = doc.getSheets();
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (isExcludedTabName(sheetName)) {
        continue;
      }

      const classClearedRows = clearSheetAssignmentColumns(sheet);
      clearedRows += classClearedRows;
      clearedTabs.push({
        sheetName,
        clearedRows: classClearedRows,
      });
    }

    return jsonResponse({
      status: "success",
      spreadsheetId: doc.getId(),
      spreadsheetName: doc.getName(),
      action: "clear_all_class_tabs",
      clearedRows,
      clearedTabs,
    });
  }

  if (action === "clear_class_tab") {
    const className = String((e && e.parameter && e.parameter.className) || "").trim();
    if (!className) {
      return jsonResponse({ status: "error", message: "className is required" });
    }

    const sheet = doc.getSheetByName(className);
    if (!sheet) {
      return jsonResponse({ status: "error", message: `Sheet not found: ${className}` });
    }

    const classClearedRows = clearSheetAssignmentColumns(sheet);
    return jsonResponse({
      status: "success",
      spreadsheetId: doc.getId(),
      spreadsheetName: doc.getName(),
      action: "clear_class_tab",
      className,
      clearedRows: classClearedRows,
    });
  }

  if (action === "sync_assignments") {
    const rawRecords = (e && e.parameter && e.parameter.records) || "[]";
    const dryRunValue = String((e && e.parameter && e.parameter.dryRun) || "false").toLowerCase();
    const dryRun = dryRunValue === "true" || dryRunValue === "1" || dryRunValue === "yes";
    const replaceExistingValue = String((e && e.parameter && e.parameter.replaceExisting) || "false").toLowerCase();
    const replaceExisting = replaceExistingValue === "true" || replaceExistingValue === "1" || replaceExistingValue === "yes";
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
    const debugMessages = [];
    const classStats = {};
    let addedRows = 0;
    let updatedRows = 0;

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

      const classRecordsFiltered = classRecords.filter((item) => String(item.assignmentName || "").trim() !== "");

      if (replaceExisting) {
        const lastRowForClear = sheet.getLastRow();
        if (!dryRun && lastRowForClear >= 2) {
          const clearCount = lastRowForClear - 1;
          sheet.getRange(2, 1, clearCount, 1).clearContent();
          sheet.getRange(2, 2, clearCount, 1).clearContent();
          sheet.getRange(2, 4, clearCount, 1).clearContent();
        }

        let classAdded = 0;
        for (const item of classRecordsFiltered) {
          const incomingDueDate = formatDueDate(item.dueDate);
          if (!dryRun) {
            const newRow = firstEmptyAssignmentRow(sheet);
            sheet.getRange(newRow, 1).setValue(item.assignmentName);
            sheet.getRange(newRow, 2).setValue(incomingDueDate);
            sheet.getRange(newRow, 4).setValue(className);
          }
          addedRows += 1;
          classAdded += 1;
        }

        classStats[className] = {
          incomingCount: classRecordsFiltered.length,
          existingNamedCount: 0,
          matchedCount: 0,
          addedCount: classAdded,
          updatedCount: 0,
          replaceMode: true,
        };

        updatedClasses.push(className);
        continue;
      }

      const lastRow = sheet.getLastRow();
      const rowCount = Math.max(0, lastRow - 1);
      const allExistingRows = rowCount > 0
        ? sheet.getRange(2, 1, rowCount, 4).getValues().map((row, index) => ({
            rowNumber: index + 2,
            assignmentName: String(row[0] || "").trim(),
            dueDate: formatDueDate(row[1]),
            dueDateKey: normalizeDueDateKey(row[1]),
            className: String(row[3] || "").trim(),
            matched: false,
          }))
        : [];

      const existingRows = allExistingRows.filter((row) => row.assignmentName !== "");
      let classAdded = 0;
      let classUpdated = 0;
      let classMatched = 0;

      for (const item of classRecordsFiltered) {
        const bestMatch = findBestMatchingRow(existingRows, item.assignmentName);
        const incomingDueDate = formatDueDate(item.dueDate);
        const incomingDueDateKey = normalizeDueDateKey(item.dueDate);

        if (bestMatch) {
          bestMatch.matched = true;
          classMatched += 1;

          if (bestMatch.dueDateKey !== incomingDueDateKey) {
            if (!dryRun) {
              sheet.getRange(bestMatch.rowNumber, 2).setValue(incomingDueDate);
            }
            debugMessages.push(
              `assignment ${item.assignmentName} date updated from ${bestMatch.dueDate || "(blank)"} to ${incomingDueDate || "(blank)"}`
            );
            bestMatch.dueDate = incomingDueDate;
            bestMatch.dueDateKey = incomingDueDateKey;
            updatedRows += 1;
            classUpdated += 1;
          }

          if (bestMatch.className !== className) {
            if (!dryRun) {
              sheet.getRange(bestMatch.rowNumber, 4).setValue(className);
            }
            bestMatch.className = className;
          }

          if (!bestMatch.assignmentName) {
            if (!dryRun) {
              sheet.getRange(bestMatch.rowNumber, 1).setValue(item.assignmentName);
            }
            bestMatch.assignmentName = item.assignmentName;
          }
        } else {
          if (!dryRun) {
            const newRow = firstEmptyAssignmentRow(sheet);
            sheet.getRange(newRow, 1).setValue(item.assignmentName);
            sheet.getRange(newRow, 2).setValue(incomingDueDate);
            sheet.getRange(newRow, 4).setValue(className);
          }
          addedRows += 1;
          classAdded += 1;
        }
      }

      classStats[className] = {
        incomingCount: classRecordsFiltered.length,
        existingNamedCount: existingRows.length,
        matchedCount: classMatched,
        addedCount: classAdded,
        updatedCount: classUpdated,
        replaceMode: false,
      };

      updatedClasses.push(className);
    }

    return jsonResponse({
      status: "success",
      spreadsheetId: doc.getId(),
      spreadsheetName: doc.getName(),
      dryRun,
      replaceExisting,
      updatedClasses,
      rowsWritten: addedRows + updatedRows,
      addedRows,
      updatedRows,
      classStats,
      debugMessages,
      requestedClasses: Object.keys(grouped),
    });
  }

  const textValue = (e && e.parameter && e.parameter.text) || "SUCCESS";
  return jsonResponse({
    status: "success",
    spreadsheetId: doc.getId(),
    spreadsheetName: doc.getName(),
    text: textValue,
  });
}

function compactName(value) {
  return String(value || "").replace(/\s+/g, "").toLowerCase();
}

function isExcludedTabName(name) {
  return EXCLUDED_TAB_NAMES.has(compactName(name));
}

function clearSheetAssignmentColumns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const clearCount = lastRow - 1;
  sheet.getRange(2, 1, clearCount, 1).clearContent();
  sheet.getRange(2, 2, clearCount, 1).clearContent();
  sheet.getRange(2, 4, clearCount, 1).clearContent();
  return clearCount;
}

function findBestMatchingRow(existingRows, assignmentName) {
  let bestRow = null;
  let bestScore = 0;

  for (const row of existingRows) {
    if (row.matched) continue;
    const score = similarityScore(row.assignmentName, assignmentName);
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }

  return bestScore >= 7 ? bestRow : null;
}

function similarityScore(existingName, incomingName) {
  const a = normalizeName(existingName);
  const b = normalizeName(incomingName);
  if (!a || !b) return 0;
  if (a === b) return 10;

  const aKey = buildAssignmentKey(a);
  const bKey = buildAssignmentKey(b);
  if (aKey && bKey && aKey === bKey) {
    return 9;
  }

  let score = 0;

  if (a.length >= 6 && b.length >= 6 && (a.includes(b) || b.includes(a))) {
    score += 3;
  }

  const aTokens = a.split(" ").filter(Boolean);
  const bTokens = b.split(" ").filter(Boolean);
  const bSet = new Set(bTokens);

  let overlap = 0;
  let numericOverlap = 0;
  for (const token of aTokens) {
    if (!bSet.has(token)) continue;
    overlap += 1;
    if (/^\d+$/.test(token)) numericOverlap += 1;
  }

  if (overlap >= 3) score += 4;
  else if (overlap === 2) score += 3;
  else if (overlap === 1) score += 0;

  if (numericOverlap >= 1) {
    score += 2;
  }

  return score;
}

function buildAssignmentKey(normalizedName) {
  const tokens = normalizedName.split(" ").filter(Boolean);
  const normalizedNumbers = tokens
    .filter((token) => /^\d+$/.test(token))
    .map((token) => String(Number(token)));

  const hasHw = tokens.includes("hw") || tokens.includes("homework");
  if (hasHw && normalizedNumbers.length >= 1) {
    return `hw:${normalizedNumbers.join("-")}`;
  }

  if (tokens.includes("attendance") && normalizedNumbers.length >= 2) {
    return `attendance:${normalizedNumbers[0]}-${normalizedNumbers[1]}`;
  }

  if (tokens.includes("quiz") && normalizedNumbers.length >= 1) {
    return `quiz:${normalizedNumbers[0]}`;
  }

  if (tokens.includes("exam") && normalizedNumbers.length >= 1) {
    return `exam:${normalizedNumbers[0]}`;
  }

  if (tokens.includes("problem") && normalizedNumbers.length >= 1) {
    return `problem:${normalizedNumbers.join("-")}`;
  }

  return "";
}

function normalizeName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/([a-z])(\d)/g, "$1 $2")
    .replace(/(\d)([a-z])/g, "$1 $2")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function parseMmDdYyyy(value) {
  const match = String(value || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return new Date(8640000000000000);

  const month = Number(match[1]) - 1;
  const day = Number(match[2]);
  const year = Number(match[3]);
  return new Date(year, month, day);
}

function parseDateValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const mmddyyyy = parseMmDdYyyy(text);
  if (!isNaN(mmddyyyy.getTime()) && mmddyyyy.getTime() !== 8640000000000000) {
    return mmddyyyy;
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
}

function normalizeDueDateKey(value) {
  const date = parseDateValue(value);
  if (!date) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDueDate(value) {
  const date = parseDateValue(value);
  if (!date) return String(value || "").trim();
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

function firstEmptyAssignmentRow(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rowCount = Math.max(1, lastRow - 1);
  const values = sheet.getRange(2, 1, rowCount, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === "") {
      return i + 2;
    }
  }

  return lastRow + 1;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeKey_(str) {
  return String(str || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function toCamelCase_(str) {
  if (!str) return "";
  const s = String(str).trim();
  if (/^[a-z][a-zA-Z0-9]*$/.test(s)) return s;
  return s
    .toLowerCase()
    .replace(/[^a-z0-9]+([a-z0-9])/g, (_, chr) => chr.toUpperCase())
    .replace(/[^a-zA-Z0-9]/g, "")
    .replace(/^[A-Z]/, (chr) => chr.toLowerCase());
}

function formatHeaderLabel_(key) {
  if (!key) return "";
  const acronyms = {
    id: "ID",
    uan: "UAN",
    pmc: "PMC",
    cbre: "CBRE",
    po: "PO",
    pono: "PO NO",
    dcno: "DC NO",
    dc: "DC",
    ppe: "PPE",
    tbt: "TBT",
    lti: "LTI",
    swms: "SWMS",
    hira: "HIRA",
    msds: "MSDS",
    sla: "SLA",
    slahours: "SLA HOURS",
    pdf: "PDF",
    pdffileid: "PDF FILE ID",
    docfileid: "DOC FILE ID",
    json: "JSON",
    payloadjson: "PAYLOAD JSON",
    safemanhours: "SAFE MAN HOURS",
    cumsafemanhours: "CUM SAFE MAN HOURS",
    totalmanhours: "TOTAL MAN HOURS",
    totalmanpower: "TOTAL MANPOWER",
    projectid: "PROJECT ID",
    employeeid: "EMPLOYEE ID",
    toemployeeid: "TO EMPLOYEE ID",
    formcode: "FORM CODE",
    entrytype: "ENTRY TYPE",
    createdat: "CREATED AT",
    createdby: "CREATED BY",
    submittedat: "SUBMITTED AT",
    submittedby: "SUBMITTED BY",
    reviewedat: "REVIEWED AT",
    reviewedby: "REVIEWED BY",
    dueat: "DUE AT",
    escalateat: "ESCALATE AT",
    auditdate: "AUDIT DATE",
    balancestock: "BALANCE STOCK",
    totalreceived: "TOTAL RECEIVED",
    workinghours: "WORKING HOURS",
    areasqft: "AREA (SQFT)",
    projectduration: "PROJECT DURATION",
    contractor: "CONTRACTOR",
    receivedby: "RECEIVED BY",
    issuedqty: "ISSUED QTY",
    totalreceivedqty: "TOTAL RECEIVED QTY",
    receivercontractor: "RECEIVER CONTRACTOR",
    receivername: "RECEIVER NAME",
    issuedate: "ISSUE DATE",
  };
  const norm = normalizeKey_(key);
  if (acronyms[norm]) return acronyms[norm];
  return String(key)
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_\-]+/g, " ")
    .trim()
    .toUpperCase();
}

let _cachedSs = null;

function ss_() {
  if (_cachedSs) {
    try {
      _cachedSs.getId();
      return _cachedSs;
    } catch (e) {
      _cachedSs = null;
    }
  }
  const id =
    String(DATABASE_SPREADSHEET_ID || "").trim() ||
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id)
    throw new Error(
      "System not initialized. Run initializeSystem() from the Apps Script editor.",
    );
  _cachedSs = SpreadsheetApp.openById(id);
  return _cachedSs;
}

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh && HEADERS[name]) {
    sh = ss.insertSheet(name);
    const cols = HEADERS[name];
    const headerLabels = cols.map((h) => formatHeaderLabel_(h));
    sh.getRange(1, 1, 1, cols.length).setValues([headerLabels]);
  }
  if (!sh)
    throw new Error(
      "Missing sheet: " +
        name +
        ". Run initializeSystem() to repair the database.",
    );
  return sh;
}

function uid_(prefix) {
  return (
    (prefix || "ID") +
    "_" +
    Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase()
  );
}

function nowIso_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );
}

function todayIso_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
}

function todayDisplay_() {
  return displayDate_(todayIso_());
}

function displayDate_(value) {
  const text = String(value == null ? "" : value).trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})(.*)$/);
  return match ? match[3] + "/" + match[2] + "/" + match[1] + match[4] : text;
}

function rowsToObjects_(sheetName) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row.join("").trim()) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, idx) => {
      const val = row[idx];
      obj[h] = val;
      const camel = toCamelCase_(h);
      if (camel && obj[camel] === undefined) obj[camel] = val;
      const norm = normalizeKey_(h);
      if (norm && obj[norm] === undefined) obj[norm] = val;
    });
    out.push(obj);
  }
  return out;
}

function appendRow_(sheetName, obj) {
  const sh = sheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map((h) => {
    if (obj[h] !== undefined && obj[h] !== null) return obj[h];
    const camel = toCamelCase_(h);
    if (camel && obj[camel] !== undefined && obj[camel] !== null)
      return obj[camel];
    const normH = normalizeKey_(h);
    const matchKey = Object.keys(obj).find((k) => normalizeKey_(k) === normH);
    if (matchKey && obj[matchKey] !== undefined && obj[matchKey] !== null)
      return obj[matchKey];
    return "";
  });
  sh.appendRow(row);
}

function updateRowByKey_(sheetName, keyName, keyValue, patch) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return false;
  const headers = values[0];
  let keyCol = headers.indexOf(keyName);
  if (keyCol < 0) {
    const normKey = normalizeKey_(keyName);
    keyCol = headers.findIndex((h) => normalizeKey_(h) === normKey);
  }
  if (keyCol < 0) {
    throw new Error(
      'Key column "' + keyName + '" not found on sheet "' + sheetName + '"',
    );
  }
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]).trim() === String(keyValue).trim()) {
      Object.keys(patch).forEach((k) => {
        let c = headers.indexOf(k);
        if (c < 0) {
          const normK = normalizeKey_(k);
          c = headers.findIndex((h) => normalizeKey_(h) === normK);
        }
        if (c >= 0) sh.getRange(i + 1, c + 1).setValue(patch[k]);
      });
      return true;
    }
  }
  return false;
}

function updateRowById_(sheetName, id, patch) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return false;
  const headers = values[0];
  let idCol = headers.indexOf("id");
  if (idCol < 0) {
    const norm = headers.map((h) => normalizeKey_(h));
    const candidates = ["id", "formcode", "employeeid", "reportno", "code"];
    for (let cand of candidates) {
      const idx = norm.indexOf(cand);
      if (idx >= 0) {
        idCol = idx;
        break;
      }
    }
  }
  if (idCol < 0) {
    if (typeof id === "number" && id >= 2 && id <= values.length) {
      Object.keys(patch).forEach((k) => {
        let c = headers.indexOf(k);
        if (c < 0)
          c = headers.findIndex((h) => normalizeKey_(h) === normalizeKey_(k));
        if (c >= 0) sh.getRange(id, c + 1).setValue(patch[k]);
      });
      return true;
    }
    throw new Error("No identifier column on " + sheetName);
  }
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]).trim() === String(id).trim()) {
      Object.keys(patch).forEach((k) => {
        let c = headers.indexOf(k);
        if (c < 0)
          c = headers.findIndex((h) => normalizeKey_(h) === normalizeKey_(k));
        if (c >= 0) sh.getRange(i + 1, c + 1).setValue(patch[k]);
      });
      return true;
    }
  }
  return false;
}

function findBy_(sheetName, key, value) {
  return rowsToObjects_(sheetName).filter(
    (r) => String(r[key]) === String(value),
  );
}

function findOne_(sheetName, key, value) {
  return findBy_(sheetName, key, value)[0] || null;
}

function writeAudit_(employeeId, action, entityType, entityId, detail) {
  appendRow_(SHEETS.AUDIT_LOG, {
    id: uid_("LOG"),
    at: nowIso_(),
    employeeId: employeeId || "",
    action: action,
    entityType: entityType || "",
    entityId: entityId || "",
    detail: detail || "",
  });
}

function jsonSafe_(v) {
  if (typeof v === "string") {
    try {
      return JSON.parse(v);
    } catch (e) {
      return v;
    }
  }
  return v;
}

function cleanAndAlignSheet_(sh, sheetName, canonicalHeaders) {
  if (!sh || !canonicalHeaders || !canonicalHeaders.length) return;
  const numCanonical = canonicalHeaders.length;
  const canonicalFormatted = canonicalHeaders.map((h) => formatHeaderLabel_(h));
  const canonicalNorm = canonicalHeaders.map((h) => normalizeKey_(h));

  // Ensure sheet has enough columns for canonical headers
  const currentMaxCols = sh.getMaxColumns();
  if (currentMaxCols < numCanonical) {
    sh.insertColumnsAfter(currentMaxCols, numCanonical - currentMaxCols);
  }

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  // If brand new or empty sheet
  if (lastCol === 0 || lastRow === 0) {
    sh.clearContents();
    sh.getRange(1, 1, 1, numCanonical).setValues([canonicalFormatted]);
    formatSheetProfessionally_(sh, sheetName, canonicalHeaders);
    return;
  }

  // Read existing header row
  const existingHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const existingNorm = existingHeaders.map((h) => normalizeKey_(h));

  // Check if existing headers match canonical headers in exact order and length
  const isIdentical =
    lastCol === numCanonical &&
    canonicalNorm.every((cn, idx) => cn === existingNorm[idx]);

  if (!isIdentical) {
    // There are extra, duplicate, or misaligned columns
    const cleanedData = [];
    if (lastRow > 1) {
      const allValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (let r = 0; r < allValues.length; r++) {
        const row = allValues[r];
        if (!row.join("").trim()) continue; // skip phantom blank rows
        const newRow = canonicalHeaders.map((ch) => {
          const normCh = normalizeKey_(ch);
          const idx = existingNorm.indexOf(normCh);
          return idx >= 0 && row[idx] !== undefined && row[idx] !== null
            ? row[idx]
            : "";
        });
        cleanedData.push(newRow);
      }
    }

    // Clear contents and re-write canonical header and clean aligned rows
    sh.clearContents();
    sh.getRange(1, 1, 1, numCanonical).setValues([canonicalFormatted]);

    if (cleanedData.length > 0) {
      const currentMaxRows = sh.getMaxRows();
      const requiredRows = cleanedData.length + 1;
      if (currentMaxRows < requiredRows) {
        sh.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
      }
      sh.getRange(2, 1, cleanedData.length, numCanonical).setValues(
        cleanedData,
      );
    }

    // Delete any excess columns beyond canonical
    const maxCols = sh.getMaxColumns();
    if (maxCols > numCanonical) {
      try {
        sh.deleteColumns(numCanonical + 1, maxCols - numCanonical);
      } catch (e) {}
    }
  } else {
    // Headers already match, just ensure capitalized labels are set
    sh.getRange(1, 1, 1, numCanonical).setValues([canonicalFormatted]);
  }

  // Trim excessive trailing empty rows beyond data
  const currentLastRow = Math.max(sh.getLastRow(), 1);
  const maxRows = sh.getMaxRows();
  if (maxRows > currentLastRow + 25) {
    try {
      sh.deleteRows(currentLastRow + 26, maxRows - (currentLastRow + 25));
    } catch (e) {}
  }

  formatSheetProfessionally_(sh, sheetName, canonicalHeaders);
}

function batchWriteObjects_(sh, headerList, objects) {
  if (!sh || !headerList || !headerList.length || !objects || !objects.length)
    return;
  const numCols = headerList.length;
  const matrix = objects.map((obj) => {
    return headerList.map((h) => {
      if (obj[h] !== undefined && obj[h] !== null) return obj[h];
      const camel = toCamelCase_(h);
      if (camel && obj[camel] !== undefined && obj[camel] !== null)
        return obj[camel];
      const normH = normalizeKey_(h);
      const matchKey = Object.keys(obj).find((k) => normalizeKey_(k) === normH);
      if (matchKey && obj[matchKey] !== undefined && obj[matchKey] !== null)
        return obj[matchKey];
      return "";
    });
  });

  const currentMaxRows = sh.getMaxRows();
  const requiredRows = matrix.length + 1;
  if (currentMaxRows < requiredRows) {
    sh.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
  }
  sh.getRange(2, 1, matrix.length, numCols).setValues(matrix);
}

function formatSheetProfessionally_(sh, sheetName, headers) {
  if (!sh || !headers || !headers.length) return;
  const numCols = headers.length;

  // 1. Column dimension guard: ensure sheet has at least numCols
  const currentMaxCols = sh.getMaxColumns();
  if (currentMaxCols < numCols) {
    sh.insertColumnsAfter(currentMaxCols, numCols - currentMaxCols);
  }

  // 2. Remove excess trailing columns beyond canonical schema
  if (sh.getMaxColumns() > numCols) {
    try {
      sh.deleteColumns(numCols + 1, sh.getMaxColumns() - numCols);
    } catch (e) {}
  }

  const lastRow = Math.max(sh.getLastRow(), 1);

  // 3. Freeze Header Row 1
  sh.setFrozenRows(1);

  // 4. Highlighted Corporate Header: Deep SESIPL Teal (#0f766e), Bold White Text, 38px height, Centered
  const formattedHeaders = headers.map((h) => formatHeaderLabel_(h));
  const headerRange = sh.getRange(1, 1, 1, numCols);
  headerRange
    .setValues([formattedHeaders])
    .setBackground("#0f766e")
    .setFontColor("#ffffff")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontWeight("bold")
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sh.setRowHeight(1, 38);

  // 5. Tab Color: Matching SESIPL Corporate Teal
  try {
    sh.setTabColor("#0f766e");
  } catch (e) {}

  // 6. AutoFilter on Header Row across all columns
  const existingFilter = sh.getFilter();
  if (existingFilter) existingFilter.remove();
  try {
    sh.getRange(1, 1, Math.max(lastRow, 2), numCols).createFilter();
  } catch (e) {}

  // 7. Data Rows Formatting
  if (lastRow > 1) {
    const numDataRows = lastRow - 1;
    // Spacious corporate row height: 28px
    sh.setRowHeights(2, numDataRows, 28);
    const dataRange = sh.getRange(2, 1, numDataRows, numCols);

    // Remove any alternating zebra banding so body remains clean & plain white
    const bandings = sh.getBandings();
    bandings.forEach((b) => {
      try {
        b.remove();
      } catch (err) {}
    });

    // Plain corporate styling: Pure white background, crisp dark slate font
    dataRange
      .setBackground("#ffffff")
      .setFontFamily("Arial")
      .setFontSize(10)
      .setFontColor("#1e293b")
      .setFontWeight("normal")
      .setVerticalAlignment("middle");

    // Subtle modern grid borders (#e2e8f0)
    try {
      dataRange.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#e2e8f0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
    } catch (e) {}

    // Column-specific alignment and number formatting
    for (let c = 1; c <= numCols; c++) {
      const h = headers[c - 1];
      const normH = normalizeKey_(h);
      const colDataRange = sh.getRange(2, c, numDataRows, 1);

      // Plain text format (@) for IDs, phones, codes to preserve exact characters
      if (
        ["employeeid", "uan", "phone", "pono", "reportno", "dcno"].indexOf(
          normH,
        ) >= 0
      ) {
        colDataRange.setNumberFormat("@").setHorizontalAlignment("center");
      } else if (
        [
          "id",
          "code",
          "projectid",
          "level",
          "version",
          "active",
          "returnable",
          "cadence",
          "entrytype",
          "slahours",
        ].indexOf(normH) >= 0
      ) {
        colDataRange.setNumberFormat("@").setHorizontalAlignment("center");
      } else if (normH === "status") {
        colDataRange.setHorizontalAlignment("center").setFontWeight("bold");
      } else if (normH.indexOf("date") >= 0) {
        colDataRange
          .setHorizontalAlignment("center")
          .setNumberFormat("yyyy-MM-dd");
      } else if (normH.indexOf("at") >= 0 || normH === "at") {
        colDataRange
          .setHorizontalAlignment("center")
          .setNumberFormat("yyyy-MM-dd HH:mm");
      } else if (normH === "percent") {
        colDataRange.setHorizontalAlignment("right").setNumberFormat('0"%"');
      } else if (normH === "grade") {
        colDataRange.setHorizontalAlignment("center").setFontWeight("bold");
      } else if (
        [
          "staff",
          "workers",
          "totalmanpower",
          "manhours",
          "safemanhours",
          "cumsafemanhours",
          "totalmanhours",
          "inductions",
          "indstaff",
          "indworkers",
          "tbtcount",
          "tbtpersons",
          "trainingpersons",
          "totalscaffold",
          "totalladder",
          "totalreceived",
          "balancestock",
          "totalscore",
          "maxscore",
          "workinghours",
          "permithot",
          "permitelectrical",
          "permitcold",
          "permitgeneral",
          "permitothers",
          "firstaid",
          "nearmiss",
          "lticount",
          "areasqft",
          "issuedqty",
          "totalreceivedqty",
          "returnedscaffold",
          "returnedladder",
        ].indexOf(normH) >= 0
      ) {
        colDataRange.setHorizontalAlignment("right").setNumberFormat("#,##0");
      }

      // Text wrapping strategies
      if (
        [
          "observation",
          "preventive",
          "remarks",
          "scope",
          "payloadjson",
          "address",
          "accidentdetails",
          "defects",
          "highlights",
          "notes",
          "detail",
          "permanentaddress",
          "presentaddress",
        ].indexOf(normH) >= 0
      ) {
        colDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      } else if (
        [
          "fileid",
          "unsafefileid",
          "rectifiedfileid",
          "pdffileid",
          "docfileid",
          "url",
        ].indexOf(normH) >= 0
      ) {
        colDataRange
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
          .setHorizontalAlignment("center");
      }
    }

    // ==========================================
    // SPECIAL CASES WITH CREATIVE EHS STYLING
    // ==========================================
    const rules = [];

    // Special Case 1: Status column (Executive soft pill styling)
    const statusIdx = headers.findIndex((h) => normalizeKey_(h) === "status");
    if (statusIdx >= 0) {
      const statusRange = sh.getRange(2, statusIdx + 1, numDataRows, 1);
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("APPROVED")
          .setBackground("#ecfdf5")
          .setFontColor("#047857")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("REJECTED")
          .setBackground("#fff1f2")
          .setFontColor("#be123c")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("OPEN")
          .setBackground("#fffbeb")
          .setFontColor("#b45309")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("PENDING")
          .setBackground("#fffbeb")
          .setFontColor("#b45309")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("SUBMITTED")
          .setBackground("#f0f9ff")
          .setFontColor("#0369a1")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("CLOSED")
          .setBackground("#f8fafc")
          .setFontColor("#475569")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("RUNNING")
          .setBackground("#ecfdf5")
          .setFontColor("#047857")
          .setRanges([statusRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("ACTIVE")
          .setBackground("#ecfdf5")
          .setFontColor("#047857")
          .setRanges([statusRange])
          .build(),
      );
    }

    // Special Case 2: Cumulative Safe Man-Hours (Prominent safety milestone in pale teal)
    const cumSafeIdx = headers.findIndex(
      (h) => normalizeKey_(h) === "cumsafemanhours",
    );
    if (cumSafeIdx >= 0) {
      const cumSafeRange = sh.getRange(2, cumSafeIdx + 1, numDataRows, 1);
      cumSafeRange
        .setBackground("#f0fdfa")
        .setFontColor("#0f766e")
        .setFontWeight("bold");
    }

    // Special Case 3: Lost Time Injuries (Zero Harm Celebration vs Accident Alert)
    const ltiIdx = headers.findIndex((h) => normalizeKey_(h) === "lticount");
    if (ltiIdx >= 0) {
      const ltiRange = sh.getRange(2, ltiIdx + 1, numDataRows, 1);
      ltiRange.setFontWeight("bold");
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberEqualTo(0)
          .setBackground("#f0fdf4")
          .setFontColor("#15803d")
          .setRanges([ltiRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground("#fee2e2")
          .setFontColor("#991b1b")
          .setRanges([ltiRange])
          .build(),
      );
    }

    // Special Case 4: Audit Grades & Percentages
    const gradeIdx = headers.findIndex((h) => normalizeKey_(h) === "grade");
    if (gradeIdx >= 0) {
      const gradeRange = sh.getRange(2, gradeIdx + 1, numDataRows, 1);
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("A")
          .setBackground("#ecfdf5")
          .setFontColor("#047857")
          .setRanges([gradeRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("B")
          .setBackground("#fffbeb")
          .setFontColor("#b45309")
          .setRanges([gradeRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("C")
          .setBackground("#fff1f2")
          .setFontColor("#be123c")
          .setRanges([gradeRange])
          .build(),
      );
    }

    const percentIdx = headers.findIndex((h) => normalizeKey_(h) === "percent");
    if (percentIdx >= 0) {
      const pctRange = sh.getRange(2, percentIdx + 1, numDataRows, 1);
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(90)
          .setBackground("#f0fdf4")
          .setFontColor("#166534")
          .setRanges([pctRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(75)
          .setBackground("#fffbeb")
          .setFontColor("#b45309")
          .setRanges([pctRange])
          .build(),
      );
    }

    // Special Case 5: Active users flag
    const activeIdx = headers.findIndex((h) => normalizeKey_(h) === "active");
    if (activeIdx >= 0) {
      const activeRange = sh.getRange(2, activeIdx + 1, numDataRows, 1);
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("TRUE")
          .setBackground("#ecfdf5")
          .setFontColor("#047857")
          .setRanges([activeRange])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("FALSE")
          .setBackground("#fff1f2")
          .setFontColor("#be123c")
          .setRanges([activeRange])
          .build(),
      );
    }

    if (rules.length > 0) {
      sh.setConditionalFormatRules(rules);
    }
  }

  // 8. Trim excessive empty rows to keep the sheet compact, clean and fast
  const maxRows = sh.getMaxRows();
  const keepRows = Math.max(lastRow + 20, 25);
  if (maxRows > keepRows) {
    try {
      sh.deleteRows(keepRows + 1, maxRows - keepRows);
    } catch (e) {}
  }

  // 9. Intelligent Corporate Column Widths (Spacious, Clear & Readable)
  try {
    sh.autoResizeColumns(1, numCols);
  } catch (e) {}

  for (let c = 1; c <= numCols; c++) {
    const h = headers[c - 1];
    const normH = normalizeKey_(h);
    let w = sh.getColumnWidth(c);
    if (w < 120) w = 120; // Minimum 120px for ample spacing

    if (
      ["observation", "preventive", "payloadjson", "accidentdetails"].indexOf(
        normH,
      ) >= 0
    ) {
      w = 350;
    } else if (
      [
        "remarks",
        "scope",
        "defects",
        "highlights",
        "notes",
        "detail",
        "permanentaddress",
        "presentaddress",
      ].indexOf(normH) >= 0
    ) {
      w = 280;
    } else if (
      [
        "name",
        "title",
        "client",
        "pmc",
        "contractor",
        "email",
        "topic",
        "trainingtopic",
      ].indexOf(normH) >= 0
    ) {
      w = Math.max(w, 200);
    } else if (
      [
        "status",
        "region",
        "level",
        "version",
        "grade",
        "active",
        "cadence",
        "entrytype",
      ].indexOf(normH) >= 0
    ) {
      w = 130;
    } else if (normH.indexOf("date") >= 0) {
      w = 125;
    } else if (normH.indexOf("at") >= 0 || normH === "at") {
      w = 155;
    } else if (
      [
        "employeeid",
        "uan",
        "phone",
        "pono",
        "reportno",
        "dcno",
        "id",
        "code",
        "projectid",
      ].indexOf(normH) >= 0
    ) {
      w = Math.max(w, 130);
    }
    if (w > 400) w = 400;
    sh.setColumnWidth(c, w);
  }
}

function deduplicateSheet_(sheetName, keyFields) {
  const sh = sheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length <= 2) return 0;
  const headers = values[0];
  const normHeaders = headers.map((h) => normalizeKey_(h));
  const keyIndices = keyFields
    .map((k) => normHeaders.indexOf(normalizeKey_(k)))
    .filter((idx) => idx >= 0);
  if (!keyIndices.length) return 0;

  const seen = {};
  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const key = keyIndices
      .map((idx) => String(row[idx] || "").trim())
      .join(":::");
    if (!key || key === ":::") continue;
    if (seen[key]) {
      rowsToDelete.push(i + 1);
    } else {
      seen[key] = true;
    }
  }

  // Delete from bottom to top to avoid shifting row indices
  for (let r = rowsToDelete.length - 1; r >= 0; r--) {
    sh.deleteRow(rowsToDelete[r]);
  }
  if (rowsToDelete.length) {
    Logger.log(
      "Deduplicated " + rowsToDelete.length + " rows from " + sheetName,
    );
  }
  return rowsToDelete.length;
}

function formatAllSheets_() {
  const ss = ss_();
  Object.keys(HEADERS).forEach((name) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    cleanAndAlignSheet_(sh, name, HEADERS[name]);
  });
}

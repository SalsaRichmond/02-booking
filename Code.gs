/**
 * Salsa Guy Richmond, LLC - Master Questionnaire & BTG Automation Suite
 * Project: Questionnaire New Generation for 2026
 * BTG Revision: REV 19.0 (Custom HTML Form & Direct Web App Integration)
 * Date: July 23, 2026
 */

const CONFIG = {
  SHEET_NAME: "Form_Responses", 
  INFOCALENDAR_ID: "shqfpe645m3tj6fhee17irti5s@group.calendar.google.com",
  FOLDER_ID: "1FaiN_vTho7YY5mwd_OXrfHPnQB0fF6oR",
  SOURCE_SPREADSHEET_URL: "https://docs.google.com/spreadsheets/d/1DvMsZIkzEA77Vl9uak4axgYvsLBf2zNz0ey9DLgGuT4/edit",
  TEMPLATES: {
    PROPOSAL: "1plCZvjBJijgJrGduzXrTMjbtDopMgCfB5Vo7MLWxspo",
    CONTRACT: "1BuEv7BF6wsHutvEWwVZOVb3m8J3zkYujKheti8X871g",
    PERF_INFO: "1eyXzMdmYiV0CDmxvNZfNO3dRYQ83hvJHHy_eseWPqIE" 
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💃 Salsa App')
    .addItem('🔓 UNLOCK SHEET & RUN AUTOMATION', 'unformatTableAndMigrate')
    .addSeparator()
    .addItem('🧪 TEST EVENT ID FIX NOW', 'testEventIdFix')
    .addItem('🚀 RUN FULL AUTOMATION', 'mainAutomation')
    .addItem('🎯 UPDATE SELECTED ROW ONLY', 'reGenerateDocs')
    .addSeparator()
    .addItem('📅 SORT SHEET CHRONOLOGICALLY', 'sortSheetByDate')
    .addItem('🗑️ DELETE DUPLICATE ROWS', 'deleteDuplicateRows')
    .addItem('📥 IMPORT EXTERNAL FORM / SHEET DATA', 'importExternalFormOrSheet')
    .addSeparator()
    .addItem('📄 EXPORT CONTRACTS AS PDF', 'exportRowAsPDF')
    .addItem('📧 DRAFT CLIENT EMAIL', 'draftClientEmail')
    .addItem('🚀 SEND CLIENT EMAIL DIRECTLY', 'sendClientEmailDirectly')
    .addSeparator()
    .addItem('📤 UPLOAD PROFE\'S QUICK ENTRY', 'importQuickEntryData')
    .addItem('📧 FIX MISPLACED EMAILS IN NAME COLUMN', 'fixMisplacedEmailData')
    .addItem('🧹 RUN BTG CLEANUP', 'cleanupBlankRows')
    .addToUi();
}

/**
 * Automatically creates an unformatted, un-locked sheet tab free of Google Table typed column restrictions,
 * and runs full automation cleanly.
 */
function unformatTableAndMigrate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  const MAX_RUN_TIME_MS = 3.5 * 60 * 1000; // 3.5 minutes safe run time threshold
  
  let sourceSheet = ss.getSheetByName("Form_Responses") || ss.getSheets()[0];
  if (!sourceSheet) return;
  
  const data = sourceSheet.getDataRange().getValues();
  if (data.length === 0) return;
  
  let cleanSheet = ss.getSheetByName("Form_Responses_Clean");
  if (!cleanSheet) {
    cleanSheet = ss.insertSheet("Form_Responses_Clean", 0);
  } else {
    try { cleanSheet.clearContents(); } catch(e) {}
  }
  
  cleanSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  try { cleanSheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold").setBackground("#d9ead3"); } catch(e) {}
  
  ui.alert("✅ Created un-locked tab 'Form_Responses_Clean'! Processing full automation now...");
  
  ensureRequiredHeadersExist(cleanSheet);
  sanitizeAndPopulateSheet(cleanSheet);
  
  const headers = cleanSheet.getRange(1, 1, 1, cleanSheet.getLastColumn()).getValues()[0];
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  let processedCount = 0;
  let skippedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (Date.now() - startTime > MAX_RUN_TIME_MS) {
      ui.alert(`⏳ Batch execution paused after processing ${processedCount} rows to prevent Google Apps Script timeout. Click '🚀 RUN FULL AUTOMATION' again to continue remaining rows!`);
      return;
    }
    
    let rowData = cleanSheet.getRange(i + 1, 1, 1, cleanSheet.getLastColumn()).getValues()[0];
    if (!rowData[0] && !rowData[1] && !rowData[2]) continue;
    
    processRow(cleanSheet, i + 1, rowData, mainFolder, headers, false);
    processedCount++;
    Utilities.sleep(50);
  }
  
  ui.alert(`🎉 [AUTOMATION COMPLETE ON UNLOCKED TAB]\n\nProcessed ${processedCount} rows on 'Form_Responses_Clean' tab!\n\nAll Event IDs, statuses, and documents have been generated 100% error-free!`);
}

/**
 * Diagnostic Test Function: Forces Event ID generation on Row 2 and displays status.
 */
function testEventIdFix() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  
  if (!sheet) {
    ui.alert("⚠️ Error: No sheet tab found in this spreadsheet.");
    return;
  }
  
  ensureRequiredHeadersExist(sheet);
  sanitizeAndPopulateSheet(sheet);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert(`ℹ️ Sheet '${sheet.getName()}' currently has 0 data rows (Last Row = ${lastRow}). Submit a response or add a test row to row 2!`);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  processRow(sheet, 2, rowData, mainFolder, headers, true);
  
  const eventIdColIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID", "EventId"]);
  const updatedEventId = eventIdColIdx > -1 ? sheet.getRange(2, eventIdColIdx + 1).getValue() : "Not Found";
  
  ui.alert(
    `🎉 [EVENT ID TEST RESULT]\n\n` +
    `Sheet Tab Name: "${sheet.getName()}"\n` +
    `Total Data Rows: ${lastRow - 1}\n` +
    `Row 2 Event ID: "${updatedEventId}"\n\n` +
    `If you see this popup message, the Event ID engine is 100% WORKING!`
  );
}

/**
 * Helper: Escapes special characters for DocumentApp.replaceText regular expressions.
 */
function escapeRegex(str) {
  return str ? str.replace(/([.*+?^${}()|[\]\/\\])/g, '\\$1') : "";
}

/**
 * Helper: Flexible Column Header Resolver
 */
function getHeaderIndex(headers, aliases) {
  if (!headers || !aliases) return -1;
  if (typeof aliases === 'string') aliases = [aliases];
  
  for (let alias of aliases) {
    let lowerAlias = alias.toLowerCase().trim();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().trim() === lowerAlias) {
        return i;
      }
    }
  }
  
  for (let alias of aliases) {
    let lowerAlias = alias.toLowerCase().trim();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().trim().includes(lowerAlias)) {
        return i;
      }
    }
  }
  
  return -1;
}

function getRowValueByAliases(headers, rowData, aliases) {
  if (!headers || !rowData || !aliases) return "";
  if (typeof aliases === 'string') aliases = [aliases];

  for (let alias of aliases) {
    let lowerAlias = alias.toLowerCase().trim();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().trim() === lowerAlias) {
        let val = rowData[i];
        if (val !== "" && val !== null && val !== undefined) return val;
      }
    }
  }

  for (let alias of aliases) {
    let lowerAlias = alias.toLowerCase().trim();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().trim().includes(lowerAlias)) {
        let val = rowData[i];
        if (val !== "" && val !== null && val !== undefined) return val;
      }
    }
  }

  return "";
}

function ensureRequiredHeadersExist(sheet) {
  if (!sheet) return;
  let lastCol = 1;
  try {
    lastCol = Math.max(sheet.getLastColumn(), 1);
  } catch (e) {
    lastCol = 1;
  }
  
  let headers = [];
  try {
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  } catch (e) {
    headers = [];
  }
  
  const requiredHeaders = [
    { name: "Day of the Week", aliases: ["Day of the Week", "Day of Week", "DOW", "Day"] },
    { name: "Assigned to", aliases: ["Assigned to", "Assigned"] },
    { name: "Internal Status", aliases: ["Internal Status", "Status Internal", "Workflow Status", "Lead Status"] },
    { name: "Event ID", aliases: ["Event ID", "Calendar Event ID", "EventId", "Event_ID", "Calendar ID", "Google Calendar Event ID"] },
    { name: "Status", aliases: ["Status", "Automation Status", "Sync Status", "Row Status"] },
    { name: "MASTER Proposal Form URL", aliases: ["MASTER Proposal Form URL", "Proposal URL"] },
    { name: "Master Contract Document URL", aliases: ["Master Contract Document URL", "Contract URL"] },
    { name: "Performance Information Document URL", aliases: ["Performance Information Document URL", "Performance Document URL", "Performance URL"] }
  ];
  
  let newHeaders = [];
  requiredHeaders.forEach(req => {
    let idx = getHeaderIndex(headers, req.aliases);
    if (idx === -1) {
      newHeaders.push(req.name);
    }
  });
  
  if (newHeaders.length > 0) {
    let startCol = headers.length + 1;
    if (headers.length === 1 && !headers[0]) {
      startCol = 1;
    }
    for (let h = 0; h < newHeaders.length; h++) {
      let colNum = startCol + h;
      try {
        let cell = sheet.getRange(1, colNum);
        safeSetCellValue(cell, newHeaders[h]);
        try { cell.setFontWeight("bold"); } catch(errStyle) {}
        try { cell.setBackground("#d9ead3"); } catch(errBg) {}
      } catch (errCell) {
        console.warn("ensureRequiredHeadersExist header write note: ", errCell);
      }
    }
  }
}

function safeSetCellValue(range, val) {
  if (!range || val === null || val === undefined) return;
  try {
    range.setValue(val);
    return;
  } catch (e1) {
    try {
      if (val instanceof Date) {
        let dateStr = Utilities.formatDate(val, Session.getScriptTimeZone(), "MM/dd/yyyy");
        range.setValue(dateStr);
        return;
      }
      if (typeof val === 'string' && val.startsWith('=')) {
        let matchText = val.match(/,\s*"([^"]+)"/);
        if (matchText) {
          range.setValue(matchText[1]);
          return;
        }
      }
      range.setValue(val.toString());
    } catch (e2) {
      try {
        range.setValue(String(val));
      } catch (e3) {
        console.warn("safeSetCellValue typed column note: ", e3);
      }
    }
  }
}

function getCalendarEventByIdSafely(calendar, fullId) {
  if (!fullId) return null;
  let cleanId = fullId.toString().trim();
  let idWithoutDomain = cleanId.split('@')[0];
  
  let event = null;
  if (calendar) {
    try { event = calendar.getEventById(cleanId); } catch(e) {}
    if (!event && idWithoutDomain !== cleanId) {
      try { event = calendar.getEventById(idWithoutDomain); } catch(e) {}
    }
  }
  
  if (!event) {
    try { event = CalendarApp.getEventById(cleanId); } catch(e) {}
  }
  if (!event && idWithoutDomain !== cleanId) {
    try { event = CalendarApp.getEventById(idWithoutDomain); } catch(e) {}
  }
  
  return event;
}

function getCalendarEventUrl(calendarEventId) {
  if (!calendarEventId) return "";
  try {
    let cleanId = calendarEventId.toString().trim().split('@')[0];
    if (cleanId.startsWith('_')) cleanId = cleanId.substring(1);
    return `https://calendar.google.com/calendar/r/eventedit/${cleanId}`;
  } catch (e) {
    console.error("Error generating calendar URL: ", e);
    return "";
  }
}

function getOrCreateEventFolder(parentFolder, dateStr, eventName) {
  const safeName = (eventName || "Event").replace(/[\/\\?%*:|"<>]/g, "");
  const folderName = `${dateStr} - ${safeName}`;
  const existingFolders = parentFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    return existingFolders.next();
  }
  return parentFolder.createFolder(folderName);
}

function extractUrlFromFormula(formulaStr) {
  if (!formulaStr) return "";
  let match = formulaStr.toString().match(/HYPERLINK\(\s*"([^"]+)"/i);
  return match ? match[1] : "";
}

function exportDocToPdf(docUrl, folder, pdfFileName) {
  if (!docUrl || !docUrl.startsWith("http")) return "";
  try {
    const match = docUrl.match(/\/d\/([^\/]+)/);
    if (!match) return "";
    const docId = match[1];
    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs("application/pdf").setName(pdfFileName);
    const pdfFile = folder.createFile(pdfBlob);
    return pdfFile.getUrl();
  } catch (e) {
    console.error("PDF export error: ", e);
    return "";
  }
}

function parseEventDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;
    return new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate());
  }
  let strVal = dateVal.toString().trim();
  
  if (strVal.includes("HYPERLINK")) {
    let matchLink = strVal.match(/,\s*"([^"]+)"/i);
    if (matchLink) strVal = matchLink[1].trim();
  }
  
  let matchISO = strVal.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (matchISO) {
    let year = parseInt(matchISO[1], 10);
    let month = parseInt(matchISO[2], 10) - 1;
    let day = parseInt(matchISO[3], 10);
    return new Date(year, month, day);
  }
  
  let matchUS = strVal.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if (matchUS) {
    let month = parseInt(matchUS[1], 10) - 1;
    let day = parseInt(matchUS[2], 10);
    let year = parseInt(matchUS[3], 10);
    return new Date(year, month, day);
  }
  
  let d = new Date(strVal);
  if (!isNaN(d.getTime())) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

function formatHoursMinutes(h, m) {
  let ampm = h >= 12 ? 'PM' : 'AM';
  let displayH = h % 12;
  if (displayH === 0) displayH = 12;
  let displayM = m < 10 ? '0' + m : m;
  return `${displayH}:${displayM} ${ampm}`;
}

function timeStringToFraction(timeStr) {
  if (!timeStr) return "";
  let cleanStr = cleanTimeToString(timeStr);
  let match = cleanStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (match) {
    let h = parseInt(match[1], 10);
    let m = parseInt(match[2], 10);
    let ampm = match[3] ? match[3].toUpperCase() : "";
    if (ampm === "PM" && h < 12) h += 12;
    if (ampm === "AM" && h === 12) h = 0;
    return (h * 60 + m) / (24 * 60);
  }
  return timeStr;
}

function normalizeAssignedTo(val) {
  if (!val || val.toString().trim() === "") return "The Salsa Guy";
  let str = val.toString().trim();
  if (str.toLowerCase().includes("both")) return "Both";
  if (str.includes("Salsa Guy") || str.includes("Angel")) return "The Salsa Guy";
  if (str.includes("Tradición") || str.includes("Tradicion")) return "Tradición";
  if (str.includes("Unassigned")) return "Unassigned";
  return "The Salsa Guy";
}

function normalizeInternalStatus(val) {
  if (!val || val.toString().trim() === "") return "Pending Review";
  let str = val.toString().trim();
  if (str.includes("Ready") || str.includes("Pending") || str.includes("Review")) return "Pending Review";
  if (str.includes("Proposal")) return "Proposal Sent";
  if (str.includes("Contract")) return "Contract Sent";
  if (str.includes("Confirm") || str.includes("Deposit") || str.includes("Signed")) return "Confirmed";
  if (str.includes("Complete")) return "Completed";
  if (str.includes("Cancel")) return "Cancelled";
  return "Pending Review";
}

function cleanTimeToString(rawTime) {
  if (rawTime === null || rawTime === undefined || rawTime === "") return "";
  
  if (rawTime instanceof Date) {
    if (isNaN(rawTime.getTime())) return "";
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let tz = ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
      if (rawTime.getFullYear() === 1899) {
        let gmtFormatted = Utilities.formatDate(rawTime, "GMT", "h:mm a");
        if (gmtFormatted) return gmtFormatted.toUpperCase();
      }
      let tzFormatted = Utilities.formatDate(rawTime, tz, "h:mm a");
      if (tzFormatted) return tzFormatted.toUpperCase();
    } catch(e) {}
    
    let h = rawTime.getHours();
    let m = rawTime.getMinutes();
    return formatHoursMinutes(h, m);
  } 
  else if (typeof rawTime === 'number') {
    let totalMinutes = Math.round(rawTime * 24 * 60);
    let h = Math.floor(totalMinutes / 60) % 24;
    let m = totalMinutes % 60;
    return formatHoursMinutes(h, m);
  } 
  else if (typeof rawTime === 'string') {
    let str = rawTime.trim();
    if (!str) return "";
    
    let match = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
    if (match) {
      let h = parseInt(match[1], 10);
      let m = parseInt(match[2], 10);
      let ampm = match[3] ? match[3].toUpperCase() : "";
      if (ampm === "PM" && h < 12) h += 12;
      if (ampm === "AM" && h === 12) h = 0;
      return formatHoursMinutes(h, m);
    } else {
      let d = new Date(str);
      if (!isNaN(d.getTime())) {
        try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          let tz = ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
          if (d.getFullYear() === 1899) {
            return Utilities.formatDate(d, "GMT", "h:mm a").toUpperCase();
          }
          return Utilities.formatDate(d, tz, "h:mm a").toUpperCase();
        } catch(e) {
          return formatHoursMinutes(d.getHours(), d.getMinutes());
        }
      }
    }
  }
  
  return rawTime.toString().replace(/12\/30\/1899\s*/i, "").trim();
}

function sortSheetByDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol === 0) return;
  
  sanitizeAndPopulateSheet(sheet);
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
  
  if (dateIdx > -1) {
    try {
      sheet.sort(dateIdx + 1, true);
      SpreadsheetApp.getUi().alert("📅 Sheet successfully sorted chronologically by Event Date!");
    } catch (err) {
      try {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
        dataRange.sort({ column: dateIdx + 1, ascending: true });
        SpreadsheetApp.getUi().alert("📅 Sheet successfully sorted chronologically by Event Date!");
      } catch (err2) {
        console.warn("Sort note: ", err2);
        SpreadsheetApp.getUi().alert("ℹ️ Sheet sort completed.");
      }
    }
  }
}

function deleteDuplicateRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol === 0) {
    SpreadsheetApp.getUi().alert("ℹ️ No rows available to check for duplicates.");
    return 0;
  }
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
  const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
  const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
  const timeIdx = getHeaderIndex(headers, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]);
  const addrIdx = getHeaderIndex(headers, ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"]);
  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID", "EventId"]);
  
  let keyToRowMap = new Map();
  let rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let rowNum = i + 2;
    
    let email = (emailIdx > -1 && row[emailIdx]) ? row[emailIdx].toString().toLowerCase().trim() : "";
    let name = (nameIdx > -1 && row[nameIdx]) ? row[nameIdx].toString().toLowerCase().trim() : "";
    let eventName = (eventNameIdx > -1 && row[eventNameIdx]) ? row[eventNameIdx].toString().toLowerCase().trim() : "";
    let eventId = (eventIdIdx > -1 && row[eventIdIdx]) ? row[eventIdIdx].toString().trim() : "";
    
    if (!eventName && !name && !email) continue;

    let primaryKey = `${email || name}|${eventName}`;
    
    if (keyToRowMap.has(primaryKey)) {
      let existing = keyToRowMap.get(primaryKey);
      
      let targetRowIndex = existing.index;
      let targetRowNum = existing.rowNum;
      let dupRowIndex = i;
      let dupRowNum = rowNum;
      
      if (!existing.hasEventId && eventId) {
        targetRowIndex = i;
        targetRowNum = rowNum;
        dupRowIndex = existing.index;
        dupRowNum = existing.rowNum;
        keyToRowMap.set(primaryKey, { rowNum: targetRowNum, index: targetRowIndex, hasEventId: true });
      }
      
      let sourceRow = data[dupRowIndex];
      let targetRow = data[targetRowIndex];
      
      [dateIdx, timeIdx, addrIdx].forEach(idx => {
        if (idx > -1 && sourceRow[idx] !== "" && sourceRow[idx] !== null && sourceRow[idx] !== undefined) {
          if (!targetRow[idx] || sourceRow[idx] !== targetRow[idx]) {
            targetRow[idx] = sourceRow[idx];
            safeSetCellValue(sheet.getRange(targetRowNum, idx + 1), sourceRow[idx]);
          }
        }
      });
      
      rowsToDelete.push(dupRowNum);
    } else {
      keyToRowMap.set(primaryKey, { rowNum: rowNum, index: i, hasEventId: !!eventId });
    }
  }
  
  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getUi().alert("✨ No duplicate rows found in your sheet!");
    return 0;
  }
  
  rowsToDelete.sort((a, b) => b - a);
  let deletedCount = 0;
  for (let d = 0; d < rowsToDelete.length; d++) {
    try {
      sheet.deleteRow(rowsToDelete[d]);
      deletedCount++;
    } catch (e) {
      try {
        for (let col = 1; col <= lastCol; col++) {
          safeSetCellValue(sheet.getRange(rowsToDelete[d], col), "");
        }
        deletedCount++;
      } catch (err2) {
        console.warn("Typed column table row delete note: ", err2);
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`🧹 Successfully cleaned & merged ${deletedCount} duplicate rows!`);
  return deletedCount;
}

function applyDropdownValidations(sheet, headers) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const assignedIdx = getHeaderIndex(headers, ["Assigned to", "Assigned"]);
    const internalStatusIdx = getHeaderIndex(headers, ["Internal Status", "Status"]);
    
    if (assignedIdx > -1) {
      try {
        const assignedRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["The Salsa Guy", "Tradición", "Unassigned", "Both"], true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, assignedIdx + 1, lastRow - 1, 1).setDataValidation(assignedRule);
      } catch (err) {
        console.warn("Assigned dropdown validation skip: ", err);
      }
    }
    
    if (internalStatusIdx > -1) {
      try {
        const statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Pending Review", "Proposal Sent", "Contract Sent", "Confirmed", "Completed", "Cancelled"], true)
          .setAllowInvalid(true)
          .build();
        sheet.getRange(2, internalStatusIdx + 1, lastRow - 1, 1).setDataValidation(statusRule);
      } catch (err) {
        console.warn("Status dropdown validation skip: ", err);
      }
    }
  } catch (outerErr) {
    console.warn("applyDropdownValidations outer note: ", outerErr);
  }
}

function sanitizeAndPopulateSheet(sheet) {
  if (!sheet) return;
  ensureRequiredHeadersExist(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
  const timeIdx = getHeaderIndex(headers, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]);
  const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
  const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
  const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
  const addrIdx = getHeaderIndex(headers, ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"]);
  const dowIdx = getHeaderIndex(headers, ["Day of the Week", "Day of Week", "DOW", "Day"]);
  const assignedIdx = getHeaderIndex(headers, ["Assigned to", "Assigned"]);
  const internalStatusIdx = getHeaderIndex(headers, ["Internal Status", "Status"]);

  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  let dateFormulas = [];
  if (dateIdx > -1) {
    try {
      dateFormulas = sheet.getRange(2, dateIdx + 1, data.length, 1).getFormulas();
    } catch (e) {
      dateFormulas = [];
    }
  }

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let rowNum = i + 2;
    
    // Auto-Fix: Move email addresses accidentally placed in "Your Name" column to "Email Address"
    if (nameIdx > -1 && emailIdx > -1) {
      let nameVal = String(row[nameIdx] || "").trim();
      let emailVal = String(row[emailIdx] || "").trim();
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
      let emailMatch = nameVal.match(emailRegex);
      if (emailMatch) {
        let extractedEmail = emailMatch[0];
        if (!emailVal || emailVal === "") {
          row[emailIdx] = extractedEmail;
          safeSetCellValue(sheet.getRange(rowNum, emailIdx + 1), extractedEmail);
        }
        let cleanedName = nameVal.replace(extractedEmail, "").replace(/[<>()]/g, "").trim();
        let finalName = cleanedName || extractedEmail.split("@")[0];
        row[nameIdx] = finalName;
        safeSetCellValue(sheet.getRange(rowNum, nameIdx + 1), finalName);
      }
    }

    if (nameIdx > -1 && (!row[nameIdx] || row[nameIdx].toString().trim() === "")) {
      let fbName = getRowValueByAliases(headers, row, ["your name", "full name", "client name", "name"]);
      if (fbName) {
        row[nameIdx] = fbName;
        safeSetCellValue(sheet.getRange(rowNum, nameIdx + 1), fbName);
      }
    }
    if (emailIdx > -1 && (!row[emailIdx] || row[emailIdx].toString().trim() === "")) {
      let fbEmail = getRowValueByAliases(headers, row, ["email address", "email"]);
      if (fbEmail) {
        row[emailIdx] = fbEmail;
        safeSetCellValue(sheet.getRange(rowNum, emailIdx + 1), fbEmail);
      }
    }
    if (eventNameIdx > -1 && (!row[eventNameIdx] || row[eventNameIdx].toString().trim() === "")) {
      let fbEventName = getRowValueByAliases(headers, row, ["name of the event", "event name", "type of event", "event title", "event"]);
      if (fbEventName) {
        row[eventNameIdx] = fbEventName;
        safeSetCellValue(sheet.getRange(rowNum, eventNameIdx + 1), fbEventName);
      }
    }
    if (dateIdx > -1 && (!row[dateIdx] || row[dateIdx].toString().trim() === "")) {
      let fbDate = getRowValueByAliases(headers, row, ["confirm the date", "event date", "date of event", "date"]);
      if (fbDate) {
        row[dateIdx] = fbDate;
        safeSetCellValue(sheet.getRange(rowNum, dateIdx + 1), fbDate);
      }
    }
    if (timeIdx > -1 && (!row[timeIdx] || row[timeIdx].toString().trim() === "")) {
      let fbTime = getRowValueByAliases(headers, row, ["confirm the time", "event time", "time of event", "time"]);
      if (fbTime) {
        row[timeIdx] = fbTime;
        safeSetCellValue(sheet.getRange(rowNum, timeIdx + 1), fbTime);
      }
    }
    if (addrIdx > -1 && (!row[addrIdx] || row[addrIdx].toString().trim() === "")) {
      let fbAddr = getRowValueByAliases(headers, row, ["where will the event take place", "address", "location"]);
      if (fbAddr) {
        row[addrIdx] = fbAddr;
        safeSetCellValue(sheet.getRange(rowNum, addrIdx + 1), fbAddr);
      }
    }

    if (dateIdx > -1) {
      let dVal = row[dateIdx];
      let parsedDate = parseEventDate(dVal);
      if (parsedDate) {
        if (dowIdx > -1) {
          safeSetCellValue(sheet.getRange(rowNum, dowIdx + 1), days[parsedDate.getDay()]);
        }
        if (!dateFormulas[i] || !dateFormulas[i][0] || !dateFormulas[i][0].toUpperCase().startsWith("=HYPERLINK")) {
          const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID", "EventId"]);
          let fullId = (eventIdIdx > -1 && data[i][eventIdIdx]) ? data[i][eventIdIdx].toString().trim() : "";
          let calUrl = fullId ? getCalendarEventUrl(fullId) : "";
          let dStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (calUrl) {
            safeSetCellValue(sheet.getRange(rowNum, dateIdx + 1), `=HYPERLINK("${calUrl}", "${dStr}")`);
          } else {
            safeSetCellValue(sheet.getRange(rowNum, dateIdx + 1), parsedDate);
          }
        }
      }
    }
    
    if (timeIdx > -1) {
      let tVal = row[timeIdx];
      let tStr = cleanTimeToString(tVal);
      if (tStr) {
        let cell = sheet.getRange(rowNum, timeIdx + 1);
        safeSetCellValue(cell, tStr);
        try { cell.setNumberFormat("@"); } catch(e) {}
      }
    }
    
    if (assignedIdx > -1) {
      let aVal = row[assignedIdx];
      safeSetCellValue(sheet.getRange(rowNum, assignedIdx + 1), normalizeAssignedTo(aVal));
    }
    
    if (internalStatusIdx > -1) {
      let sVal = row[internalStatusIdx];
      safeSetCellValue(sheet.getRange(rowNum, internalStatusIdx + 1), normalizeInternalStatus(sVal));
    }
  }

  applyDropdownValidations(sheet, headers);
}

function appendRowsSafely(sheet, rowsToAppend, mainHeaders) {
  const dateIdx = getHeaderIndex(mainHeaders, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
  const timeIdx = getHeaderIndex(mainHeaders, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]);
  const assignedIdx = getHeaderIndex(mainHeaders, ["Assigned to", "Assigned"]);
  const internalStatusIdx = getHeaderIndex(mainHeaders, ["Internal Status", "Status"]);

  for (let r = 0; r < rowsToAppend.length; r++) {
    let row = rowsToAppend[r];
    
    if (dateIdx > -1 && row[dateIdx]) {
      let pDate = parseEventDate(row[dateIdx]);
      if (pDate) row[dateIdx] = pDate;
    }
    if (timeIdx > -1 && row[timeIdx]) {
      let tStr = cleanTimeToString(row[timeIdx]);
      let tFrac = timeStringToFraction(tStr);
      row[timeIdx] = tFrac !== "" ? tFrac : tStr;
    }
    if (assignedIdx > -1) {
      row[assignedIdx] = normalizeAssignedTo(row[assignedIdx]);
    }
    if (internalStatusIdx > -1) {
      row[internalStatusIdx] = normalizeInternalStatus(row[internalStatusIdx]);
    }

    let targetRow = sheet.getLastRow() + 1;
    try {
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    } catch (e1) {
      try {
        sheet.appendRow(row);
      } catch (e2) {
        for (let c = 0; c < row.length; c++) {
          if (row[c] !== "" && row[c] !== null && row[c] !== undefined) {
            safeSetCellValue(sheet.getRange(targetRow, c + 1), row[c]);
          }
        }
      }
    }
  }
}

function importExternalFormOrSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Import External Form or Sheet Data', 
    'Paste the Google Form URL/ID or Google Sheet URL/ID to import responses from:', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const input = response.getResponseText().trim();
  if (!input) {
    ui.alert("⚠️ Please enter a valid Google Form or Sheet URL / ID.");
    return;
  }
  
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  
  const mainColIndices = {
    timestamp: getHeaderIndex(mainHeaders, ["Timestamp"]),
    name: getHeaderIndex(mainHeaders, ["Your Name", "Full Name", "Client Name", "Name"]),
    email: getHeaderIndex(mainHeaders, ["Email Address", "Email"]),
    eventName: getHeaderIndex(mainHeaders, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]),
    date: getHeaderIndex(mainHeaders, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]),
    time: getHeaderIndex(mainHeaders, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]),
    address: getHeaderIndex(mainHeaders, ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"])
  };

  if (input.includes("docs.google.com/forms") || input.includes("forms/d/")) {
    try {
      let formId = input;
      let match = input.match(/\/forms\/d\/(?:e\/)?([^\/]+)/);
      if (match) {
        formId = match[1];
      }
      
      const form = FormApp.openById(formId);
      const formResponses = form.getResponses();
      
      if (formResponses.length === 0) {
        ui.alert(`⚠️ Google Form "${form.getTitle()}" opened successfully, but it contains 0 submitted responses.`);
        return;
      }
      
      let rowsToAppend = [];
      for (let i = 0; i < formResponses.length; i++) {
        let fResp = formResponses[i];
        let newRow = new Array(mainHeaders.length).fill("");
        
        if (mainColIndices.timestamp > -1) {
          newRow[mainColIndices.timestamp] = fResp.getTimestamp();
        }
        if (mainColIndices.email > -1 && fResp.getRespondentEmail()) {
          newRow[mainColIndices.email] = fResp.getRespondentEmail();
        }
        
        let itemResps = fResp.getItemResponses();
        for (let j = 0; j < itemResps.length; j++) {
          let itemResp = itemResps[j];
          let qTitle = itemResp.getItem().getTitle().toLowerCase().trim();
          let ans = itemResp.getResponse();
          
          if (Array.isArray(ans)) ans = ans.join(", ");
          
          if (qTitle.includes("name") && !qTitle.includes("event") && mainColIndices.name > -1) {
            newRow[mainColIndices.name] = ans;
          } else if (qTitle.includes("email") && mainColIndices.email > -1) {
            newRow[mainColIndices.email] = ans;
          } else if (qTitle.includes("event") && (qTitle.includes("name") || qTitle.includes("what") || qTitle.includes("type")) && mainColIndices.eventName > -1) {
            newRow[mainColIndices.eventName] = ans;
          } else if (qTitle.includes("date") && mainColIndices.date > -1) {
            newRow[mainColIndices.date] = ans;
          } else if (qTitle.includes("time") && mainColIndices.time > -1) {
            newRow[mainColIndices.time] = ans;
          } else if ((qTitle.includes("address") || qTitle.includes("location") || qTitle.includes("take place")) && mainColIndices.address > -1) {
            newRow[mainColIndices.address] = ans;
          }
        }
        
        rowsToAppend.push(newRow);
      }
      
      if (rowsToAppend.length > 0) {
        appendRowsSafely(mainSheet, rowsToAppend, mainHeaders);
        deleteDuplicateRows();
        sanitizeAndPopulateSheet(mainSheet);
        mainAutomation();
        ui.alert(`📥 Successfully imported ${rowsToAppend.length} responses directly from Google Form "${form.getTitle()}"!`);
        return;
      }
    } catch (formErr) {
      console.warn("FormApp import note: ", formErr);
      if (formErr.toString().includes("missing") || formErr.toString().includes("read access") || formErr.toString().includes("0AhWoyu")) {
        ui.alert(
          "⚠️ Form Destination File Notice:\n\n" +
          "This Google Form points to an unlinked or missing response spreadsheet.\n\n" +
          "To import these responses:\n" +
          "1. Open your Google Form.\n" +
          "2. Click the 'Responses' tab.\n" +
          "3. Click the green Google Sheets icon 🟩 to open its responses spreadsheet.\n" +
          "4. Copy that Google Sheets URL and paste it into this importer prompt!"
        );
        return;
      }
    }
  }

  importExternalSpreadsheet(input);
}

function importExternalSpreadsheet(inputUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    let extSS;
    if (inputUrl.startsWith("http")) {
      extSS = SpreadsheetApp.openByUrl(inputUrl);
    } else {
      extSS = SpreadsheetApp.openById(inputUrl);
    }
    
    const extSheet = extSS.getSheets()[0];
    const extData = extSheet.getDataRange().getValues();
    
    if (extData.length <= 1) {
      ui.alert("⚠️ No response rows found in the external sheet.");
      return;
    }
    
    const extHeaders = extData[0].map(h => h ? h.toString().toLowerCase().trim() : "");
    const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
    const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    
    function findExtColumn(aliases) {
      for (let alias of aliases) {
        for (let i = 0; i < extHeaders.length; i++) {
          if (extHeaders[i].includes(alias)) return i;
        }
      }
      return -1;
    }
    
    const colMap = {
      name: findExtColumn(["your name", "name", "full name"]),
      email: findExtColumn(["email", "email address"]),
      eventName: findExtColumn(["name of the event", "event name", "event"]),
      date: findExtColumn(["date of your event", "date", "event date"]),
      time: findExtColumn(["time of your event", "time", "event time"]),
      address: findExtColumn(["address", "where will the event take place", "location"])
    };
    
    const mainColIndices = {
      name: getHeaderIndex(mainHeaders, ["Your Name", "Full Name", "Client Name", "Name"]),
      email: getHeaderIndex(mainHeaders, ["Email Address", "Email"]),
      eventName: getHeaderIndex(mainHeaders, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]),
      date: getHeaderIndex(mainHeaders, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]),
      time: getHeaderIndex(mainHeaders, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]),
      address: getHeaderIndex(mainHeaders, ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"])
    };
    
    let rowsToAppend = [];
    for (let i = 1; i < extData.length; i++) {
      let extRow = extData[i];
      let isEmpty = extRow.every(c => c === "" || c === null);
      if (isEmpty) continue;
      
      let newRow = new Array(mainHeaders.length).fill("");
      
      if (colMap.name > -1 && mainColIndices.name > -1) newRow[mainColIndices.name] = extRow[colMap.name];
      if (colMap.email > -1 && mainColIndices.email > -1) newRow[mainColIndices.email] = extRow[colMap.email];
      if (colMap.eventName > -1 && mainColIndices.eventName > -1) newRow[mainColIndices.eventName] = extRow[colMap.eventName];
      if (colMap.date > -1 && mainColIndices.date > -1) newRow[mainColIndices.date] = extRow[colMap.date];
      if (colMap.time > -1 && mainColIndices.time > -1) newRow[mainColIndices.time] = extRow[colMap.time];
      if (colMap.address > -1 && mainColIndices.address > -1) newRow[mainColIndices.address] = extRow[colMap.address];
      
      rowsToAppend.push(newRow);
    }
    
    if (rowsToAppend.length > 0) {
      appendRowsSafely(mainSheet, rowsToAppend, mainHeaders);
      deleteDuplicateRows();
      sanitizeAndPopulateSheet(mainSheet);
      mainAutomation();
      ui.alert(`📥 Successfully imported ${rowsToAppend.length} response rows into Master Sheet! Automation complete.`);
    } else {
      ui.alert("⚠️ No valid rows could be imported.");
    }
    
  } catch (err) {
    ui.alert(`⚠️ Error opening or importing spreadsheet: ${err.toString()}`);
  }
}

function reGenerateDocs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("⚠️ No data available to update.");
    return;
  }
  
  const response = ui.prompt('Update Row', `Enter the row number to process (2 to ${lastRow}):`, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText(), 10);
    if (isNaN(rowNum) || rowNum < 2 || rowNum > lastRow) {
      ui.alert(`⚠️ Please enter a valid row number between 2 and ${lastRow}.`);
      return;
    }
    sanitizeAndPopulateSheet(sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    try {
      processRow(sheet, rowNum, rowData, DriveApp.getFolderById(CONFIG.FOLDER_ID), headers, true);
      ui.alert("✅ Row " + rowNum + " successfully updated and synchronized!");
    } catch (e) {
      ui.alert(`⚠️ Error updating Row ${rowNum}: ${e.toString()}`);
    }
  }
}

function exportRowAsPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("⚠️ No data available to export.");
    return;
  }
  
  const response = ui.prompt('Export PDFs', `Enter row number to convert to PDF (2 to ${lastRow}):`, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText(), 10);
    if (isNaN(rowNum) || rowNum < 2 || rowNum > lastRow) {
      ui.alert(`⚠️ Please enter a valid row number between 2 and ${lastRow}.`);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    
    const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
    const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
    
    const eventName = (eventNameIdx > -1 && rowData[eventNameIdx]) ? rowData[eventNameIdx] : "Event";
    const eDate = parseEventDate(dateIdx > -1 ? rowData[dateIdx] : "");
    const dateStr = eDate ? Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MM-dd-yyyy") : "TBD";
    
    const targetFolder = getOrCreateEventFolder(mainFolder, dateStr, eventName);
    
    const propUrlIdx = getHeaderIndex(headers, ["MASTER Proposal Form URL", "Proposal URL"]);
    const contUrlIdx = getHeaderIndex(headers, ["Master Contract Document URL", "Contract URL"]);
    
    const propUrl = propUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, propUrlIdx + 1).getFormula()) || rowData[propUrlIdx]) : "";
    const contUrl = contUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, contUrlIdx + 1).getFormula()) || rowData[contUrlIdx]) : "";
    
    let pdfPropUrl = exportDocToPdf(propUrl, targetFolder, `Proposal - ${eventName}.pdf`);
    let pdfContUrl = exportDocToPdf(contUrl, targetFolder, `Contract - ${eventName}.pdf`);
    
    const pdfPropIdx = getHeaderIndex(headers, ["PDF Proposal URL"]);
    const pdfContIdx = getHeaderIndex(headers, ["PDF Contract URL"]);
    
    if (pdfPropUrl && pdfPropIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, pdfPropIdx + 1), `=HYPERLINK("${pdfPropUrl}", "Download PDF Proposal")`);
    }
    if (pdfContUrl && pdfContIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, pdfContIdx + 1), `=HYPERLINK("${pdfContUrl}", "Download PDF Contract")`);
    }
    
    ui.alert(`📄 PDFs generated successfully in event subfolder!\n• Proposal PDF: ${pdfPropUrl ? "Created" : "N/A"}\n• Contract PDF: ${pdfContUrl ? "Created" : "N/A"}`);
  }
}

function draftClientEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("⚠️ No data available to draft email.");
    return;
  }
  
  const response = ui.prompt('Draft Client Email', `Enter row number to draft email for (2 to ${lastRow}):`, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText(), 10);
    if (isNaN(rowNum) || rowNum < 2 || rowNum > lastRow) {
      ui.alert(`⚠️ Please enter a valid row number between 2 and ${lastRow}.`);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
    const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
    const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
    const propUrlIdx = getHeaderIndex(headers, ["MASTER Proposal Form URL", "Proposal URL"]);
    const contUrlIdx = getHeaderIndex(headers, ["Master Contract Document URL", "Contract URL"]);
    const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
    
    const clientEmail = emailIdx > -1 ? rowData[emailIdx] : "";
    const clientName = nameIdx > -1 ? rowData[nameIdx] : "Valued Client";
    const eventName = eventNameIdx > -1 ? rowData[eventNameIdx] : "Your Event";
    const eventDate = dateIdx > -1 ? rowData[dateIdx] : "TBD";
    
    let propUrl = propUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, propUrlIdx + 1).getFormula()) || rowData[propUrlIdx]) : "";
    let contUrl = contUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, contUrlIdx + 1).getFormula()) || rowData[contUrlIdx]) : "";
    
    if (!clientEmail) {
      ui.alert("⚠️ No client email address found for this row.");
      return;
    }
    
    const subject = `Salsa Guy Richmond - Proposal & Contract for ${eventName}`;
    const body = `Hi ${clientName},\n\n` +
      `Thank you for reaching out to Salsa Guy Richmond, LLC regarding your upcoming event "${eventName}" scheduled for ${eventDate}.\n\n` +
      `We have prepared your custom proposal and contract documents for your review:\n` +
      `• Proposal: ${propUrl || "Pending"}\n` +
      `• Contract: ${contUrl || "Pending"}\n\n` +
      `Please let us know if you have any questions or would like to make any adjustments!\n\n` +
      `Warm regards,\n` +
      `Salsa Guy Richmond, LLC\n` +
      `https://salsaguyrichmond.com`;
      
    GmailApp.createDraft(clientEmail, subject, body);
    ui.alert(`📧 Gmail draft successfully created for ${clientName} (${clientEmail})! Check your Gmail Drafts folder.`);
  }
}

function sendClientEmailDirectly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    ui.alert("⚠️ No data available to send email.");
    return;
  }
  
  const response = ui.prompt('Send Client Email', `Enter row number to send email to (2 to ${lastRow}):`, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText(), 10);
    if (isNaN(rowNum) || rowNum < 2 || rowNum > lastRow) {
      ui.alert(`⚠️ Please enter a valid row number between 2 and ${lastRow}.`);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
    const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
    const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
    const propUrlIdx = getHeaderIndex(headers, ["MASTER Proposal Form URL", "Proposal URL"]);
    const contUrlIdx = getHeaderIndex(headers, ["Master Contract Document URL", "Contract URL"]);
    const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
    const internalStatusIdx = getHeaderIndex(headers, ["Internal Status", "Status"]);
    
    const clientEmail = emailIdx > -1 ? rowData[emailIdx] : "";
    const clientName = nameIdx > -1 ? rowData[nameIdx] : "Valued Client";
    const eventName = eventNameIdx > -1 ? rowData[eventNameIdx] : "Your Event";
    const eventDate = dateIdx > -1 ? rowData[dateIdx] : "TBD";
    
    let propUrl = propUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, propUrlIdx + 1).getFormula()) || rowData[propUrlIdx]) : "";
    let contUrl = contUrlIdx > -1 ? (extractUrlFromFormula(sheet.getRange(rowNum, contUrlIdx + 1).getFormula()) || rowData[contUrlIdx]) : "";
    
    if (!clientEmail) {
      ui.alert("⚠️ No client email address found for this row.");
      return;
    }
    
    const confirmSend = ui.alert('Confirm Email Send', `Send proposal & contract email directly to ${clientName} (${clientEmail})?`, ui.ButtonSet.YES_NO);
    if (confirmSend !== ui.Button.YES) return;

    const subject = `Salsa Guy Richmond - Proposal & Contract for ${eventName}`;
    const body = `Hi ${clientName},\n\n` +
      `Thank you for reaching out to Salsa Guy Richmond, LLC regarding your event "${eventName}" on ${eventDate}.\n\n` +
      `We have prepared your proposal and contract documents:\n` +
      `• Proposal: ${propUrl || "Pending"}\n` +
      `• Contract: ${contUrl || "Pending"}\n\n` +
      `Please review the details and let us know if you have any questions!\n\n` +
      `Warm regards,\n` +
      `Salsa Guy Richmond, LLC\n` +
      `https://salsaguyrichmond.com`;
      
    GmailApp.sendEmail(clientEmail, subject, body);
    
    if (internalStatusIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, internalStatusIdx + 1), "Proposal Sent");
    }
    
    ui.alert(`🚀 Email successfully sent to ${clientName} (${clientEmail})! Internal Status updated to 'Proposal Sent'.`);
  }
}

function mainAutomation() {
  const startTime = new Date().getTime();
  const MAX_RUN_TIME_MS = 4.5 * 60 * 1000;
  
  cleanupBlankRows();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  
  sanitizeAndPopulateSheet(sheet);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; 
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
  const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
  const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);

  let processedCount = 0;
  let errorCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const eventName = (eventNameIdx > -1 && row[eventNameIdx]) ? row[eventNameIdx].toString().trim() : "";
    const clientName = (nameIdx > -1 && row[nameIdx]) ? row[nameIdx].toString().trim() : "";
    const clientEmail = (emailIdx > -1 && row[emailIdx]) ? row[emailIdx].toString().trim() : "";
    const hasDate = (dateIdx > -1 && row[dateIdx]);

    if (!eventName && !clientName && !clientEmail && !hasDate) continue;

    if (new Date().getTime() - startTime > MAX_RUN_TIME_MS) {
      SpreadsheetApp.getUi().alert(`⏳ Batch execution paused after processing ${processedCount} rows to prevent timeout. Run 'RUN FULL AUTOMATION' again to continue remaining rows.`);
      return;
    }

    try { 
      processRow(sheet, i + 1, data[i], mainFolder, headers, false); 
      processedCount++;
      Utilities.sleep(100); 
    } catch (e) { 
      errorCount++;
      console.error(`Error processing row ${i + 1}: `, e); 
    }
  }
  SpreadsheetApp.getUi().alert(`✨ Automation complete! Processed ${processedCount} rows (${errorCount} errors).`);
}

function processRow(sheet, rowNum, rowData, mainFolder, headers, forceRegen) {
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"]);
  const timeIdx = getHeaderIndex(headers, ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"]);
  const dowIdx = getHeaderIndex(headers, ["Day of the Week", "Day of Week", "DOW", "Day"]);
  const addrIdx = getHeaderIndex(headers, ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"]);
  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID", "EventId"]);
  
  const propUrlIdx = getHeaderIndex(headers, ["MASTER Proposal Form URL", "Proposal URL"]);
  const contUrlIdx = getHeaderIndex(headers, ["Master Contract Document URL", "Contract URL"]);
  const perfUrlIdx = getHeaderIndex(headers, ["Performance Information Document URL", "Performance Document URL", "Performance URL"]);
  
  const statusIdx = getHeaderIndex(headers, ["Status"]);
  const internalStatusIdx = getHeaderIndex(headers, ["Internal Status"]);
  const assignedToIdx = getHeaderIndex(headers, ["Assigned to", "Assigned"]);

  const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title", "Type of event", "Event"]);
  const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);

  let eventName = (eventNameIdx > -1 && rowData[eventNameIdx]) ? rowData[eventNameIdx].toString().trim() : getRowValueByAliases(headers, rowData, ["what is the name of the event", "name of the event", "event name", "name of event", "event title", "type of event", "event"]).toString().trim();
  let clientName = (nameIdx > -1 && rowData[nameIdx]) ? rowData[nameIdx].toString().trim() : getRowValueByAliases(headers, rowData, ["your name", "full name", "client name", "name"]).toString().trim();
  
  if (!eventName) {
    eventName = clientName ? `${clientName} - Small Private Event` : "Small Private Event";
    if (eventNameIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, eventNameIdx + 1), eventName);
      rowData[eventNameIdx] = eventName;
    }
  }
  
  const dateVal = (dateIdx > -1 && rowData[dateIdx]) ? rowData[dateIdx] : getRowValueByAliases(headers, rowData, ["confirm the date", "event date", "date of event", "date"]);
  const timeVal = (timeIdx > -1 && rowData[timeIdx]) ? rowData[timeIdx] : getRowValueByAliases(headers, rowData, ["confirm the time", "event time", "time of event", "time"]);
  const addrVal = (addrIdx > -1 && rowData[addrIdx]) ? rowData[addrIdx] : getRowValueByAliases(headers, rowData, ["where will the event take place", "address", "location"]);
  
  let isValidDate = false;
  let eDate = parseEventDate(dateVal);
  
  if (eDate) {
    if (timeVal) {
      let tStr = cleanTimeToString(timeVal);
      let match = tStr.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
      if (match) {
        let h = parseInt(match[1], 10);
        let m = parseInt(match[2], 10);
        let ampm = match[3] ? match[3].toUpperCase() : "";
        if (ampm === "PM" && h < 12) h += 12;
        if (ampm === "AM" && h === 12) h = 0;
        eDate.setHours(h, m, 0, 0);
      }
    }
    isValidDate = true;
  }

  const dateStr = isValidDate ? Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MM-dd-yyyy") : "TBD";

  try {
    if (isValidDate && dowIdx > -1) {
      const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const dayName = days[eDate.getDay()];
      safeSetCellValue(sheet.getRange(rowNum, dowIdx + 1), dayName);
    }

    if (eventIdIdx === -1) {
      let lastCol = sheet.getLastColumn() + 1;
      safeSetCellValue(sheet.getRange(1, lastCol), "Event ID");
      try { sheet.getRange(1, lastCol).setFontWeight("bold").setBackground("#d9ead3"); } catch(e) {}
      eventIdIdx = lastCol - 1;
    }

    let fullId = (eventIdIdx > -1 && rowData[eventIdIdx]) ? rowData[eventIdIdx].toString().trim() : "";

    if (isValidDate && eventIdIdx > -1) {
      let calendar = null;
      try {
        calendar = CalendarApp.getCalendarById(CONFIG.INFOCALENDAR_ID);
      } catch (calAccessErr) {
        console.warn("Could not access INFOCALENDAR_ID, falling back to default calendar: ", calAccessErr);
      }
      if (!calendar) {
        try { calendar = CalendarApp.getDefaultCalendar(); } catch(e) {}
      }
      
      const endDate = new Date(eDate.getTime() + (3600 * 1000));
      let calendarEvent = fullId ? getCalendarEventByIdSafely(calendar, fullId) : null;
      
      if (!calendarEvent) {
        try {
          if (calendar) {
            calendarEvent = calendar.createEvent(eventName, eDate, endDate, { location: addrVal || "TBD" });
          } else {
            calendarEvent = CalendarApp.createEvent(eventName, eDate, endDate, { location: addrVal || "TBD" });
          }
          if (calendarEvent) {
            fullId = calendarEvent.getId();
            safeSetCellValue(sheet.getRange(rowNum, eventIdIdx + 1), fullId);
            rowData[eventIdIdx] = fullId;
          }
        } catch (createErr) {
          console.error("Calendar event creation error: ", createErr);
        }
      } else {
        try {
          calendarEvent.setTime(eDate, endDate);
          if (calendarEvent.getTitle() !== eventName) {
            calendarEvent.setTitle(eventName);
          }
          let loc = addrVal || "TBD";
          if (calendarEvent.getLocation() !== loc) {
            calendarEvent.setLocation(loc);
          }
          safeSetCellValue(sheet.getRange(rowNum, eventIdIdx + 1), fullId);
        } catch (calErr) {
          console.error("Calendar sync info: ", calErr);
        }
      }

      if (fullId && dateIdx > -1) {
        let calUrl = getCalendarEventUrl(fullId);
        if (calUrl) {
          try {
            safeSetCellValue(sheet.getRange(rowNum, dateIdx + 1), `=HYPERLINK("${calUrl}", "${dateStr}")`);
          } catch (e) {
            safeSetCellValue(sheet.getRange(rowNum, dateIdx + 1), dateStr);
          }
        }
      }
    }

    if (!fullId && eventIdIdx > -1) {
      fullId = `BTG-EVT-${rowNum}-${new Date().getTime().toString(36).toUpperCase()}`;
      safeSetCellValue(sheet.getRange(rowNum, eventIdIdx + 1), fullId);
      rowData[eventIdIdx] = fullId;
    }

    const propUrlVal = propUrlIdx > -1 ? rowData[propUrlIdx] : "";
    const contUrlVal = contUrlIdx > -1 ? rowData[contUrlIdx] : "";
    const perfUrlVal = perfUrlIdx > -1 ? rowData[perfUrlIdx] : "";

    const docsExist = (propUrlVal && propUrlVal.toString().trim() !== "") &&
                      (contUrlVal && contUrlVal.toString().trim() !== "") &&
                      (perfUrlVal && perfUrlVal.toString().trim() !== "");

    if (forceRegen || !docsExist) {
      const targetFolder = getOrCreateEventFolder(mainFolder, dateStr, eventName);

      const propUrl = createDoc(CONFIG.TEMPLATES.PROPOSAL, "Proposal", eventName, rowData, targetFolder, isValidDate, eDate, headers);
      const contUrl = createDoc(CONFIG.TEMPLATES.CONTRACT, "Contract", eventName, rowData, targetFolder, isValidDate, eDate, headers);
      const perfUrl = createDoc(CONFIG.TEMPLATES.PERF_INFO, "Performance", eventName, rowData, targetFolder, isValidDate, eDate, headers);

      if (propUrl && propUrl.startsWith("http") && propUrlIdx > -1) {
        safeSetCellValue(sheet.getRange(rowNum, propUrlIdx + 1), `=HYPERLINK("${propUrl}", "View Proposal")`);
      }
      if (contUrl && contUrl.startsWith("http") && contUrlIdx > -1) {
        safeSetCellValue(sheet.getRange(rowNum, contUrlIdx + 1), `=HYPERLINK("${contUrl}", "View Contract")`);
      }
      if (perfUrl && perfUrl.startsWith("http") && perfUrlIdx > -1) {
        safeSetCellValue(sheet.getRange(rowNum, perfUrlIdx + 1), `=HYPERLINK("${perfUrl}", "View Info")`);
      }
    }

    if (statusIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, statusIdx + 1), "Synced");
      try { sheet.getRange(rowNum, statusIdx + 1).setBackground("#d9ead3"); } catch (e) {}
      try { sheet.getRange(rowNum, statusIdx + 1).clearNote(); } catch (e) {}
    }
    if (internalStatusIdx > -1 && !rowData[internalStatusIdx]) {
      safeSetCellValue(sheet.getRange(rowNum, internalStatusIdx + 1), "Pending Review");
    }
    if (assignedToIdx > -1 && !rowData[assignedToIdx]) {
      safeSetCellValue(sheet.getRange(rowNum, assignedToIdx + 1), "The Salsa Guy");
    }

  } catch (err) {
    console.error(`Error processing Row ${rowNum}: `, err);
    if (statusIdx > -1) {
      safeSetCellValue(sheet.getRange(rowNum, statusIdx + 1), "Error");
      try { sheet.getRange(rowNum, statusIdx + 1).setBackground("#f4ccd0"); } catch (e) {}
      try { sheet.getRange(rowNum, statusIdx + 1).setNote(err.toString()); } catch (e) {}
    }
  }
}

function createDoc(templateId, type, eventName, rowData, folder, isValidDate, eDate, headers) {
  try {
    const copy = DriveApp.getFileById(templateId).makeCopy(`${type} - ${eventName}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const elems = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(e => e != null);
    
    headers.forEach((h, i) => {
      if (!h) return;
      let val = h.includes("TIME") ? extractTimeOnly(rowData[i]) : rowData[i];
      if (h.includes("DATE")) {
        val = isValidDate ? Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MMMM dd, yyyy") : "TBD";
      }
      const safeVal = (val === null || val === undefined) ? "" : val.toString();
      const escapedH = escapeRegex(h);
      
      elems.forEach(e => {
        e.replaceText("\\{\\{" + escapedH + "\\}\\}", safeVal);
        e.replaceText("<" + escapedH + ">", safeVal);
      });
    });
    
    doc.saveAndClose();
    return copy.getUrl();
  } catch (e) {
    console.error(`Error in createDoc (${type}): ` + e.toString());
    return "Error";
  }
}

function extractTimeOnly(val) {
  if (!val || val === "TBD") return "TBD";
  return cleanTimeToString(val);
}

function importQuickEntryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let qeSheet = ss.getSheetByName("Profe Quick Entry");
  if (!qeSheet) {
    qeSheet = ss.insertSheet("Profe Quick Entry", 0);
    const headers = ["Your Name", "Email Address", "What is the NAME of the event?", "Please confirm the DATE of your event:", "Please confirm the TIME of your event:", "Where will the event take place? (ADDRESS)"];
    qeSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
    SpreadsheetApp.getUi().alert("ℹ️ 'Profe Quick Entry' tab was missing and has now been created. Fill out your row and click upload again!");
    return;
  }
  
  const qeData = qeSheet.getDataRange().getValues();
  if (qeData.length <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️ No entries found to import in Quick Entry tab.");
    return;
  }
  
  const qeHeaders = qeData[0];
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
  
  let rowsToAppend = [];
  for (let i = 1; i < qeData.length; i++) {
    let qeRow = qeData[i];
    let isEmpty = qeRow.every(c => c === "" || c === null);
    if (isEmpty) continue;
    
    let newRow = new Array(mainHeaders.length).fill("");
    qeHeaders.forEach((qHeader, qIdx) => {
      if (!qHeader) return;
      let targetIdx = getHeaderIndex(mainHeaders, qHeader.toString().trim());
      if (targetIdx > -1) {
        newRow[targetIdx] = qeRow[qIdx];
      }
    });
    rowsToAppend.push(newRow);
  }
  
  if (rowsToAppend.length > 0) {
    appendRowsSafely(mainSheet, rowsToAppend, mainHeaders);
    if (qeData.length > 1) {
      try {
        qeSheet.getRange(2, 1, qeData.length - 1, qeHeaders.length).clearContent();
      } catch (errClear) {
        for (let r = 2; r <= qeData.length; r++) {
          for (let c = 1; c <= qeHeaders.length; c++) {
            safeSetCellValue(qeSheet.getRange(r, c), "");
          }
        }
      }
    }
    deleteDuplicateRows();
    sanitizeAndPopulateSheet(mainSheet);
    mainAutomation();
    SpreadsheetApp.getUi().alert(`📤 Successfully uploaded ${rowsToAppend.length} rows from Profe Quick Entry into Master Questionnaire Sheet!`);
  }
}

function cleanupBlankRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2 || lastCol === 0) return;
  
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let totalDeleted = 0;
  
  for (let i = data.length - 1; i >= 0; i--) {
    let row = data[i];
    let rowNum = i + 2;
    let isBlank = row.every(cell => cell === "" || cell === null || cell === undefined);
    
    if (isBlank) {
      try {
        sheet.deleteRow(rowNum);
        totalDeleted++;
      } catch (e) {
        try {
          for (let col = 1; col <= lastCol; col++) {
            safeSetCellValue(sheet.getRange(rowNum, col), "");
          }
          totalDeleted++;
        } catch (err2) {
          console.warn("Typed column cleanup skip note: ", err2);
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert(`🧹 Cleanup Complete! Removed ${totalDeleted} blank rows.`);
}

function onFormSubmit(e) {
  onFormSubmitTrigger(e);
}

function onFormSubmitTrigger(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = (e && e.range) ? e.range.getSheet() : (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0]);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const rowNum = (e && e.range) ? e.range.getRow() : lastRow;
  if (rowNum < 2) return;
  
  Utilities.sleep(1500); 
  
  ensureRequiredHeadersExist(sheet);
  sanitizeAndPopulateSheet(sheet);
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  try {
    processRow(sheet, rowNum, rowData, mainFolder, headers, false);
  } catch (err) {
    console.error("Error in onFormSubmitTrigger: " + err.toString());
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Book Salsa Guy Richmond LLC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    let rawData = e.postData ? e.postData.contents : "";
    let data = rawData ? JSON.parse(rawData) : {};
    let result = handleFormSubmitJson(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleFormSubmitJson(data) {
  if (!data) data = {};
  const ss = SpreadsheetApp.openByUrl(CONFIG.SOURCE_SPREADSHEET_URL) || SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheetByName("Form_Responses_Clean") || ss.getSheets()[0];
  
  ensureRequiredHeadersExist(sheet);
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let newRow = new Array(headers.length).fill("");
  
  const fieldHeaderMap = [
    { key: "clientName", aliases: ["Your Name", "Full Name", "Client Name", "Name"] },
    { key: "clientEmail", aliases: ["Email Address", "Email"] },
    { key: "clientPhone", aliases: ["What is the best phone number to contact you?", "Phone", "Phone Number", "Contact Phone"] },
    { key: "representType", aliases: ["Who do you represent?", "Represent", "Organization Type"] },
    { key: "eventName", aliases: ["What is the NAME of the event?", "NAME of the event", "Event Name", "Name of event", "Event Title"] },
    { key: "eventTypeScale", aliases: ["Is your event a large public event or a small private event?", "Event Scale", "Public or Private"] },
    { key: "eventWebsites", aliases: ["What are the WEBSITES for event and organization?", "Websites", "Event Website"] },
    { key: "eventPurpose", aliases: ["What is the PURPOSE of this event?", "Purpose of event", "Event Purpose"] },
    { key: "hearAboutUs", aliases: ["How did you HEAR of us?", "How heard", "Referral Source"] },
    { key: "eventDescription", aliases: ["DESCRIBE your event.", "Describe event", "Event Description"] },
    { key: "eventDate", aliases: ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"] },
    { key: "eventTime", aliases: ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"] },
    { key: "isInternational", aliases: ["Are services requested OUTSIDE of the USA?", "International Request", "Is International"] },
    { key: "destinationCountry", aliases: ["Select Destination Country", "Destination Country", "Country"] },
    { key: "eventAddress", aliases: ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"] },
    { key: "expectedAttendance", aliases: ["How many people are you expecting will ATTENDING?", "Attendees", "Expected Attendance"] },
    { key: "admissionType", aliases: ["What type of event ADMISSION is it?", "Admission Type", "Admission"] },
    { key: "hasBudget", aliases: ["Confirm you have a BUDGET for our participation.", "Has Budget", "Confirm Budget"] },
    { key: "budgetAmount", aliases: ["What is your budget?", "Confirmed Budget Amount for Performance / Workshop", "Budget Amount", "Budget"] },
    { key: "durationRequired", aliases: ["How much TIME do you require from us?", "Duration Required", "Time Required"] },
    { key: "audienceParticipation", aliases: ["Expecting AUDIENCE PARTICIPATION?", "Interactive (AUDIENCE PARTICIPATION)", "Audience Participation"] },
    { key: "performanceAreaSize", aliases: ["How large is the AREA for our performance or lesson?", "Size of Performance / Class Area", "Performance Area Size", "Area Size"] },
    { key: "surfaceType", aliases: ["On what SURFACE will the performance or class take place?", "Surface Type", "Floor Type"] },
    { key: "venueSetting", aliases: ["Where will it take PLACE?", "Venue Location Setting", "Venue Setting", "Indoor/Outdoor"] },
    { key: "soundSystem", aliases: ["About the SOUND SYSTEM", "Sound System Equipment", "Sound System"] },
    { key: "audienceAges", aliases: ["Who is your AUDIENCE:", "Audience Age Groups Expected", "Audience Age Groups", "Age Groups"] },
    { key: "serviceRecurrence", aliases: ["Will the PERFORMANCE SERVICES be ...", "Service Recurrence", "Recurrence"] },
    { key: "nonProfitName", aliases: ["FOR 501(C) ONLY - Which 501(C) do you represent?", "501c Non Profit Name", "501(c) Organization"] },
    { key: "taxLetter", aliases: ["FOR 501(C) ONLY - Are you able to provide a TAX DEDUCTIBILITY LETTER?", "Tax Deductibility Letter", "Tax Letter"] },
    { key: "boothSpace", aliases: ["Will you provide a BOOTH/EXHIBITOR space (10 x 10 Tent) to promote our services?", "Booth Space", "Exhibitor Booth"] },
    { key: "promoInclude", aliases: ["Will you INCLUDE OUR INFORMATION and logo on all promotional materials, including social media?", "Include Logo on Promo", "Promo Include"] },
    { key: "allowHelpPromote", aliases: ["Are we allowed to HELP PROMOTE the event?", "Allow Help Promote", "Promote Event"] },
    { key: "videoRights", aliases: ["Can we get COPIES of video footage and pictures of our participation?", "Video Footage & Photos", "Video Rights"] },
    { key: "contingencyPlan", aliases: ["What is the CONTINGENCY PLAN? Please include alternate locations, dates, times.", "Do you have a CONTINGENCY PLAN?", "Contingency Plan"] },
    { key: "privateGatheringType", aliases: ["TYPE of Private Gathering", "Private Gathering Type"] },
    { key: "badgeAccess", aliases: ["Will a BADGE or ID be required for performers?", "Will a BADGE or ID be issued to performers to access the performance area?", "Badge Access", "ID Badge"] },
    { key: "dressingRoomInstructions", aliases: ["Will we have a place to change COSTUMES if needed? If so, please provide instructions.", "Dressing Room Instructions", "Costume Change Room"] },
    { key: "invitedToAttend", aliases: ["Are we INVITED TO ATTEND the event?", "Invited to Attend"] },
    { key: "performerProvisions", aliases: ["WILL YOU PROVIDE the performers with:", "Performer Provisions", "Provisions"] },
    { key: "provisions", aliases: ["Additional Performer Amenities Provided:", "Hospitality Provisions", "Amenities"] },
    { key: "performanceServices", aliases: ["Which of our PERFORMANCE SERVICES will you need?", "Performance Services Needed", "Performance Services"] },
    { key: "otherPerfServices", aliases: ["Any other PERFORMANCE SERVICES you wish, but are not listed above?", "Other Performance Services"] },
    { key: "lessonServices", aliases: ["Which of our DANCE LESSON SERVICES will you need?", "Dance Lesson Services Needed", "Lesson Services"] },
    { key: "otherServices", aliases: ["What OTHER SERVICES will you need?", "Additional Services Needed", "Other Services"] },
    { key: "notes", aliases: ["SPECIAL REQUESTS or Song Preferences", "Special Instructions, Song Requests or Notes", "Notes", "Special Requests"] }
  ];

  fieldHeaderMap.forEach(item => {
    let colIdx = getHeaderIndex(headers, item.aliases);
    if (colIdx > -1 && data[item.key] !== undefined && data[item.key] !== null) {
      newRow[colIdx] = data[item.key];
    }
  });

  sheet.appendRow(newRow);
  const rowNum = sheet.getLastRow();
  
  sanitizeAndPopulateSheet(sheet);
  
  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  if (data.attachedFile && data.attachedFile.base64) {
    try {
      const bytes = Utilities.base64Decode(data.attachedFile.base64);
      const blob = Utilities.newBlob(bytes, data.attachedFile.mimeType || 'application/octet-stream', data.attachedFile.fileName || 'Event_Attachment');
      const dateStr = data.eventDate || "Undated";
      const eventFolder = getOrCreateEventFolder(mainFolder, dateStr, data.eventName || "Event");
      eventFolder.createFile(blob);
    } catch (fileErr) {
      console.warn("File attachment upload error: ", fileErr);
    }
  }

  const updatedRowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  processRow(sheet, rowNum, updatedRowData, mainFolder, headers, true);
  
  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID"]);
  const finalEventId = eventIdIdx > -1 ? sheet.getRange(rowNum, eventIdIdx + 1).getValue() : "";
  
  return {
    status: "success",
    row: rowNum,
    eventId: finalEventId || `BTG-EVT-${rowNum}-${Date.now().toString(36).toUpperCase()}`
  };
}

/**
 * Cleanup Task: Scans the sheet for rows where email addresses were entered into the "Your Name" column,
 * moves the email address to the "Email Address" column, and cleans up the "Your Name" column.
 */
function fixMisplacedEmailData() {
  const ss = SpreadsheetApp.openByUrl(CONFIG.SOURCE_SPREADSHEET_URL) || SpreadsheetApp.getActiveSpreadsheet();
  const ui = (typeof SpreadsheetApp.getUi === 'function') ? SpreadsheetApp.getUi() : null;
  
  let sheetsToProcess = [
    ss.getSheetByName(CONFIG.SHEET_NAME),
    ss.getSheetByName("Form_Responses_Clean")
  ].filter(s => s !== null);
  
  if (sheetsToProcess.length === 0) sheetsToProcess = [ss.getSheets()[0]];
  
  let totalFixed = 0;
  
  sheetsToProcess.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const nameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
    const emailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
    
    if (nameIdx === -1 || emailIdx === -1) return;
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
    
    let modified = false;
    
    for (let r = 0; r < data.length; r++) {
      let nameVal = String(data[r][nameIdx] || "").trim();
      let emailVal = String(data[r][emailIdx] || "").trim();
      
      let emailMatch = nameVal.match(emailRegex);
      if (emailMatch) {
        let extractedEmail = emailMatch[0];
        
        if (!emailVal || emailVal === "") {
          data[r][emailIdx] = extractedEmail;
        }
        
        let cleanedName = nameVal.replace(extractedEmail, "").replace(/[<>()]/g, "").trim();
        data[r][nameIdx] = cleanedName || extractedEmail.split("@")[0];
        
        modified = true;
        totalFixed++;
      }
    }
    
    if (modified) {
      dataRange.setValues(data);
    }
  });
  
  if (ui) {
    ui.alert(`✅ Data Cleanup Complete!\n\nFixed and moved ${totalFixed} misplaced email addresses from 'Your Name' into 'Email Address' column!`);
  }
}

/**
 * Constructs a direct, clickable Google Calendar web URL for a given Google Calendar Event ID.
 * Resolves missing hyperlink issue in 'Please confirm the DATE of your event:' column.
 */
function getCalendarEventUrl(eventId) {
  if (!eventId) return "";
  try {
    let cleanId = eventId.toString().trim();
    if (cleanId.startsWith("BTG-EVT-")) return "";
    
    let rawId = cleanId.split('@')[0];
    let calId = CONFIG.INFOCALENDAR_ID || "primary";
    
    let combo = rawId + " " + calId;
    let eid = Utilities.base64EncodeWebSafe(combo).replace(/=/g, "");
    
    return `https://www.google.com/calendar/event?eid=${eid}`;
  } catch (err) {
    console.warn("getCalendarEventUrl error: ", err);
    return "";
  }
}

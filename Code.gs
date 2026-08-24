/**
 * Salsa Guy Richmond LLC - Automation Suite v20.71
 * Production Script for Google Sheets, Forms, Calendar, Drive & Document Automation
 */

const CONFIG = {
  SHEET_NAME: "Form Responses 1", 
  SPREADSHEET_ID: "1ssJPBdSTOUzq1S9b_kHYuBLMoUbxwrDap2FIkfHOX4k",
  SPREADSHEET_URL: "https://docs.google.com/spreadsheets/d/1ssJPBdSTOUzq1S9b_kHYuBLMoUbxwrDap2FIkfHOX4k/edit",
  INFOCALENDAR_ID: "shqfpe645m3tj6fhee17irti5s@group.calendar.google.com",
  FOLDER_ID: "1FaiN_vTho7YY5mwd_OXrfHPnQB0fF6oR",
  TEMPLATES: {
    PROPOSAL: "1plCZvjBJijgJrGduzXrTMjbtDopMgCfB5Vo7MLWxspo",
    CONTRACT: "1BuEv7BF6wsHutvEWwVZOVb3m8J3zkYujKheti8X871g",
    PERF_INFO: "1eyXzMdmYiV0CDmxvNZfNO3dRYQ83hvJHHy_eseWPqIE" 
  }
};

/**
 * Returns active spreadsheet or opens via target SPREADSHEET_ID.
 */
function getSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && CONFIG.SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {
      console.error("Failed to open spreadsheet by ID: " + e.message);
    }
  }
  return ss;
}

/**
 * Initializes the custom Google Sheets UI dropdown menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💃 Salsa App')
    .addItem('🛠️ SETUP ALL QUESTIONNAIRE HEADERS', 'setupMasterHeaders')
    .addSeparator()
    .addItem('🚀 RUN FULL AUTOMATION', 'mainAutomation')
    .addItem('🎯 UPDATE SELECTED ROW ONLY', 'reGenerateDocs')
    .addSeparator()
    .addItem('🎨 COLOR CODE HEADERS', 'colorCodeHeaders')
    .addItem('🔍 DIAGNOSE SHEET (Check Headers)', 'diagnoseSheet')
    .addSeparator()
    .addItem("📝 1. SETUP PROFE'S QUICK ENTRY", 'setupQuickEntryTab')
    .addItem("📤 2. UPLOAD PROFE'S QUICK ENTRY", 'importQuickEntryData')
    .addToUi();
}

/**
 * Generates Row 1 headers from the complete 7-Section Questionnaire and tracking fields,
 * renames Sheet1 to CONFIG.SHEET_NAME if needed, and applies section color coding.
 */
function setupMasterHeaders() {
  const ss = getSpreadsheet();
  if (!ss) {
    SpreadsheetApp.getUi().alert("⚠️ Unable to access spreadsheet.");
    return;
  }
  
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    const activeSheet = ss.getActiveSheet();
    if (activeSheet && activeSheet.getName() === "Sheet1") {
      sheet = activeSheet;
      sheet.setName(CONFIG.SHEET_NAME);
    } else {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME, 0);
    }
  }

  const masterHeaders = [
    // 1. Contact & Organizer Information
    "Submission Date",
    "Your Name",
    "Email Address",
    "Best Contact Phone Number",
    "Who do you represent? (Organization / Business / Self)",
    "How did you HEAR of us?",

    // 2. Event Overview & Schedule
    "What is the NAME of the event?",
    "Purpose of this Event",
    "Websites for Event & Organization",
    "Describe Your Event",
    "Please confirm the DATE of your event:",
    "Please confirm the TIME of your event:",
    "Expected Number of Attendees",
    "Audience Age Groups Expected",
    "Type of Event Admission",

    // 3. Location & Event Classification Routing
    "Is this event located within the Commonwealth of Virginia (USA)?",
    "Where will the event take place? (ADDRESS)",
    "Select Event Classification:",

    // 4A. Conditional: Large Public Event Details
    "Will the PERFORMANCE SERVICES be...",
    "501(c) Non-Profit Name (If Applicable)",
    "Can you provide a Tax Deductibility Letter?",
    "Provide a Booth/Exhibitor Space (10×10 Tent)?",
    "Include our logo on promo materials & social media?",
    "Allowed to help promote the event?",
    "Provide copies of video footage and photos?",
    "Weather / Contingency Plan",

    // 4B. Conditional: Small Private Event Specifications
    "Type of Private Gathering",
    "Are performers invited to attend/stay for the event?",

    // 4C. Conditional: International Logistics
    "International: Specific Country, City, & Venue Name",
    "International: Travel & Lodging Logistics",
    "International: Visa & Legal Documentation Support",
    "International: Preferred Currency & Payment Terms",
    "International: Costumes, Props & Customs Considerations",

    // 5A. Repertoire Requests
    "Repertoire: Mexico / North American Dances",
    "Repertoire: Caribbean Dances (Cuba & Puerto Rico)",
    "Repertoire: Central American Dances (El Salvador)",
    "Repertoire: South American Dances (Colombia & Argentina)",
    "Repertoire: Theatrical, Parade & Live Singing Performances",
    "Any other PERFORMANCE SERVICES you wish, but are not listed above?",

    // 5B. Dance Instruction & Formats
    "Which of our DANCE LESSON SERVICES will you need?",
    "Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?",
    "Additional Services Needed (MC, DJ, Lecture)",
    "General Formats (Stage, Opening, Headliner, Main Act, Background)",
    "How much TIME do you require from us?",

    // 6. Technical, Venue & Hospitality Logistics
    "Venue Location Setting",
    "On what SURFACE will the performance or class take place?",
    "Size of Performance / Class Area",
    "Sound System Equipment",
    "Will a BADGE or ID be required for performers?",
    "WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)",
    "Dressing Room / Costume Changing Instructions",

    // 7. Budget, Documents & Terms
    "Confirm you have a BUDGET for our participation",
    "Confirmed Budget Amount for Performance / Workshop",
    "Upload Event Document / Attachment",
    "Special Instructions, Song Requests or Notes",
    "Notice: Hiring Similar Performers Disclosure",
    "Terms of Service & Privacy Policy Agreement",

    // Automation & Management Tracking Columns
    "Event ID",
    "Day of the Week",
    "MASTER Proposal Form URL",
    "Master Contract Document URL",
    "Performance Information Document URL",
    "Assigned to",
    "Status",
    "Internal Status"
  ];

  sheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]).setFontWeight("bold");
  colorCodeHeaders();
  diagnoseSheet();
}

/**
 * Validates that critical response & tracking columns exist in the active header row.
 */
function diagnoseSheet() {
  const ss = getSpreadsheet();
  const sheet = ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null;
  if (!sheet) return;
  const headers = getUniqueHeaders(sheet);
  const ui = SpreadsheetApp.getUi();

  const required = [
    "What is the NAME of the event?",
    "Please confirm the DATE of your event:",
    "Please confirm the TIME of your event:",
    "Event ID",
    "Status"
  ];

  let report = "📊 Header Diagnosis Report:\n\n";
  let allFound = true;

  required.forEach(req => {
    const found = headers.includes(req);
    report += `${found ? "✅" : "❌"} ${req}\n`;
    if (!found) allFound = false;
  });

  report += `\nActive Sheet Name: "${sheet.getName()}"\n`;
  report += allFound ? "\nAll critical headers found successfully!" : "\n⚠️ Warning: Some required headers were not found.";
  
  ui.alert(report);
}

/**
 * Scans Row 1, deduplicates headers, and updates duplicate header cells with numeric suffixes.
 */
function getUniqueHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const rawHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const seen = {};
  return rawHeaders.map((h, index) => {
    let headerName = h ? h.toString().trim() : `Column_${index + 1}`;
    if (seen[headerName]) {
      seen[headerName]++;
      const uniqueName = `${headerName}_${seen[headerName]}`;
      sheet.getRange(1, index + 1).setValue(uniqueName);
      return uniqueName;
    } else {
      seen[headerName] = 1;
      return headerName;
    }
  });
}

/**
 * Visually categorizes header cells with section background colors.
 */
function colorCodeHeaders() {
  const ss = getSpreadsheet();
  const sheet = ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null;
  if (!sheet) return;
  const headers = getUniqueHeaders(sheet);
  
  headers.forEach((h, i) => {
    const colNum = i + 1;
    const cell = sheet.getRange(1, colNum);
    
    if (h.includes("Submission Date") || h.includes("Timestamp") || h.includes("Name") || h.includes("Email") || h.includes("event?")) {
      cell.setBackground("#fce5cd"); // Soft Peach
    } else if (h.includes("URL") || h.includes("Document") || h.includes("Event ID") || h.includes("Upload") || h.includes("Attachment") || h.includes("Flyer")) {
      cell.setBackground("#f4cccc"); // Soft Red
    } else if (h.includes("DATE") || h.includes("TIME") || h.includes("Day of the Week")) {
      cell.setBackground("#cfe2f3"); // Soft Blue
    } else if (h.includes("Status") || h.includes("Assigned")) {
      cell.setBackground("#d9ead3"); // Soft Green
    } else {
      cell.setBackground("#fff2cc"); // Soft Yellow
    }
  });
  
  SpreadsheetApp.getUi().alert("🎨 Headers have been successfully color-coded and deduplicated!");
}

/**
 * Prompts user for a specific row number to re-process.
 */
function reGenerateDocs() {
  const ss = getSpreadsheet();
  const sheet = ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null;
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt('Update Row', 'Enter the row number to process (e.g., 41):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const rowNum = parseInt(response.getResponseText());
    if (isNaN(rowNum) || rowNum < 2) {
      ui.alert("⚠️ Please enter a valid row number (2 or higher).");
      return;
    }
    const headers = getUniqueHeaders(sheet);
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    processRow(sheet, rowNum, rowData, DriveApp.getFolderById(CONFIG.FOLDER_ID), headers);
    ui.alert(`✅ Row ${rowNum} processed successfully!`);
  }
}

/**
 * Batch-processes all active, unsynced rows in the sheet.
 */
function mainAutomation() {
  const ss = getSpreadsheet();
  const sheet = ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null;
  if (!sheet) return;
  const headers = getUniqueHeaders(sheet);
  const data = sheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  
  const statusIdx = headers.indexOf("Status");
  const eventNameIdx = headers.indexOf("What is the NAME of the event?");

  for (let i = 1; i < data.length; i++) {
    const eventName = data[i][eventNameIdx] || "";
    const status = statusIdx > -1 ? data[i][statusIdx] : "";
    
    if (eventName === "" || status === "Synced") continue;
    
    try { 
      processRow(sheet, i + 1, data[i], folder, headers); 
      Utilities.sleep(300);
    } catch (e) { 
      console.error(`Error processing row ${i + 1}: ` + e.message); 
    }
  }
}

/**
 * Processes a single row: formatting attachments, syncing Google Calendar, generating documents, and updating status.
 */
function processRow(sheet, rowNum, rowData, folder, headers) {
  const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "Event Date", "Date of event", "Date"]);
  const timeIdx = getHeaderIndex(headers, ["Please confirm the TIME of your event:", "Event Time", "Time of event", "Time"]);
  const dowIdx = getHeaderIndex(headers, ["Day of the Week", "Day of Week", "DOW"]);
  const addrIdx = getHeaderIndex(headers, ["Where will the event take place? (ADDRESS)", "Event Address", "Address", "Location"]);
  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID"]);
  
  const statusIdx = getHeaderIndex(headers, ["Status"]);
  const internalStatusIdx = getHeaderIndex(headers, ["Internal Status"]);
  const assignedToIdx = getHeaderIndex(headers, ["Assigned to", "Assigned To"]);

  const propUrlIdx = getHeaderIndex(headers, ["MASTER Proposal Form URL", "Proposal Document URL", "Proposal URL"]);
  const contUrlIdx = getHeaderIndex(headers, ["Master Contract Document URL", "Contract Document URL", "Contract URL"]);
  const perfUrlIdx = getHeaderIndex(headers, ["Performance Information Document URL", "Performance Document URL", "Performance Info URL"]);

  const eventNameIdx = getHeaderIndex(headers, ["What is the NAME of the event?", "Event Name", "Name of Event"]);
  const eventName = (eventNameIdx > -1 && rowData[eventNameIdx]) ? rowData[eventNameIdx] : "Unnamed Event";
  
  // Format Uploaded File Links (Drive Attachment Link Handler)
  headers.forEach((h, i) => {
    if (h.toLowerCase().includes("upload") || h.toLowerCase().includes("attachment") || h.toLowerCase().includes("flyer")) {
      const fileUrlVal = rowData[i];
      if (fileUrlVal && typeof fileUrlVal === "string") {
        const urls = fileUrlVal.split(",").map(u => u.trim()).filter(u => u.startsWith("http"));
        if (urls.length > 0) {
          const formula = urls.length === 1 
            ? `=HYPERLINK("${urls[0]}", "View Uploaded File")`
            : `=HYPERLINK("${urls[0]}", "View Uploaded File (1 of ${urls.length})")`;
          sheet.getRange(rowNum, i + 1).setFormula(formula);
        }
      }
    }
  });

  let isValidDate = false;
  let eDate = null;
  if (rowData[dateIdx]) {
    eDate = new Date(rowData[dateIdx]);
    if (rowData[timeIdx] instanceof Date) { 
      eDate.setHours(rowData[timeIdx].getHours()); 
      eDate.setMinutes(rowData[timeIdx].getMinutes()); 
    }
    isValidDate = !isNaN(eDate.getTime());
  }

  if (isValidDate && dowIdx > -1) {
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[eDate.getDay()];
    sheet.getRange(rowNum, dowIdx + 1).setValue(dayName);
  }

  let calendarEventUrl = "";
  if (isValidDate) {
    const calendar = CalendarApp.getCalendarById(CONFIG.INFOCALENDAR_ID);
    if (calendar) {
      let calendarEvent = null;
      let fullId = eventIdIdx > -1 ? (rowData[eventIdIdx] ? rowData[eventIdIdx].toString().trim() : "") : "";

      if (fullId) {
        try {
          calendarEvent = calendar.getEventById(fullId);
        } catch (e) {}
      }

      if (!calendarEvent) {
        try {
          const eventsOnDay = calendar.getEventsForDay(eDate);
          calendarEvent = eventsOnDay.find(ev => ev.getTitle().toLowerCase() === eventName.toLowerCase());
        } catch (e) {}
      }

      if (!calendarEvent) {
        try {
          const endDate = new Date(eDate.getTime() + (3600 * 1000));
          calendarEvent = calendar.createEvent(eventName, eDate, endDate, { location: rowData[addrIdx] || "TBD" });
        } catch (e) {}
      }

      if (calendarEvent) {
        const eventId = calendarEvent.getId();
        const splitEventId = eventId.split('@')[0];
        const eid = Utilities.base64Encode(`${splitEventId} ${CONFIG.INFOCALENDAR_ID}`).replace(/=+$/, '');
        calendarEventUrl = `https://www.google.com/calendar/event?eid=${eid}`;
        
        if (eventIdIdx > -1) {
          sheet.getRange(rowNum, eventIdIdx + 1).setValue(eventId);
        }
      }
    }
  }

  if (isValidDate && dateIdx > -1) {
    const formattedDateStr = Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
    if (calendarEventUrl) {
      sheet.getRange(rowNum, dateIdx + 1).setFormula(`=HYPERLINK("${calendarEventUrl}", "${formattedDateStr}")`);
    } else {
      sheet.getRange(rowNum, dateIdx + 1).setValue(formattedDateStr);
    }
  }

  const propUrl = createDoc(CONFIG.TEMPLATES.PROPOSAL, "Proposal", eventName, rowData, folder, isValidDate, eDate, headers);
  const contUrl = createDoc(CONFIG.TEMPLATES.CONTRACT, "Contract", eventName, rowData, folder, isValidDate, eDate, headers);
  const perfUrl = createDoc(CONFIG.TEMPLATES.PERF_INFO, "Performance", eventName, rowData, folder, isValidDate, eDate, headers);

  if (propUrl !== "Error" && propUrlIdx > -1) sheet.getRange(rowNum, propUrlIdx + 1).setFormula(`=HYPERLINK("${propUrl}", "View Proposal")`);
  if (contUrl !== "Error" && contUrlIdx > -1) sheet.getRange(rowNum, contUrlIdx + 1).setFormula(`=HYPERLINK("${contUrl}", "View Contract")`);
  if (perfUrl !== "Error" && perfUrlIdx > -1) sheet.getRange(rowNum, perfUrlIdx + 1).setFormula(`=HYPERLINK("${perfUrl}", "View Info")`);

  if (statusIdx > -1) {
    sheet.getRange(rowNum, statusIdx + 1).setValue("Synced").setBackground("#d9ead3");
  }
  if (internalStatusIdx > -1 && !rowData[internalStatusIdx]) {
    sheet.getRange(rowNum, internalStatusIdx + 1).setValue("Ready for Review");
  }
  if (assignedToIdx > -1 && !rowData[assignedToIdx]) {
    sheet.getRange(rowNum, assignedToIdx + 1).setValue("The Salsa Guy");
  }
}

/**
 * Creates individual document copy from master template and replaces placeholders with row data.
 */
function createDoc(templateId, type, eventName, rowData, folder, isValidDate, eDate, headers) {
  try {
    const copy = DriveApp.getFileById(templateId).makeCopy(`${type} - ${eventName}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const elems = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(e => e != null);
    
    headers.forEach((h, i) => {
      let val = (h.includes("TIME")) ? extractTimeOnly(rowData[i]) : rowData[i];
      if (h.includes("DATE")) val = (isValidDate) ? Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MMMM dd, yyyy") : "TBD";
      const safeVal = (val === null || val === undefined) ? "" : val.toString();
      
      const escapedH = h.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      
      elems.forEach(e => {
        e.replaceText(escapedH, safeVal);
        e.replaceText("{{" + escapedH + "}}", safeVal);
        e.replaceText("<" + escapedH + ">", safeVal);
      });
    });
    doc.saveAndClose();
    return copy.getUrl();
  } catch (e) { 
    console.error("Error creating document: " + e.message);
    return "Error"; 
  }
}

/**
 * Formats time values into clean 12-hour format (h:mm a).
 */
function extractTimeOnly(val) {
  if (!val || val === "TBD") return "TBD";
  try {
    const d = new Date(val);
    return isNaN(d.getTime()) ? val.toString() : Utilities.formatDate(d, Session.getScriptTimeZone(), "h:mm a");
  } catch (e) {
    return "TBD";
  }
}

/**
 * Generates staging tab 'Profe Quick Entry' for manual fast data entry.
 */
function setupQuickEntryTab() {
  const ss = getSpreadsheet();
  if (!ss) return;
  let qeSheet = ss.getSheetByName("Profe Quick Entry");
  if (!qeSheet) {
    qeSheet = ss.insertSheet("Profe Quick Entry", 0);
    const headers = [
      "Your Name", 
      "Email Address", 
      "What is the NAME of the event?", 
      "Please confirm the DATE of your event:", 
      "Please confirm the TIME of your event:", 
      "Where will the event take place? (ADDRESS)"
    ];
    qeSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
    SpreadsheetApp.getUi().alert("✅ Staging tab 'Profe Quick Entry' initialized!");
  } else {
    SpreadsheetApp.getUi().alert("ℹ️ Tab 'Profe Quick Entry' already exists.");
  }
}

/**
 * Imports staging rows from 'Profe Quick Entry' tab into main response sheet and runs automation.
 */
function importQuickEntryData() {
  const ss = getSpreadsheet();
  if (!ss) return;
  const qeSheet = ss.getSheetByName("Profe Quick Entry");
  const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
  if (!qeSheet) return;
  const qeData = qeSheet.getDataRange().getValues();
  if (qeData.length <= 1) return;
  const qeHeaders = qeData[0];
  const mainHeaders = getUniqueHeaders(mainSheet);
  
  let addedCount = 0;
  for (let i = 1; i < qeData.length; i++) {
    let rowData = qeData[i];
    if (!rowData[0] && !rowData[2]) continue;
    let newRow = new Array(mainHeaders.length).fill("");
    qeHeaders.forEach((h, j) => {
      let idx = mainHeaders.indexOf(h);
      if (idx > -1) newRow[idx] = rowData[j];
    });
    mainSheet.appendRow(newRow);
    addedCount++;
  }
  qeSheet.getRange(2, 1, Math.max(qeSheet.getLastRow() - 1, 1), qeSheet.getLastColumn()).clearContent();
  SpreadsheetApp.getUi().alert(`📥 Imported ${addedCount} row(s) into '${mainSheet.getName()}'. Running automation now...`);
  mainAutomation();
}

/**
 * Processes JSON form submissions received via Web App or Google Script API,
 * mapping JSON keys to Google Sheet header aliases and appending the new response.
 */
function handleFormSubmitJson(data) {
  const ss = getSpreadsheet();
  const sheet = ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null;
  if (!sheet) throw new Error("Target sheet tab not found.");

  const headers = getUniqueHeaders(sheet);
  let newRow = new Array(headers.length).fill("");

  const fieldHeaderMap = [
    { key: "timestamp", aliases: ["Submission Date", "Timestamp", "Date Submitted", "Date", "Submission Timestamp"] },
    { key: "clientName", aliases: ["Your Name", "Full Name", "Client Name", "Name"] },
    { key: "organizationName", aliases: ["Organization / Company Name", "Organization Name", "Organization", "Company"] },
    { key: "clientEmail", aliases: ["Email Address", "Email"] },
    { key: "clientPhone", aliases: ["Best Contact Phone Number", "Phone Number", "Phone", "Telephone"] },
    { key: "contactPreference", aliases: ["Preferred Contact Method", "Contact Preference"] },
    { key: "representType", aliases: ["Who do you represent? (Organization / Business / Self)", "Who do you represent?", "Represent Type", "Representing"] },
    { key: "hearAboutUs", aliases: ["How did you HEAR of us?", "How did you HEAR ABOUT US?", "How did you hear about us", "Referral Source"] },
    { key: "eventName", aliases: ["What is the NAME of the event?", "Event Name", "Name of Event"] },
    { key: "eventPurpose", aliases: ["Purpose of this Event", "What is the PURPOSE / THEME of your event?", "Event Purpose", "Purpose / Theme", "Theme"] },
    { key: "eventWebsites", aliases: ["Websites for Event & Organization", "Event Website(s) / Social Media", "Event Website", "Website", "Websites"] },
    { key: "eventDescription", aliases: ["Describe Your Event", "DESCRIBE your event.", "Describe event", "Event Description"] },
    { key: "eventDate", aliases: ["Please confirm the DATE of your event:", "confirm the DATE", "Event Date", "Date of event", "Date"] },
    { key: "eventTime", aliases: ["Please confirm the TIME of your event:", "confirm the TIME", "Event Time", "Time of event", "Time"] },
    { key: "expectedAttendance", aliases: ["Expected Number of Attendees", "How many people are you expecting will ATTENDING?", "Attendees", "Expected Attendance"] },
    { key: "audienceAges", aliases: ["Audience Age Groups Expected", "Who is your AUDIENCE:", "Audience Age Groups", "Age Groups"] },
    { key: "admissionType", aliases: ["Type of Event Admission", "What type of event ADMISSION is it?", "Admission Type", "Admission"] },
    { key: "regionalLocation", aliases: ["Is this event located within the Commonwealth of Virginia (USA)?", "Regional Location & Jurisdiction Check", "Regional Location", "Virginia Check"] },
    { key: "eventAddress", aliases: ["Where will the event take place? (ADDRESS)", "Where will the event take place", "Address", "Location"] },
    { key: "eventTypeScale", aliases: ["Select Event Classification:", "Event Type Scale", "Scale of Event", "Event Scale", "Service Category / Scope", "Service Category"] },
    { key: "serviceRecurrence", aliases: ["Will the PERFORMANCE SERVICES be...", "Will the PERFORMANCE SERVICES be ...", "Service Recurrence", "Recurrence"] },
    { key: "nonProfitName", aliases: ["501(c) Non-Profit Name (If Applicable)", "FOR 501(C) ONLY - Which 501(C) do you represent?", "501c Non Profit Name", "501(c) Organization"] },
    { key: "taxLetter", aliases: ["Can you provide a Tax Deductibility Letter?", "FOR 501(C) ONLY - Are you able to provide a TAX DEDUCTIBILITY LETTER?", "Tax Deductibility Letter", "Tax Letter"] },
    { key: "boothSpace", aliases: ["Provide a Booth/Exhibitor Space (10×10 Tent)?", "Provide a Booth/Exhibitor Space (10x10 Tent)?", "Will you provide a BOOTH/EXHIBITOR space (10 x 10 Tent) to promote our services?", "Booth Space"] },
    { key: "promoInclude", aliases: ["Include our logo on promo materials & social media?", "Will you INCLUDE OUR INFORMATION and logo on all promotional materials, including social media?", "Include Logo on Promo", "Promo Include"] },
    { key: "allowHelpPromote", aliases: ["Allowed to help promote the event?", "Are we allowed to HELP PROMOTE the event?", "Allow Help Promote", "Promote Event"] },
    { key: "videoRights", aliases: ["Provide copies of video footage and photos?", "Can we get COPIES of video footage and pictures of our participation?", "Video Footage & Photos", "Video Rights"] },
    { key: "contingencyPlan", aliases: ["Weather / Contingency Plan", "What is the CONTINGENCY PLAN? Please include alternate locations, dates, times.", "Do you have a CONTINGENCY PLAN?", "Contingency Plan"] },
    { key: "privateGatheringType", aliases: ["Type of Private Gathering", "TYPE of Private Gathering", "Private Gathering Type"] },
    { key: "invitedToAttend", aliases: ["Are performers invited to attend/stay for the event?", "Are we INVITED TO ATTEND the event?", "Invited to Attend"] },
    { key: "intlVenueName", aliases: ["International: Specific Country, City, & Venue Name", "Specific Country, City, & Venue Name", "International Venue Name"] },
    { key: "intlLogistics", aliases: ["International: Travel & Lodging Logistics", "International Travel & Lodging Logistics", "International Logistics"] },
    { key: "intlVisaSupport", aliases: ["International: Visa & Legal Documentation Support", "Visa & Legal Documentation Support", "Visa Support", "Visa Legal"] },
    { key: "intlPaymentTerms", aliases: ["International: Preferred Currency & Payment Terms", "Preferred Currency & Payment Terms", "Preferred Currency", "Payment Terms"] },
    { key: "intlCustomsConsiderations", aliases: ["International: Costumes, Props & Customs Considerations", "Costumes, Props & Customs Considerations", "Customs Notes", "Customs Considerations"] },
    { key: "repertoireMexico", aliases: ["Repertoire: Mexico / North American Dances"] },
    { key: "repertoireCaribbean", aliases: ["Repertoire: Caribbean Dances (Cuba & Puerto Rico)"] },
    { key: "repertoireCentralAmerica", aliases: ["Repertoire: Central American Dances (El Salvador)"] },
    { key: "repertoireSouthAmerica", aliases: ["Repertoire: South American Dances (Colombia & Argentina)"] },
    { key: "repertoireTheatrical", aliases: ["Repertoire: Theatrical, Parade & Live Singing Performances"] },
    { key: "performanceServices", aliases: ["Which of our PERFORMANCE SERVICES will you need?", "Performance Services Needed", "Performance Services"] },
    { key: "otherPerfServices", aliases: ["Any other PERFORMANCE SERVICES you wish, but are not listed above?", "Other Performance Services"] },
    { key: "lessonServices", aliases: ["Which of our DANCE LESSON SERVICES will you need?", "Dance Lesson Services Needed", "Lesson Services"] },
    { key: "audienceParticipation", aliases: ["Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?", "Expecting AUDIENCE PARTICIPATION?", "Audience Participation"] },
    { key: "durationRequired", aliases: ["How much TIME do you require from us?", "DURATION of Service Required", "Duration Required", "Duration"] },
    { key: "otherServices", aliases: ["Additional Services Needed (MC, DJ, Lecture)", "Additional Services Needed", "Other Services"] },
    { key: "generalFormats", aliases: ["General Formats (Stage, Opening, Headliner, Main Act, Background)", "General Formats", "Formats"] },
    { key: "venueSetting", aliases: ["Venue Location Setting", "Where will it take PLACE?", "Venue Setting", "Indoor/Outdoor"] },
    { key: "performanceSurface", aliases: ["On what SURFACE will the performance or class take place?", "Performance Surface", "Surface Type", "Floor Type"] },
    { key: "performanceArea", aliases: ["Size of Performance / Class Area", "How large is the AREA for our performance or lesson?", "Performance Area Size", "Area Size"] },
    { key: "soundEquipment", aliases: ["Sound System Equipment", "About the SOUND SYSTEM", "Sound Equipment Provided", "Sound System"] },
    { key: "badgeRequired", aliases: ["Will a BADGE or ID be required for performers?", "Will a BADGE or ID be issued to performers to access the performance area?", "Badge Access", "ID Badge"] },
    { key: "hospitalityProvided", aliases: ["WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)", "WILL YOU PROVIDE the performers with:", "Hospitality Provided", "Performer Provisions", "Provisions"] },
    { key: "dressingRoomDetails", aliases: ["Dressing Room / Costume Changing Instructions", "Will we have a place to change COSTUMES if needed? If so, please provide instructions.", "Dressing Room Instructions", "Costume Change Room"] },
    { key: "hasBudget", aliases: ["Confirm you have a BUDGET for our participation", "Confirm Budget", "Has Budget"] },
    { key: "confirmedBudget", aliases: ["Confirmed Budget Amount for Performance / Workshop", "Confirmed Budget Amount", "Confirmed Budget", "Budget Amount"] },
    { key: "notes", aliases: ["Special Instructions, Song Requests or Notes", "SPECIAL REQUESTS or Song Preferences", "Notes", "Special Requests"] },
    { key: "termsAgreed", aliases: ["Terms of Service & Privacy Policy Agreement", "Terms Agreed", "Terms of Service"] },
    { key: "sendEmailReceipt", aliases: ["Send Email Receipt", "Email Receipt", "Receipt"] },
    { key: "serviceCategory", aliases: ["Select Event Classification:", "Service Category / Scope", "Service Category", "EventType Scale"] }
  ];

  fieldHeaderMap.forEach(item => {
    let colIdx = getHeaderIndex(headers, item.aliases);
    if (colIdx > -1 && data[item.key] !== undefined && data[item.key] !== null) {
      let val = data[item.key];
      if (Array.isArray(val)) val = val.join(", ");
      newRow[colIdx] = val;
    }
  });

  // Guarantee complete Timestamp / Submission Date based on exact time of submittal
  const timestampIdx = getHeaderIndex(headers, ["Submission Date", "Timestamp", "Date Submitted", "Date", "Submission Timestamp"]);
  if (timestampIdx > -1 && (!newRow[timestampIdx] || newRow[timestampIdx] === "")) {
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "America/New_York";
    newRow[timestampIdx] = Utilities.formatDate(new Date(), tz, "M/d/yyyy H:mm:ss");
  }

  sheet.appendRow(newRow);
  const rowNum = sheet.getLastRow();

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
  try {
    processRow(sheet, rowNum, updatedRowData, mainFolder, headers, true);
  } catch (procErr) {
    console.warn("processRow non-fatal error during submit: ", procErr);
  }

  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID"]);
  const finalEventId = eventIdIdx > -1 ? sheet.getRange(rowNum, eventIdIdx + 1).getValue() : "";

  return {
    status: "success",
    row: rowNum,
    eventId: finalEventId || `BTG-EVT-${rowNum}-${Date.now().toString(36).toUpperCase()}`
  };
}

/**
 * Web App HTTP GET endpoint.
 */
function doGet(e) {
  return HtmlService.createHtmlOutput("Salsa Guy Richmond LLC Automation Suite Web App Endpoint (v20.71) Active.");
}

/**
 * Web App HTTP POST endpoint for receiving custom form submissions (JSON/Form payload).
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        payload = e.parameter || {};
      }
    } else if (e.parameter) {
      payload = e.parameter;
    }

    const result = handleFormSubmitJson(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper function to find column index in headers array matching a list of aliases.
 */
function getHeaderIndex(headers, aliases) {
  if (!headers || !aliases || !Array.isArray(aliases)) return -1;
  const normHeaders = headers.map(h => (h ? h.toString().toLowerCase().trim() : ""));
  
  for (let alias of aliases) {
    const normAlias = alias.toLowerCase().trim();
    let idx = normHeaders.indexOf(normAlias);
    if (idx > -1) return idx;
    
    idx = normHeaders.findIndex(h => h && (h.includes(normAlias) || normAlias.includes(h)));
    if (idx > -1) return idx;
  }
  return -1;
}

/**
 * Helper function to find or create an event subfolder inside the main Drive folder.
 */
function getOrCreateEventFolder(parentFolder, dateStr, eventName) {
  const folderName = `${dateStr} - ${eventName}`.replace(/[/\\?%*:|"<>]/g, '_');
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}



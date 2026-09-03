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
  ADMIN_EMAILS: ["salsaguyrichmond@gmail.com"],
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
    .addItem('📧 SEND TEST NOTIFICATION EMAIL', 'sendTestNotificationEmail')
    .addItem('🔔 SETUP AUTO-SUBMISSION TRIGGER', 'installFormSubmitTrigger')
    .addItem('🔍 DIAGNOSE NOTIFICATION SYSTEM', 'diagnoseNotificationSystem')
    .addSeparator()
    .addItem('🚀 RUN FULL AUTOMATION', 'mainAutomation')
    .addItem('🎯 UPDATE SELECTED ROW ONLY', 'reGenerateDocs')
    .addSeparator()
    .addItem('📄 UPDATE MASTER DOC TEMPLATES', 'updateMasterTemplates')
    .addSeparator()
    .addItem('🎨 COLOR CODE HEADERS', 'colorCodeHeaders')
    .addItem('🔍 DIAGNOSE SHEET (Check Headers)', 'diagnoseSheet')
    .addSeparator()
    .addItem("📝 1. SETUP PROFE'S QUICK ENTRY", 'setupQuickEntryTab')
    .addItem("📤 2. UPLOAD PROFE'S QUICK ENTRY", 'importQuickEntryData')
    .addSeparator()
    .addItem("📄 CREATE REVIEW GOOGLE DOC", 'createQuestionnaireReviewGoogleDoc')
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
    "Request ID",
    "Event ID",
    "Your Name",
    "Email Address",
    "Best Contact Phone Number",
    "Who do you represent? (Organization / Business / Self)",
    "How did you HEAR of us?",

    // 2. Event Overview & Schedule
    "What is the NAME of the event?",
    "Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title",
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
    "Out-of-State Travel & Logistics Arrangement",
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
    "Service Type Requested",
    "Troupe Headcount / Ensemble Size",
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
 * Includes a 5-minute safety timeout guard to prevent Google Apps Script limit errors.
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

  const startTime = Date.now();
  let processedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const eventName = data[i][eventNameIdx] || "";
    const status = statusIdx > -1 ? data[i][statusIdx] : "";
    
    if (eventName === "" || status === "Synced") continue;
    
    // Safety guard: Stop at 5 minutes (300,000 ms) to stay well under Apps Script's 6-minute hard limit
    if (Date.now() - startTime > 300000) {
      console.warn("Script approaching 5-minute execution limit. Stopping batch loop gracefully.");
      SpreadsheetApp.getUi().alert(`⏳ Processed ${processedCount} row(s). Batch paused to prevent Google limit error. Run '🚀 RUN FULL AUTOMATION' again to process remaining rows.`);
      return;
    }

    try { 
      processRow(sheet, i + 1, data[i], folder, headers); 
      processedCount++;
      Utilities.sleep(100);
    } catch (e) { 
      console.error(`Error processing row ${i + 1}: ` + e.message); 
    }
  }

  if (processedCount > 0) {
    SpreadsheetApp.getUi().alert(`🚀 Automation complete! Successfully processed ${processedCount} row(s).`);
  }
}

/**
 * Processes a single row: formatting attachments, syncing Google Calendar, generating documents, updating status, and sending notifications.
 */
function processRow(sheet, rowNum, rowData, folder, headers, isNewSubmit = false, directFileUrl = "", skipEmailNotification = false) {
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
  try {
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
  } catch (attErr) {
    console.warn("Attachment hyperlink formatting skipped: " + attErr.message);
  }

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
    try {
      const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const dayName = days[eDate.getDay()];
      sheet.getRange(rowNum, dowIdx + 1).setValue(dayName);
    } catch (dowErr) {}
  }

  let calendarEventUrl = "";
  if (isValidDate) {
    try {
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
    } catch (calErr) {
      console.warn("Calendar synchronization skipped or failed: " + calErr.message);
    }
  }

  if (isValidDate && dateIdx > -1) {
    try {
      const formattedDateStr = Utilities.formatDate(eDate, Session.getScriptTimeZone(), "MM-dd-yyyy");
      if (calendarEventUrl) {
        sheet.getRange(rowNum, dateIdx + 1).setFormula(`=HYPERLINK("${calendarEventUrl}", "${formattedDateStr}")`);
      } else {
        sheet.getRange(rowNum, dateIdx + 1).setValue(formattedDateStr);
      }
    } catch (dErr) {}
  }

  let propUrl = "Error";
  let contUrl = "Error";
  let perfUrl = "Error";

  try {
    propUrl = createDoc(CONFIG.TEMPLATES.PROPOSAL, "Proposal", eventName, rowData, folder, isValidDate, eDate, headers);
    if (propUrl !== "Error" && propUrlIdx > -1) sheet.getRange(rowNum, propUrlIdx + 1).setFormula(`=HYPERLINK("${propUrl}", "View Proposal")`);
  } catch (pErr) {
    console.warn("Proposal document generation error: " + pErr.message);
  }

  try {
    contUrl = createDoc(CONFIG.TEMPLATES.CONTRACT, "Contract", eventName, rowData, folder, isValidDate, eDate, headers);
    if (contUrl !== "Error" && contUrlIdx > -1) sheet.getRange(rowNum, contUrlIdx + 1).setFormula(`=HYPERLINK("${contUrl}", "View Contract")`);
  } catch (cErr) {
    console.warn("Contract document generation error: " + cErr.message);
  }

  try {
    perfUrl = createDoc(CONFIG.TEMPLATES.PERF_INFO, "Performance", eventName, rowData, folder, isValidDate, eDate, headers);
    if (perfUrl !== "Error" && perfUrlIdx > -1) sheet.getRange(rowNum, perfUrlIdx + 1).setFormula(`=HYPERLINK("${perfUrl}", "View Info")`);
  } catch (perfErr) {
    console.warn("Performance Info document generation error: " + perfErr.message);
  }

  try {
    if (statusIdx > -1) {
      sheet.getRange(rowNum, statusIdx + 1).setValue("Synced").setBackground("#d9ead3");
    }
    if (internalStatusIdx > -1 && !rowData[internalStatusIdx]) {
      sheet.getRange(rowNum, internalStatusIdx + 1).setValue("Ready for Review");
    }
    if (assignedToIdx > -1 && !rowData[assignedToIdx]) {
      sheet.getRange(rowNum, assignedToIdx + 1).setValue("The Salsa Guy");
    }
  } catch (sErr) {}

  // Handle Email Notifications for Sheet/Google Form submissions or re-runs (if not already sent)
  if (isNewSubmit && !skipEmailNotification) {
    try {
      const clientNameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
      const clientEmailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
      const clientPhoneIdx = getHeaderIndex(headers, ["Best Contact Phone Number", "Phone Number", "Phone", "Telephone"]);
      const representTypeIdx = getHeaderIndex(headers, ["Who do you represent? (Organization / Business / Self)", "Who do you represent?", "Represent Type", "Representing"]);
      const budgetIdx = getHeaderIndex(headers, ["Confirmed Budget Amount for Performance / Workshop", "Confirmed Budget Amount", "Confirmed Budget", "Budget Amount"]);
      const serviceTypeIdx = getHeaderIndex(headers, ["Service Type Requested", "Service Category / Scope", "Service Category", "Primary Service"]);
      const troupeIdx = getHeaderIndex(headers, ["Troupe Headcount / Ensemble Size", "Troupe Headcount", "Ensemble Size"]);
      const durationIdx = getHeaderIndex(headers, ["How much TIME do you require from us?", "DURATION of Service Required", "Duration Required", "Duration"]);
      const notesIdx = getHeaderIndex(headers, ["Special Instructions, Song Requests or Notes", "SPECIAL REQUESTS or Song Preferences", "Notes", "Special Requests"]);
      const sendReceiptIdx = getHeaderIndex(headers, ["Send Email Receipt", "Email Receipt", "Receipt"]);

      const clientName = (clientNameIdx > -1 && rowData[clientNameIdx]) ? rowData[clientNameIdx] : "Valued Client";
      const clientEmail = (clientEmailIdx > -1 && rowData[clientEmailIdx]) ? rowData[clientEmailIdx] : "";
      const clientPhone = (clientPhoneIdx > -1 && rowData[clientPhoneIdx]) ? rowData[clientPhoneIdx] : "N/A";
      const representType = (representTypeIdx > -1 && rowData[representTypeIdx]) ? rowData[representTypeIdx] : "N/A";
      const budgetAmount = (budgetIdx > -1 && rowData[budgetIdx]) ? rowData[budgetIdx] : "Pending Confirmation";
      const serviceType = (serviceTypeIdx > -1 && rowData[serviceTypeIdx]) ? rowData[serviceTypeIdx] : "Dance Booking";
      const troupeHeadcount = (troupeIdx > -1 && rowData[troupeIdx]) ? rowData[troupeIdx] : "Standard Ensemble";
      const durationRequired = (durationIdx > -1 && rowData[durationIdx]) ? rowData[durationIdx] : "As specified";
      const notes = (notesIdx > -1 && rowData[notesIdx]) ? rowData[notesIdx] : "";
      const eventAddress = (addrIdx > -1 && rowData[addrIdx]) ? rowData[addrIdx] : "TBD";
      const eventTimeStr = (timeIdx > -1 && rowData[timeIdx]) ? extractTimeOnly(rowData[timeIdx]) : "TBD";
      const eventDateStr = isValidDate && eDate ? formatDateValue(eDate) : ((dateIdx > -1 && rowData[dateIdx]) ? rowData[dateIdx] : "TBD");
      
      const fullId = eventIdIdx > -1 ? (rowData[eventIdIdx] ? rowData[eventIdIdx].toString().trim() : "") : "";
      const finalEventId = fullId || `BTG-EVT-${rowNum}-${Date.now().toString(36).toUpperCase()}`;

      let fileUrl = directFileUrl || "";
      if (!fileUrl) {
        headers.forEach((h, i) => {
          if (h.toLowerCase().includes("upload") || h.toLowerCase().includes("attachment") || h.toLowerCase().includes("flyer")) {
            const val = rowData[i];
            if (val && typeof val === "string") {
              const urls = val.split(",").map(u => u.trim()).filter(u => u.startsWith("http"));
              if (urls.length > 0) fileUrl = urls[0];
            }
          }
        });
      }

      let eventFolderUrl = "";
      if (folder) {
        try {
          const dateFolderStr = rowData[dateIdx] || "Undated";
          const subFolders = folder.getFoldersByName(`${dateFolderStr} - ${eventName}`.replace(/[/\\?%*:|"<>]/g, '_'));
          if (subFolders.hasNext()) {
            eventFolderUrl = subFolders.next().getUrl();
          }
        } catch (fErr) {}
      }

      const notificationDetails = {
        rowNum: rowNum,
        eventId: finalEventId,
        clientName: clientName,
        clientEmail: clientEmail,
        clientPhone: clientPhone,
        representType: representType,
        budgetAmount: budgetAmount,
        serviceTypeRequested: serviceType,
        troupeHeadcount: troupeHeadcount,
        durationRequired: durationRequired,
        notes: notes,
        eventName: eventName,
        eventDate: eventDateStr,
        eventTime: eventTimeStr,
        eventAddress: eventAddress,
        calendarEventUrl: calendarEventUrl,
        eventFolderUrl: eventFolderUrl,
        propUrl: propUrl,
        contUrl: contUrl,
        perfUrl: perfUrl,
        fileUrl: fileUrl
      };

      // 1. Send Admin Notification Email
      sendAdminSubmittalNotification(notificationDetails);

      // 2. Send Client Receipt Email if requested
      const shouldSendReceipt = sendReceiptIdx > -1 ? rowData[sendReceiptIdx] : true;
      if (shouldSendReceipt && shouldSendReceipt !== "No" && shouldSendReceipt !== false) {
        sendClientReceiptNotification(notificationDetails);
      }
    } catch (notifErr) {
      console.warn("Error triggering submittal email notifications in processRow: ", notifErr);
    }
  }
}

/**
 * Safely dispatches email with fallback from MailApp to GmailApp.
 * Includes plain-text fallback generation and sender display name.
 */
function sendEmailSafely(toEmail, subject, htmlBody) {
  if (!toEmail || typeof toEmail !== "string" || toEmail.trim() === "") return false;
  const target = toEmail.trim();

  // Create clean plain-text version for email clients that don't render HTML
  const plainBody = htmlBody
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();

  const senderName = "Salsa Guy Richmond LLC";

  try {
    MailApp.sendEmail({
      to: target,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: senderName
    });
    console.log("Successfully sent email via MailApp to: " + target);
    return true;
  } catch (err1) {
    console.warn("MailApp.sendEmail failed for " + target + " (" + err1.message + "), trying GmailApp fallback...");
    try {
      GmailApp.sendEmail(target, subject, plainBody, {
        htmlBody: htmlBody,
        name: senderName
      });
      console.log("Successfully sent email via GmailApp to: " + target);
      return true;
    } catch (err2) {
      console.error("GmailApp.sendEmail fallback also failed for " + target + ": " + err2.message);
      return false;
    }
  }
}

/**
 * Sends a rich, instant admin notification email to all configured admin recipients.
 */
function sendAdminSubmittalNotification(details) {
  if (!CONFIG.ADMIN_EMAILS || CONFIG.ADMIN_EMAILS.length === 0) return;

  const cleanPhone = (details.clientPhone || '').replace(/[^\d+]/g, '');
  const intlPhone = cleanPhone.startsWith('1') ? cleanPhone : ('1' + cleanPhone.replace(/^0+/, ''));

  const budgetDisplay = details.budgetAmount && details.budgetAmount !== "Pending Confirmation" 
    ? (details.budgetAmount.startsWith("$") ? details.budgetAmount : `$${details.budgetAmount}`)
    : "Pending Confirmation";

  const subject = `🔥 NEW BOOKING: ${details.eventName} - ${budgetDisplay} - ${details.clientName} (ID: ${details.eventId})`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; background-color: #f8fafc; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0;">
      <div style="background-color: #dc2626; padding: 18px 24px; border-radius: 8px 8px 0 0; color: #ffffff;">
        <h2 style="margin: 0; font-size: 20px;">💃 Salsa Guy Richmond LLC - New Submittal Alert</h2>
        <p style="margin: 4px 0 0 0; font-size: 14px; opacity: 0.9;">Instant Booking Inquiry Notification & Lead Details</p>
      </div>

      <!-- Quick 1-Tap Mobile Action Bar for Angel -->
      <div style="background-color: #fef3c7; border-left: 4px solid #f59e0b; padding: 14px 18px; margin-top: 0; margin-bottom: 0; border-right: 1px solid #e2e8f0; border-left: 1px solid #e2e8f0;">
        <p style="margin: 0 0 10px 0; font-size: 14px; font-weight: bold; color: #92400e;">⚡ Fast-Track Lead Follow-Up Actions:</p>
        <div style="display: flex; gap: 10px; flex-wrap: wrap;">
          <a href="tel:${cleanPhone}" style="background-color: #2563eb; color: #ffffff; padding: 9px 16px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 13px; display: inline-block;">📞 1-Tap Call (${details.clientPhone})</a>
          <a href="https://wa.me/${intlPhone}?text=Hello%20${encodeURIComponent(details.clientName)}%2C%20this%20is%20Angel%20from%20Salsa%20Guy%20Richmond%20following%20up%20on%20your%20event%20booking%20for%20${encodeURIComponent(details.eventName)}" target="_blank" style="background-color: #22c55e; color: #ffffff; padding: 9px 16px; border-radius: 6px; text-decoration: none; font-weight: bold; font-size: 13px; display: inline-block;">💬 1-Tap WhatsApp</a>
        </div>
      </div>

      <div style="background-color: #ffffff; padding: 24px; border-radius: 0 0 8px 8px; border: 1px solid #e2e8f0; border-top: none;">
        
        <!-- STEP 1: Response Logging & Client Details -->
        <div style="margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid #f1f5f9;">
          <h3 style="color: #b45309; margin: 0 0 10px 0; font-size: 16px;">📌 Client & Lead Contact</h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #334155;">
            <tr><td style="padding: 4px 0; font-weight: bold; width: 150px;">Request ID:</td><td><span style="background-color: #fef3c7; color: #92400e; padding: 2px 8px; border-radius: 4px; font-weight: bold;">${details.eventId}</span></td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Confirmed Budget:</td><td><span style="background-color: #dcfce7; color: #166534; padding: 2px 8px; border-radius: 4px; font-weight: bold; font-size: 14px;">${budgetDisplay}</span></td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Master Sheet Row:</td><td>Row #${details.rowNum} (<a href="${CONFIG.SPREADSHEET_URL}" target="_blank" style="color: #2563eb; text-decoration: none;">View Master Google Sheet</a>)</td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Client Name:</td><td>${details.clientName}</td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Email:</td><td><a href="mailto:${details.clientEmail}" style="color: #2563eb; font-weight: bold;">${details.clientEmail}</a></td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Phone:</td><td><a href="tel:${cleanPhone}" style="color: #2563eb; font-weight: bold;">${details.clientPhone}</a></td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Representing:</td><td>${details.representType}</td></tr>
          </table>
        </div>

        <!-- STEP 2: Service & Event Specifications -->
        <div style="margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid #f1f5f9;">
          <h3 style="color: #1d4ed8; margin: 0 0 10px 0; font-size: 16px;">🎭 Event & Service Requirements</h3>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #334155;">
            <tr><td style="padding: 4px 0; font-weight: bold; width: 150px;">Event Name:</td><td><strong>${details.eventName}</strong></td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Date & Time:</td><td>${details.eventDate} at ${details.eventTime}</td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Location Address:</td><td>${details.eventAddress}</td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Service Requested:</td><td>${details.serviceTypeRequested || 'Dance Booking'}</td></tr>
            <tr><td style="padding: 4px 0; font-weight: bold;">Troupe Headcount:</td><td>${details.troupeHeadcount || 'Standard Ensemble'}</td></tr>
            ${details.performanceServices ? `<tr><td style="padding: 4px 0; font-weight: bold;">Repertoire/Dances:</td><td>${details.performanceServices}</td></tr>` : ''}
            ${details.lessonServices ? `<tr><td style="padding: 4px 0; font-weight: bold;">Lessons Needed:</td><td>${details.lessonServices}</td></tr>` : ''}
            <tr><td style="padding: 4px 0; font-weight: bold;">Duration Required:</td><td>${details.durationRequired || 'As requested'}</td></tr>
            ${details.notes ? `<tr><td style="padding: 4px 0; font-weight: bold;">Special Notes:</td><td style="background-color: #f1f5f9; padding: 6px 10px; border-radius: 4px; font-style: italic;">${details.notes}</td></tr>` : ''}
            <tr><td style="padding: 4px 0; font-weight: bold;">Google Calendar:</td><td>${details.calendarEventUrl ? `<a href="${details.calendarEventUrl}" target="_blank" style="color: #2563eb; font-weight: bold;">📆 View Calendar Event</a>` : 'Date TBD / Not Scheduled'}</td></tr>
          </table>
        </div>

        <!-- STEP 3: Google Docs & Drive Links -->
        <div style="margin-bottom: 20px; padding-bottom: 16px; border-bottom: 1px solid #f1f5f9;">
          <h3 style="color: #15803d; margin: 0 0 10px 0; font-size: 16px;">📂 Automated Google Docs & Drive</h3>
          <ul style="margin: 0; padding-left: 20px; font-size: 14px; color: #334155; line-height: 1.8;">
            ${details.eventFolderUrl ? `<li>📁 <strong>Drive Event Folder:</strong> <a href="${details.eventFolderUrl}" target="_blank" style="color: #2563eb;">Open Google Drive Folder</a></li>` : ''}
            <li>📄 <strong>Proposal Doc:</strong> ${details.propUrl && details.propUrl !== "Error" ? `<a href="${details.propUrl}" target="_blank" style="color: #2563eb;">View Proposal Document</a>` : 'Ready in Google Drive'}</li>
            <li>📜 <strong>Contract Doc:</strong> ${details.contUrl && details.contUrl !== "Error" ? `<a href="${details.contUrl}" target="_blank" style="color: #2563eb;">View Contract Document</a>` : 'Ready in Google Drive'}</li>
            <li>🎭 <strong>Performance Info Doc:</strong> ${details.perfUrl && details.perfUrl !== "Error" ? `<a href="${details.perfUrl}" target="_blank" style="color: #2563eb;">View Performance Info Document</a>` : 'Ready in Google Drive'}</li>
          </ul>
        </div>

        <!-- STEP 4: Attachments -->
        <div>
          <h3 style="color: #6b21a8; margin: 0 0 10px 0; font-size: 16px;">📎 File Attachments</h3>
          <p style="font-size: 14px; color: #334155; margin: 0;">
            ${details.fileUrl ? `📎 <strong>Uploaded Attachment:</strong> <a href="${details.fileUrl}" target="_blank" style="color: #2563eb; font-weight: bold;">View Attached Document</a>` : 'No document attached by client.'}
          </p>
        </div>

      </div>

      <div style="text-align: center; margin-top: 16px; font-size: 12px; color: #64748b;">
        Salsa Guy Richmond LLC Automation Suite • Instant Admin Notification
      </div>
    </div>
  `;

  CONFIG.ADMIN_EMAILS.forEach(email => {
    sendEmailSafely(email, subject, htmlBody);
  });
}

/**
 * Sends an instant confirmation receipt email to the client if requested.
 */
function sendClientReceiptNotification(details) {
  if (!details.clientEmail || details.clientEmail.trim() === "") return;

  const subject = `Confirmation Receipt: Your Event Booking Request - ${details.eventName} (Request ID: ${details.eventId})`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 24px; border-radius: 12px; border: 1px solid #e2e8f0;">
      <div style="text-align: center; margin-bottom: 20px;">
        <h2 style="color: #dc2626; margin: 0; font-size: 22px;">💃 Salsa Guy Richmond LLC</h2>
        <p style="color: #64748b; margin: 4px 0 0 0; font-size: 14px;">Thank you for your Event Request!</p>
      </div>

      <p style="font-size: 15px; color: #334155;">Dear <strong>${details.clientName}</strong>,</p>
      
      <p style="font-size: 14px; color: #334155; line-height: 1.6;">
        We have received your event booking request! Your official Request Tracking ID is below. Our team is currently reviewing your details and preparing your proposal.
      </p>

      <div style="background-color: #fffbeb; border: 1px dashed #f59e0b; padding: 16px; border-radius: 8px; text-align: center; margin: 20px 0;">
        <span style="font-size: 12px; color: #b45309; text-transform: uppercase; letter-spacing: 1px; display: block; margin-bottom: 4px;">Your Request Tracking ID</span>
        <span style="font-size: 20px; font-weight: bold; color: #92400e;">${details.eventId}</span>
      </div>

      <h3 style="font-size: 15px; color: #0f172a; margin-bottom: 8px;">Submission Summary:</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 14px; color: #334155; margin-bottom: 20px;">
        <tr><td style="padding: 4px 0; font-weight: bold; width: 140px;">Event Name:</td><td>${details.eventName}</td></tr>
        <tr><td style="padding: 4px 0; font-weight: bold;">Date & Time:</td><td>${details.eventDate} at ${details.eventTime}</td></tr>
        <tr><td style="padding: 4px 0; font-weight: bold;">Location:</td><td>${details.eventAddress}</td></tr>
        <tr><td style="padding: 4px 0; font-weight: bold;">Service Type:</td><td>${details.serviceTypeRequested || 'Dance Event'}</td></tr>
      </table>

      <p style="font-size: 14px; color: #334155; line-height: 1.6;">
        If you have any questions or need to make updates, please contact us directly at <a href="mailto:salsaguyrichmond@gmail.com" style="color: #dc2626; font-weight: bold;">salsaguyrichmond@gmail.com</a> or call/text <a href="tel:8045550199" style="color: #2563eb; font-weight: bold;">(804) 555-0199</a>.
      </p>

      <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 24px 0 16px 0;" />
      
      <div style="text-align: center; font-size: 12px; color: #94a3b8;">
        Salsa Guy Richmond LLC • Bring the Passion of Dance to Your Event
      </div>
    </div>
  `;

  sendEmailSafely(details.clientEmail, subject, htmlBody);
}

/**
 * Test function that sends a verification email to all admin emails in CONFIG.
 * Verifies email sending permissions and displays daily email quota.
 */
function sendTestNotificationEmail() {
  const ui = (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getUi) ? SpreadsheetApp.getUi() : null;
  const testDetails = {
    rowNum: 99,
    eventId: "BTG-TEST-" + Date.now().toString(36).toUpperCase(),
    clientName: "Angel A. Rodriguez (Test Notification)",
    clientEmail: "salsaguyrichmond@gmail.com",
    clientPhone: "804-555-0199",
    representType: "Salsa Guy Richmond LLC",
    eventName: "Salsa & Bachata Gala Showcase 2026",
    eventDate: "October 15, 2026",
    eventTime: "6:00 PM",
    eventAddress: "123 E Franklin St, Richmond, VA 23219",
    budgetAmount: "$1,500.00",
    serviceTypeRequested: "Performance & Dance Instruction",
    troupeHeadcount: "4 to 6 Performers",
    performanceServices: "Dances from Disney COCO, Salsa Rueda (Cuba), El Baile de los Viejitos (Mexico)",
    lessonServices: "Salsa (Beginner / Intermediate), Bachata",
    durationRequired: "2 Hours",
    soundEquipment: "Venue provides PA system & microphones",
    notes: "Special test run to verify email notification dispatch and instant mobile alert delivery.",
    calendarEventUrl: "https://calendar.google.com",
    eventFolderUrl: CONFIG.SPREADSHEET_URL,
    propUrl: "",
    contUrl: "",
    perfUrl: "",
    fileUrl: ""
  };

  sendAdminSubmittalNotification(testDetails);
  
  let quota = -1;
  try {
    quota = MailApp.getRemainingDailyQuota();
  } catch (qErr) {}

  const msg = `✅ Test notification email successfully dispatched to:\n${CONFIG.ADMIN_EMAILS.join(", ")}\n\n` +
    (quota >= 0 ? `📧 Remaining Daily Email Quota: ${quota} emails\n\n` : "") +
    `Please check your inbox (and Spam/Promotions folder) for the "🔥 NEW BOOKING" alert.`;

  if (ui) {
    ui.alert("📧 Email Notification Test", msg, ui.ButtonSet.OK);
  } else {
    console.log(msg);
  }
}

/**
 * Runs a complete health check on the entire notification and automation pipeline.
 */
function diagnoseNotificationSystem() {
  const ui = SpreadsheetApp.getUi();
  let report = "🔍 Notification & Automation System Diagnosis:\n\n";

  // 1. Admin Emails Check
  if (CONFIG.ADMIN_EMAILS && CONFIG.ADMIN_EMAILS.length > 0) {
    report += `✅ Admin Recipients: ${CONFIG.ADMIN_EMAILS.join(", ")}\n`;
  } else {
    report += `❌ No admin emails configured in CONFIG.ADMIN_EMAILS!\n`;
  }

  // 2. Email Quota Check
  try {
    const quota = MailApp.getRemainingDailyQuota();
    report += `✅ Google MailApp Daily Quota Remaining: ${quota} emails\n`;
  } catch (e) {
    report += `⚠️ MailApp Quota check failed: ${e.message}\n`;
  }

  // 3. Spreadsheet Access Check
  const ss = getSpreadsheet();
  if (ss) {
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet();
    report += `✅ Spreadsheet Connected: "${sheet ? sheet.getName() : 'Active Sheet'}" (${ss.getId()})\n`;
  } else {
    report += `❌ Failed to connect to spreadsheet!\n`;
  }

  // 4. Drive Folder Check
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    report += `✅ Google Drive Folder Connected: "${folder.getName()}"\n`;
  } catch (e) {
    report += `⚠️ Drive Folder Access warning: ${e.message}\n`;
  }

  // 5. Calendar Check
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.INFOCALENDAR_ID);
    if (cal) {
      report += `✅ InfoCalendar Connected: "${cal.getName()}"\n`;
    } else {
      report += `⚠️ InfoCalendar not found by ID (${CONFIG.INFOCALENDAR_ID})\n`;
    }
  } catch (e) {
    report += `⚠️ Calendar check error: ${e.message}\n`;
  }

  // 6. Form Submit Trigger Check
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const formTrigger = triggers.find(t => t.getHandlerFunction() === "onFormSubmit");
    if (formTrigger) {
      report += `✅ Form Submit Trigger Active: onFormSubmit (${formTrigger.getEventType()})\n`;
    } else {
      report += `ℹ️ Form Submit Trigger: Not installed yet (Use '🔔 SETUP AUTO-SUBMISSION TRIGGER' if using Google Forms)\n`;
    }
  } catch (e) {
    report += `⚠️ Trigger check error: ${e.message}\n`;
  }

  ui.alert("System Health Report", report, ui.ButtonSet.OK);
}

/**
 * Installs an installable onFormSubmit trigger on the active spreadsheet if not already installed.
 */
function installFormSubmitTrigger() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = getSpreadsheet();
    if (!ss) {
      ui.alert("⚠️ Unable to access spreadsheet to install trigger.");
      return;
    }

    const triggers = ScriptApp.getProjectTriggers();
    const existing = triggers.find(t => t.getHandlerFunction() === "onFormSubmit");
    if (existing) {
      ui.alert("ℹ️ Form Submit Trigger is already installed and active!");
      return;
    }

    ScriptApp.newTrigger("onFormSubmit")
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();

    ui.alert("✅ Form Submit Trigger successfully installed!\n\nNew Google Forms or Sheet submittals will now automatically trigger email notifications.");
  } catch (e) {
    ui.alert("❌ Error installing trigger: " + e.message);
  }
}

/**
 * Handles installable onFormSubmit event from Google Sheets or linked Google Form.
 */
function onFormSubmit(e) {
  try {
    const ss = getSpreadsheet();
    const sheet = (e && e.range) ? e.range.getSheet() : (ss ? (ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getActiveSheet()) : null);
    if (!sheet) return;
    
    const rowNum = e && e.range ? e.range.getRow() : sheet.getLastRow();
    const headers = getUniqueHeaders(sheet);
    const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    let folder = null;
    try {
      folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    } catch (fErr) {}
    
    processRow(sheet, rowNum, rowData, folder, headers, true, "", false);
  } catch (err) {
    console.error("Error in onFormSubmit trigger: " + err.message);
  }
}

/**
 * Safely replaces all occurrences of searchText with literal replacementValue in a DocumentApp element.
 * Escapes Java regex special characters in search text and escapes $ and \ in replacement text.
 */
function safeReplaceText(element, searchText, replacementValue) {
  if (!element || !searchText) return;
  const safeValue = (replacementValue === null || replacementValue === undefined) ? "" : replacementValue.toString();
  // Escape all 14 regex special characters in search text
  const escapedSearch = searchText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  // Escape $ and \ in replacement value for Java regex replacement
  const safeReplacement = safeValue.replace(/\\/g, "\\\\").replace(/\$/g, "\\$");
  try {
    element.replaceText(escapedSearch, safeReplacement);
  } catch (e) {
    console.warn("replaceText error for '" + searchText + "': " + e.message);
  }
}

/**
 * Formats any date input (Date object, timestamp, or date string) into "MMM dd, yyyy" format (e.g. Aug 25, 2026).
 */
function formatDateValue(val) {
  if (!val) return "TBD";
  let d = null;
  if (val instanceof Date) {
    d = val;
  } else if (typeof val === "string" || typeof val === "number") {
    const str = val.toString().trim();
    if (!str || str.toLowerCase() === "tbd") return "TBD";
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      d = parsed;
    }
  }
  
  if (d && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM dd, yyyy");
  }
  return val.toString();
}

/**
 * Creates individual document copy from master template and replaces placeholders with row data.
 * Formats all dates as "MMM dd, yyyy" and eliminates all brackets.
 */
function createDoc(templateId, type, eventName, rowData, folder, isValidDate, eDate, headers) {
  try {
    const copy = DriveApp.getFileById(templateId).makeCopy(`${type} - ${eventName}`, folder);
    const doc = DocumentApp.openById(copy.getId());
    const elems = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(e => e != null);

    // Build key-value map of header to formatted row value
    const dataMap = {};
    headers.forEach((h, i) => {
      let val = rowData[i];
      if (h.toUpperCase().includes("TIME")) {
        val = extractTimeOnly(val);
      } else if (h.toUpperCase().includes("DATE") || val instanceof Date) {
        val = formatDateValue(val);
      } else if (val === null || val === undefined) {
        val = "";
      } else {
        val = val.toString();
      }
      dataMap[h] = val;
    });

    const formattedToday = formatDateValue(new Date());

    // Dictionary of all tag aliases mapping to values or headers
    const tagMap = {
      // Dates
      "Today's Date": formattedToday,
      "Submission Date": formatDateValue(dataMap["Submission Date"]) || formattedToday,

      // Contact & Coordinator
      "Your Name": dataMap["Your Name"] || "",
      "Your Name 2": dataMap["Your Name"] || "",
      "Client Name": dataMap["Your Name"] || "",
      "Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Who is the COORDINATOR and or decision maker for this event? Name and Title:": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Event Planner / Coordinator": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Event Planner/Coordinator": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Event Planner": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Coordinator": dataMap["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title"] || dataMap["Your Name"] || "",
      "Email Address": dataMap["Email Address"] || "",
      "Best Contact Phone Number": dataMap["Best Contact Phone Number"] || "",
      "Who do you represent? (Organization / Business / Self)": dataMap["Who do you represent? (Organization / Business / Self)"] || "",
      "Organization Name": dataMap["Who do you represent? (Organization / Business / Self)"] || "",

      // Event Details
      "What is the NAME of the event?": dataMap["What is the NAME of the event?"] || eventName,
      "Event Name": dataMap["What is the NAME of the event?"] || eventName,
      "NAME of the event": dataMap["What is the NAME of the event?"] || eventName,
      "Event": dataMap["What is the NAME of the event?"] || eventName,
      "Purpose of this Event": dataMap["Purpose of this Event"] || "",
      "Websites for Event & Organization": dataMap["Websites for Event & Organization"] || "",
      "Website 2": dataMap["Websites for Event & Organization"] || "",
      "Describe Your Event": dataMap["Describe Your Event"] || "",
      "Please confirm the DATE of your event:": formatDateValue(dataMap["Please confirm the DATE of your event:"]) || (isValidDate ? formatDateValue(eDate) : "TBD"),
      "Event Date": formatDateValue(dataMap["Please confirm the DATE of your event:"]) || (isValidDate ? formatDateValue(eDate) : "TBD"),
      "Please confirm the TIME of your event:": dataMap["Please confirm the TIME of your event:"] || "TBD",
      "Event Time": dataMap["Please confirm the TIME of your event:"] || "TBD",

      // Attendance & Audience
      "Expected Number of Attendees": dataMap["Expected Number of Attendees"] || "",
      "How many people are you expecting will ATTEND?": dataMap["Expected Number of Attendees"] || "",
      "How many people are you expecting will ATTENDING?": dataMap["Expected Number of Attendees"] || "",
      "Expected attendance": dataMap["Expected Number of Attendees"] || "",
      "Audience Age Groups Expected": dataMap["Audience Age Groups Expected"] || "",
      "Who is your AUDIENCE:": dataMap["Audience Age Groups Expected"] || "",
      "Audience": dataMap["Audience Age Groups Expected"] || "",

      // Admission
      "Type of Event Admission": dataMap["Type of Event Admission"] || "",
      "What type of event ADMISSION is it?": dataMap["Type of Event Admission"] || "",
      "Type of Event": dataMap["Type of Event Admission"] || "",

      // Location
      "Where will the event take place? (ADDRESS)": dataMap["Where will the event take place? (ADDRESS)"] || "",
      "Event Address": dataMap["Where will the event take place? (ADDRESS)"] || "",
      "Address": dataMap["Where will the event take place? (ADDRESS)"] || "",

      // Classification & Recurrence
      "Select Event Classification:": dataMap["Select Event Classification:"] || "",
      "Will the PERFORMANCE SERVICES be...": dataMap["Will the PERFORMANCE SERVICES be..."] || "",

      // Promotions, Media & 501c
      "501(c) Non-Profit Name (If Applicable)": dataMap["501(c) Non-Profit Name (If Applicable)"] || "",
      "Can you provide a Tax Deductibility Letter?": dataMap["Can you provide a Tax Deductibility Letter?"] || "",
      "Provide a Booth/Exhibitor Space (10×10 Tent)?": dataMap["Provide a Booth/Exhibitor Space (10×10 Tent)?"] || "",
      "Will you provide a BOOTH/EXHIBITOR space (10 x 10 Tent) to promote our services?": dataMap["Provide a Booth/Exhibitor Space (10×10 Tent)?"] || "",
      "Will you provide a BOOTH/EXHIBITOR space (10x10 Tent) to promote our services?": dataMap["Provide a Booth/Exhibitor Space (10×10 Tent)?"] || "",
      "Include our logo on promo materials & social media?": dataMap["Include our logo on promo materials & social media?"] || "",
      "Allowed to help promote the event?": dataMap["Allowed to help promote the event?"] || "",
      "Are we allowed to HELP PROMOTE the event?": dataMap["Allowed to help promote the event?"] || "",
      "Provide copies of video footage and photos?": dataMap["Provide copies of video footage and photos?"] || "",
      "Who will be taking PICTURES AND VIDEO?": dataMap["Provide copies of video footage and photos?"] || "",
      "Can we get COPIES of video footage and pictures of our participation?": dataMap["Provide copies of video footage and photos?"] || "",
      "Weather / Contingency Plan": dataMap["Weather / Contingency Plan"] || "",

      // Private Gathering
      "Type of Private Gathering": dataMap["Type of Private Gathering"] || "",
      "Are performers invited to attend/stay for the event?": dataMap["Are performers invited to attend/stay for the event?"] || "",
      "Are we INVITED TO ATTEND the event?": dataMap["Are performers invited to attend/stay for the event?"] || "",

      // Repertoire & Services
      "Repertoire: Mexico / North American Dances": dataMap["Repertoire: Mexico / North American Dances"] || "",
      "Repertoire: Caribbean Dances (Cuba & Puerto Rico)": dataMap["Repertoire: Caribbean Dances (Cuba & Puerto Rico)"] || "",
      "Repertoire: Central American Dances (El Salvador)": dataMap["Repertoire: Central American Dances (El Salvador)"] || "",
      "Repertoire: South American Dances (Colombia & Argentina)": dataMap["Repertoire: South American Dances (Colombia & Argentina)"] || "",
      "Service Type Requested": dataMap["Service Type Requested"] || "",
      "Troupe Headcount / Ensemble Size": dataMap["Troupe Headcount / Ensemble Size"] || "",
      "Troupe Headcount": dataMap["Troupe Headcount / Ensemble Size"] || "",
      "Ensemble Size": dataMap["Troupe Headcount / Ensemble Size"] || "",
      "Repertoire: Theatrical, Parade & Live Singing Performances": dataMap["Repertoire: Theatrical, Parade & Live Singing Performances"] || "",
      "Which of our DANCE LESSON SERVICES will you need?": dataMap["Which of our DANCE LESSON SERVICES will you need?"] || "",
      "Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?": dataMap["Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?"] || "",
      "Expecting AUDIENCE PARTICIPATION?": dataMap["Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?"] || "",
      "How much TIME do you require from us?": dataMap["How much TIME do you require from us?"] || "",
      "Anticipated Performance Time": dataMap["How much TIME do you require from us?"] || "",
      "Are there OTHER ENTERTAINERS sharing the time with us?": dataMap["Additional Services Needed (MC, DJ, Lecture)"] || dataMap["Special Instructions, Song Requests or Notes"] || "N/A",
      "Which of our PERFORMANCE SERVICES will you need?": [
        dataMap["Repertoire: Mexico / North American Dances"],
        dataMap["Repertoire: Caribbean Dances (Cuba & Puerto Rico)"],
        dataMap["Repertoire: Central American Dances (El Salvador)"],
        dataMap["Repertoire: South American Dances (Colombia & Argentina)"],
        dataMap["Repertoire: Theatrical, Parade & Live Singing Performances"],
        dataMap["Any other PERFORMANCE SERVICES you wish, but are not listed above?"]
      ].filter(Boolean).join("; ") || "As specified in agreement",

      // Venue & Logistics
      "Venue Location Setting": dataMap["Venue Location Setting"] || "",
      "Where will it take PLACE?": dataMap["Venue Location Setting"] || "",
      "On what SURFACE will the performance or class take place?": dataMap["On what SURFACE will the performance or class take place?"] || "",
      "Size of Performance / Class Area": dataMap["Size of Performance / Class Area"] || "",
      "Size of the performance Area": dataMap["Size of Performance / Class Area"] || "",
      "Sound System Equipment": dataMap["Sound System Equipment"] || "",
      "About the SOUND SYSTEM": dataMap["Sound System Equipment"] || "",
      "ABOUT THE SOUND SYSTEM": dataMap["Sound System Equipment"] || "",
      "Will a BADGE or ID be required for performers?": dataMap["Will a BADGE or ID be required for performers?"] || "",
      "Will a BADGE or ID be issued to performers to access the performance area?": dataMap["Will a BADGE or ID be required for performers?"] || "",
      "WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)": dataMap["WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)"] || "",
      "WILL YOU PROVIDE the performers with:": dataMap["WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)"] || "",
      "Dressing Room / Costume Changing Instructions": dataMap["Dressing Room / Costume Changing Instructions"] || "",
      "Will we have a place to change COSTUMES if needed? If so, please provide instructions.": dataMap["Dressing Room / Costume Changing Instructions"] || "",
      "Will we have a designated place for PARKING? If so, please provide instructions": dataMap["Special Instructions, Song Requests or Notes"] || "See event instructions",

      // Budget & Notes
      "Confirm you have a BUDGET for our participation": dataMap["Confirm you have a BUDGET for our participation"] || "",
      "Confirmed Budget Amount for Performance / Workshop": dataMap["Confirmed Budget Amount for Performance / Workshop"] || "",
      "Special Instructions, Song Requests or Notes": dataMap["Special Instructions, Song Requests or Notes"] || "",

      // Digital E-Signature & Agreement Acceptance
      "Signature": "____________________________________ (Digital Acceptance on File)",
      "Client Signature": `${dataMap["Your Name"] || "Client Authorized Signatory"} (Digital Verification on File)`,
      "Signature Date": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy"),
      "Electronic Acceptance": "Verified & Electronically Acknowledged via Salsa Guy Richmond LLC 2026 Portal",
      "Agreement Status": "PENDING CONFIRMATION / DIGITALLY SUBMITTED",

      // Tracking
      "Event ID": dataMap["Event ID"] || ""
    };

    elems.forEach(e => {
      const text = e.getText();
      if (!text || text.trim() === "") return;

      // 1. Replace {{Placeholder}} tags
      const curlyMatches = text.match(/\{\{([^}]+)\}\}/g);
      if (curlyMatches) {
        const uniqueCurly = [...new Set(curlyMatches)];
        uniqueCurly.forEach(fullTag => {
          const innerTag = fullTag.substring(2, fullTag.length - 2).trim();
          let val = (tagMap[innerTag] !== undefined && tagMap[innerTag] !== "") ? tagMap[innerTag] : (dataMap[innerTag] !== undefined ? dataMap[innerTag] : null);
          if (val === null) {
            const foundH = headers.find(h => h.toLowerCase() === innerTag.toLowerCase());
            if (foundH) val = dataMap[foundH];
          }
          if (val instanceof Date) val = formatDateValue(val);
          safeReplaceText(e, fullTag, val !== null ? val : "");
        });
      }

      // 2. Replace <Placeholder> tags
      const angleMatches = text.match(/<([^>]+)>/g);
      if (angleMatches) {
        const uniqueAngle = [...new Set(angleMatches)];
        uniqueAngle.forEach(fullTag => {
          const innerTag = fullTag.substring(1, fullTag.length - 1).trim();
          let val = (tagMap[innerTag] !== undefined && tagMap[innerTag] !== "") ? tagMap[innerTag] : (dataMap[innerTag] !== undefined ? dataMap[innerTag] : null);
          if (val === null) {
            const foundH = headers.find(h => h.toLowerCase() === innerTag.toLowerCase());
            if (foundH) val = dataMap[foundH];
          }
          if (val instanceof Date) val = formatDateValue(val);
          safeReplaceText(e, fullTag, val !== null ? val : "");
        });
      }

      // 3. Replace exact headers without brackets if present in text
      headers.forEach(h => {
        if (h && h.length > 6 && text.includes(h)) {
          let val = dataMap[h] !== undefined ? dataMap[h] : "";
          if (val instanceof Date) val = formatDateValue(val);
          safeReplaceText(e, h, val);
        }
      });

      // Cleanup step: Strip any remaining unreplaced {{...}} or <...> brackets from document text
      const remainingCurly = e.getText().match(/\{\{([^}]+)\}\}/g);
      if (remainingCurly) {
        [...new Set(remainingCurly)].forEach(leftover => {
          safeReplaceText(e, leftover, "");
        });
      }
      const remainingAngle = e.getText().match(/<([^>]+)>/g);
      if (remainingAngle) {
        [...new Set(remainingAngle)].forEach(leftover => {
          safeReplaceText(e, leftover, "");
        });
      }
    });

    doc.saveAndClose();

    // Auto-Export companion PDF to the event's Google Drive folder
    try {
      if (folder) {
        const pdfBlob = copy.getAs("application/pdf");
        pdfBlob.setName(`${type} - ${eventName}.pdf`);
        folder.createFile(pdfBlob);
      }
    } catch (pdfErr) {
      console.warn("Companion PDF auto-export skipped for " + type + ": " + pdfErr.message);
    }

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

  // Auto-categorize selected repertoire items into regional buckets if unified array/string provided
  if (data.performanceServices) {
    const perfList = Array.isArray(data.performanceServices) 
      ? data.performanceServices 
      : String(data.performanceServices).split(",").map(s => s.trim());
    
    if (!data.repertoireMexico || data.repertoireMexico.length === 0) {
      data.repertoireMexico = perfList.filter(item => /Jalisco|Veracruz|Michoac|Viejitos|Huasteco|Chiapas|Sinaloa|Norte|Nayarit|Oaxaca|Mexico|Mexicana/i.test(item));
    }
    if (!data.repertoireCaribbean || data.repertoireCaribbean.length === 0) {
      data.repertoireCaribbean = perfList.filter(item => /Puerto Rico|Bomba|Plena|Salsa|Cuba|Son|Rueda|Dominican|Bachata|Merengue|Caribbean/i.test(item));
    }
    if (!data.repertoireCentralAmerica || data.repertoireCentralAmerica.length === 0) {
      data.repertoireCentralAmerica = perfList.filter(item => /Salvador|Carbonero|Cortadoras|Guatemala|Honduras|Nicaragua|Costa Rica|Panam|Central/i.test(item));
    }
    if (!data.repertoireSouthAmerica || data.repertoireSouthAmerica.length === 0) {
      data.repertoireSouthAmerica = perfList.filter(item => /Colombia|Cumbia|Mapal|Bullerengue|Argentina|Tango|Chacarera|Per|Venezuela|Joropo|South/i.test(item));
    }
    if (!data.repertoireTheatrical || data.repertoireTheatrical.length === 0) {
      data.repertoireTheatrical = perfList.filter(item => /Llorona|Muertos|Cantor|Singer|Ranchera|Mariachi|Bolero|Desfile|Parade|Carnival|Parranda|Navide/i.test(item));
    }
  }

  const fieldHeaderMap = [
    { key: "timestamp", aliases: ["Submission Date", "Timestamp", "Date Submitted", "Date", "Submission Timestamp"] },
    { key: "requestId", aliases: ["Request ID", "Request Tracking ID", "Tracking ID"] },
    { key: "clientName", aliases: ["Your Name", "Full Name", "Client Name", "Name"] },
    { key: "organizationName", aliases: ["Organization / Company Name", "Organization Name", "Organization", "Company"] },
    { key: "clientEmail", aliases: ["Email Address", "Email"] },
    { key: "clientPhone", aliases: ["Best Contact Phone Number", "Phone Number", "Phone", "Telephone"] },
    { key: "contactPreference", aliases: ["Preferred Contact Method", "Contact Preference"] },
    { key: "representType", aliases: ["Who do you represent? (Organization / Business / Self)", "Who do you represent?", "Represent Type", "Representing"] },
    { key: "hearAboutUs", aliases: ["How did you HEAR of us?", "How did you HEAR ABOUT US?", "How did you hear about us", "Referral Source"] },
    { key: "eventName", aliases: ["What is the NAME of the event?", "Event Name", "Name of Event"] },
    { key: "eventCoordinator", aliases: ["Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title", "Who is the COORDINATOR and or decision maker for this event? Name and Title:", "Event Planner / Coordinator", "Event Coordinator", "Coordinator Name", "Decision Maker"] },
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
    { key: "outOfStateLogistics", aliases: ["Out-of-State Travel & Logistics Arrangement", "Out-of-State Travel & Logistics Arrangement (Conditional: Out-of-State)", "Out-of-State Logistics", "Travel Logistics", "Lodging & Travel Logistics", "Out of State Logistics"] },
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
    { key: "serviceTypeRequested", aliases: ["Service Type Requested", "Service Category / Scope", "Service Category", "Primary Service", "Requested Service", "Service Selection"] },
    { key: "troupeHeadcount", aliases: ["Troupe Headcount / Ensemble Size", "Ensemble Size", "Performer Headcount", "Troupe Size", "Number of Performers", "Troupe Headcount"] },
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
    { key: "confirmedBudget", aliases: ["Confirmed Budget Amount for Performance / Workshop", "Confirmed Budget Amount", "Confirmed Budget", "Budget Amount", "budgetAmount"] },
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

  const timestampIdx = getHeaderIndex(headers, ["Submission Date", "Timestamp", "Date Submitted", "Date", "Submission Timestamp"]);
  if (timestampIdx > -1 && (!newRow[timestampIdx] || newRow[timestampIdx] === "")) {
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "America/New_York";
    newRow[timestampIdx] = Utilities.formatDate(new Date(), tz, "M/d/yyyy H:mm:ss");
  }

  sheet.appendRow(newRow);
  const rowNum = sheet.getLastRow();
  const finalRequestId = `BTG-REQ-${rowNum}-${Date.now().toString(36).toUpperCase()}`;

  const reqIdIdx = getHeaderIndex(headers, ["Request ID", "Request Tracking ID", "Tracking ID"]);
  if (reqIdIdx > -1) {
    sheet.getRange(rowNum, reqIdIdx + 1).setValue(finalRequestId);
  }

  const eventIdIdx = getHeaderIndex(headers, ["Event ID", "Calendar Event ID"]);
  if (eventIdIdx > -1) {
    const curVal = sheet.getRange(rowNum, eventIdIdx + 1).getValue();
    if (!curVal) {
      sheet.getRange(rowNum, eventIdIdx + 1).setValue(finalRequestId);
    }
  }

  const mainFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  let uploadedFileUrl = "";
  if (data.attachedFile && data.attachedFile.base64) {
    try {
      const bytes = Utilities.base64Decode(data.attachedFile.base64);
      const blob = Utilities.newBlob(bytes, data.attachedFile.mimeType || 'application/octet-stream', data.attachedFile.fileName || 'Event_Attachment');
      const dateStr = data.eventDate || "Undated";
      const eventFolder = getOrCreateEventFolder(mainFolder, dateStr, data.eventName || "Event");
      const createdFile = eventFolder.createFile(blob);
      try {
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingErr) {}
      uploadedFileUrl = createdFile.getUrl();
      const fileColIdx = getHeaderIndex(headers, ["Upload Event Document / Attachment", "Upload Event Document", "Attachment", "Uploaded File"]);
      if (fileColIdx > -1) {
        sheet.getRange(rowNum, fileColIdx + 1).setFormula(`=HYPERLINK("${uploadedFileUrl}", "View Uploaded File")`);
      }
    } catch (fileErr) {
      console.warn("File attachment upload error: ", fileErr);
    }
  }

  const updatedRowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (uploadedFileUrl) {
    const fileColIdx = getHeaderIndex(headers, ["Upload Event Document / Attachment", "Upload Event Document", "Attachment", "Uploaded File"]);
    if (fileColIdx > -1) updatedRowData[fileColIdx] = uploadedFileUrl;
  }
  if (reqIdIdx > -1) updatedRowData[reqIdIdx] = finalRequestId;
  if (eventIdIdx > -1 && (!updatedRowData[eventIdIdx] || updatedRowData[eventIdIdx] === "")) {
    updatedRowData[eventIdIdx] = finalRequestId;
  }

  // --- IMMEDIATE EMAIL NOTIFICATIONS (Guaranteed Dispatch) ---
  try {
    const clientNameIdx = getHeaderIndex(headers, ["Your Name", "Full Name", "Client Name", "Name"]);
    const clientEmailIdx = getHeaderIndex(headers, ["Email Address", "Email"]);
    const clientPhoneIdx = getHeaderIndex(headers, ["Best Contact Phone Number", "Phone Number", "Phone", "Telephone"]);
    const representTypeIdx = getHeaderIndex(headers, ["Who do you represent? (Organization / Business / Self)", "Who do you represent?", "Represent Type", "Representing"]);
    const dateIdx = getHeaderIndex(headers, ["Please confirm the DATE of your event:", "Event Date", "Date of event", "Date"]);
    const timeIdx = getHeaderIndex(headers, ["Please confirm the TIME of your event:", "Event Time", "Time of event", "Time"]);
    const addrIdx = getHeaderIndex(headers, ["Where will the event take place? (ADDRESS)", "Event Address", "Address", "Location"]);
    const budgetIdx = getHeaderIndex(headers, ["Confirmed Budget Amount for Performance / Workshop", "Confirmed Budget Amount", "Confirmed Budget", "Budget Amount"]);
    const serviceTypeIdx = getHeaderIndex(headers, ["Service Type Requested", "Service Category / Scope", "Service Category", "Primary Service"]);
    const troupeIdx = getHeaderIndex(headers, ["Troupe Headcount / Ensemble Size", "Troupe Headcount", "Ensemble Size"]);
    const durationIdx = getHeaderIndex(headers, ["How much TIME do you require from us?", "DURATION of Service Required", "Duration Required", "Duration"]);
    const notesIdx = getHeaderIndex(headers, ["Special Instructions, Song Requests or Notes", "SPECIAL REQUESTS or Song Preferences", "Notes", "Special Requests"]);

    const clientName = data.clientName || (clientNameIdx > -1 && updatedRowData[clientNameIdx]) || "Valued Client";
    const clientEmail = data.clientEmail || (clientEmailIdx > -1 && updatedRowData[clientEmailIdx]) || "";
    const clientPhone = data.clientPhone || (clientPhoneIdx > -1 && updatedRowData[clientPhoneIdx]) || "N/A";
    const representType = data.representType || (representTypeIdx > -1 && updatedRowData[representTypeIdx]) || "N/A";
    const budgetAmount = data.confirmedBudget || data.budgetAmount || (budgetIdx > -1 && updatedRowData[budgetIdx]) || "Pending Confirmation";
    const serviceType = data.serviceTypeRequested || (serviceTypeIdx > -1 && updatedRowData[serviceTypeIdx]) || "Dance Booking";
    const troupeHeadcount = data.troupeHeadcount || (troupeIdx > -1 && updatedRowData[troupeIdx]) || "Standard Ensemble";
    const durationRequired = data.durationRequired || (durationIdx > -1 && updatedRowData[durationIdx]) || "As specified";
    const notes = data.notes || (notesIdx > -1 && updatedRowData[notesIdx]) || "";
    const eventAddress = data.eventAddress || (addrIdx > -1 && updatedRowData[addrIdx]) || "TBD";
    const eventTimeStr = data.eventTime ? extractTimeOnly(data.eventTime) : ((timeIdx > -1 && updatedRowData[timeIdx]) ? extractTimeOnly(updatedRowData[timeIdx]) : "TBD");
    const eventDateStr = data.eventDate ? formatDateValue(data.eventDate) : ((dateIdx > -1 && updatedRowData[dateIdx]) ? formatDateValue(updatedRowData[dateIdx]) : "TBD");

    const directNotificationDetails = {
      rowNum: rowNum,
      eventId: finalRequestId,
      clientName: clientName,
      clientEmail: clientEmail,
      clientPhone: clientPhone,
      representType: representType,
      budgetAmount: budgetAmount,
      serviceTypeRequested: serviceType,
      troupeHeadcount: troupeHeadcount,
      performanceServices: Array.isArray(data.performanceServices) ? data.performanceServices.join(", ") : (data.performanceServices || ""),
      lessonServices: data.lessonServices || "",
      durationRequired: durationRequired,
      notes: notes,
      eventName: data.eventName || "Event Booking Request",
      eventDate: eventDateStr,
      eventTime: eventTimeStr,
      eventAddress: eventAddress,
      calendarEventUrl: "",
      eventFolderUrl: "",
      propUrl: "",
      contUrl: "",
      perfUrl: "",
      fileUrl: uploadedFileUrl || ""
    };

    // 1. Immediately send admin notification
    sendAdminSubmittalNotification(directNotificationDetails);

    // 2. Immediately send client receipt notification if requested
    const shouldSendReceipt = data.sendEmailReceipt !== undefined ? (data.sendEmailReceipt === "Yes" || data.sendEmailReceipt === true) : true;
    if (shouldSendReceipt && clientEmail) {
      sendClientReceiptNotification(directNotificationDetails);
    }
  } catch (emailErr) {
    console.error("Critical error in immediate submittal email dispatch: " + emailErr.message);
  }

  // --- BACKGROUND DRIVE / DOCS / CALENDAR PROCESSING ---
  try {
    processRow(sheet, rowNum, updatedRowData, mainFolder, headers, true, uploadedFileUrl, true);
  } catch (procErr) {
    console.warn("processRow secondary background processing warning: ", procErr);
  }

  return {
    status: "success",
    row: rowNum,
    eventId: finalRequestId
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
 * Helper function to normalize text strings for flexible header matching.
 * Strips punctuation (colons, slashes, question marks), replaces "and/or" with "and or",
 * and collapses whitespace.
 */
function cleanHeaderStr(str) {
  if (!str) return "";
  return str.toString()
    .toLowerCase()
    .replace(/and\/or/g, "and or")
    .replace(/[/\\?%*:|"<>(),._-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Helper function to find column index in headers array matching a list of aliases.
 * Strictly prioritizes exact raw header matching, then normalized matching (ignoring slashes/punctuation),
 * then specific non-generic substring matching.
 */
function getHeaderIndex(headers, aliases) {
  if (!headers || !aliases || !Array.isArray(aliases)) return -1;
  
  const rawHeaders = headers.map(h => (h ? h.toString().toLowerCase().trim() : ""));
  const cleanedHeaders = headers.map(h => cleanHeaderStr(h));
  
  // Pass 1: Raw Exact match across all aliases (highest priority)
  for (let alias of aliases) {
    if (!alias) continue;
    const rawAlias = alias.toLowerCase().trim();
    let idx = rawHeaders.indexOf(rawAlias);
    if (idx > -1) return idx;
  }

  // Pass 2: Cleaned Normalized Match (ignores slashes vs spaces, colons, question marks)
  for (let alias of aliases) {
    if (!alias) continue;
    const cleanedAlias = cleanHeaderStr(alias);
    if (!cleanedAlias) continue;
    let idx = cleanedHeaders.indexOf(cleanedAlias);
    if (idx > -1) return idx;
  }
  
  // Pass 3: Substring matching ONLY for explicit, non-generic aliases (length > 3 and not generic terms)
  const genericTerms = ["date", "time", "name", "phone", "type", "event", "address", "notes", "receipt", "company"];
  for (let alias of aliases) {
    if (!alias) continue;
    const cleanedAlias = cleanHeaderStr(alias);
    if (cleanedAlias.length > 3 && !genericTerms.includes(cleanedAlias)) {
      let idx = cleanedHeaders.findIndex(h => h && h.length > 3 && (h.includes(cleanedAlias) || cleanedAlias.includes(h)));
      if (idx > -1) return idx;
    }
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





/**
 * Creates a formatted Google Doc in Google Drive containing all Questionnaire Sections,
 * Questions, Fields, and Responsive Dropdown Options for Review.
 */
function createQuestionnaireReviewGoogleDoc() {
  const docTitle = "Salsa Guy Richmond LLC - 2026 Event Booking Questionnaire (Full Review)";
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Document Title Header
  const titlePara = body.appendParagraph(docTitle);
  titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
  titlePara.setForegroundColor("#DC2626");

  const subPara = body.appendParagraph("Complete Responsive Field & Dropdown Specification for Review");
  subPara.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  subPara.setForegroundColor("#475569");

  body.appendHorizontalRule();

  const questionnaireData = [
  {
    "num": "1",
    "title": "1. Location & Logistics",
    "fields": [
      {
        "type": "select",
        "name": "regionalLocation",
        "label": "Is this in Virginia? (If you are located outside of Virginia, please provide details on travel and logistics so we can accommodate your needs)",
        "required": true,
        "options": [
          "Yes (Within Virginia, USA)",
          "No, Out-of-State (Within USA, outside VA) ➔ (Triggers Travel Logistics)",
          "No, International (Outside USA) ➔ (Triggers International Logistics)"
        ]
      },
      {
        "type": "text",
        "name": "eventAddress",
        "label": "Event Location (The full address or venue name)",
        "required": true,
        "placeholder": "e.g. 123 Main St, Richmond, VA 23220 or Venue Name"
      },
      {
        "type": "textarea",
        "name": "outOfStateLogistics",
        "label": "Out-of-State Travel & Logistics Arrangement (Conditional: Out-of-State)",
        "required": true,
        "placeholder": "Describe flight/driving logistics, hotel accommodations, per diem, and arrival schedule..."
      },
      {
        "type": "text",
        "name": "intlVenueName",
        "label": "International: Specific Country, City, & Venue Name (Conditional: International)",
        "required": true,
        "placeholder": "e.g. San José, Costa Rica - Teatro Nacional"
      },
      {
        "type": "textarea",
        "name": "intlLogistics",
        "label": "International: Travel & Lodging Logistics (Flights, ground transport, hotel arrangements)",
        "required": true,
        "placeholder": "Describe flight arrangements, airport transfers, hotel accommodations..."
      },
      {
        "type": "select",
        "name": "intlVisaSupport",
        "label": "International: Visa & Legal Documentation Support",
        "required": true,
        "options": [
          "Yes - Full Legal & Visa Support Provided",
          "No - Assistance Needed",
          "Work Visa Handled",
          "P-3 / Cultural Performer Visa Support Needed",
          "Not Applicable"
        ]
      },
      {
        "type": "text",
        "name": "intlPaymentTerms",
        "label": "International: Preferred Currency & Payment Terms",
        "required": true,
        "placeholder": "e.g. USD / Wire Transfer / 50% Deposit"
      },
      {
        "type": "textarea",
        "name": "intlCustomsConsiderations",
        "label": "International: Costumes, Props & Customs Considerations (Wardrobe transport or local provisions)",
        "required": true,
        "placeholder": "Details regarding wardrobe customs clearance, instruments, sound props..."
      }
    ]
  },
  {
    "num": "2",
    "title": "2. Event Basics",
    "fields": [
      {
        "type": "text",
        "name": "eventName",
        "label": "Name of the Event (What should we call your celebration or gathering?)",
        "required": true,
        "placeholder": "e.g. Richmond Salsa Festival 2026"
      },
      {
        "type": "text",
        "name": "eventWebsites",
        "label": "Websites (If your event or organization has a website, Facebook page, or public profile)",
        "required": false,
        "placeholder": "Website URL, Facebook page, or N/A"
      },
      {
        "type": "date",
        "name": "eventDate",
        "label": "Event Date (When does the magic happen?)",
        "required": true,
        "placeholder": "YYYY-MM-DD"
      },
      {
        "type": "time",
        "name": "eventTime",
        "label": "Start Time (When does the magic happen?)",
        "required": true,
        "placeholder": "HH:MM AM/PM"
      },
      {
        "type": "textarea",
        "name": "eventDescription",
        "label": "Description (Tell us about the theme, vibe, and flow of your program)",
        "required": true,
        "placeholder": "Tell us about the theme, vibe, and flow of your program..."
      },
      {
        "type": "select",
        "name": "eventTypeScale",
        "label": "Private vs. Public Classification",
        "required": true,
        "options": [
          "Option A: Large / Public Event (Please be prepared to provide details on booth space, promotional help, and contingency plans)",
          "Option B: Small / Private Gathering (Let us know the type, e.g., Birthday, Wedding, Quinceañera)"
        ]
      },
      {
        "type": "select",
        "name": "serviceRecurrence",
        "label": "Large/Public Event: Will the PERFORMANCE SERVICES be... (Conditional: Option A)",
        "required": true,
        "options": [
          "Nonrecurring - One time",
          "Recurring",
          "Recurring Weekly",
          "Recurring Monthly",
          "Recurring Yearly"
        ]
      },
      {
        "type": "text",
        "name": "nonProfitName",
        "label": "Large/Public Event: 501(c) Non-Profit Name (If Applicable)",
        "required": false,
        "placeholder": "Organization Tax Name"
      },
      {
        "type": "select",
        "name": "taxLetter",
        "label": "Large/Public Event: Can you provide a Tax Deductibility Letter?",
        "required": false,
        "options": [
          "Yes",
          "No",
          "Not Applicable"
        ]
      },
      {
        "type": "select",
        "name": "boothSpace",
        "label": "Large/Public Event: Provide a Booth/Exhibitor Space (10×10 Tent)?",
        "required": true,
        "options": [
          "Yes",
          "No",
          "Not Applicable"
        ]
      },
      {
        "type": "select",
        "name": "promoInclude",
        "label": "Large/Public Event: Include our logo on promo materials & social media?",
        "required": true,
        "options": [
          "Yes",
          "No",
          "Not Applicable"
        ]
      },
      {
        "type": "select",
        "name": "allowHelpPromote",
        "label": "Large/Public Event: Allowed to help promote the event?",
        "required": true,
        "options": [
          "Yes",
          "No",
          "Not Applicable"
        ]
      },
      {
        "type": "select",
        "name": "videoRights",
        "label": "Large/Public Event: Provide copies of video footage and photos?",
        "required": true,
        "options": [
          "Yes",
          "No",
          "Not Applicable"
        ]
      },
      {
        "type": "select",
        "name": "invitedToAttend3a",
        "label": "Large/Public Event: Are performers invited to attend/stay for the event?",
        "required": true,
        "options": [
          "Yes",
          "No"
        ]
      },
      {
        "type": "textarea",
        "name": "contingencyPlan",
        "label": "Large/Public Event: Weather / Contingency Plan (Alternate Location, Rain Date, Time adjustments)",
        "required": true,
        "placeholder": "Describe backup indoor location, rain date, time adjustments..."
      },
      {
        "type": "select",
        "name": "privateGatheringType",
        "label": "Small/Private Event: Type of Private Gathering (Conditional: Option B)",
        "required": true,
        "options": [
          "Birthday Party",
          "Wedding / Anniversary",
          "Quinceañera / Sweet 15",
          "House Gathering",
          "Private Dance Lesson",
          "Other Private Event"
        ]
      },
      {
        "type": "select",
        "name": "invitedToAttend",
        "label": "Small/Private Event: Are performers invited to attend/stay for the event?",
        "required": true,
        "options": [
          "Yes",
          "No"
        ]
      },
      {
        "type": "select",
        "name": "admissionType",
        "label": "Admission Type (Is this a free, ticketed, or private/invitation-only event?)",
        "required": true,
        "options": [
          "Free / Open to Public",
          "Ticketed / Paid Admission",
          "Private Invitation Only"
        ]
      }
    ]
  },
  {
    "num": "3",
    "title": "3. What You Need From Us",
    "fields": [
      {
        "type": "select",
        "name": "serviceTypeRequested",
        "label": "Services Requested (Choose from Performance, Dance Instruction, or both)",
        "required": true,
        "options": [
          "Performance Only (Dance Show / Live Music / Characters)",
          "Dance Instruction / Workshop Only",
          "Both Performance & Dance Instruction"
        ]
      },
      {
        "type": "checkbox_group",
        "name": "perfServices",
        "label": "Repertoire: Performance Dance Repertoire (Conditional: Performance)",
        "options": [
          "🇲🇽 Mexico & North America: A Selection of Dances from the Disney Movie COCO (Mexico)",
          "🇲🇽 Mexico & North America: El Baile de los Viejitos (The Dance of The Old Men) (Mexico)",
          "🇲🇽 Mexico & North America: The Waltz of La Llorona (Weeping Woman) (Mexico)",
          "🇲🇽 Mexico & North America: Día de los Muertos - Day of the Dead Celebration (Mexico)",
          "🇨🇺 🇵🇷 Caribbean: Salsa Rueda (Cuba)",
          "🇨🇺 🇵🇷 Caribbean: Bomba (Puerto Rico)",
          "🇨🇺 🇵🇷 Caribbean: Plena (Puerto Rico)",
          "🇨🇺 🇵🇷 Caribbean: Seis (Puerto Rico)",
          "🇨🇺 🇵🇷 Caribbean: Danza (Puerto Rico)",
          "🇸🇻 Central America: Cumbia (El Salvador)",
          "🇸🇻 Central America: Las Comaleras (El Salvador)",
          "🇸🇻 Central America: Las Cortadoras (El Salvador) COMING SOON",
          "🇨🇴 🇦🇷 South America: Cumbia (Colombia)",
          "🇨🇴 🇦🇷 South America: Cumbia \"La Pollera Colorá\" (Colombia)",
          "🇨🇴 🇦🇷 South America: El Mapalé (Colombia)",
          "🇨🇴 🇦🇷 South America: Tango (Argentina)",
          "🎭 Theatrical & Singing: Salsa Performance Group (Ladies in Red)",
          "🎭 Theatrical & Singing: USA - Line Dance Show & Animation",
          "🎭 Theatrical & Singing: Singer - Alma Ranchera (Mariachi & Ranchera Classics)",
          "🎭 Theatrical & Singing: Los Cantores de Tradición (Puerto Rican Christmas Parranda)",
          "🎭 Theatrical & Singing: Folk Characters Meet & Greet (Vejigantes & Catrinas / Charros)"
        ]
      },
      {
        "type": "text",
        "name": "otherPerfServices",
        "label": "Any other PERFORMANCE SERVICES you wish, but are not listed above?",
        "required": false,
        "placeholder": "Describe any custom performance services..."
      },
      {
        "type": "select",
        "name": "lessonServicesSelect",
        "label": "Which of our DANCE LESSON SERVICES will you need? (Conditional: Instruction)",
        "required": false,
        "options": [
          "None",
          "Salsa Guy's Line Dance lesson Salsa, Merengue, Bachata and Chacha",
          "Puerto Rican Folk - Bomba y Plena (Live Music)",
          "Merengue - Dominican Republic",
          "Bachata - Dominican Republic",
          "Salsa - Puerto Rico",
          "Cuba - Rueda",
          "Choreographies for Wedding or Sweet Fifteen (Quinceañero)",
          "All Salsa, Merengue and Bachata",
          "for Team Building",
          "Virtual",
          "Wedding 1st Dance",
          "for Children"
        ]
      },
      {
        "type": "select",
        "name": "audienceParticipation",
        "label": "Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?",
        "required": false,
        "options": [
          "Yes",
          "No"
        ]
      },
      {
        "type": "select",
        "name": "serviceTimeSelect",
        "label": "How much TIME do you require from us?",
        "required": false,
        "options": [
          "30 minutes",
          "1 hour",
          "1.5 hours",
          "2 hours",
          "Half Day (3-4 hours)",
          "Full Day (5+ hours)"
        ]
      },
      {
        "type": "checkbox_group",
        "name": "generalFormats",
        "label": "General Formats",
        "options": [
          "Stage Performance",
          "Opening Act",
          "Festival Headliner",
          "Main Act / Showcase",
          "Background Performance"
        ]
      },
      {
        "type": "select",
        "name": "soundEquipment",
        "label": "Sound (Will you provide the PA system, or should we bring our own?)",
        "required": true,
        "options": [
          "Provided by Venue (Venue provides sound / PA system)",
          "Performer Must Provide PA (We bring our own sound equipment)"
        ]
      },
      {
        "type": "checkbox_group",
        "name": "otherServices",
        "label": "Additional Needs (Let us know if you need an MC, DJ services, or a cultural presentation)",
        "options": [
          "Master of Ceremonies (MC)",
          "DJ Services",
          "Cultural Presentation / Lecture"
        ]
      }
    ]
  },
  {
    "num": "4",
    "title": "4. Venue & Hospitality",
    "fields": [
      {
        "type": "select",
        "name": "venueSetting",
        "label": "Setting (Indoor / Outdoor)",
        "required": true,
        "options": [
          "Indoor",
          "Outdoor Covered",
          "Outdoor Open"
        ]
      },
      {
        "type": "select",
        "name": "performanceSurface",
        "label": "Surface (Wood floor / Concrete / etc.)",
        "required": true,
        "options": [
          "Wood Dance Floor",
          "Stage",
          "Concrete",
          "Carpet",
          "Grass"
        ]
      },
      {
        "type": "select",
        "name": "stageSizes",
        "label": "Size of Area (Approximate size of the performance/class area)",
        "required": true,
        "options": [
          "8 ft x 8 ft (64 sq ft): Best for a solo speaker, DJ booth, or solo acoustic performer.",
          "8 ft x 12 ft (96 sq ft): Fits a single speaker with a podium or small ceremony setup.",
          "12 ft x 16 ft (192 sq ft): Ideal for a basic wedding band or a 3- to 6-piece group.",
          "16 ft x 20 ft (320 sq ft): The most common size for a full 5- to 8-piece band.",
          "20 ft x 24 ft (480 sq ft): Accommodates high-energy bands with added horn sections.",
          "24 ft x 32 ft+ (768+ sq ft): Built for major concerts and large corporate productions."
        ]
      }
    ]
  },
  {
    "num": "5",
    "title": "5. Your Contact Details",
    "fields": [
      {
        "type": "text",
        "name": "clientName",
        "label": "Your Name (The primary person coordinating the booking)",
        "required": true,
        "placeholder": "e.g. Maria Santos"
      },
      {
        "type": "email",
        "name": "clientEmail",
        "label": "Email Address (For official confirmation and correspondence)",
        "required": true,
        "placeholder": "maria@example.com"
      },
      {
        "type": "tel",
        "name": "clientPhone",
        "label": "Best Contact Phone Number (So we can reach you quickly)",
        "required": true,
        "placeholder": "(804) 555-0199"
      },
      {
        "type": "select",
        "name": "representType",
        "label": "Who do you represent? (Organization / Business / Self)",
        "required": false,
        "options": [
          "Organization / Business / Self",
          "Dance Studio",
          "Educational Institution",
          "Event Planning Company",
          "Government",
          "I'm a Dancer Instructor",
          "Nonprofit Organization 501(C)",
          "Private Sector Business",
          "Religious Institution",
          "Restaurant",
          "Talent Agency",
          "Yourself"
        ]
      },
      {
        "type": "text",
        "name": "eventCoordinator",
        "label": "Coordinator / Decision Maker (Who is the main contact we should speak to regarding final plans?)",
        "required": true,
        "placeholder": "e.g. Maria Santos, Lead Coordinator / Decision Maker"
      },
      {
        "type": "select",
        "name": "hasBudget",
        "label": "Confirm you have a BUDGET for our participation",
        "required": true,
        "options": [
          "Yes - Budget Confirmed",
          "In Progress / Pending Approval",
          "No / Grant Funded"
        ]
      },
      {
        "type": "text",
        "name": "budgetAmount",
        "label": "Confirmed Budget Amount for Performance / Workshop",
        "required": true,
        "placeholder": "e.g. $1,500 USD"
      }
    ]
  },
  {
    "num": "6",
    "title": "6. Final Steps",
    "fields": [
      {
        "type": "number",
        "name": "expectedAttendance",
        "label": "Expected Attendance (A rough headcount helps us prepare properly)",
        "required": true,
        "placeholder": "Estimate"
      },
      {
        "type": "select",
        "name": "audienceAgesSelect",
        "label": "Audience Ages (Let us know if this is a family-friendly, youth, or adult-only event)",
        "required": true,
        "options": [
          "All Ages / Family-Friendly / Adults Only / Youth",
          "Every Age Group (All Ages Welcome)",
          "Toddlers (1-3 yrs)",
          "Preschoolers (3-5 yrs)",
          "School-age (6-12 yrs)",
          "Adolescents (13-19 yrs)",
          "Young Adults (20-39 yrs)",
          "Middle-aged (40-59 yrs)",
          "Seniors (60+ yrs)",
          "Adults Only (18+ / 21+)"
        ]
      },
      {
        "type": "select",
        "name": "hearAboutUs",
        "label": "How did you hear about us? (How did you discover Salsa Guy Richmond?)",
        "required": true,
        "options": [
          "Attended Past Salsa Guy Event",
          "Word of Mouth / Referral",
          "Search Engine (Google/Bing)",
          "Social Media (Instagram/Facebook)",
          "Community Flyer / Poster",
          "Radio / TV Feature",
          "Corporate/University Directory",
          "Youtube Ads"
        ]
      },
      {
        "type": "select",
        "name": "hospitalityProvided",
        "label": "Hospitality (Water, meals, or a green room for costume changes)",
        "required": true,
        "options": [
          "Water",
          "Hospitality",
          "Meal",
          "Green Room",
          "Full Hospitality (Water, Meal, Green Room)"
        ]
      },
      {
        "type": "select",
        "name": "badgeRequired",
        "label": "Badges/ID (Let us know if our team needs special identification to enter the venue)",
        "required": true,
        "options": [
          "Yes",
          "No"
        ]
      },
      {
        "type": "textarea",
        "name": "dressingRoomDetails",
        "label": "Dressing Room / Costume Changing Instructions",
        "required": true,
        "placeholder": "e.g., Private Green Room on 2nd Floor with secure storage"
      },
      {
        "type": "textarea",
        "name": "notes",
        "label": "Special Instructions (Feel free to include song requests, specific notes, or event guidelines)",
        "required": false,
        "placeholder": "Tell us any special requests, song choices, or event guidelines..."
      },
      {
        "type": "file",
        "name": "fileInput",
        "label": "Attachments (PDF, PNG, JPG - Max 10MB)",
        "required": false,
        "placeholder": "Click or Drag & Drop File Here to Attach to Drive"
      },
      {
        "type": "notice",
        "name": "txtGeneralInfo",
        "label": "Notice: Hiring Similar Performers Disclosure",
        "required": false
      },
      {
        "type": "checkbox",
        "name": "termsCheck",
        "label": "Agreement: I agree to the Terms of Service & Privacy Policy",
        "required": true
      },
      {
        "type": "checkbox",
        "name": "sendEmailReceipt",
        "label": "Send instant Request ID confirmation receipt to my email",
        "required": false
      }
    ]
  }
];

  questionnaireData.forEach(function(sec) {
    const secHeader = body.appendParagraph(sec.title);
    secHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    secHeader.setForegroundColor("#F59E0B");

    sec.fields.forEach(function(f, idx) {
      const reqText = f.required ? " [REQUIRED *]" : " [Optional]";
      const qPara = body.appendParagraph((idx + 1) + ". " + f.label + reqText);
      qPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      qPara.setForegroundColor("#0F172A");

      const typePara = body.appendParagraph("Field Type: " + f.type + " | ID: " + f.name);
      typePara.setItalic(true);
      typePara.setForegroundColor("#64748B");

      if (f.options && f.options.length > 0) {
        body.appendParagraph("Dropdown Option Selections:");
        secList = f.options.forEach(function(opt) {
          const item = body.appendListItem("☐  " + opt);
          item.setGlyphType(DocumentApp.GlyphType.BULLET);
        });
      }
      body.appendParagraph(""); // spacing
    });

    body.appendHorizontalRule();
  });

  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  if (CONFIG && CONFIG.FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
      file.moveTo(folder);
    } catch(e) {
      console.warn("Moved doc error: " + e.message);
    }
  }

  console.log("Created Google Doc URL: " + doc.getUrl());
  return doc.getUrl();
}

/**
 * Opens and fixes/updates the 3 Master Google Doc Templates:
 * 1. Proposal Template (1plCZvjBJijgJrGduzXrTMjbtDopMgCfB5Vo7MLWxspo)
 * 2. Contract Template (1BuEv7BF6wsHutvEWwVZOVb3m8J3zkYujKheti8X871g)
 * 3. Performance Info Template (1eyXzMdmYiV0CDmxvNZfNO3dRYQ83hvJHHy_eseWPqIE)
 * 
 * Replaces mismatched, buggy, and legacy placeholders with clean, standardized tags.
 */
function updateMasterTemplates() {
  const ui = SpreadsheetApp.getUi();
  let log = "📝 Master Doc Templates Update Log:\n\n";

  // 1. Update Proposal Template
  try {
    const docProp = DocumentApp.openById(CONFIG.TEMPLATES.PROPOSAL);
    const bodyProp = docProp.getBody();
    
    safeReplaceText(bodyProp, "{{Today's Date}}", "{{Submission Date}}");
    safeReplaceText(bodyProp, "{{Who is the COORDINATOR and or decision maker for this event? Name and Title:}}", "{{Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title}}");
    safeReplaceText(bodyProp, "{{Event Planner/Coordinator}}", "{{Who is the Event Planner/Coordinator and or decision maker for this event? Name and Title}}");
    safeReplaceText(bodyProp, "{{Event Name}}", "{{What is the NAME of the event?}}");
    safeReplaceText(bodyProp, "{{Event}}", "{{What is the NAME of the event?}}");
    safeReplaceText(bodyProp, "{{Website 2}}", "{{Websites for Event & Organization}}");
    safeReplaceText(bodyProp, "{{Your Name 2}}", "{{Your Name}}");
    safeReplaceText(bodyProp, "Hello Ms. {{", "Hello {{");
    
    docProp.saveAndClose();
    log += "✅ Proposal Master Doc updated successfully!\n";
  } catch (e) {
    log += "❌ Error updating Proposal Doc: " + e.message + "\n";
  }

  // 2. Update Contract Template
  try {
    const docCont = DocumentApp.openById(CONFIG.TEMPLATES.CONTRACT);
    const bodyCont = docCont.getBody();
    
    safeReplaceText(bodyCont, "{{Today's Date}}", "{{Submission Date}}");
    safeReplaceText(bodyCont, "{{Who is the COORDINATOR and or decision maker for this event? Name and Title:}}", "{{Your Name}}");
    safeReplaceText(bodyCont, "Hello Mr. {{Your Name 2}},", "Hello {{Your Name}},");
    safeReplaceText(bodyCont, "Hello Mr. {{Your Name}},", "Hello {{Your Name}},");
    safeReplaceText(bodyCont, "{{Your Name 2}}", "{{Your Name}}");
    
    docCont.saveAndClose();
    log += "✅ Contract Master Doc updated successfully!\n";
  } catch (e) {
    log += "❌ Error updating Contract Doc: " + e.message + "\n";
  }

  // 3. Update Performance Info Template
  try {
    const docPerf = DocumentApp.openById(CONFIG.TEMPLATES.PERF_INFO);
    const bodyPerf = docPerf.getBody();
    
    // Fix critical bugs (wrong tags in fields)
    safeReplaceText(bodyPerf, "Size of the performance Area: {{Will a BADGE or ID be issued to performers to access the performance area?}}", "Size of the performance Area: {{Size of Performance / Class Area}}");
    safeReplaceText(bodyPerf, "We can take pictures and videos:  {{Are we allowed to HELP PROMOTE the event?}}", "We can take pictures and videos:  {{Provide copies of video footage and photos?}}");
    safeReplaceText(bodyPerf, "We can take pictures and videos: {{Are we allowed to HELP PROMOTE the event?}}", "We can take pictures and videos: {{Provide copies of video footage and photos?}}");

    // Standardize legacy & alternative phrasing tags to exact sheet header tags
    safeReplaceText(bodyPerf, "{{What type of event ADMISSION is it?}}", "{{Type of Event Admission}}");
    safeReplaceText(bodyPerf, "{{How many people are you expecting will ATTEND?}}", "{{Expected Number of Attendees}}");
    safeReplaceText(bodyPerf, "{{Expecting AUDIENCE PARTICIPATION?}}", "{{Interactive (AUDIENCE PARTICIPATION / Mini-Lesson)?}}");
    safeReplaceText(bodyPerf, "{{Are we INVITED TO ATTEND the event?}}", "{{Are performers invited to attend/stay for the event?}}");
    safeReplaceText(bodyPerf, "{{WILL YOU PROVIDE the performers with:}}", "{{WILL YOU PROVIDE the performers with (Water, Hospitality, Meal, Green Room)}}");
    safeReplaceText(bodyPerf, "{{Who is your AUDIENCE:}}", "{{Audience Age Groups Expected}}");
    safeReplaceText(bodyPerf, "{{Event Address}}", "{{Where will the event take place? (ADDRESS)}}");
    safeReplaceText(bodyPerf, "{{Where will it take PLACE?}}", "{{Venue Location Setting}}");
    safeReplaceText(bodyPerf, "{{Who is the COORDINATOR and or decision maker for this event? Name and Title:}}", "{{Your Name}}");
    safeReplaceText(bodyPerf, "{{Will we have a place to change COSTUMES if needed? If so, please provide instructions.}}", "{{Dressing Room / Costume Changing Instructions}}");
    safeReplaceText(bodyPerf, "{{Will a BADGE or ID be issued to performers to access the performance area?}}", "{{Will a BADGE or ID be required for performers?}}");
    safeReplaceText(bodyPerf, "{{ABOUT THE SOUND SYSTEM}}", "{{Sound System Equipment}}");
    safeReplaceText(bodyPerf, "{{Will you provide a BOOTH/EXHIBITOR space (10 x 10 Tent) to promote our services?}}", "{{Provide a Booth/Exhibitor Space (10×10 Tent)?}}");
    safeReplaceText(bodyPerf, "{{Will you provide a BOOTH/EXHIBITOR space (10x10 Tent) to promote our services?}}", "{{Provide a Booth/Exhibitor Space (10×10 Tent)?}}");
    safeReplaceText(bodyPerf, "{{Are we allowed to HELP PROMOTE the event?}}", "{{Allowed to help promote the event?}}");
    safeReplaceText(bodyPerf, "{{Who will be taking PICTURES AND VIDEO?}}", "{{Provide copies of video footage and photos?}}");
    
    docPerf.saveAndClose();
    log += "✅ Performance Info Master Doc updated successfully!\n";
  } catch (e) {
    log += "❌ Error updating Performance Info Doc: " + e.message + "\n";
  }

  if (ui) ui.alert(log);
}

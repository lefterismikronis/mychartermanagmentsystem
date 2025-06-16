// --- Configuration Variables ---
// IMPORTANT: Customize these values to match your setup exactly.
// Ensure these sheet names EXACTLY match the names of your tabs in 'Yacht Charter Management System' spreadsheet.
const MASTER_LOG_SHEET_NAME = "Master Charter Log";
const SETTINGS_SHEET_NAME = "Settings"; // This sheet stores global settings like current charter ID
 
const CSB_RESPONSES_SHEET_NAME = "CS&B Responses"; 
const DAILY_LOG_RESPONSES_SHEET_NAME = "Daily Log Responses";
const PREDEPARTURE_RESPONSES_SHEET_NAME = "Pre-Departure Responses"; // CRITICAL FOR LOGGING DAY
const BEFORE_APPROACH_RESPONSES_SHEET_NAME = "Before Approach Responses";
const BEFORE_BED_RESPONSES_SHEET_NAME = "Before Bed Responses"; 
const END_CHARTER_RESPONSES_SHEET_NAME = "End of Charter Responses";
 
// Form IDs for generating Prefilled URLs - CONFIRM THESE ARE CORRECT FOR YOUR FORMS
const DAILY_LOG_FORM_ID = "1hFzj7PXlI_neBqM38Px7aa68_hLHwQ438LY2wyu4W6w";
const PRE_DEPARTURE_FORM_ID = "1WjJHLbQ9RmNNXmXEDvVNq1qJ-cSUpy7uAQ-TpWaGoDo";
const BEFORE_APPROACH_FORM_ID = "1oR0FsCGoan_F_ci9OvdkVL9I3H9wVW4fq1E7Zwa1PcA";
const BEFORE_BED_FORM_ID = "198wYxv-vAU1KVyfTRLn4cX36v-HXI4jpSuTenkybd2E";
const END_CHARTER_FORM_ID = "18WnC687NGMNAWqtkM5OmAxjNuEUlvkLDHTbxtS0ZzeM";
 
const ADMIN_EMAIL_FOR_NOTIFICATIONS = "mikronislefteris@gmail.com"; 
 
// Define mandatory daily forms for the "Previous Day Completion" gate (checked by Pre-Departure)
const MANDATORY_DAILY_FORMS_FOR_GATE = ["Daily Log", "Before Bed", "Before Approach"]; // Pre-Departure is implicitly checked as it's the gate itself
 
// --- Define the Cross-Midnight Logging Cutoff Hour ---
// Any form submitted between 00:00 (midnight) and this hour will be considered part of the *previous calendar day's* log.
// Submissions from this hour onwards belong to the *current calendar day's* log.
const DAILY_LOG_CUTOFF_HOUR = 3; // Example: 3 AM EEST (03:00)
 
// --- Universal Form Submit Handler ---
// This function acts as a router for all form submissions.
// All "On form submit" triggers will point to this single function.
function onFormSubmit(e) {
  Logger.log("--- onFormSubmit triggered ---");
 
  if (!e || !e.source || typeof e.source.getName !== 'function') {
    Logger.log("Error: Event 'e' is not from a valid spreadsheet context or e.source is missing getName() method. Exiting.");
    Logger.log("Event object keys: " + (e ? Object.keys(e).join(', ') : "null event"));
    return;
  }
 
  const sourceSpreadsheet = e.source; 
  const responseSheetName = e.range.getSheet().getName(); 
  Logger.log("Submission received on sheet: \"" + responseSheetName + "\"");
 
  switch (responseSheetName) {
    case CSB_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as CS&B Form submission (from sheet: " + responseSheetName + ").");
      onCSBSubmit(e);
      break;
    case DAILY_LOG_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as Daily Log Form submission (from sheet: " + responseSheetName + ").");
      onDailyLogSubmit(e);
      break;
    case PREDEPARTURE_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as Pre-Departure Checklist submission (from sheet: " + responseSheetName + ").");
      onChecklistSubmit(e); // This will handle the new "gate" logic for Predeparture
      break;
    case BEFORE_APPROACH_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as Before Approach Form submission (from sheet: " + responseSheetName + ").");
      onChecklistSubmit(e);
      break;
    case BEFORE_BED_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as Before Sleep Checklist Form submission (from sheet: " + responseSheetName + ").");
      onChecklistSubmit(e);
      break;
    case END_CHARTER_RESPONSES_SHEET_NAME:
      Logger.log("Router: Identified as End of Charter Form submission (from sheet: " + responseSheetName + ").");
      onEndCharterSubmit(e);
      break;
    default:
      Logger.log("Router: No specific handler for response sheet: \"" + responseSheetName + "\". No action taken.");
      break;
  }
}
 
// --- Core Function: Triggered on Charter Start & Briefings Form Submission (via onFormSubmit router) ---
function onCSBSubmit(e) {
  Logger.log("[onCSBSubmit] Function started.");
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const csbResponsesSheet = ss.getSheetByName(CSB_RESPONSES_SHEET_NAME); 
 
  if (!logSheet) { Logger.log("Error: Master Log Sheet not found!"); return; }
  if (!settingsSheet) { Logger.log("Error: Settings Sheet not found!"); return; }
  if (!csbResponsesSheet) { Logger.log("Error: CS&B Responses Sheet not found!"); return; }
 
  // --- CRITICAL CHECK: Is there an active charter? ---
  const currentActiveCharterID = settingsSheet.getRange("B4").getValue();
  if (currentActiveCharterID) {
    const errorMessage = "ERROR: An active charter (" + currentActiveCharterID + ") is already in progress. New charter submission blocked.";
    Logger.log("[onCSBSubmit] " + errorMessage);
 
    // Mark as "INVALID" in CS&B Responses sheet for this submission
    const responseRow = e.range.getRow(); 
    const headers = csbResponsesSheet.getRange(1, 1, 1, csbResponsesSheet.getLastColumn()).getValues()[0];
    const charterIdColIndexInResponses = headers.indexOf("Charter ID"); 
 
    if (charterIdColIndexInResponses !== -1) {
        csbResponsesSheet.getRange(responseRow, charterIdColIndexInResponses + 1).setValue("INVALID");
        Logger.log("[onCSBSubmit] Marked Charter ID in CS&B Responses row " + responseRow + " as \"INVALID\".");
    } else {
        Logger.log("[onCSBSubmit] Warning: 'Charter ID' column not found in CS&B Responses sheet. Cannot mark as \"INVALID\".");
    }
 
    // Send email notification to ADMIN_EMAIL_FOR_NOTIFICATIONS
    const emailSubject = "Blocked New Charter Submission - Action Required";
    const emailBody = "A new \"Charter Start & Briefings\" form submission was just blocked.\n\n" +
                      "Reason: Charter ID \"" + currentActiveCharterID + "\" is currently marked as active in the system.\n\n" +
                      "Please ensure the active charter is completed via the \"End of Charter Form\" before attempting to start a new one.\n\n" +
                      "Submission details (from form): Timestamp: " + e.namedValues["Timestamp"][0] + ", Boat: " + e.namedValues["Boat Name"][0] + ", From: " + e.namedValues["Starting Port"][0] + ".\n\n" +
                      "You can review details in the CS&B Responses sheet.";
    MailApp.sendEmail(ADMIN_EMAIL_FOR_NOTIFICATIONS, emailSubject, emailBody);
    Logger.log("Sent email alert to " + ADMIN_EMAIL_FOR_NOTIFICATIONS + " regarding blocked submission.");
 
    return; // Stop the function execution
  }
  Logger.log("[onCSBSubmit] No active charter detected. Proceeding with new charter creation.");
 
  // Extract data from the CS&B Form submission
  const csbData = e.namedValues;
  const submissionTimestamp = new Date(csbData["Timestamp"][0]); 
 
  // --- Read the selected Charter Type from the dropdown (first question) ---
  const selectedCharterType = csbData["Charter ID"] ? csbData["Charter ID"][0] : ""; // This now holds "Charter ID" or "Special ID"
 
  const boatModel = csbData["Boat Model"] ? csbData["Boat Model"][0] : "";
  const boatName = csbData["Boat Name"] ? csbData["Boat Name"][0] : "";
  const guests = csbData["Charter Guests Number"] ? csbData["Charter Guests Number"][0] : "";
  const startingBase = csbData["Starting Port"] ? csbData["Starting Port"][0] : "";
 
  let generatedCharterID = ""; // Variable to hold the generated ID
 
  // --- Call the appropriate ID generator based on user selection ---
  if (selectedCharterType === "Charter ID") { // User selected "Charter ID" for regular
      Logger.log("[onCSBSubmit] Generating Regular Charter ID.");
      generatedCharterID = generateCharterID(settingsSheet);
  } else if (selectedCharterType === "Special ID") { // User selected "Special ID"
      Logger.log("[onCSBSubmit] Generating Special ID.");
      generatedCharterID = generateSpecialID(settingsSheet);
  } else {
      Logger.log("[onCSBSubmit] Error: Unrecognized Charter ID type selected: \"" + selectedCharterType + "\". Exiting.");
      return; // Stop if selection is invalid
  }
  Logger.log("[onCSBSubmit] Generated ID: " + generatedCharterID + ". Type selected: " + selectedCharterType);
 
  // 1. Fill Generated ID in CS&B Responses Sheet (Column B, where "Charter ID" dropdown was) ---
  const responseRow = e.range.getRow(); 
  const headers = csbResponsesSheet.getRange(1, 1, 1, csbResponsesSheet.getLastColumn()).getValues()[0];
  const charterIdColIndexInResponses = headers.indexOf("Charter ID"); 
 
  if (charterIdColIndexInResponses !== -1) {
      csbResponsesSheet.getRange(responseRow, charterIdColIndexInResponses + 1).setValue(generatedCharterID); // Write generated ID back
      Logger.log("[onCSBSubmit] Generated ID " + generatedCharterID + " written back to CS&B Responses sheet in row " + responseRow + ".");
  } else {
      Logger.log("[onCSBSubmit] Warning: 'Charter ID' column not found in CS&B Responses sheet. Cannot write back ID.");
  }
 
  // --- Insert new charter block into Master Log with precise spacing and coloring ---
 
  let currentInsertionRow = logSheet.getLastRow(); // This is the 1-based index of the last row before we start inserting.
  Logger.log("[onCSBSubmit] Separator block start. Initial currentInsertionRow: " + currentInsertionRow);
 
  // 1. Insert first blank row (after existing content/headers)
  logSheet.insertRowAfter(currentInsertionRow); 
  currentInsertionRow++; // Now points to the newly inserted blank row (e.g., Row 2)
  logSheet.getRange(currentInsertionRow, 1, 1, logSheet.getLastColumn()).clearFormat(); // Explicitly clear format for this blank row
  Logger.log("[onCSBSubmit] Inserted first blank row. Row: " + currentInsertionRow + ". Format cleared.");
 
  // 2. Insert separator row
  logSheet.insertRowAfter(currentInsertionRow); // Insert AFTER the blank row we just added
  currentInsertionRow++; // Now points to the newly inserted separator row (e.g., Row 3)
  const separatorRowIndex = currentInsertionRow; // Store the exact row index of the separator
  Logger.log("[onCSBSubmit] Inserted separator row. Row: " + currentInsertionRow + ". Separator is at row: " + separatorRowIndex + ".");
 
  let separatorRowData = Array(logSheet.getLastColumn()).fill(""); 
  separatorRowData[0] = "–––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––––"; 
  logSheet.getRange(separatorRowIndex, 1, 1, logSheet.getLastColumn()).setValues([separatorRowData]); 
  logSheet.getRange(separatorRowIndex, 1, 1, logSheet.getLastColumn()).setBackground("#cfe2f3"); // Apply blue background only to the separator row
  Logger.log("[onCSBSubmit] Separator text and color applied to row: " + separatorRowIndex + ". Color: #cfe2f3.");
 
  // 3. Insert second blank row (after separator)
  logSheet.insertRowAfter(currentInsertionRow); 
  currentInsertionRow++; // Now points to the newly inserted second blank row (e.g., Row 4)
  logSheet.getRange(currentInsertionRow, 1, 1, logSheet.getLastColumn()).clearFormat(); // Explicitly clear format for this blank row
  Logger.log("[onCSBSubmit] Inserted second blank row. Row: " + currentInsertionRow + ". Format cleared.");
 
  // 4. Insert Charter Header Row (e.g., Row 5)
  logSheet.insertRowAfter(currentInsertionRow);
  currentInsertionRow++; // Now points to the newly inserted charter header row
  const charterHeaderRowValues = [
    submissionTimestamp.toLocaleDateString('el-GR'), 
    generatedCharterID, // Use the generated ID here
    boatModel,
    boatName,
    guests,
    startingBase,
    "", "", "", "", "", "", "", "", "" // Remaining columns empty
  ];
  logSheet.getRange(currentInsertionRow, 1, 1, charterHeaderRowValues.length).setValues([charterHeaderRowValues]); 
 
  // Apply formatting to the header row (bold entire range, green ONLY on Charter ID cell)
  logSheet.getRange(currentInsertionRow, 1, 1, logSheet.getLastColumn()).clearFormat(); // CLEAR WHOLE ROW'S BACKGROUND BEFORE SETTING BOLD/COLOR
  logSheet.getRange(currentInsertionRow, 1, 1, charterHeaderRowValues.length).setFontWeight("bold"); // Apply bold to the filled portion
  logSheet.getRange(currentInsertionRow, 2).setBackground("#d9ead3"); // Apply green color to Charter ID cell (Col B)
  Logger.log("[onCSBSubmit] Appended charter header. Row: " + currentInsertionRow + ". Charter ID color: #d9ead3.");
 
  // Store the current active Charter ID in the Settings sheet
  settingsSheet.getRange("B4").setValue(generatedCharterID); 
  Logger.log("[onCSBSubmit] currentActiveCharterID (B4) set to: " + generatedCharterID);
  Logger.log("New Charter Created: " + generatedCharterID + " for Boat: " + boatName);
 
  // --- Generate Prefilled URLs for all forms and send in email ---
  // List of all forms to generate prefilled links for the email
  const allPrefillFormsForEmail = [
    { name: "Pre-Departure Checklist", id: PRE_DEPARTURE_FORM_ID, question: "Charter ID" },
    { name: "Daily Log Form", id: DAILY_LOG_FORM_ID, question: "Charter ID" },
    { name: "Before Approach Form", id: BEFORE_APPROACH_FORM_ID, question: "Charter ID" },
    { name: "Before Sleep Checklist Form", id: BEFORE_BED_FORM_ID, question: "Charter ID" },
    { name: "End of Charter Form", id: END_CHARTER_FORM_ID, question: "Charter ID" }
  ];
 
  let prefilledLinksBody = "Please use the following links to access the forms for Charter " + generatedCharterID + ":\n\n";
  allPrefillFormsForEmail.forEach(formInfo => {
      const prefillData = { [formInfo.question]: generatedCharterID };
      const link = getPrefilledUrl(formInfo.id, prefillData);
      prefilledLinksBody += formInfo.name + ": " + link + "\n";
      Logger.log("Generated Link for " + formInfo.name + ": " + link);
  });
 
  // --- Send confirmation email with ALL links ---
  const emailSubject = "✅ New Charter Started: " + generatedCharterID + " (" + boatName + ")";
  const emailBody = "A new charter has been successfully started:\n\n" +
                    "Charter ID: " + generatedCharterID + "\n" +
                    "Boat: " + boatName + "\n" +
                    "Guests: " + guests + "\n" +
                    "Starting Base: " + startingBase + "\n" +
                    "Start Date: " + submissionTimestamp.toLocaleDateString('el-GR') + "\n\n" +
                    prefilledLinksBody + // Includes all links generated above
                    "\nYou can view the Master Charter Log here: " + ss.getUrl() + "\n\n" +
                    "Thank you.";
  MailApp.sendEmail(ADMIN_EMAIL_FOR_NOTIFICATIONS, emailSubject, emailBody);
  Logger.log("Sent confirmation email with all prefilled links to " + ADMIN_EMAIL_FOR_NOTIFICATIONS);
}
 
 
// --- Helper Function: Generate a unique Charter ID (e.g., S25C1) ---
function generateCharterID(settingsSheet) {
  Logger.log("[generateCharterID] Function started.");
  const currentYear = new Date().getFullYear().toString().slice(-2); // This will be a STRING like "25"
 
  let lastCharterNumber = parseInt(settingsSheet.getRange("B2").getValue() || "0"); // B2 is 'lastCharterNumber'
  let lastSeason = String(settingsSheet.getRange("B3").getValue() || ""); // B3 is 'lastSeason'
 
  Logger.log("[generateCharterID] READ from Settings: lastCharterNumber=" + lastCharterNumber + " (Type: " + (typeof lastCharterNumber) + "), lastSeason=\"" + lastSeason + "\" (Type: " + (typeof lastSeason) + ")");
 
  if (lastSeason !== currentYear) { 
    Logger.log(" [generateCharterID] Mismatch detected: lastSeason (\"" + lastSeason + "\") !== currentYear (\"" + currentYear + "\"). Resetting charter count for new season.");
    lastCharterNumber = 0; 
    lastSeason = currentYear; 
    Logger.log(" [generateCharterID] Reset: lastCharterNumber set to 0, lastSeason set to \"" + lastSeason + "\".");
  }
 
  lastCharterNumber++; 
 
  settingsSheet.getRange("B2").setValue(lastCharterNumber);
  settingsSheet.getRange("B3").setValue(lastSeason);
  Logger.log(" [generateCharterID] WROTE to Settings: B2 (lastCharterNumber)=" + settingsSheet.getRange("B2").getValue() + ", B3 (lastSeason)=" + settingsSheet.getRange("B3").getValue());
 
  const newCharterID = "S" + currentYear + "C" + lastCharterNumber;
  Logger.log("[generateCharterID] Generated new Charter ID: " + newCharterID);
  return newCharterID;
}
 
// --- Helper Function: Generate a unique Special ID (e.g., SPC1) ---
function generateSpecialID(settingsSheet) {
  Logger.log("[generateSpecialID] Function started.");
  const currentYear = new Date().getFullYear().toString().slice(-2); // e.g., '25'
 
  // Read special charter number and season from Settings sheet
  // B5 is 'lastSpecialCharterNumber', B6 is 'lastSpecialSeason' (from your new Settings table)
  let lastSpecialCharterNumber = parseInt(settingsSheet.getRange("B5").getValue() || "0"); 
  let lastSpecialSeason = String(settingsSheet.getRange("B6").getValue() || ""); 
 
  Logger.log("[generateSpecialID] READ from Settings: lastSpecialCharterNumber=" + lastSpecialCharterNumber + " (Type: " + (typeof lastSpecialCharterNumber) + "), lastSpecialSeason=\"" + lastSpecialSeason + "\" (Type: " + (typeof lastSpecialCharterNumber) + ")");
 
  // Check if it's a new season for special charters
  if (lastSpecialSeason !== currentYear) { 
    Logger.log(" [generateSpecialID] Mismatch detected: lastSpecialSeason (\"" + lastSpecialSeason + "\") !== currentYear (\"" + currentYear + "\"). Resetting special charter count for new season.");
    lastSpecialCharterNumber = 0; // Reset counter for new season
    lastSpecialSeason = currentYear; // Update season to current year (as a string)
    Logger.log(" [generateSpecialID] Reset: lastSpecialCharterNumber set to 0, lastSpecialSeason set to \"" + lastSpecialSeason + "\".");
  }
 
  lastSpecialCharterNumber++; // Increment the special charter number
 
  // Write the updated values back to Settings sheet
  settingsSheet.getRange("B5").setValue(lastSpecialCharterNumber);
  settingsSheet.getRange("B6").setValue(lastSpecialSeason);
  Logger.log(" [generateSpecialID] WROTE to Settings: B5 (lastSpecialCharterNumber)=" + settingsSheet.getRange("B5").getValue() + ", B6 (lastSpecialSeason)=" + settingsSheet.getRange("B6").getValue());
 
  const newSpecialID = "SPC" + lastSpecialCharterNumber; // Format as SPC1, SPC2, etc.
  Logger.log("[generateSpecialID] Generated new Special ID: " + newSpecialID);
  return newSpecialID;
}
 
// --- Helper Function: Generates a prefilled Google Form URL ---
// formId: The ID of the Google Form
// prefillData: An object where keys are form question titles and values are the prefilled text
function getPrefilledUrl(formId, prefillData) {
  Logger.log("Generating URL for form ID: " + formId + ", Data: " + JSON.stringify(prefillData)); 
  let form = FormApp.openById(formId);
  let prefilledResponse = form.createResponse(); 

  for (let questionTitle in prefillData) {
    if (prefillData.hasOwnProperty(questionTitle)) {
      let item = form.getItems().find(item => item.getTitle() === questionTitle); // Find item by title 
      if (!item) {
        Logger.log(`Warning: Question '${questionTitle}' not found in form ${formId}. Skipping.`);
        continue;
      }
      
      let value = prefillData[questionTitle];
      Logger.log(`Processing question '${questionTitle}', Type: ${item.getType()}, Value: '${value}'`);

      try {
        if (item.getType() === FormApp.ItemType.TEXT) {
          if (typeof value === 'string' && value.trim() !== '') {
            prefilledResponse.withItemResponse(item.asTextItem().createResponse(value));
            Logger.log(`Prefilled TEXT question '${questionTitle}' with value '${value}'`);
          } else {
            Logger.log(`Skipping TEXT question '${questionTitle}' due to invalid or empty value: '${value}'`);
          }
        } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE || item.getType() === FormApp.ItemType.LIST) {
          let mcItem = item.asMultipleChoiceItem();
          let choices = mcItem.getChoices().map(c => c.getValue());
          Logger.log(`Choices for '${questionTitle}': [${choices.join(", ")}]`);

          let choice = mcItem.getChoices().find(c => c.getValue() === value);
          if (choice) {
            prefilledResponse.withItemResponse(mcItem.createResponse(choice.getValue()));
            Logger.log(`Prefilled MULTIPLE_CHOICE/LIST question '${questionTitle}' with value '${value}'`);
          } else {
            Logger.log(`Warning: Choice '${value}' not found for question '${questionTitle}'. Skipping.`);
          }
        } else {
          Logger.log(`Warning: Unsupported question type '${item.getType()}' for question '${questionTitle}'. Skipping.`);
        }
      } catch (e) {
        Logger.log(`Exception when prefilling question '${questionTitle}': ${e.message}`);
      }
    }
  }

  try {
    const url = prefilledResponse.toPrefilledUrl();
    Logger.log("Generated URL: " + url); 
    return url;
  } catch(e) {
    Logger.log("Exception when generating prefilled URL: " + e.message);
    throw e;
  }
}

 
// --- Helper Function: Determines the correct 'logging day' for a submission ---
// This function will find the most recent Pre-Departure submission for the active charter
// and use its date as the logging date. If no Pre-Departure yet, uses submission date.
// This also handles cross-midnight submissions: forms before DAILY_LOG_CUTOFF_HOUR on next calendar day are backdated.
function getLoggingDate(charterID, submissionTimestamp) {
  Logger.log("[getLoggingDate] Determining logging date for " + charterID + " at " + submissionTimestamp.toLocaleString('el-GR') + " (Actual Submission Time).");
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const preDepartureResponsesSheet = ss.getSheetByName(PREDEPARTURE_RESPONSES_SHEET_NAME);
 
  // Check if Pre-Departure Responses sheet exists
  if (!preDepartureResponsesSheet) {
    Logger.log("[getLoggingDate] Warning: Pre-Departure Responses sheet not found. Falling back to submission timestamp's date.");
    // Apply cross-midnight adjustment to submission date as fallback
    let adjustedDate = new Date(submissionTimestamp.getTime());
    // Assuming DAILY_LOG_CUTOFF_HOUR is defined globally (e.g., 3 AM)
    if (adjustedDate.getHours() < DAILY_LOG_CUTOFF_HOUR) { // If submitted early morning (e.g., 1 AM, cutoff is 3 AM)
      adjustedDate.setDate(adjustedDate.getDate() - 1); // Subtract one day
      Logger.log("[getLoggingDate] (Fallback) Applied cross-midnight adjustment. Logging date adjusted to: " + adjustedDate.toLocaleDateString('el-GR') + ".");
    }
    return adjustedDate.toLocaleDateString('el-GR');
  }
 
  const pdData = preDepartureResponsesSheet.getDataRange().getValues();
  let mostRecentPredepartureTimestampForCharter = null; // Stores the Date object of the latest relevant Pre-Departure
 
  // Iterate through Pre-Departure submissions to find the most recent one for this charter,
  // that occurred on or before the current submission.
  for (let i = 1; i < pdData.length; i++) { // Skip headers
    const rowPdTimestamp = new Date(pdData[i][0]); // Timestamp of a Pre-Departure submission
    const rowPdCharterID = pdData[i][1]; // Charter ID in Pre-Departure Responses
 
    if (rowPdCharterID === charterID && rowPdTimestamp.getTime() <= submissionTimestamp.getTime()) {
      if (!mostRecentPredepartureTimestampForCharter || rowPdTimestamp.getTime() > mostRecentPredepartureTimestampForCharter.getTime()) {
        mostRecentPredepartureTimestampForCharter = rowPdTimestamp;
      }
    }
  }
 
  // If a Pre-Departure was found, that's our base for the logging day.
  if (mostRecentPredepartureTimestampForCharter) {
    Logger.log("[getLoggingDate] Found most recent Pre-Departure at: " + mostRecentPredepartureTimestampForCharter.toLocaleString('el-GR') + " for " + charterID + ". Using its date.");
    return mostRecentPredepartureTimestampForCharter.toLocaleDateString('el-GR');
  } else {
    // Fallback: If no Pre-Departure has been submitted yet for this charter,
    // or if the current submission's timestamp is BEFORE any Pre-Departure:
    // Use the submission's own date, but apply the cross-midnight cutoff logic.
    Logger.log("[getLoggingDate] No relevant Pre-Departure found for " + charterID + ". Applying cutoff to submission date.");
 
    let adjustedDate = new Date(submissionTimestamp.getTime()); // Copy the date
    if (adjustedDate.getHours() < DAILY_LOG_CUTOFF_HOUR) { // If submitted early morning (e.g., 1 AM, cutoff is 3 AM)
      adjustedDate.setDate(adjustedDate.getDate() - 1); // Subtract one day
      Logger.log("[getLoggingDate] (Fallback) Applied cross-midnight adjustment. Logging date adjusted to: " + adjustedDate.toLocaleDateString('el-GR') + ".");
    }
    return adjustedDate.toLocaleDateString('el-GR');
  }
}

function isFirstPredeparture(charterID) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Charter Log");
  const data = sheet.getDataRange().getValues();

  let inCurrentCharter = false;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Start of current charter
    if (row[0] === "CHARTER" && row[1] === charterID) {
      inCurrentCharter = true;
      continue;
    }

    // Stop when we reach next charter
    if (inCurrentCharter && row[0] === "CHARTER" && row[1] !== charterID) break;

    if (inCurrentCharter) {
      const predepartureStatus = row[10]; // Column K
      if (predepartureStatus === "✅") {
        return false; // Found existing predeparture
      }
    }
  }

  return true; // No Predeparture yet — first one
}

 
// --- Helper Function: Finds the correct daily row in the Master Log or creates it ---
// This function now returns the 0-based row index (for data array) of the target row.
function getOrCreateDailyRow(logSheet, charterID, dateStr) {
    Logger.log("[getOrCreateDailyRow] Searching for or creating daily row for " + charterID + " on logging date: " + dateStr);
    const data = logSheet.getDataRange().getValues(); // Get all data from Master Log
    let charterHeaderRowIndex = -1; // 0-indexed row of the charter header
    let lastRowInCharterBlock = -1; // 0-indexed row of the last row in the charter block before a separator/summary
 
    // 1. Find the charter's header row
    for (let i = 0; i < data.length; i++) {
        // Charter ID is in Col B (index 1) of header row, starts with S or SPC
        if (typeof data[i][1] === "string" && (data[i][1].startsWith("S") || data[i][1].startsWith("SPC")) && data[i][1].includes(charterID)) { 
            charterHeaderRowIndex = i; 
            lastRowInCharterBlock = i; // Initial assumption for last row in block if no daily entries exist yet
            break;
        }
    }
 
    if (charterHeaderRowIndex === -1) {
        Logger.log("[getOrCreateDailyRow] ERROR: Charter ID " + charterID + " header not found in Master Log. Cannot log daily data.");
        return null;
    }
    Logger.log("[getOrCreateDailyRow] Found charter header at index: " + charterHeaderRowIndex);
 
    // 2. Search for an existing daily row for this specific logging date and charter
    let foundDailyRowIndex = -1; // 0-indexed row of the found daily row
    // Iterate from immediately after the charter header (charterHeaderRowIndex + 1 is the first potential daily row)
    for (let i = charterHeaderRowIndex + 1; i < data.length; i++) { 
        const row = data[i];
        const rowDate = row[0] ? new Date(row[0]).toLocaleDateString('el-GR') : ""; 
        const rowCharterID = row[1]; 
 
        // Check if this row is an existing daily log for the current logging date and charter ID
        if (rowDate === dateStr && rowCharterID === charterID) {
            foundDailyRowIndex = i; 
            Logger.log("[getOrCreateDailyRow] Found existing daily row at index " + foundDailyRowIndex + " for " + charterID + " on logging date " + dateStr + ". Will UPDATE this row.");
            break;
        }
 
        // If we hit a separator or summary (for this or the next charter), stop searching within this block
        // This defines the end of the current charter's daily log block.
        if (typeof row[0] === "string" && (row[0].includes("–––––––––") || row[0].includes("Summary") || (row[1] && (row[1].startsWith("S") || row[1].startsWith("SPC"))))) {
            lastRowInCharterBlock = i - 1; // The row *before* the separator/summary is the true last data row
            break;
        }
        lastRowInCharterBlock = i; // Keep tracking the last data row in this block
    }
 
    // 3. If no existing daily row found, create a new one
    if (foundDailyRowIndex === -1) {
        // Determine insertion point:
        // If lastRowInCharterBlock is still -1 (meaning no daily rows yet, just header), insert right after header (charterHeaderRowIndex + 1).
        // Otherwise, insert after the last row in the charter block.
        let insertAtRow = (lastRowInCharterBlock !== -1) ? (lastRowInCharterBlock + 1) : (charterHeaderRowIndex + 1); 
 
        logSheet.insertRowAfter(insertAtRow); 
        foundDailyRowIndex = insertAtRow + 1; // The new row's 0-indexed position (actual row number is +1)
 
        // Initialize new daily row with date and Charter ID, and default '❌' for checklists
        const emptyDailyRowValues = Array(logSheet.getRange(1, logSheet.getLastColumn()).getValues()[0].length).fill(""); // Create array matching sheet width
        emptyDailyRowValues[0] = dateStr; // Col A: Date
        emptyDailyRowValues[1] = charterID; // Col B: Charter ID
        emptyDailyRowValues[10] = "❌"; // Col K: Predeparture (index 10)
        emptyDailyRowValues[11] = "❌"; // Col L: Daily Log (index 11)
        emptyDailyRowValues[12] = "❌"; // Col M: Before Approach (index 12)
        emptyDailyRowValues[13] = "❌"; // Col N: Before Bed (index 13)
 
        // Write the new row's data. Note: foundDailyRowIndex is 0-indexed, getRange is 1-based.
        logSheet.getRange(foundDailyRowIndex + 1, 1, 1, emptyDailyRowValues.length).setValues([emptyDailyRowValues]); 
        Logger.log("[getOrCreateDailyRow] Created NEW daily row at index " + foundDailyRowIndex + " for " + charterID + " on logging date " + dateStr + ". Will populate this new row.");
    }
 
    return foundDailyRowIndex; // Return the 0-indexed row number of the daily row (found or newly created)
}
 
 
// --- Helper Function: Reads all response sheets to get current checklist statuses ---
// This function needs to check against the logging date, not just submission date
function getChecklistStatusForDate(dateStr, charterID) {
  Logger.log("[getChecklistStatusForDate] Checking statuses for " + charterID + " on logging date: " + dateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
 
  const status = {
    "Predeparture": "❌",
    "Daily Log": "❌", 
    "Before Approach": "❌",
    "Before Bed": "❌"
  };
 
  const checklistSheetsMap = {
    "Predeparture": PREDEPARTURE_RESPONSES_SHEET_NAME,
    "Daily Log": DAILY_LOG_RESPONSES_SHEET_NAME,
    "Before Approach": BEFORE_APPROACH_RESPONSES_SHEET_NAME,
    "Before Bed": BEFORE_BED_RESPONSES_SHEET_NAME
  };
 
  for (let key in checklistSheetsMap) {
    const sheetName = checklistSheetsMap[key];
    const sheet = ss.getSheetByName(sheetName); 
    if (!sheet) {
      Logger.log(" [getChecklistStatusForDate] Warning: Response sheet '" + sheetName + "' not found.");
      continue; 
    }
    const data = sheet.getDataRange().getValues(); 
 
    // Assuming Timestamp is Column A (index 0) and Charter ID is Column B (index 1) in all response sheets
    for (let i = 1; i < data.length; i++) { // Start from 1 to skip headers
      const entryTimestamp = new Date(data[i][0]);
      const formCharterID = data[i][1]; 
 
      // Use getLoggingDate to determine if this entry belongs to the target logging date
      const entryLoggingDate = getLoggingDate(formCharterID, entryTimestamp); 
 
      if (entryLoggingDate === dateStr && formCharterID === charterID) {
        status[key] = "✅"; 
        Logger.log(" [getChecklistStatusForDate] Found " + key + " for " + charterID + " on logging date " + dateStr + ".");
        break; 
      }
    }
  }
  Logger.log("[getChecklistStatusForDate] Final statuses: " + JSON.stringify(status));
  return status;
}
 
 
// --- Helper Function: Gets boat name from Master Log for a given charter ID ---
function getBoatNameFromMasterLog(charterID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  if (!logSheet) {
      Logger.log("[getBoatNameFromMasterLog] Master Log Sheet not found for boat name lookup.");
      return "Unknown Boat";
  }
 
  const data = logSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    // Find charter header row (Col B starts with S or SPC and contains the ID)
    // Boat Name is Column D (index 3) in the header row
    if (typeof data[i][1] === "string" && (data[i][1].startsWith("S") || data[i][1].startsWith("SPC")) && data[i][1].includes(charterID)) {
      return data[i][3] || "Unknown Boat"; // Return boat name, or fallback
    }
  }
  return "Unknown Boat"; // If charter header not found
}
 
// --- Function: Triggered by Daily Log Form Submission (via onFormSubmit router) ---
function onDailyLogSubmit(e) {
  Logger.log("[onDailyLogSubmit] Function started.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
 
  if (!logSheet) { Logger.log("Error: Master Log Sheet not found!"); return; }
  if (!settingsSheet) { Logger.log("Error: Settings Sheet not found!"); return; }
 
  const formResponses = e.namedValues;
  const submittedCharterID = formResponses["Charter ID"] ? formResponses["Charter ID"][0] : "";
  const submissionTimestamp = new Date(formResponses["Timestamp"][0]);
  // --- Determine logging date ---
  const dateStr = getLoggingDate(submittedCharterID, submissionTimestamp); // Use the new logging date helper
 
  const currentActiveCharterID = settingsSheet.getRange("B4").getValue(); 
 
  if (!submittedCharterID || submittedCharterID !== currentActiveCharterID) {
    Logger.log("[onDailyLogSubmit] Submitted Charter ID '" + submittedCharterID + "' does not match active charter '" + currentActiveCharterID + "' or is empty. Submission ignored.");
    return;
  }
 
  // --- NEW VALIDATION: Daily Log cannot be submitted without a Pre-Departure for the same logging day ---
  const preDepartureStatusForToday = getChecklistStatusForDate(dateStr, submittedCharterID)["Predeparture"];
  if (preDepartureStatusForToday !== "✅") {
      const errorMessage = "Daily Log for Charter ID " + submittedCharterID + " on logging day " + dateStr + " blocked. Pre-Departure Checklist for this day is missing or incomplete.";
      Logger.log("[onDailyLogSubmit] ERROR: " + errorMessage);
 
      // Mark as "BLOCKED_NO_PREDEPARTURE" in Daily Log Responses sheet for this submission
      const dlResponsesSheet = ss.getSheetByName(DAILY_LOG_RESPONSES_SHEET_NAME); 
      const responseRow = e.range.getRow(); 
      const headers = dlResponsesSheet.getRange(1, 1, 1, dlResponsesSheet.getLastColumn()).getValues()[0];
      const charterIdColIndexInResponses = headers.indexOf("Charter ID"); 
 
      if (charterIdColIndexInResponses !== -1) {
          dlResponsesSheet.getRange(responseRow, charterIdColIndexInResponses + 1).setValue("BLOCKED_NO_PREDEPARTURE");
          Logger.log("[onDailyLogSubmit] Marked Charter ID in Daily Log Responses row " + responseRow + " as \"BLOCKED_NO_PREDEPARTURE\".");
      } else {
          Logger.log("[onDailyLogSubmit] Warning: 'Charter ID' column not found in Daily Log Responses sheet. Cannot mark as \"BLOCKED_NO_PREDEPARTURE\".");
      }
 
      // Send email notification to ADMIN_EMAIL_FOR_NOTIFICATIONS
      const boatNameForEmail = getBoatNameFromMasterLog(submittedCharterID); 
      const emailSubject = "⛔ Action Required: Daily Log Blocked - No Pre-Departure";
      const emailBody = "A Daily Log form submission for Charter ID " + submittedCharterID + " (Boat: " + boatNameForEmail + ") on logging day " + dateStr + " was blocked.\n\n" +
                        "Reason: The Pre-Departure Checklist for this day is missing or incomplete.\n\n" +
                        "Please ensure the Pre-Departure Checklist is submitted for this day before attempting to log daily activities.\n\n" +
                        "Submission details (from form): Timestamp: " + e.namedValues["Timestamp"][0] + ", From: " + (formResponses["Start Port/ Marina"] ? formResponses["Start Port/ Marina"][0] : "N/A") + ".\n\n" +
                        "You can review details in the Daily Log Responses sheet.\n\n" +
                        "Master Log: " + ss.getUrl(); 
      MailApp.sendEmail(ADMIN_EMAIL_FOR_NOTIFICATIONS, emailSubject, emailBody);
      Logger.log("[onDailyLogSubmit] Sent email alert regarding blocked Daily Log submission.");
 
      return; // Stop the function execution
  }
  Logger.log("[onDailyLogSubmit] Pre-Departure check passed for Daily Log. Proceeding.");
 
  const dailyRowIndex = getOrCreateDailyRow(logSheet, submittedCharterID, dateStr); 
  if (dailyRowIndex === null) {
      Logger.log("[onDailyLogSubmit] Could not find or create daily row for Daily Log submission. Exiting.");
      return;
  }
 
  const miles = parseFloat(formResponses["Distance in NM"] ? formResponses["Distance in NM"][0] : 0);
  const startPort = formResponses["Start Port/ Marina"] ? formResponses["Start Port/ Marina"][0] : "";
  const endPort = formResponses["End Port/Marina"] ? formResponses["End Port/Marina"][0] : "";
 
  const departureTimeStr = formResponses["Departure Time"] ? formResponses["Departure Time"][0] : "00:00";
  const arrivalTimeStr = formResponses["Arrival Time"] ? formResponses["Arrival Time"][0] : "00:00";
  let timeUnderCommand = "";
  try {
    const preDepartureResponsesSheet = ss.getSheetByName(PREDEPARTURE_RESPONSES_SHEET_NAME);
    let preDepartureTimestamp = null;
    if (preDepartureResponsesSheet) {
        const pdData = preDepartureResponsesSheet.getDataRange().getValues();
        for (let i = 1; i < pdData.length; i++) {
            const pdEntryTimestamp = new Date(pdData[i][0]);
            const pdCharterID = pdData[i][1];
            if (pdCharterID === submittedCharterID && getLoggingDate(pdCharterID, pdEntryTimestamp) === dateStr) {
                preDepartureTimestamp = pdEntryTimestamp;
                break; 
            }
        }
    } else {
        Logger.log("[onDailyLogSubmit] Warning: Pre-Departure Responses sheet not found for TUC calculation.");
    }
 
    if (preDepartureTimestamp && !isNaN(preDepartureTimestamp.getTime())) {
        const arrivalDateTime = new Date(dateStr + " " + arrivalTimeStr); 
        if (arrivalDateTime && !isNaN(arrivalDateTime.getTime())) {
            let diffMs = arrivalDateTime.getTime() - preDepartureTimestamp.getTime();
            timeUnderCommand = (diffMs / (1000 * 60 * 60)).toFixed(2);
        } else {
            Logger.log("[onDailyLogSubmit] Warning: Daily Log Arrival Time is invalid.");
            timeUnderCommand = "N/A";
        }
    } else {
        Logger.log("[onDailyLogSubmit] Warning: No Pre-Departure found for TUC calculation for " + submittedCharterID + " on logging date " + dateStr + ". TUC set to N/A.");
        timeUnderCommand = "N/A";
    }
  } catch (err) {
    Logger.log("[onDailyLogSubmit] Error calculating Time Under Command: " + err.message);
    timeUnderCommand = "N/A";
  }
 
  const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0]; 
 
  const colMap = {
    'Start Port': headerRow.indexOf('Start Port') + 1,
    'End Port': headerRow.indexOf('End Port') + 1,
    'Miles': headerRow.indexOf('Miles') + 1,
    'Time Under Command': headerRow.indexOf('Time Under Command') + 1,
    'Daily Log ✅': headerRow.indexOf('Daily Log ✅') + 1, 
    'Predeparture ✅': headerRow.indexOf('Predeparture ✅') + 1,
    'Before Approach ✅': headerRow.indexOf('Before Approach ✅') + 1,
    'Before Bed ✅': headerRow.indexOf('Before Bed ✅') + 1
  };
 
  const checklistStatuses = getChecklistStatusForDate(dateStr, submittedCharterID);
 
  const rangeToUpdate = logSheet.getRange(dailyRowIndex + 1, 1, 1, headerRow.length); 
  let rowValues = rangeToUpdate.getValues()[0]; 
 
  if (colMap['Start Port'] > 0) rowValues[colMap['Start Port'] - 1] = startPort;
  if (colMap['End Port'] > 0) rowValues[colMap['End Port'] - 1] = endPort;
  if (colMap['Miles'] > 0) rowValues[colMap['Miles'] - 1] = miles;
  if (colMap['Time Under Command'] > 0) rowValues[colMap['Time Under Command'] - 1] = timeUnderCommand;
 
  if (colMap['Daily Log ✅'] > 0) rowValues[colMap['Daily Log ✅'] - 1] = "✅"; 
 
  if (colMap['Predeparture ✅'] > 0) rowValues[colMap['Predeparture ✅'] - 1] = checklistStatuses["Predeparture"];
  if (colMap['Before Approach ✅'] > 0) rowValues[colMap['Before Approach ✅'] - 1] = checklistStatuses["Before Approach"];
  if (colMap['Before Bed ✅'] > 0) rowValues[colMap['Before Bed ✅'] - 1] = checklistStatuses["Before Bed"];
 
  rangeToUpdate.setValues([rowValues]); 
 
  Logger.log("[onDailyLogSubmit] Daily Log data updated for " + submittedCharterID + " on logging date " + dateStr + ".");
}
 
 
// --- Functions: Triggered by other daily Checklist Form Submissions (via onFormSubmit router) ---
function onChecklistSubmit(e) {
  Logger.log("[onChecklistSubmit] Function started.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);

  if (!logSheet) { Logger.log("Error: Master Log Sheet not found!"); return; }
  if (!settingsSheet) { Logger.log("Error: Settings Sheet not found!"); return; }

  const formResponses = e.namedValues;
  const submittedCharterID = formResponses["Charter ID"] ? formResponses["Charter ID"][0] : "";
  const submissionTimestamp = new Date(formResponses["Timestamp"][0]);
  const dateStr = getLoggingDate(submittedCharterID, submissionTimestamp); // Use the new logging date helper

  const responseSheetName = e.range.getSheet().getName(); 
  let formTitleForLog = "UNKNOWN_CHECKLIST_FORM"; 
  let checklistColHeader = "";

  if (responseSheetName === PREDEPARTURE_RESPONSES_SHEET_NAME) {
      formTitleForLog = "Pre-Departure Checklist";
      checklistColHeader = "Predeparture ✅";
  } else if (responseSheetName === BEFORE_APPROACH_RESPONSES_SHEET_NAME) {
      formTitleForLog = "Before Approach Form";
      checklistColHeader = "Before Approach ✅";
  } else if (responseSheetName === BEFORE_BED_RESPONSES_SHEET_NAME) {
      formTitleForLog = "Before Sleep Checklist Form";
      checklistColHeader = "Before Bed ✅";
  } else {
      Logger.log("[onChecklistSubmit] Warning: Called from unrecognized sheet: " + responseSheetName + ". Exiting.");
      return; 
  }

  Logger.log("[onChecklistSubmit] Called by: " + formTitleForLog + " from sheet: " + responseSheetName);

  const currentActiveCharterID = settingsSheet.getRange("B4").getValue(); 

  if (!submittedCharterID || submittedCharterID !== currentActiveCharterID) {
    Logger.log("[onChecklistSubmit] " + formTitleForLog + ": Submitted Charter ID '" + submittedCharterID + "' does not match active charter '" + currentActiveCharterID + "' or is empty. Submission ignored.");
    return;
  }

  // --- Pre-Departure Specific Logic with First-Submission Exception ---
  if (responseSheetName === PREDEPARTURE_RESPONSES_SHEET_NAME) {
      Logger.log("[onChecklistSubmit] Pre-Departure Checklist submitted. Initiating first-submission check.");
      const isFirst = isFirstPredeparture(submittedCharterID);
      Logger.log("[onChecklistSubmit] Pre-Departure: Is first for charter? " + isFirst);

      if (!isFirst) {
          const previousDayCompletionStatus = checkPreviousDayCompletionStatus(submittedCharterID, dateStr, logSheet, ss);
          if (!previousDayCompletionStatus.isComplete) {
              const blockMessage = "BLOCKING Pre-Departure for " + submittedCharterID + " on " + dateStr + ". Previous logging day (" + previousDayCompletionStatus.previousLoggingDate + ") incomplete: " + previousDayCompletionStatus.missingForms.join(', ') + ".";
              Logger.log("[onChecklistSubmit] ERROR: " + blockMessage);

              // Mark as BLOCKED_PREVIOUS_DAY_INCOMPLETE in Pre-Departure Responses sheet
              const pdResponsesSheet = ss.getSheetByName(PREDEPARTURE_RESPONSES_SHEET_NAME);
              const responseRow = e.range.getRow();
              const pdHeaders = pdResponsesSheet.getRange(1, 1, 1, pdResponsesSheet.getLastColumn()).getValues()[0];
              const charterIdColIndex = pdHeaders.indexOf("Charter ID");
              if (charterIdColIndex !== -1) {
                  pdResponsesSheet.getRange(responseRow, charterIdColIndex + 1).setValue("BLOCKED_PREVIOUS_DAY_INCOMPLETE");
                  Logger.log("[onChecklistSubmit] Pre-Departure submission marked as BLOCKED_PREVIOUS_DAY_INCOMPLETE.");
              }

              // Send email notification
              const boatName = getBoatNameFromMasterLog(submittedCharterID); 
              const emailSubject = "⛔ Action Required: Cannot Start New Day for " + submittedCharterID + " (" + boatName + ")";
              const emailBody = "A Pre-Departure Checklist for Charter ID " + submittedCharterID + " on " + dateStr + " was blocked.\n\n" +
                                "Reason: The previous logging day (" + previousDayCompletionStatus.previousLoggingDate + ") is incomplete.\n" +
                                "Missing forms: " + previousDayCompletionStatus.missingForms.join(', ') + "\n\n" +
                                "Please submit the missing forms for " + previousDayCompletionStatus.previousLoggingDate + " to proceed.\n\n" +
                                "Master Log: " + ss.getUrl(); 
              MailApp.sendEmail(ADMIN_EMAIL_FOR_NOTIFICATIONS, emailSubject, emailBody);
              Logger.log("[onChecklistSubmit] Email sent about blocked Pre-Departure due to incomplete previous day.");

              return; // BLOCK THE SUBMISSION
          }
      }
      Logger.log("[onChecklistSubmit] Pre-Departure check passed. Proceeding.");
  }

  // --- Log ✅ in Master Log ---
  const dailyRowIndex = getOrCreateDailyRow(logSheet, submittedCharterID, dateStr); 
  if (dailyRowIndex === null) {
      Logger.log("[onChecklistSubmit] Could not find or create daily row for " + formTitleForLog + " submission. Exiting.");
      return;
  }

  const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const checklistColIndex = headerRow.indexOf(checklistColHeader);

  if (checklistColIndex !== -1) { 
      const rangeToUpdate = logSheet.getRange(dailyRowIndex + 1, 1, 1, headerRow.length);
      let rowValues = rangeToUpdate.getValues()[0];
      rowValues[checklistColIndex] = "✅"; 
      rangeToUpdate.setValues([rowValues]); 
      Logger.log("[onChecklistSubmit] " + formTitleForLog + " updated to ✅ for " + submittedCharterID + " on logging date " + dateStr + ".");
  } else {
      Logger.log("[onChecklistSubmit] Error: Checklist column '" + checklistColHeader + "' not found in Master Log.");
  }
}
 
// --- Helper function for Previous Day Completion Gate ---
// Checks if mandatory forms for the previous logging day are complete.
function checkPreviousDayCompletionStatus(charterID, currentPredepartureLoggingDateStr, logSheet, ss) {
    Logger.log("[checkPreviousDayCompletionStatus] Checking previous day completion for " + charterID + " before " + currentPredepartureLoggingDateStr);
    const headerRowValues = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0]; // Master Log headers
 
    const preDepartureResponsesSheet = ss.getSheetByName(PREDEPARTURE_RESPONSES_SHEET_NAME);
    if (!preDepartureResponsesSheet) {
        Logger.log("[checkPreviousDayCompletionStatus] Warning: Pre-Departure Responses sheet not found. Cannot check previous day completion.");
        return { isComplete: true, missingForms: [], previousLoggingDate: null }; 
    }
 
    const pdData = preDepartureResponsesSheet.getDataRange().getValues();
    let currentPredepartureTimestamp = new Date(currentPredepartureLoggingDateStr); 
    let previousPredepartureTimestamp = null; 
 
    for (let i = 1; i < pdData.length; i++) {
        const rowPdTimestamp = new Date(pdData[i][0]);
        const rowPdCharterID = pdData[i][1];
        if (rowPdCharterID === charterID && rowPdTimestamp.getTime() < currentPredepartureTimestamp.getTime()) { 
            if (!previousPredepartureTimestamp || rowPdTimestamp.getTime() > previousPredepartureTimestamp.getTime()) {
                previousPredepartureTimestamp = rowPdTimestamp;
            }
        }
    }
 
    if (!previousPredepartureTimestamp) {
        Logger.log("[checkPreviousDayCompletionStatus] No previous Pre-Departure found. Assuming first logging day of charter. No previous day to check.");
        return { isComplete: true, missingForms: [], previousLoggingDate: null }; 
    }
 
    const previousLoggingDate = getLoggingDate(charterID, previousPredepartureTimestamp); 
 
    Logger.log("[checkPreviousDayCompletionStatus] Checking previous logging day: " + previousLoggingDate);
 
    const masterLogData = logSheet.getDataRange().getValues();
    let previousDayRowInMasterLog = null; 
 
    for (let i = 0; i < masterLogData.length; i++) {
        const row = masterLogData[i];
        if (row[0] && new Date(row[0]).toLocaleDateString('el-GR') === previousLoggingDate && row[1] === charterID) {
            previousDayRowInMasterLog = row;
            break;
        }
        if (typeof row[0] === "string" && (row[0].includes("–––––––––") || row[0].includes("Summary") || (row[1] && (row[1].startsWith("S") || row[1].startsWith("SPC"))))) {
            break; 
        }
    }
 
    if (!previousDayRowInMasterLog) {
        Logger.log("[checkPreviousDayCompletionStatus] Previous logging day row (" + previousLoggingDate + ") not found in Master Log. Marking as incomplete.");
        return { isComplete: false, missingForms: ["Daily Log", "Before Bed", "Before Approach", "Pre-Departure (Previous Day)"], previousLoggingDate: previousLoggingDate }; 
    }
 
    let missingForms = [];
    for (const formName of MANDATORY_DAILY_FORMS_FOR_GATE) { 
        const colIndex = headerRowValues.indexOf(formName + " ✅");
        if (colIndex !== -1 && previousDayRowInMasterLog[colIndex] !== "✅") {
            missingForms.push(formName);
        }
    }
    // Also explicitly check Predeparture for the previous day, as it's the gate.
    // The previous PD's status is crucial for its own day.
    const pdColIndex = headerRowValues.indexOf("Predeparture ✅");
    if (pdColIndex !== -1 && previousDayRowInMasterLog[pdColIndex] !== "✅") {
        if (!missingForms.includes("Pre-Departure")) { // Avoid duplicates if already added
            missingForms.push("Pre-Departure");
        }
    }
 
    const isComplete = missingForms.length === 0;
    Logger.log("[checkPreviousDayCompletionStatus] Previous day (" + previousLoggingDate + ") completion status: " + isComplete + ", Missing: " + JSON.stringify(missingForms));
 
    return { isComplete: isComplete, missingForms: missingForms, previousLoggingDate: previousLoggingDate };
}
 
 
// --- Function: Triggered on End of Charter Form Submission (via onFormSubmit router) ---
function onEndCharterSubmit(e) {
  Logger.log("[onEndCharterSubmit] Function started.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
 
  if (!logSheet) { Logger.log("Error: Master Log Sheet not found!"); return; }
  if (!settingsSheet) { Logger.log("Error: Settings Sheet not found!"); return; }
 
  const formResponses = e.namedValues;
  const submittedCharterID = formResponses["Charter ID"] ? formResponses["Charter ID"][0] : "";
  const submissionTimestamp = new Date(formResponses["Timestamp"][0]);
  // Use submission timestamp for end date as it represents final action time
  const endDateStr = submissionTimestamp.toLocaleDateString('el-GR'); 
 
  const currentActiveCharterID = settingsSheet.getRange("B4").getValue();
 
  if (!submittedCharterID || submittedCharterID !== currentActiveCharterID) {
    Logger.log("[onEndCharterSubmit] Submitted Charter ID '" + submittedCharterID + "' does not match active charter '" + currentActiveCharterID + "' or is empty. Submission ignored.");
    return;
  }
 
  const finalEnginePort = parseFloat(formResponses["Final Engine Hours Port"] ? formResponses["Final Engine Hours Port"][0] : 0);
  const finalEngineST = parseFloat(formResponses["Final Engine Hours ST"] ? formResponses["Final Engine Hours ST"][0] : 0);
  const finalGenerator = parseFloat(formResponses["Generator Hours"] ? formResponses["Generator Hours"][0] : 0);
  const dieselConsumed = parseFloat(formResponses["Liters of Diese Consumed"] ? formResponses["Liters of Diese Consumed"][0] : 0);
 
  const data = logSheet.getDataRange().getValues();
  let charterHeaderRowIndex = -1; 
  let lastDailyRowIndex = -1; 
 
  for (let i = 0; i < data.length; i++) {
    if (typeof data[i][1] === "string" && (data[i][1].startsWith("S") || data[i][1].startsWith("SPC")) && data[i][1].includes(submittedCharterID)) { 
      charterHeaderRowIndex = i;
      for (let j = i + 1; j < data.length; j++) { 
              const row = data[j];
              if (typeof row[0] === "string" && (row[0].includes("–––––––––") || row[0].includes("Summary") || (row[1] && (row[1].startsWith("S") || row[1].startsWith("SPC"))))) {
                  lastDailyRowIndex = j - 1; 
                  break;
              }
              lastDailyRowIndex = j; 
          }
          break; 
        }
      }
 
      if (charterHeaderRowIndex === -1) {
        Logger.log("[onEndCharterSubmit] Charter ID " + submittedCharterID + " header not found in Master Log. Exiting.");
        return;
      }
 
      let totalMiles = 0;
      let totalTimeUnderCommand = 0;
      let daysCount = 0;
 
      if (lastDailyRowIndex !== -1) { 
        for (let i = charterHeaderRowIndex + 1; i <= lastDailyRowIndex; i++) { 
          const dailyRow = data[i];
          if (dailyRow[8] && typeof dailyRow[8] === 'number') totalMiles += dailyRow[8]; 
          if (dailyRow[9] && typeof dailyRow[9] === 'number') totalTimeUnderCommand += dailyRow[9]; 
          if (dailyRow[0] && dailyRow[1] === submittedCharterID) daysCount++; 
        }
      }
      Logger.log("[onEndCharterSubmit] Calculated Totals: Miles=" + totalMiles + ", Time=" + totalTimeUnderCommand + ", Days=" + daysCount);
 
      const summaryRowIndex = lastDailyRowIndex + 1; 
      logSheet.insertRowAfter(summaryRowIndex); 
      const actualSummaryRowIndex = summaryRowIndex + 1;
 
 
      const headerRowForSummaryPlacement = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
      const milesColIndex = headerRowForSummaryPlacement.indexOf('Miles'); 
      const timeCmdColIndex = headerRowForSummaryPlacement.indexOf('Time Under Command'); 
      const predepartureColIndex = headerRowForSummaryPlacement.indexOf('Predeparture ✅'); 
 
      let finalSummaryRowValues = Array(headerRowForSummaryPlacement.length).fill(""); 
      finalSummaryRowValues[0] = "Summary"; 
      finalSummaryRowValues[1] = submittedCharterID; 
 
      if (milesColIndex !== -1) finalSummaryRowValues[milesColIndex] = "Total Miles: " + totalMiles.toFixed(0);
      if (timeCmdColIndex !== -1) finalSummaryRowValues[timeCmdColIndex] = "Total Time: " + totalTimeUnderCommand.toFixed(2) + "h";
 
      let currentSummaryCol = predepartureColIndex; 
      if (currentSummaryCol !== -1) {
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "Port Eng: " + finalEnginePort + "h";
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "Stbd Eng: " + finalEngineST + "h";
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "Gen: " + finalGenerator + "h";
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "Diesel: " + dieselConsumed + "L";
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "End Date: " + endDateStr;
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol++] = "Days: " + daysCount;
          if (currentSummaryCol < finalSummaryRowValues.length) finalSummaryRowValues[currentSummaryCol] = "COMPLETED"; 
      }
 
      logSheet.getRange(actualSummaryRowIndex, 1, 1, finalSummaryRowValues.length).setValues([finalSummaryRowValues]);
      logSheet.getRange(actualSummaryRowIndex, 1, 1, finalSummaryRowValues.length).setBackground("#e0f2f7"); 
      logSheet.getRange(actualSummaryRowIndex, 1, 1, finalSummaryRowValues.length).setFontWeight("bold");
 
      logSheet.getRange(charterHeaderRowIndex + 1, 2).setBackground("#cfe2f3"); // Color Charter ID cell in header blue
 
      settingsSheet.getRange("B4").clearContent(); 
      Logger.log("[onEndCharterSubmit] Charter " + submittedCharterID + " marked as COMPLETED. Settings updated.");
    }
 
 
    // --- Daily Scheduled Function: Check for Missing Checklists & Send Email ---
    function checkMissingChecklistsDaily() {
      Logger.log("[checkMissingChecklistsDaily] Function started.");
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName(MASTER_LOG_SHEET_NAME);
      const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
 
      if (!logSheet) { Logger.log("Error: Master Log Sheet not found!"); return; }
      if (!settingsSheet) { Logger.log("Error: Settings Sheet not found!"); return; }
 
      const activeCharterID = settingsSheet.getRange("B4").getValue(); 
 
      if (!activeCharterID) {
        Logger.log("[checkMissingChecklistsDaily] No active charter to check for missing checklists. Exiting.");
        return; 
      }
      Logger.log("Active Charter ID: " + activeCharterID);
 
      const today = new Date();
      // --- Determine logging date for today's check ---
      const dateStr = getLoggingDate(activeCharterID, today); // Use logging date helper for current check
 
      const data = logSheet.getDataRange().getValues();
      let charterHeaderRowIndex = -1; 
      let lastDailyRowIndex = -1; 
 
      for (let i = 0; i < data.length; i++) {
        if (typeof data[i][1] === "string" && (data[i][1].startsWith("S") || data[i][1].startsWith("SPC")) && data[i][1].includes(activeCharterID)) {
          charterHeaderRowIndex = i;
          for (let j = i + 1; j < data.length; j++) { 
              const row = data[j];
              if (typeof row[0] === "string" && (row[0].includes("–––––––––") || row[0].includes("Summary") || (row[1] && (row[1].startsWith("S") || row[1].startsWith("SPC"))))) {
                  lastDailyRowIndex = j - 1; 
                  break;
              }
              lastDailyRowIndex = j; 
          }
          break; 
        }
      }
 
      if (charterHeaderRowIndex === -1) {
        Logger.log("[checkMissingChecklistsDaily] Active charter " + activeCharterID + " header not found for daily checks. Exiting.");
        return;
      }
      Logger.log("Charter header found at row " + charterHeaderRowIndex + ". Last daily row at " + lastDailyRowIndex + ".");
 
      let todayRowIndex = -1; 
      for (let i = charterHeaderRowIndex + 1; i <= lastDailyRowIndex; i++) { 
        if (new Date(data[i][0]).toLocaleDateString('el-GR') === dateStr) {
          todayRowIndex = i;
          Logger.log("Found today's daily row at index " + todayRowIndex + ".");
          break;
        }
      }
 
      if (todayRowIndex === -1) {
          Logger.log("No daily row found for " + activeCharterID + " for today (" + dateStr + "). Cannot perform checklist checks.");
          return; 
      }
 
      const dailyRowValues = data[todayRowIndex]; 
      const headerRowValues = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0]; 
 
      const checklistTypesToCheck = ["Predeparture", "Daily Log", "Before Approach", "Before Bed"]; 
      let missingChecklists = [];
      let valuesToUpdate = [...dailyRowValues]; 
 
      const updatedStatuses = getChecklistStatusForDate(dateStr, activeCharterID); 
      Logger.log("Updated statuses from forms: " + JSON.stringify(updatedStatuses));
 
      for (const type of checklistTypesToCheck) {
        const colIndex = headerRowValues.indexOf(type + " ✅"); 
        if (colIndex !== -1) {
          if (updatedStatuses[type] === "❌") { 
            missingChecklists.push(type);
            valuesToUpdate[colIndex] = "❌"; 
          } else { 
            valuesToUpdate[colIndex] = "✅"; 
          }
        }
      }
      logSheet.getRange(todayRowIndex + 1, 1, 1, valuesToUpdate.length).setValues([valuesToUpdate]); 
 
      for (const type of checklistTypesToCheck) {
        const colIndex = headerRowValues.indexOf(type + " ✅");
        if (colIndex !== -1) {
          const cell = logSheet.getRange(todayRowIndex + 1, colIndex + 1); 
          if (cell.getValue() === "✅") {
            cell.setBackground("#b6d7a8"); 
          } else if (cell.getValue() === "❌") {
            cell.setBackground("#ea9999"); 
          } else {
            cell.setBackground(null); 
          }
        }
      }
 
      if (missingChecklists.length > 0) {
        const boatName = logSheet.getRange(charterHeaderRowIndex + 1, 4).getValue(); 
        const emailSubject = "⚠️ Missing Checklists for Charter " + activeCharterID + " (" + boatName + ") - " + dateStr;
        const emailBody = "The following checklists are missing for Charter " + activeCharterID + " (" + boatName + ") today (" + dateStr + "):\n\n- " + missingChecklists.join('\n- ') + "\n\nPlease ensure these are completed.";
 
        MailApp.sendEmail(ADMIN_EMAIL_FOR_NOTIFICATIONS, emailBody, emailSubject); 
        Logger.log("Sent email alert to " + ADMIN_EMAIL_FOR_NOTIFICATIONS + " for " + activeCharterID + ".");
      } else {
        Logger.log("All required checklists completed for " + activeCharterID + " on " + dateStr + ". No alert sent.");
      }
    }

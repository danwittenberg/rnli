/**
 * Global Configuration Settings
 * Ensure your Spreadsheet sheet tab names perfectly match these constants.
 */
const SHEET_CREW = "Crew";
const SHEET_SESSIONS = "Sessions";
const SHEET_CONFIG = "Config";
const SHEET_SIGNUPS = "Signups"; 

/**
 * Core GET Request Handler Interface
 */
function doGet(e) {
  var payload = getMatrixDataPayload();
  
  if (e && (e.parameter || e.queryString)) {
    return ContentService.createTextOutput(JSON.stringify(payload))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  return payload;
}

/**
 * Core POST Request Handler Interface
 */
function doPost(e) {
  var payloadString;
  
  if (typeof e === 'string') {
    payloadString = e;
  } else if (e && e.postData && e.postData.contents) {
    payloadString = e.postData.contents;
  } else {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No data payload package discovered." }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    var payload = JSON.parse(payloadString);
    processActionExecution(payload);
    
    var responseObj = { status: "success", message: "Ledger transaction updated successfully." };
    if (typeof e === 'string') {
      return responseObj;
    } else {
      return ContentService.createTextOutput(JSON.stringify(responseObj))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    var errorObj = { status: "error", message: error.toString() };
    if (typeof e === 'string') {
      return errorObj;
    } else {
      return ContentService.createTextOutput(JSON.stringify(errorObj))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }
}

/**
 * Helper: Dynamically maps column indices by scanning the header row of the Sessions sheet.
 */
function getSessionsColumnMapping(sheet) {
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 5)).getValues()[0];
  // Safe defaults (0-based indices)
  var mapping = { id: 0, month: 1, topic: 2, date: 3, time: 4 };
  
  for (var i = 0; i < headers.length; i++) {
    var headerText = headers[i].toString().toLowerCase().trim();
    if (headerText.includes("id")) {
      mapping.id = i;
    } else if (headerText.includes("month")) {
      mapping.month = i;
    } else if (headerText.includes("date")) {
      mapping.date = i;
    } else if (headerText.includes("time")) {
      mapping.time = i;
    } else if (headerText.includes("name") || headerText.includes("topic") || headerText.includes("session")) {
      if (!headerText.includes("id") && !headerText.includes("date") && !headerText.includes("time") && !headerText.includes("month")) {
        mapping.topic = i;
      }
    }
  }
  return mapping;
}

/**
 * Compiles and returns all spreadsheet tables translated into responsive JSON structures.
 */
function getMatrixDataPayload() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  
  // 1. Pull Crew names array (from Column A)
  var crewSheet = ss.getSheetByName(SHEET_CREW);
  var crew = [];
  if (crewSheet && crewSheet.getLastRow() >= 2) {
    crew = crewSheet.getRange(2, 1, crewSheet.getLastRow() - 1, 1).getValues().flat().filter(Boolean);
  }
  
  // 2. Pull Sessions array details with automated Calendar Month grouping
  var sessionSheet = ss.getSheetByName(SHEET_SESSIONS);
  var sessions = [];
  if (sessionSheet && sessionSheet.getLastRow() >= 2) {
    var lastRow = sessionSheet.getLastRow();
    var colMapping = getSessionsColumnMapping(sessionSheet);
    var maxCol = Math.max(colMapping.id, colMapping.month, colMapping.topic, colMapping.date, colMapping.time) + 1;
    
    var sessionRows = sessionSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
    var sessionDisplayRows = sessionSheet.getRange(2, 1, lastRow - 1, maxCol).getDisplayValues();
    
    sessions = sessionRows.map(function(row, idx) {
      var displayRow = sessionDisplayRows[idx];
      var dateObj = row[colMapping.date]; // Dynamically targets the real Date column
      
      var calendarMonthText = "";
      var dayNameText = "Training Day";
      
      if (dateObj instanceof Date) {
        calendarMonthText = Utilities.formatDate(dateObj, tz, "MMMM"); 
        dayNameText = Utilities.formatDate(dateObj, tz, "EEEE");        
      } else if (dateObj) {
        try {
          var pDate = new Date(dateObj);
          if (!isNaN(pDate.getTime())) {
            calendarMonthText = Utilities.formatDate(pDate, tz, "MMMM");
            dayNameText = Utilities.formatDate(pDate, tz, "EEEE");
          }
        } catch(err) {}
      }
      
      if (!calendarMonthText) {
        calendarMonthText = displayRow[colMapping.month] ? displayRow[colMapping.month].toString().trim() : "Unknown Month";
      }
      
      return {
        id: row[colMapping.id].toString().trim(),
        month: calendarMonthText, 
        topic: displayRow[colMapping.topic] ? displayRow[colMapping.topic].toString().trim() : "",
        date: displayRow[colMapping.date] ? displayRow[colMapping.date].toString().trim() : "",
        dayName: dayNameText,
        time: displayRow[colMapping.time] ? displayRow[colMapping.time].toString().trim() : ""
      };
    });
  }
  
  // 3. Pull operational drop-down values from Config sheet
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  var categories = [];
  var roles = [];
  var adminCategories = [];
  var subCategories = [];
  
  if (configSheet && configSheet.getLastRow() >= 2) {
    var configMax = configSheet.getLastRow();
    categories = configSheet.getRange(2, 1, configMax - 1, 1).getValues().flat().filter(Boolean);
    roles = configSheet.getRange(2, 2, configMax - 1, 1).getValues().flat().filter(Boolean);
    adminCategories = configSheet.getRange(2, 4, configMax - 1, 1).getValues().flat().filter(Boolean);
    subCategories = configSheet.getRange(2, 5, configMax - 1, 1).getValues().flat().filter(Boolean);
  }
  
  // 4. Pull existing assignments from Signups sheet mapping to precise columns A-K
  var signupSheet = ss.getSheetByName(SHEET_SIGNUPS);
  var signups = [];
  if (signupSheet && signupSheet.getLastRow() >= 2) {
    var signupRows = signupSheet.getRange(2, 1, signupSheet.getLastRow() - 1, 11).getValues();
    signups = signupRows.map(function(row) {
      return {
        name: row[3] ? row[3].toString().trim() : "",                  // Col D: CrewName
        sessionId: row[9] ? row[9].toString().trim() : "",             // Col J: SessionID
        preferred: row[4] ? row[4].toString().trim() : "",             // Col E: PreferredChoice
        notes: row[5] ? row[5].toString().trim() : "",                 // Col F: Notes
        allocatedCategory: row[6] ? row[6].toString().trim() : "",     // Col G: AllocatedCategory
        allocatedRole: row[7] ? row[7].toString().trim() : "",         // Col H: AllocatedRole
        allocatedSubCategory: row[10] ? row[10].toString().trim() : "" // Col K: Allocated Activity
      };
    });
  }
  
  return {
    crew: crew,
    sessions: sessions,
    categories: categories,
    roles: roles,
    adminCategories: adminCategories,
    subCategories: subCategories, 
    signups: signups
  };
}

/**
 * Maps incoming GUI actions and writes them directly to explicit column keys.
 * Implements LockService to handle simultaneous user race-conditions.
 */
function processActionExecution(payload) {
  // A. Establish the atomic Script Lock
  var lock = LockService.getScriptLock();
  
  try {
    // B. Block execution threads here, waiting up to 10 seconds for concurrent instances to complete
    lock.waitLock(10000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_SIGNUPS);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SIGNUPS);
      sheet.appendRow(["Session Date", "Session Time", "Session Name", "CrewName", "PreferredChoice", "Notes", "AllocatedCategory", "AllocatedRole", "Timestamp", "SessionID", "Allocated Activity"]);
    }
    
    var crewName = payload.crewName;
    var sessionId = payload.sessionId;
    var action = payload.action;
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var targetRowIndex = -1;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][3].toString().trim() === crewName && values[i][9].toString().trim() === sessionId.toString()) {
        targetRowIndex = i + 1; 
        break;
      }
    }
    
    if (action === "remove") {
      if (targetRowIndex !== -1) {
        sheet.deleteRow(targetRowIndex);
      }
      // C. Flush cache to complete rows structural compression instantly
      SpreadsheetApp.flush();
      return;
    }
    
    var sessionContext = getSessionMetaContext(sessionId);
    var formattedTimestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    if (action === "signup") {
      if (targetRowIndex !== -1) {
        sheet.getRange(targetRowIndex, 5).setValue(payload.preferredRole);
        sheet.getRange(targetRowIndex, 6).setValue(payload.notes);
        sheet.getRange(targetRowIndex, 9).setValue(formattedTimestamp);
      } else {
        var nextRow = sheet.getLastRow() + 1;
        var rowData = [
          sessionContext.date,   // A: Session Date
          sessionContext.time,   // B: Session Time
          sessionContext.topic,  // C: Session Name
          crewName,              // D: CrewName
          payload.preferredRole, // E: PreferredChoice
          payload.notes,         // F: Notes
          "",                    // G: AllocatedCategory
          "",                    // H: AllocatedRole
          formattedTimestamp,    // I: Timestamp
          sessionId,             // J: SessionID
          ""                     // K: Allocated Activity
        ];
        sheet.getRange(nextRow, 1, 1, 11).setValues([rowData]);
      }
    } 
    else if (action === "allocate") {
      if (targetRowIndex !== -1) {
        sheet.getRange(targetRowIndex, 7).setValue(payload.roleType);
        sheet.getRange(targetRowIndex, 8).setValue(payload.specificRole);
        sheet.getRange(targetRowIndex, 9).setValue(formattedTimestamp);
        sheet.getRange(targetRowIndex, 11).setValue(payload.subCategory);
      } else {
        var nextRow = sheet.getLastRow() + 1;
        var rowData = [
          sessionContext.date,   // A: Session Date
          sessionContext.time,   // B: Session Time
          sessionContext.topic,  // C: Session Name
          crewName,              // D: CrewName
          "",                    // E: PreferredChoice
          "",                    // F: Notes
          payload.roleType,      // G: AllocatedCategory
          payload.specificRole,  // H: AllocatedRole
          formattedTimestamp,    // I: Timestamp
          sessionId,             // J: SessionID
          payload.subCategory    // K: Allocated Activity
        ];
        sheet.getRange(nextRow, 1, 1, 11).setValues([rowData]);
      }
    }
    
    // D. CRITICAL: Force cache writing to cells BEFORE lock-release occurs
    SpreadsheetApp.flush();
    
  } catch (error) {
    // Bubble database error message up gracefully to doPost handler
    throw new Error("Concurrency Database error: " + error.toString());
  } finally {
    // E. Release the lock pipeline so the next queued request can process
    lock.releaseLock();
  }
}

/**
 * Private Helper: Resolves metadata properties using dynamic column header queries
 */
function getSessionMetaContext(sessionId) {
  var fallback = { month: "", topic: "", date: "", time: "" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SESSIONS);
  if (!sheet || sheet.getLastRow() < 2) return fallback;
  
  var colMapping = getSessionsColumnMapping(sheet);
  var maxCol = Math.max(colMapping.id, colMapping.month, colMapping.topic, colMapping.date, colMapping.time) + 1;
  
  var displayRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, maxCol).getDisplayValues();
  for (var i = 0; i < displayRows.length; i++) {
    if (displayRows[i][colMapping.id].toString().trim() === sessionId.toString()) {
      return {
        month: displayRows[i][colMapping.month] ? displayRows[i][colMapping.month].toString().trim() : "",
        topic: displayRows[i][colMapping.topic] ? displayRows[i][colMapping.topic].toString().trim() : "",
        date: displayRows[i][colMapping.date] ? displayRows[i][colMapping.date].toString().trim() : "",
        time: displayRows[i][colMapping.time] ? displayRows[i][colMapping.time].toString().trim() : ""
      };
    }
  }
  return fallback;
}
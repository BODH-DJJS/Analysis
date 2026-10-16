// ==========================================
// BODH Reporting Form - Apps Script Backend
// ==========================================

var MAIN_DRIVE_FOLDER_ID = '1Z6PKvQbNXo2RJ7DGZsCaWgoO1nzZiPgw';
var SHEET_ID = '1OvVqqDI5JwTMJA4svMKXUb2JuOMQUaPgH7Un-UMum14';
var SHEET_TAB_NAME = 'BODH Responses';

// Serve the HTML form
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.userEmail = (e && e.parameter && e.parameter.email) ? e.parameter.email : '';
  return template.evaluate()
      .setTitle('BODH Event Reporting Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// Helper: Get or Create a subfolder by name
// ==========================================
function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

// ==========================================
// Auto Folder & Doc Creation
// ==========================================
// ----------------------------------------------------------------------
// DYNAMIC FOLDER CREATION
// ----------------------------------------------------------------------

function getPathToDateFolder(formData) {
  var state = formData.state; var branch = formData.branch; var dateStr = formData.date;
  if (!state || !branch || !dateStr) throw new Error('State, Branch, and Date are required.');

  var d = new Date(dateStr + 'T00:00:00'); var year = d.getFullYear().toString();
  var fullMonths = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var shortMonths = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  var monthName = fullMonths[d.getMonth()];
  var day = d.getDate();
  var dayStr = (day < 10 ? '0' : '') + day;
  
  var baseName = dayStr + '-' + shortMonths[d.getMonth()] + '-' + year;
  var safeVenue = formData.venue ? formData.venue.trim().replace(/[\\/:*?"<>|]/g, '') : 'NoVenue';

  var mainDrive = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);
  var eventsFolder = getOrCreateFolder(mainDrive, '2 - Events Reports - Year Wise');
  var stateFolder = getOrCreateFolder(eventsFolder, state);
  var branchFolder = getOrCreateFolder(stateFolder, branch);
  var yearFolder = getOrCreateFolder(branchFolder, year);
  var monthFolder = getOrCreateFolder(yearFolder, monthName);
  
  // Smart Venue Collision Logic
  var folders = monthFolder.getFoldersByName(baseName);
  var targetFolder = null;
  
  if (folders.hasNext()) {
    var firstFolder = folders.next();
    var desc = firstFolder.getDescription();
    
    if (!desc) {
      // Claim legacy un-described folder
      firstFolder.setDescription(safeVenue);
      targetFolder = firstFolder;
    } else if (desc === safeVenue) {
      // Belongs to this venue
      targetFolder = firstFolder;
    } else {
      // Belongs to a different venue! Create suffixed folder.
      var suffixedName = baseName + ' ' + safeVenue;
      targetFolder = getOrCreateFolder(monthFolder, suffixedName);
      targetFolder.setDescription(safeVenue);
    }
  } else {
    // 1st event of the day!
    targetFolder = monthFolder.createFolder(baseName);
    targetFolder.setDescription(safeVenue);
  }
  
  return targetFolder;
}

// Called real-time when user clicks "Yes" for a specific media type
function createSpecificMediaFolder(formData, typeName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Prevent duplicate folders from concurrent clicks
    var dateFolder = getPathToDateFolder(formData);
    var rawFolder = getOrCreateFolder(dateFolder, 'Raw');
    var specificFolder = getOrCreateFolder(rawFolder, typeName);
    return { success: true, url: specificFolder.getUrl() };
  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Called during final submit to ensure core structure and docs exist
function finalizeFolderStructure(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var dateFolder = getPathToDateFolder(formData);
    
    var rawFolder = getOrCreateFolder(dateFolder, 'Raw');
    
    // Create Thumbnail under Photos if Photos folder exists
    var photosIter = rawFolder.getFoldersByName('Photos');
    if (photosIter.hasNext()) {
      getOrCreateFolder(photosIter.next(), 'Thumbnail');
    }
    
    getOrCreateFolder(dateFolder, 'Selected');
    getOrCreateFolder(dateFolder, 'Selected (Edited)');

    var docNames = ['Writing', 'Editing', 'Proofreading', 'Crosscheck'];
    var existingFiles = dateFolder.getFilesByType(MimeType.GOOGLE_DOCS);
    var foundDocs = {};
    while (existingFiles.hasNext()) foundDocs[existingFiles.next().getName()] = true;
    
    for (var i = 0; i < docNames.length; i++) {
      if (!foundDocs[docNames[i]]) {
        var newDoc = DocumentApp.create(docNames[i]);
        DriveApp.getFileById(newDoc.getId()).moveTo(dateFolder);
      }
    }
    return dateFolder.getUrl();
  } catch (e) {
    console.error(e);
    return '';
  } finally {
    lock.releaseLock();
  }
}

// Called real-time when user changes from "Yes" to "No"
function deleteSpecificMediaFolder(formData, typeName) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var dateFolder = getPathToDateFolder(formData);
    if (!dateFolder) return { success: true };
    var rawIter = dateFolder.getFoldersByName('Raw');
    if (!rawIter.hasNext()) return { success: true };
    var rawFolder = rawIter.next();
    
    var specIter = rawFolder.getFoldersByName(typeName);
    if (specIter.hasNext()) {
      specIter.next().setTrashed(true);
    }
    return { success: true };
  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// Submit Report -> Append to Google Sheet
// ==========================================
function submitReport(formData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    // Check file limits before creating core structure or writing to sheet
    var dateFolderForCheck = getPathToDateFolder(formData);
    
    var rawIter = dateFolderForCheck.getFoldersByName('Raw');
    var rawFolder = rawIter.hasNext() ? rawIter.next() : null;
    
    var limitErrors = [];
    
    function checkLimit(fName, min, isSelected) {
       var count = 0;
       if (rawFolder) {
         var iter = rawFolder.getFoldersByName(fName);
         if (iter.hasNext()) {
            var mediaFolder = iter.next();
            // Count files
            var files = mediaFolder.getFiles();
            while(files.hasNext()) { files.next(); count++; }
            // Count subfolders
            var folders = mediaFolder.getFolders();
            while(folders.hasNext()) { folders.next(); count++; }
         }
       }
       if (isSelected === 'Yes') {
         if (count < min) limitErrors.push(fName + ": You must upload at least " + min + " file(s) or folder(s).");
         // Note: Max limits (40 for photos, 15 for videos) cannot be enforced by blocking submission 
         // because users on Shared Drives cannot delete files if they accidentally upload too many!
       }
    }
    
    checkLimit('Photos', 1, formData.photos);
    checkLimit('Videos', 0, formData.videos);
    checkLimit('Press Release', 0, formData.pressRelease);
    checkLimit('Appreciation', 0, formData.appreciation);
    
    if (limitErrors.length > 0) {
      lock.releaseLock();
      return { error: "FILE_LIMIT_ERROR", messages: limitErrors.join('\n') };
    }

    // 1. Finalize Folder Structure silently
    var folderLink = finalizeFolderStructure(formData);

    // 2. Save to Sheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheet) throw new Error('Sheet tab "' + SHEET_TAB_NAME + '" not found.');

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Read email passed from frontend and remove any JSON stringify quotes
    var autoEmail = '';
    if (formData.email) {
      autoEmail = formData.email.replace(/^"|"$/g, '');
    }

    // Format time range
    var timeString = formatTimeRange(formData.timeFrom, formData.timeTo);

    var rowData = [];
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toString().trim();

      if (h === 'Timestamp') {
        rowData.push(new Date());
      }
      else if (h === 'Email Address') {
        rowData.push(autoEmail);
      }
      else if (h.indexOf('कार्यकारी टीम') > -1 || h.indexOf('Executing Team') > -1) {
        rowData.push(formData.executingTeam || '');
      }
      else if (h === 'राज्य | State') {
        rowData.push(formData.state || '');
      }
      // Match the correct state-specific branch column
      else if (h.indexOf('शाखा | Branch') > -1) {
        if (h.indexOf('(' + formData.state + ')') > -1) {
          rowData.push(formData.branch || '');
        } else {
          rowData.push('');
        }
      }
      else if (h.indexOf('स्थान | Venue') > -1) {
        rowData.push(formData.venue || '');
      }
      else if (h.indexOf('विषय / थीम') > -1) {
        rowData.push(formData.theme || '');
      }
      else if (h.indexOf('दिनांक / Date') > -1) {
        rowData.push(formData.date || '');
      }
      else if (h.indexOf('समय / Timings') > -1) {
        rowData.push(timeString || '');
      }
      else if (h.indexOf('पुरुष / Male') > -1) {
        rowData.push(formData.male || '');
      }
      else if (h.indexOf('महिलाएं / Female') > -1) {
        rowData.push(formData.female || '');
      }
      else if (h.indexOf('बच्चे') > -1) {
        rowData.push(formData.children || '');
      }
      else if (h.indexOf('आयोजक / Organizer') > -1) {
        rowData.push(formData.organizer || '');
      }
      else if (h.indexOf('भागीदार / प्रायोजक') > -1) {
        rowData.push(formData.partner || '');
      }
      else if (h.indexOf('अतिथि / वक्ता') > -1) {
        rowData.push(''); // Left blank since we mapped to Name
      }
      else if (h === 'नाम | Name') {
        rowData.push(formData.guestName || '');
      }
      else if (h.indexOf('पद / Designation') > -1) {
        rowData.push(formData.guestDesignation || '');
      }
      else if (h.indexOf('सहयोगी संगठन') > -1) {
        rowData.push(formData.guestOrg || '');
      }
      else if (h.indexOf('संचालित गतिविधियाँ') > -1) {
        rowData.push(formData.activities || '');
      }
      else if (h.indexOf('कोई विशेष जानकारी') > -1) {
        rowData.push(formData.observation || '');
      }
      else if (h.indexOf('फोटोज़') > -1 || h.indexOf('फोटोज़') > -1 || h.indexOf('Photos') > -1 || h.indexOf('Photographs') > -1) {
        rowData.push(formData.photos || '');
      }
      else if (h.indexOf('वीडियो') > -1 || h.indexOf('Videos') > -1) {
        rowData.push(formData.videos || '');
      }
      else if (h.indexOf('प्रैस विज्ञप्ति') > -1 || h.indexOf('Press Release') > -1 || h.indexOf('Press Coverage') > -1) {
        rowData.push(formData.pressRelease || '');
      }
      else if (h.indexOf('प्रशंसा पत्र') > -1 || h.indexOf('Appreciation Letter') > -1) {
        rowData.push(formData.appreciation || '');
      }
      else if (h.indexOf('प्रचार माध्यम') > -1) {
        rowData.push(formData.publicity || '');
      }
      else if (h.indexOf('ज्ञान दीक्षा') > -1) {
        rowData.push(formData.initiatedCount || '');
      }
      else if (h.indexOf('फीडबैक') > -1) {
        rowData.push(formData.feedback || '');
      }
      else if (h === 'Link') {
        rowData.push(folderLink || '');
      }
      else {
        rowData.push('');
      }
    }

    sheet.appendRow(rowData);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function formatTimeRange(start, end) {
  if (!start && !end) return '';
  function toAmPm(time24) {
    if (!time24) return '';
    var parts = time24.split(':');
    var h = parseInt(parts[0], 10);
    var m = parts[1];
    var ampm = h >= 12 ? 'pm' : 'am';
    h = h % 12;
    h = h ? h : 12;
    return h + ':' + m + ' ' + ampm;
  }
  var s = toAmPm(start);
  var e = toAmPm(end);
  if (s && e) return s + ' to ' + e;
  return s || e;
}

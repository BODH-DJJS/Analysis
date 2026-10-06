// Configuration
const SPREADSHEET_ID = '1OvVqqDI5JwTMJA4svMKXUb2JuOMQUaPgH7Un-UMum14';
const SHEET_NAME = 'Master data';
const RESPONSES_SHEET_NAME = 'POC Team Leaders';

/**
 * Web app entry point for GET requests
 */
function doGet(e) {
  try {
    // ==========================================
    // NEW ADDITION: Route to the HTML form if ?p=form
    // ==========================================
    if (e && e.parameter && e.parameter.p === 'form') {
      var template = HtmlService.createTemplateFromFile('index');
      template.userEmail = (e && e.parameter && e.parameter.email) ? e.parameter.email : '';
      return template.evaluate()
          .setTitle('BODH Event Reporting Form')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    // ==========================================

    const data = getAnalyticsData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader('Cache-Control', 'no-cache, no-store, must-revalidate')
      .addHeader('Pragma', 'no-cache')
      .addHeader('Expires', '0');
  } catch (error) {
    console.error('Error in doGet:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON)
      .setStatusCode(500);
  }
}

/**
 * Web app entry point for POST requests (Form Submissions)
 */
function doPost(e) {
  try {
    let formData;
    
    // Parse the incoming JSON data
    if (e.postData && e.postData.contents) {
      formData = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No form data received');
    }

    console.log('Received form data:', JSON.stringify(formData));

    // Open the spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found');

    // Get or create the Responses sheet
    let responseSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!responseSheet) {
      responseSheet = ss.insertSheet(RESPONSES_SHEET_NAME);
      // Add headers
      const headers = [
        'Timestamp',
        'Full Name',
        'Primary Contact Number',
        'Email Id',
        'Age',
        'Date of Birth',
        'Gyan Diksha Date',
        'State',
        'Branch Name',
        'Branch AR Name',
        'Branch AR Number',
        'Brother / Sister',
        'Role in BODH Team',
        'Highest Education',
        'Current Occupation',
        'Profession',
        'Years of Experience',
        'Company Name',
        'Languages Known',
        'Current Branch Responsibilities',
        'Available Time Commitment',
        'HO Level Sewa Yes/No',
        'HO Level Sewa Responsibilities'
      ];
      responseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // Style the header row
      const headerRange = responseSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1a237e');
      headerRange.setFontColor('#ffffff');
      headerRange.setHorizontalAlignment('center');
      
      // Auto-resize columns
      for (let i = 1; i <= headers.length; i++) {
        responseSheet.autoResizeColumn(i);
      }
      
      // Freeze header row
      responseSheet.setFrozenRows(1);
    }

    // Prepare the row data
    const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
    const rowData = [
      timestamp,
      formData.fullName || '',
      formData.contact || '',
      formData.email || '',
      formData.age || '',
      formData.dob || '',
      formData.gyanDikshaDate || '',
      formData.state || '',
      formData.branchName || '',
      formData.branchArName || '',
      formData.branchArNumber || '',
      '', // Brother / Sister is removed from form, so we leave this blank to maintain column alignment
      formData.bodhRole || '',
      formData.education || '',
      formData.occupation || '',
      formData.profession || '',
      formData.experience || '',
      formData.companyName || '',
      formData.languages || '',
      formData.responsibilities || '',
      (formData.timeCommitValue && formData.timeCommitPeriod) ? `${formData.timeCommitValue} ${formData.timeCommitPeriod}` : '',
      formData.hoSewaYesNo || '',
      formData.hoSewa || ''
    ];

    // Append the row
    responseSheet.appendRow(rowData);

    console.log('Form data saved successfully');

    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Form submitted successfully',
        timestamp: timestamp
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: error.message 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Main function to get analytics data
 */
function getAnalyticsData() {
  try {
    console.log('Opening spreadsheet...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found');
    
    console.log('Getting sheet:', SHEET_NAME);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' not found`);
    
    console.log('Fetching data...');
    const data = sheet.getDataRange().getValues();
    console.log(`Fetched ${data.length} rows of data`);
    
    if (data.length < 2) {
      throw new Error('Not enough data in the sheet');
    }
    
    // First row is empty, second row contains headers
    const headers = data[1];
    const dataRows = data.slice(2).filter(row => row.some(cell => cell !== ''));
    
    console.log(`Processing ${dataRows.length} data rows`);
    console.log('Headers:', headers);
    
    // Find column indices for important fields
    const columnIndices = {
      state: findColumnIndex(headers, 'State'),
      branch: findColumnIndex(headers, 'Branch Full'),
      date: findColumnIndex(headers, 'Date'),
      month: findColumnIndex(headers, 'Month'),
      year: findColumnIndex(headers, 'Year'),
      programType: findColumnIndex(headers, 'Type of Program'),
      activityTypes: [
        { name: 'Lecture', index: findColumnIndex(headers, 'Lecture') },
        { name: 'Skit', index: findColumnIndex(headers, 'Skit') },
        { name: 'Motivational Songs', index: findColumnIndex(headers, 'Motivational Songs') },
        { name: 'Choreography', index: findColumnIndex(headers, 'Choreography') },
        { name: 'Pledge', index: findColumnIndex(headers, 'Pledge') },
        { name: 'Slogan Writing', index: findColumnIndex(headers, 'Slogan Writing') },
        { name: 'Rally', index: findColumnIndex(headers, 'Rally') },
        { name: 'Nukkad Natak', index: findColumnIndex(headers, 'Nukkad Natak') },
        { name: 'Poster Making', index: findColumnIndex(headers, 'Poster Making') }
      ].filter(activity => activity.index !== -1),
      beneficiaries: {
        men: findColumnIndex(headers, 'Men'),
        women: findColumnIndex(headers, 'Women'),
        students: findColumnIndex(headers, 'Students'),
        teachers: findColumnIndex(headers, 'Teachers'),
        children: findColumnIndex(headers, 'Children')
      },
      media: {
        photos: findColumnIndex(headers, 'DR Photographs'),
        videos: findColumnIndex(headers, 'Videos'),
        press: findColumnIndex(headers, 'Press Coverage'),
        appreciation: findColumnIndex(headers, 'Appreciation Letter')
      },
      campaignName: findColumnIndex(headers, 'Campaign Name')
    };

    // Process the data
    const result = {
      summary: getSummaryStats(dataRows, columnIndices),
      byState: groupByState(dataRows, columnIndices),
      byActivity: getActivityStats(dataRows, columnIndices),
      byProgramType: getProgramTypeStats(dataRows, columnIndices),
      byMonth: getMonthlyStats(dataRows, columnIndices),
      totalBeneficiaries: getTotalBeneficiaries(dataRows, columnIndices)
    };

    return JSON.stringify(result);
  } catch (error) {
    console.error('Error in getAnalyticsData:', error);
    return JSON.stringify({ error: error.toString() });
  }
}

/**
 * Helper function to find column index by header name
 */
function findColumnIndex(headers, name) {
  return headers.findIndex(header => 
    header && header.toString().trim().toLowerCase() === name.toLowerCase()
  );
}

/**
 * Group data by state
 */
function groupByState(data, columns) {
  const result = {};
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    
    const state = String(row[columns.state] || '').trim();
    if (!state) continue;
    
    if (!result[state]) {
      result[state] = {
        count: 0,
        branches: new Set(),
        activities: {}
      };
    }
    
    result[state].count++;
    
    const branch = String(row[columns.branch] || '').trim();
    if (branch) result[state].branches.add(branch);
    
    if (columns.activityTypes && Array.isArray(columns.activityTypes)) {
      columns.activityTypes.forEach(activity => {
        if (activity && activity.index !== undefined) {
          const value = row[activity.index];
          if (value && String(value).toLowerCase() === 'yes') {
            const activityName = activity.name || 'activity';
            result[state].activities[activityName] = (result[state].activities[activityName] || 0) + 1;
          }
        }
      });
    }
  }
  
  Object.keys(result).forEach(state => {
    result[state].branches = Array.from(result[state].branches);
  });
  
  return result;
}

/**
 * Get statistics by activity type
 */
function getActivityStats(data, columns) {
  const result = {};
  
  if (columns.activityTypes && Array.isArray(columns.activityTypes)) {
    columns.activityTypes.forEach(activity => {
      if (activity && activity.name) {
        result[activity.name] = 0;
      }
    });
  }
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    
    if (columns.activityTypes && Array.isArray(columns.activityTypes)) {
      columns.activityTypes.forEach(activity => {
        if (activity && activity.index !== undefined) {
          const value = row[activity.index];
          if (value && String(value).toLowerCase() === 'yes') {
            result[activity.name] = (result[activity.name] || 0) + 1;
          }
        }
      });
    }
  }
  
  return result;
}

/**
 * Get statistics by program type
 */
function getProgramTypeStats(data, columns) {
  const result = {};
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    
    const programType = String(row[columns.programType] || '').trim();
    if (!programType) continue;
    
    result[programType] = (result[programType] || 0) + 1;
  }
  
  return result;
}

/**
 * Get monthly statistics
 */
function getMonthlyStats(data, columns) {
  const result = {};
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    
    const month = String(row[columns.month] || '').trim();
    const year = String(row[columns.year] || '').trim();
    
    if (!month || !year) continue;
    
    const monthYear = `${month} ${year}`;
    result[monthYear] = (result[monthYear] || 0) + 1;
  }
  
  return result;
}

/**
 * Calculate summary statistics
 */
function getSummaryStats(data, columns) {
  const states = new Set();
  const branches = new Set();
  let startDate = null;
  let endDate = null;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    
    const branch = String(row[columns.branch] || '').trim();
    const state = String(row[columns.state] || '').trim();
    const dateStr = row[columns.date];
    let date = null;
    
    if (dateStr) {
      date = new Date(dateStr);
      if (isNaN(date.getTime())) date = null;
    }
    
    if (branch) branches.add(branch);
    if (state) states.add(state);
    
    if (date) {
      if (!startDate || date < startDate) startDate = date;
      if (!endDate || date > endDate) endDate = date;
    }
  }
  
  return {
    totalEvents: data.length,
    totalBranches: branches.size,
    totalStates: states.size,
    startDate: startDate ? startDate.toISOString().split('T')[0] : null,
    endDate: endDate ? endDate.toISOString().split('T')[0] : null
  };
}

/**
 * Calculate total beneficiaries
 */
function getTotalBeneficiaries(data, columns) {
  const result = {
    men: 0,
    women: 0,
    students: 0,
    teachers: 0,
    children: 0,
    total: 0
  };
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || !columns.beneficiaries) continue;
    
    Object.entries(columns.beneficiaries).forEach(([key, colIndex]) => {
      if (colIndex !== -1) {
        const value = parseInt(row[colIndex]) || 0;
        result[key] += value;
        result.total += value;
      }
    });
  }
  
  return result;
}


// ==========================================
// NEW REPORT FORM BACKEND FUNCTIONS (APPENDED)
// ==========================================
var MAIN_DRIVE_FOLDER_ID = '1Z6PKvQbNXo2RJ7DGZsCaWgoO1nzZiPgw';
var FORM_SHEET_ID = '1OvVqqDI5JwTMJA4svMKXUb2JuOMQUaPgH7Un-UMum14';
var FORM_SHEET_TAB_NAME = 'BODH Responses';

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function generateFolderStructure(formData) {
  try {
    var state = formData.state; var branch = formData.branch; var dateStr = formData.date;
    if (!state || !branch || !dateStr) throw new Error('State, Branch, and Date are required.');

    var d = new Date(dateStr + 'T00:00:00'); var year = d.getFullYear().toString();
    var fullMonths = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var shortMonths = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    var monthName = fullMonths[d.getMonth()];
    var day = d.getDate();
    var dayStr = (day < 10 ? '0' : '') + day;
    var formattedDateFolder = dayStr + '-' + shortMonths[d.getMonth()] + '-' + year;

    var mainDrive = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);
    var eventsFolder = getOrCreateFolder(mainDrive, '2 - Events Reports - Year Wise');
    var stateFolder = getOrCreateFolder(eventsFolder, state);
    var branchFolder = getOrCreateFolder(stateFolder, branch);
    var yearFolder = getOrCreateFolder(branchFolder, year);
    var monthFolder = getOrCreateFolder(yearFolder, monthName);
    var dateFolder = getOrCreateFolder(monthFolder, formattedDateFolder);

    var rawFolder = getOrCreateFolder(dateFolder, 'Raw');
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
    return { rawFolderUrl: rawFolder.getUrl(), dateFolderUrl: dateFolder.getUrl() };
  } catch (e) {
    return { error: e.message };
  }
}

function submitReport(formData) {
  try {
    var ss = SpreadsheetApp.openById(FORM_SHEET_ID);
    var sheet = ss.getSheetByName(FORM_SHEET_TAB_NAME);
    if (!sheet) throw new Error('Sheet tab "' + FORM_SHEET_TAB_NAME + '" not found.');

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var autoEmail = Session.getActiveUser().getEmail() || '';
    if(!autoEmail) autoEmail = Session.getEffectiveUser().getEmail() || '';

    var sParts = formData.timeFrom ? formData.timeFrom.split(':') : null;
    var eParts = formData.timeTo ? formData.timeTo.split(':') : null;
    var sTime = ''; var eTime = '';
    if(sParts) { var h = parseInt(sParts[0],10); var ampm = h>=12?'pm':'am'; h = h%12; h = h?h:12; sTime = h+':'+sParts[1]+' '+ampm; }
    if(eParts) { var h = parseInt(eParts[0],10); var ampm = h>=12?'pm':'am'; h = h%12; h = h?h:12; eTime = h+':'+eParts[1]+' '+ampm; }
    var timeString = (sTime && eTime) ? (sTime + ' to ' + eTime) : (sTime || eTime);

    var rowData = [];
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toString().trim();
      if (h === 'Timestamp') rowData.push(new Date());
      else if (h === 'Email Address') rowData.push(autoEmail);
      else if (h.indexOf('कार्यकारी टीम') > -1 || h.indexOf('Executing Team') > -1) rowData.push(formData.executingTeam || '');
      else if (h === 'राज्य | State') rowData.push(formData.state || '');
      else if (h.indexOf('शाखा | Branch') > -1) {
        if (h.indexOf('(' + formData.state + ')') > -1) rowData.push(formData.branch || '');
        else rowData.push('');
      }
      else if (h.indexOf('स्थान | Venue') > -1) rowData.push(formData.venue || '');
      else if (h.indexOf('विषय / थीम') > -1) rowData.push(formData.theme || '');
      else if (h.indexOf('दिनांक / Date') > -1) rowData.push(formData.date || '');
      else if (h.indexOf('समय / Timings') > -1) rowData.push(timeString || '');
      else if (h.indexOf('पुरुष / Male') > -1) rowData.push(formData.male || '');
      else if (h.indexOf('महिलाएं / Female') > -1) rowData.push(formData.female || '');
      else if (h.indexOf('बच्चे') > -1) rowData.push(formData.children || '');
      else if (h.indexOf('आयोजक / Organizer') > -1) rowData.push(formData.organizer || '');
      else if (h.indexOf('भागीदार / प्रायोजक') > -1) rowData.push(formData.partner || '');
      else if (h.indexOf('अतिथि / वक्ता') > -1) rowData.push(''); 
      else if (h === 'नाम | Name') rowData.push(formData.guestName || '');
      else if (h.indexOf('पद / Designation') > -1) rowData.push(formData.guestDesignation || '');
      else if (h.indexOf('सहयोगी संगठन') > -1) rowData.push(formData.guestOrg || '');
      else if (h.indexOf('संचालित गतिविधियाँ') > -1) rowData.push(formData.activities || '');
      else if (h.indexOf('कोई विशेष जानकारी') > -1) rowData.push(formData.observation || '');
      else if (h.indexOf('फोटोज़') > -1) rowData.push(formData.photos || '');
      else if (h.indexOf('वीडियो') > -1) rowData.push(formData.videos || '');
      else if (h.indexOf('प्रैस विज्ञप्ति') > -1) rowData.push(formData.pressRelease || '');
      else if (h.indexOf('प्रशंसा पत्र') > -1) rowData.push(formData.appreciation || '');
      else if (h.indexOf('प्रचार माध्यम') > -1) rowData.push(formData.publicity || '');
      else if (h.indexOf('ज्ञान दीक्षा') > -1) rowData.push(formData.initiatedCount || '');
      else if (h.indexOf('फीडबैक') > -1) rowData.push(formData.feedback || '');
      else if (h === 'Link') rowData.push(formData.folderLink || '');
      else rowData.push('');
    }
    sheet.appendRow(rowData);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

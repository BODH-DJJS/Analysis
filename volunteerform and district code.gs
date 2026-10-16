// ==============================================================================
// BODH DASHBOARD, VOLUNTEER FORM & DISTRICT INFORMATION BACKEND (Code.gs)
// ==============================================================================

// Configuration
const SPREADSHEET_ID = '1OvVqqDI5JwTMJA4svMKXUb2JuOMQUaPgH7Un-UMum14';
const SHEET_NAME = 'Master data';
const RESPONSES_SHEET_NAME = 'POC Team Leaders';
const DISTRICT_RESPONSES_SHEET_NAME = 'District Information';

/**
 * Web app entry point for GET requests (Analytics Data & Month Folder Resolution)
 */
function doGet(e) {
  try {
    // Action: Resolve month folder -> Raw folder
    if (e && e.parameter && e.parameter.action === 'raw') {
      const monthUrl = e.parameter.link || '';
      const dateStr = e.parameter.date || '';
      let targetUrl = monthUrl;
      try {
        if (monthUrl && dateStr) {
          const match = monthUrl.match(/folders\/([a-zA-Z0-9_-]+)/);
          if (match && match[1]) {
            const monthFolder = DriveApp.getFolderById(match[1]);
            const d = new Date(dateStr + 'T00:00:00');
            const shortMonths = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            const day = d.getDate();
            const dayStr = (day < 10 ? '0' : '') + day;
            const folderBase = dayStr + '-' + shortMonths[d.getMonth()] + '-' + d.getFullYear();
            
            const dateFolders = monthFolder.getFoldersByName(folderBase);
            if (dateFolders.hasNext()) {
              const df = dateFolders.next();
              const rawFolders = df.getFoldersByName('Raw');
              if (rawFolders.hasNext()) {
                targetUrl = rawFolders.next().getUrl();
              } else {
                targetUrl = df.getUrl();
              }
            }
          }
        }
      } catch (err) {
        console.error('Error resolving raw folder:', err);
      }
      return ContentService.createTextOutput(JSON.stringify({ url: targetUrl }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default GET: Return Analytics Data for Dashboard
    const data = getAnalyticsData();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error('Error in doGet:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ 
        error: error.message,
        stack: error.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Web app entry point for POST requests (Volunteer Form & District Information)
 */
function doPost(e) {
  try {
    let formData;
    if (e.postData && e.postData.contents) {
      formData = JSON.parse(e.postData.contents);
    } else {
      throw new Error('No form data received');
    }

    console.log('Received form data:', JSON.stringify(formData));

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found');

    // -------------------------------------------------------------
    // 1. DISTRICT INFORMATION FORM SUBMISSION
    // -------------------------------------------------------------
    if (formData.formType === 'districtInfo') {
      let districtSheet = ss.getSheetByName(DISTRICT_RESPONSES_SHEET_NAME);
      if (!districtSheet) {
        districtSheet = ss.insertSheet(DISTRICT_RESPONSES_SHEET_NAME);
        const dHeaders = [
          'Timestamp',
          'User Email',
          'Branch Name',
          'States Covered',
          'Total States Covered',
          'Districts Covered Breakdown',
          'Total Districts Covered'
        ];
        districtSheet.getRange(1, 1, 1, dHeaders.length).setValues([dHeaders]);
        const dHeaderRange = districtSheet.getRange(1, 1, 1, dHeaders.length);
        dHeaderRange.setFontWeight('bold');
        dHeaderRange.setBackground('#0f766e');
        dHeaderRange.setFontColor('#ffffff');
        dHeaderRange.setHorizontalAlignment('center');
        for (let i = 1; i <= dHeaders.length; i++) {
          districtSheet.autoResizeColumn(i);
        }
        districtSheet.setFrozenRows(1);
      }

      const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
      const dRowData = [
        timestamp,
        formData.userEmail || '',
        formData.branchName || '',
        Array.isArray(formData.statesCovered) ? formData.statesCovered.join(', ') : (formData.statesCovered || ''),
        formData.totalStates || 0,
        formData.districtsBreakdown || '',
        formData.totalDistricts || 0
      ];

      districtSheet.appendRow(dRowData);
      console.log('District data saved successfully in tab: ' + DISTRICT_RESPONSES_SHEET_NAME);

      return ContentService
        .createTextOutput(JSON.stringify({ 
          status: 'success', 
          message: 'District Information submitted successfully',
          timestamp: timestamp
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // 2. VOLUNTEER REGISTRATION FORM SUBMISSION
    // -------------------------------------------------------------
    let responseSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!responseSheet) {
      responseSheet = ss.insertSheet(RESPONSES_SHEET_NAME);
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
        'Name - Bodh Branch Head (Brothers/ Sisters) - BBBH/ BSBH',
        'Contact - Bodh Branch Head (Brothers/ Sisters) - BBBH/ BSBH',
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
      const headerRange = responseSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#1a237e');
      headerRange.setFontColor('#ffffff');
      headerRange.setHorizontalAlignment('center');
      for (let i = 1; i <= headers.length; i++) {
        responseSheet.autoResizeColumn(i);
      }
      responseSheet.setFrozenRows(1);
    }

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
      '', // Brother / Sister blank to preserve original column indexing
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

    responseSheet.appendRow(rowData);
    console.log('Volunteer Form data saved successfully in tab: ' + RESPONSES_SHEET_NAME);

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
 * Main function to get analytics data for Dashboard
 */
function getAnalyticsData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found');
    
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' not found`);
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('Not enough data in the sheet');
    }
    
    // First row is empty, second row contains headers
    const headers = data[1];
    const dataRows = data.slice(2).filter(row => row.some(cell => cell !== ''));
    
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

function findColumnIndex(headers, name) {
  return headers.findIndex(header => 
    header && header.toString().trim().toLowerCase() === name.toLowerCase()
  );
}

function groupByState(data, columns) {
  const result = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 2) continue;
    const state = String(row[columns.state] || '').trim();
    if (!state) continue;
    if (!result[state]) {
      result[state] = { count: 0, branches: new Set(), activities: {} };
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

function getActivityStats(data, columns) {
  const result = {};
  if (columns.activityTypes && Array.isArray(columns.activityTypes)) {
    columns.activityTypes.forEach(activity => {
      if (activity && activity.name) result[activity.name] = 0;
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

function getTotalBeneficiaries(data, columns) {
  const result = { men: 0, women: 0, students: 0, teachers: 0, children: 0, total: 0 };
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

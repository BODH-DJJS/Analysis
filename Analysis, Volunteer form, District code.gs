// Configuration
const SPREADSHEET_ID = '1OvVqqDI5JwTMJA4svMKXUb2JuOMQUaPgH7Un-UMum14';
const SHEET_NAME = 'Master data';
const RESPONSES_SHEET_NAME = 'POC Team Leaders';
const DISTRICT_RESPONSES_SHEET_NAME = 'District Information';
const REMARKS_SHEET_NAME = 'Branch Remarks';

/**
 * Helper test function you can run directly inside the Google Apps Script Editor!
 * Select 'testCheckBranchRoles' from the function dropdown and click Run.
 * Check View -> Logs to see exactly what was found for Kota!
 */
function testCheckBranchRoles() {
  const result = getBranchRolesCount('Kota');
  Logger.log('TEST RESULT FOR KOTA: ' + JSON.stringify(result));
  return result;
}

/**
 * Returns comprehensive tracker data for ALL branches across both:
 * 1. POC Team Leaders (Roles BBBH, BBBM, BSBH, BSBM)
 * 2. District Information (whether filled or not)
 */
function getAllBranchTrackerData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const branchMap = {}; // normalizedKey -> { branchName, state, districtCount, BBBH, BBBM, BSBH, BSBM, districts: [] }
  const districtSubmissions = []; // [ { branch: '...', states: '...', breakdown: '...', totalDistricts: 0 } ]

  // 1. Process District Information Sheet
  try {
    const distSheet = ss.getSheetByName(DISTRICT_RESPONSES_SHEET_NAME);
    if (distSheet) {
      const dData = distSheet.getDataRange().getValues();
      if (dData.length > 1) {
        // find columns
        const dHeaders = dData[0].map(h => String(h || '').trim().toLowerCase());
        let bCol = dHeaders.findIndex(h => h.includes('branch'));
        let breakdownCol = dHeaders.findIndex(h => h.includes('breakdown') || h.includes('district'));
        let statesCol = dHeaders.findIndex(h => h.includes('states covered'));
        let totalDistsCol = dHeaders.findIndex(h => h.includes('total districts'));

        if (bCol === -1) bCol = 2; // Default Col C
        if (breakdownCol === -1) breakdownCol = 5; // Default Col F
        if (statesCol === -1) statesCol = 3; // Default Col D
        if (totalDistsCol === -1) totalDistsCol = 6; // Default Col G

        for (let i = 1; i < dData.length; i++) {
          const bName = String(dData[i][bCol] || '').trim();
          if (!bName) continue;
          const key = bName.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!branchMap[key]) {
            branchMap[key] = { branchName: bName, districtCount: 0, BBBH: 0, BBBM: 0, BSBH: 0, BSBM: 0, districts: [] };
          }
          branchMap[key].districtCount++;

          const breakdownStr = String(dData[i][breakdownCol] || '').trim();
          const statesStr = String(dData[i][statesCol] || '').trim();
          const totalD = parseInt(dData[i][totalDistsCol]) || 0;

          if (breakdownStr) {
            branchMap[key].breakdown = breakdownStr;
            districtSubmissions.push({
              branch: bName,
              states: statesStr,
              breakdown: breakdownStr,
              totalDistricts: totalD
            });
          }
        }
      }
    }
  } catch(e) {
    console.error('Error reading District Information:', e);
  }

  // 2. Process POC Team Leaders Sheet
  try {
    let pocSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!pocSheet) {
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        const sName = s.getName().toLowerCase();
        if (sName.includes('poc') || sName.includes('leader') || sName.includes('volunteer')) {
          pocSheet = s;
          break;
        }
      }
    }

    if (pocSheet) {
      const pData = pocSheet.getDataRange().getValues();
      if (pData.length > 1) {
        let headerRowIdx = 0;
        let branchCol = -1;
        let roleCol = -1;
        let nameCol = -1;
        let contactCol = -1;

        for (let r = 0; r < Math.min(5, pData.length); r++) {
          const rowHeaders = pData[r].map(h => String(h || '').trim().toLowerCase());
          let bIdx = rowHeaders.findIndex(h => h === 'branch name' || h === 'branch');
          if (bIdx === -1) bIdx = rowHeaders.findIndex(h => h.includes('branch') && !h.includes('ar'));
          const rIdx = rowHeaders.findIndex(h => h.includes('role'));
          const nIdx = rowHeaders.findIndex(h => h === 'full name' || h === 'name');
          const cIdx = rowHeaders.findIndex(h => h.includes('contact') || h.includes('phone') || h.includes('mobile'));
          if (bIdx !== -1 && rIdx !== -1) {
            headerRowIdx = r;
            branchCol = bIdx;
            roleCol = rIdx;
            if (nIdx !== -1) nameCol = nIdx;
            if (cIdx !== -1) contactCol = cIdx;
            break;
          }
        }

        if (branchCol === -1) branchCol = 8;
        if (roleCol === -1) roleCol = 12;
        if (nameCol === -1) nameCol = 1;
        if (contactCol === -1) contactCol = 2;

        for (let i = headerRowIdx + 1; i < pData.length; i++) {
          const bName = String(pData[i][branchCol] || '').trim();
          if (!bName) continue;
          const key = bName.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (!branchMap[key]) {
            branchMap[key] = { branchName: bName, districtCount: 0, BBBH: 0, BBBM: 0, BSBH: 0, BSBM: 0, members: [] };
          }
          if (!branchMap[key].members) {
            branchMap[key].members = [];
          }

          const directRole = String(pData[i][roleCol] || '').trim().toUpperCase();
          let matchedRole = '';
          if (directRole.includes('BBBH') || directRole.includes('BSBH') || directRole.includes('BBBM') || directRole.includes('BSBM')) {
            matchedRole = directRole;
          } else {
            const rowStr = pData[i].join(' ').toUpperCase();
            if (rowStr.includes('BBBH')) matchedRole = 'BBBH';
            else if (rowStr.includes('BSBH')) matchedRole = 'BSBH';
            else if (rowStr.includes('BBBM')) matchedRole = 'BBBM';
            else if (rowStr.includes('BSBM')) matchedRole = 'BSBM';
          }

          let standardRole = '';
          if (matchedRole.includes('BBBH')) {
            branchMap[key].BBBH++;
            standardRole = 'BBBH';
          } else if (matchedRole.includes('BBBM')) {
            branchMap[key].BBBM++;
            standardRole = 'BBBM';
          } else if (matchedRole.includes('BSBH')) {
            branchMap[key].BSBH++;
            standardRole = 'BSBH';
          } else if (matchedRole.includes('BSBM')) {
            branchMap[key].BSBM++;
            standardRole = 'BSBM';
          }

          const personName = String(pData[i][nameCol] || '').trim();
          const personContact = String(pData[i][contactCol] || '').trim();
          if (personName || standardRole) {
            branchMap[key].members.push({
              name: personName || 'N/A',
              role: standardRole || directRole || 'Member',
              contact: personContact || '',
              row: i + 1
            });
          }
        }
      }
    }
  } catch(e) {
    console.error('Error reading POC Team Leaders:', e);
  }

  // 3. Process Branch Remarks Sheet
  try {
    let remSheet = ss.getSheetByName(REMARKS_SHEET_NAME);
    if (remSheet) {
      const remData = remSheet.getDataRange().getValues();
      if (remData.length > 1) {
        for (let i = 1; i < remData.length; i++) {
          const bName = String(remData[i][0] || '').trim();
          if (!bName) continue;
          const key = bName.toLowerCase().replace(/[^a-z0-9]/g, '');
          if (branchMap[key]) {
            branchMap[key].remark = String(remData[i][1] || '').trim();
            branchMap[key].remarkUpdated = String(remData[i][2] || '').trim();
          }
        }
      }
    }
  } catch(e) {
    console.error('Error reading Branch Remarks:', e);
  }

  return { status: 'success', trackerData: branchMap, districtSubmissions: districtSubmissions };
}

function saveBranchRemark(branchName, remarkText) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let remSheet = ss.getSheetByName(REMARKS_SHEET_NAME);
    if (!remSheet) {
      remSheet = ss.insertSheet(REMARKS_SHEET_NAME);
      const headers = ['Branch Name', 'Remark', 'Last Updated'];
      remSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      const hRange = remSheet.getRange(1, 1, 1, headers.length);
      hRange.setFontWeight('bold');
      hRange.setBackground('#0f172a');
      hRange.setFontColor('#ffffff');
      remSheet.setFrozenRows(1);
    }

    const data = remSheet.getDataRange().getValues();
    const bNorm = branchName.toLowerCase().replace(/[^a-z0-9]/g, '');
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      const rowBNorm = String(data[i][0] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      if (rowBNorm === bNorm) {
        foundRow = i + 1; // 1-based row index
        break;
      }
    }

    const timestamp = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

    if (foundRow !== -1) {
      remSheet.getRange(foundRow, 2).setValue(remarkText);
      remSheet.getRange(foundRow, 3).setValue(timestamp);
    } else {
      remSheet.appendRow([branchName, remarkText, timestamp]);
    }

    return { status: 'success', branch: branchName, remark: remarkText, updated: timestamp };
  } catch (err) {
    console.error('Error saving remark:', err);
    return { status: 'error', message: err.message };
  }
}

/**
 * Updates a member's role directly in the POC Team Leaders sheet tab.
 * Can find the row by exact rowIndex (1-based sheet row) or by matching (branchName, memberName).
 */
function updateMemberRoleInSheet(branchName, memberName, newRole, targetRow) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
    if (!sheet) {
      const allSheets = ss.getSheets();
      for (let s of allSheets) {
        const sName = s.getName().toLowerCase();
        if (sName.includes('poc') || sName.includes('leader') || sName.includes('volunteer')) {
          sheet = s;
          break;
        }
      }
    }

    if (!sheet) {
      return { status: 'error', message: 'POC Team Leaders sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { status: 'error', message: 'No records in sheet' };
    }

    let headerRowIdx = 0;
    let branchCol = -1;
    let roleCol = -1;
    let nameCol = -1;

    for (let r = 0; r < Math.min(5, data.length); r++) {
      const rowHeaders = data[r].map(h => String(h || '').trim().toLowerCase());
      let bIdx = rowHeaders.findIndex(h => h === 'branch name' || h === 'branch');
      if (bIdx === -1) bIdx = rowHeaders.findIndex(h => h.includes('branch') && !h.includes('ar'));
      const rIdx = rowHeaders.findIndex(h => h.includes('role'));
      const nIdx = rowHeaders.findIndex(h => h === 'full name' || h === 'name');
      if (bIdx !== -1 && rIdx !== -1) {
        headerRowIdx = r;
        branchCol = bIdx;
        roleCol = rIdx;
        if (nIdx !== -1) nameCol = nIdx;
        break;
      }
    }

    if (branchCol === -1) branchCol = 8;
    if (roleCol === -1) roleCol = 12;
    if (nameCol === -1) nameCol = 1;

    let targetSheetRow = -1;
    const tRow = parseInt(targetRow, 10);

    // If targetRow is provided and valid, verify it matches
    if (!isNaN(tRow) && tRow > headerRowIdx + 1 && tRow <= data.length) {
      const rIdx = tRow - 1;
      const rBranchNorm = String(data[rIdx][branchCol] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      const bNorm = String(branchName || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      const rNameNorm = String(data[rIdx][nameCol] || '').trim().toLowerCase().replace(/\s+/g, '');
      const mNameNorm = String(memberName || '').trim().toLowerCase().replace(/\s+/g, '');

      if ((!bNorm || rBranchNorm === bNorm) && (!mNameNorm || rNameNorm === mNameNorm)) {
        targetSheetRow = tRow;
      }
    }

    // If not found by row index, search by branch & name
    if (targetSheetRow === -1 && branchName && memberName) {
      const bNorm = branchName.toLowerCase().replace(/[^a-z0-9]/g, '');
      const mNorm = memberName.toLowerCase().replace(/\s+/g, '');

      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const rowBranchNorm = String(data[i][branchCol] || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        const rowNameNorm = String(data[i][nameCol] || '').trim().toLowerCase().replace(/\s+/g, '');
        if (rowBranchNorm === bNorm && rowNameNorm === mNorm) {
          targetSheetRow = i + 1; // 1-based sheet row
          break;
        }
      }
    }

    if (targetSheetRow === -1) {
      return { status: 'error', message: 'Member record not found in sheet for branch: ' + branchName + ', member: ' + memberName };
    }

    // Update the role in the sheet (roleCol is 0-based, so +1 for Sheets 1-based column)
    const roleCell = sheet.getRange(targetSheetRow, roleCol + 1);
    const oldRole = String(roleCell.getValue() || '');
    roleCell.setValue(newRole);

    return {
      status: 'success',
      row: targetSheetRow,
      branch: branchName,
      name: memberName,
      oldRole: oldRole,
      newRole: newRole
    };
  } catch (err) {
    console.error('Error updating member role:', err);
    return { status: 'error', message: err.message };
  }
}

function getBranchRolesCount(branchName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const counts = { BBBH: 0, BBBM: 0, BSBH: 0, BSBM: 0 };
  const members = [];
  let debugInfo = { branchName: branchName, rowCount: 0, branchCol: -1, roleCol: -1, matchedRows: 0 };

  // Look for the responses sheet tab
  let sheet = ss.getSheetByName(RESPONSES_SHEET_NAME);
  if (!sheet) {
    const allSheets = ss.getSheets();
    for (let s of allSheets) {
      const sName = s.getName().toLowerCase();
      if (sName.includes('poc') || sName.includes('leader') || sName.includes('volunteer')) {
        sheet = s;
        break;
      }
    }
  }

  if (sheet && branchName) {
    debugInfo.sheetName = sheet.getName();
    const data = sheet.getDataRange().getValues();
    debugInfo.rowCount = data.length;
    if (data.length > 0) {
      let headerRowIdx = 0;
      let branchCol = -1;
      let roleCol = -1;
      let nameCol = -1;
      let contactCol = -1;

      for (let r = 0; r < Math.min(5, data.length); r++) {
        const rowHeaders = data[r].map(h => String(h || '').trim().toLowerCase());
        let bIdx = rowHeaders.findIndex(h => h === 'branch name' || h === 'branch');
        if (bIdx === -1) bIdx = rowHeaders.findIndex(h => h.includes('branch') && !h.includes('ar'));
        const rIdx = rowHeaders.findIndex(h => h.includes('role'));
        const nIdx = rowHeaders.findIndex(h => h === 'full name' || h === 'name');
        const cIdx = rowHeaders.findIndex(h => h.includes('contact') || h.includes('phone') || h.includes('mobile'));
        if (bIdx !== -1 && rIdx !== -1) {
          headerRowIdx = r;
          branchCol = bIdx;
          roleCol = rIdx;
          if (nIdx !== -1) nameCol = nIdx;
          if (cIdx !== -1) contactCol = cIdx;
          break;
        }
      }

      if (branchCol === -1) branchCol = 8;
      if (roleCol === -1) roleCol = 12;
      if (nameCol === -1) nameCol = 1;
      if (contactCol === -1) contactCol = 2;

      debugInfo.headerRowIdx = headerRowIdx;
      debugInfo.branchCol = branchCol;
      debugInfo.roleCol = roleCol;

      const bNorm = branchName.toLowerCase().replace(/[^a-z0-9]/g, '');
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const rowBranch = String(data[i][branchCol] || '').trim();
        const rowBranchNorm = rowBranch.toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (rowBranchNorm === bNorm) {
          debugInfo.matchedRows++;
          const directRole = String(data[i][roleCol] || '').trim().toUpperCase();
          let matchedRole = '';
          if (directRole.includes('BBBH') || directRole.includes('BSBH') || directRole.includes('BBBM') || directRole.includes('BSBM')) {
            matchedRole = directRole;
          } else {
            const rowStr = data[i].join(' ').toUpperCase();
            if (rowStr.includes('BBBH')) matchedRole = 'BBBH';
            else if (rowStr.includes('BSBH')) matchedRole = 'BSBH';
            else if (rowStr.includes('BBBM')) matchedRole = 'BBBM';
            else if (rowStr.includes('BSBM')) matchedRole = 'BSBM';
          }

          let standardRole = '';
          if (matchedRole.includes('BBBH')) {
            counts.BBBH++;
            standardRole = 'BBBH';
          } else if (matchedRole.includes('BBBM')) {
            counts.BBBM++;
            standardRole = 'BBBM';
          } else if (matchedRole.includes('BSBH')) {
            counts.BSBH++;
            standardRole = 'BSBH';
          } else if (matchedRole.includes('BSBM')) {
            counts.BSBM++;
            standardRole = 'BSBM';
          }

          const pName = String(data[i][nameCol] || '').trim();
          const pContact = String(data[i][contactCol] || '').trim();
          if (pName || standardRole) {
            members.push({
              name: pName || 'N/A',
              role: standardRole || directRole || 'Member',
              contact: pContact || '',
              row: i + 1
            });
          }
        }
      }
    }
  }
  return { status: 'success', branch: branchName, counts: counts, members: members, debug: debugInfo };
}

/**
 * Web app entry point for GET requests
 */
function doGet(e) {
  try {
    // ==========================================
    // Route to the HTML form if ?p=form
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
    // ACTION: Check branch roles count for POC Team Leaders
    // ==========================================
    if (e && e.parameter && e.parameter.action === 'checkBranchRoles') {
      const branchName = (e.parameter.branch || '').trim();
      const callback = (e.parameter.callback || '').trim();
      const resultObj = getBranchRolesCount(branchName);

      if (callback) {
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(resultObj) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.action === 'getAllBranchTracker') {
      const callback = (e.parameter.callback || '').trim();
      const resultObj = getAllBranchTrackerData();
      if (callback) {
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(resultObj) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.action === 'saveBranchRemark') {
      const branchName = (e.parameter.branch || '').trim();
      const remarkText = (e.parameter.remark || '').trim();
      const callback = (e.parameter.callback || '').trim();
      const resultObj = saveBranchRemark(branchName, remarkText);
      if (callback) {
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(resultObj) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (e && e.parameter && e.parameter.action === 'updateMemberRole') {
      const branchName = (e.parameter.branch || '').trim();
      const memberName = (e.parameter.name || '').trim();
      const newRole = (e.parameter.role || '').trim();
      const targetRow = (e.parameter.row || '').trim();
      const callback = (e.parameter.callback || '').trim();
      const resultObj = updateMemberRoleInSheet(branchName, memberName, newRole, targetRow);
      if (callback) {
        return ContentService
          .createTextOutput(callback + '(' + JSON.stringify(resultObj) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // ==========================================

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

    // Support checkBranchRoles via POST (bypasses browser CORS & redirect issues completely)
    if (formData && formData.action === 'checkBranchRoles') {
      const branchName = (formData.branch || '').trim();
      const resultObj = getBranchRolesCount(branchName);
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Support getAllBranchTracker via POST
    if (formData && formData.action === 'getAllBranchTracker') {
      const resultObj = getAllBranchTrackerData();
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Support saveBranchRemark via POST
    if (formData && formData.action === 'saveBranchRemark') {
      const branchName = (formData.branch || '').trim();
      const remarkText = (formData.remark || '').trim();
      const resultObj = saveBranchRemark(branchName, remarkText);
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Support updateMemberRole via POST
    if (formData && formData.action === 'updateMemberRole') {
      const branchName = (formData.branch || '').trim();
      const memberName = (formData.name || '').trim();
      const newRole = (formData.role || '').trim();
      const targetRow = (formData.row || '').toString().trim();
      const resultObj = updateMemberRoleInSheet(branchName, memberName, newRole, targetRow);
      return ContentService
        .createTextOutput(JSON.stringify(resultObj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    console.log('Received form data:', JSON.stringify(formData));

    // Open the spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found');

    // Check if this is a District Information form submission
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
        dHeaderRange.setBackground('#0f766e'); // Teal background
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
      console.log('District data saved successfully');

      return ContentService
        .createTextOutput(JSON.stringify({ 
          status: 'success', 
          message: 'District Information submitted successfully',
          timestamp: timestamp
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: Volunteer POC Team Leaders submission
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
    var selectedEditedFolder = getOrCreateFolder(dateFolder, 'Selected (Edited)');
    getOrCreateFolder(selectedEditedFolder, 'Edited Thumbnail');

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

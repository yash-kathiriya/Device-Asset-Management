// ═══════════════════════════════════════════════════════════════════════
// DEVICE ASSET MANAGEMENT - APPS SCRIPT (UPDATED)
// Search by Emp ID across ALL sheets - no Device Type needed for fetch/delete
// 3 Separate Sheets: PC Records | Laptop Records | Store Records
// ═══════════════════════════════════════════════════════════════════════

// ─── HEADERS per sheet ───
const HEADERS_BASE = ['Serial No.', 'Plant Name', 'Emp ID', 'Emp Name', 'Department', 'Location', 'Device Type'];

const PC_HEADERS = HEADERS_BASE.concat([
  'Device Serial No.', 'Device Brand', 'Display Model No.', 'Display Serial No.',
  'Keyboard Brand', 'Keyboard Serial No.', 'Mouse Brand', 'Mouse Serial No.',
  'Date', 'Time'
]);

const LAPTOP_HEADERS = HEADERS_BASE.concat([
  'Laptop Model No.', 'Laptop Serial No.',
  'Date', 'Time'
]);

const STORE_HEADERS = HEADERS_BASE.concat([
  'Store Model No.', 'Store Serial No.', 'Store Material Type',
  'Store PR No.', 'Store PO No.',
  'Date', 'Time'
]);

function getHeadersForType(dtype) {
  if (dtype === 'Laptop') return LAPTOP_HEADERS;
  if (dtype === 'Store')  return STORE_HEADERS;
  return PC_HEADERS;
}

// ─── GET OR CREATE SHEET ───
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    const hr = sheet.getRange(1, 1, 1, headers.length);
    if (sheetName === 'PC Records')     hr.setBackground('#1565C0');
    if (sheetName === 'Laptop Records') hr.setBackground('#1b5e20');
    if (sheetName === 'Store Records')  hr.setBackground('#4a148c');
    hr.setFontColor('#FFFFFF');
    hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 160);
  }
  return sheet;
}

function getSheetForType(dtype) {
  return getOrCreateSheet(dtype + ' Records', getHeadersForType(dtype));
}

// ─── FIND EMP IN ALL SHEETS ───
// Returns {sheet, headers, rowNum, dtype, rowData} or null
function findEmpIdInAllSheets(empId) {
  const empIdLower = String(empId).trim().toLowerCase();
  const types = ['PC', 'Laptop', 'Store'];

  for (const dtype of types) {
    const sheet = getSheetForType(dtype);
    const headers = getHeadersForType(dtype);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    const empCol = headers.indexOf('Emp ID') + 1; // 1-based
    const empData = sheet.getRange(2, empCol, lastRow - 1, 1).getValues();

    for (let i = 0; i < empData.length; i++) {
      if (String(empData[i][0]).trim().toLowerCase() === empIdLower) {
        const rowNum = i + 2; // 1-based, skip header
        const rowData = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
        const obj = {};
        headers.forEach((h, idx) => obj[h] = rowData[idx]);
        return { sheet, headers, rowNum, dtype, rowData: obj };
      }
    }
  }
  return null;
}

// ─── MAIN HANDLER ───
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.__action || 'insert';

    // ── GET NEXT SERIAL (auto-increment per device type) ──
    if (action === 'get_next_serial') {
      const dtype = data['Device Type'];
      const sheet = getSheetForType(dtype);
      const count = Math.max(0, sheet.getLastRow() - 1);
      const prefix = dtype === 'PC' ? 'PC' : dtype === 'Laptop' ? 'LT' : 'ST';
      const next = prefix + '-' + String(count + 1).padStart(4, '0');
      return respond({ status: 'success', next_serial: next });
    }

    // ── INSERT ──
    if (action === 'insert') {
      const dtype = data['Device Type'];
      const headers = getHeadersForType(dtype);
      const sheet = getSheetForType(dtype);
      const row = headers.map(h => data[h] !== undefined ? data[h] : '');
      sheet.appendRow(row);
      return respond({ status: 'success', rows: sheet.getLastRow() - 1 });
    }

    // ── FETCH (search ALL sheets by Emp ID only) ──
    if (action === 'fetch') {
      const empId = data['Emp ID'];
      const found = findEmpIdInAllSheets(empId);
      if (!found) return respond({ status: 'error', message: 'Record not found for Emp ID: ' + empId });
      return respond({ status: 'success', data: found.rowData, dtype: found.dtype });
    }

    // ── UPDATE (search ALL sheets, update in same or move to new) ──
    if (action === 'update') {
      const empId    = data['Emp ID'];
      const newDtype = data['Device Type'];
      const found    = findEmpIdInAllSheets(empId);
      const newHeaders = getHeadersForType(newDtype);
      const newSheet   = getSheetForType(newDtype);
      const newRow     = newHeaders.map(h => data[h] !== undefined ? data[h] : '');

      if (!found) {
        // Not found - insert as new
        newSheet.appendRow(newRow);
        return respond({ status: 'success', message: 'Not found - inserted as new record' });
      }

      if (found.dtype === newDtype) {
        // Same sheet - update in place
        found.sheet.getRange(found.rowNum, 1, 1, newHeaders.length).setValues([newRow]);
      } else {
        // Device type changed - delete from old, insert in new
        found.sheet.deleteRow(found.rowNum);
        newSheet.appendRow(newRow);
      }
      return respond({ status: 'success', message: 'Record updated successfully' });
    }

    // ── DELETE (search ALL sheets by Emp ID only) ──
    if (action === 'delete') {
      const empId = data['Emp ID'];
      const found = findEmpIdInAllSheets(empId);
      if (!found) return respond({ status: 'error', message: 'Record not found for Emp ID: ' + empId });
      found.sheet.deleteRow(found.rowNum);
      return respond({ status: 'success', rows: found.sheet.getLastRow() - 1, message: 'Deleted from ' + found.dtype + ' Records' });
    }


    // ── FETCH BY ROW RANGE (Bulk Barcode maate) ──
    if (action === 'fetchByRowRange') {
      const fromRow = parseInt(data.fromRow) || 1;
      const toRow   = parseInt(data.toRow)   || 1;
      const dtype   = data['Device Type']    || 'PC';
      const sheet   = getSheetForType(dtype);
      const headers = getHeadersForType(dtype);
      const lastRow = sheet.getLastRow();
      const allRecords = [];

      if (lastRow > 1) {
        // fromRow/toRow are 1-based data rows (1 = first row after header)
        const sheetStart = Math.min(Math.max(fromRow, 1) + 1, lastRow); // +1 skip header
        const sheetEnd   = Math.min(toRow + 1, lastRow);
        if (sheetStart <= sheetEnd) {
          const numRows = sheetEnd - sheetStart + 1;
          const values  = sheet.getRange(sheetStart, 1, numRows, headers.length).getValues();
          values.forEach(row => {
            const obj = {};
            headers.forEach((h, i) => obj[h] = row[i]);
            if (obj['Emp ID'] && String(obj['Emp ID']).trim() !== '') {
              allRecords.push(obj);
            }
          });
        }
      }
      return respond({ status: 'success', data: allRecords, total: allRecords.length });
    }

    return respond({ status: 'error', message: 'Unknown action: ' + action });

  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  return respond({ status: 'ok', message: 'Device Asset API v2 - Search by Emp ID across all sheets' });
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====================================================================
// Google Apps Script - Material Delivery Plan API
// ====================================================================
// วิธีใช้:
// 1. เปิด Google Sheets สร้างชีตใหม่
// 2. สร้าง 3 แท็บ: "SKUs", "Deliveries", "Plans"
// 3. ใส่ Header แต่ละแท็บ:
//    - SKUs:        A1=id, B1=name, C1=month, D1=forecast
//    - Deliveries:  A1=id, B1=date, C1=sku, D1=lot, E1=qty
//    - Plans:       A1=id, B1=date, C1=sku, D1=qty
// 4. ไปที่ Extensions > Apps Script
// 5. ลบโค้ดเดิม แล้ววางโค้ดนี้ทั้งหมด
// 6. กด Deploy > New Deployment > Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 7. คัดลอก URL ที่ได้ไปใส่ใน index.html (ตัวแปร GAS_URL)
// ====================================================================

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
    const action = e.parameter.action;
    const sheet = e.parameter.sheet;

    try {
        if (action === 'read') {
            return respond(readSheet(sheet));
        }
        return respond({ error: 'Invalid action' });
    } catch (err) {
        return respond({ error: err.message });
    }
}

function doPost(e) {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const sheet = body.sheet;

    try {
        if (action === 'save') {
            return respond(saveData(sheet, body.data));
        }
        if (action === 'add') {
            return respond(addRow(sheet, body.row));
        }
        if (action === 'delete') {
            return respond(deleteRow(sheet, body.id));
        }
        if (action === 'update') {
            return respond(updateRow(sheet, body.id, body.row));
        }
        return respond({ error: 'Invalid action' });
    } catch (err) {
        return respond({ error: err.message });
    }
}

// ===== READ ALL DATA FROM SHEET =====
function readSheet(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(sheetName);
    if (!ws) return [];

    const data = ws.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
        const obj = {};
        headers.forEach((h, j) => {
            obj[h] = data[i][j];
        });
        rows.push(obj);
    }
    return rows;
}

// ===== SAVE (OVERWRITE) ALL DATA =====
function saveData(sheetName, dataArray) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName(sheetName);
    if (!ws) {
        ws = ss.insertSheet(sheetName);
    }

    // Get headers from first row or from data
    const headers = getHeaders(sheetName);

    // Clear all data except header
    ws.clearContents();
    ws.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (dataArray.length > 0) {
        const rows = dataArray.map(obj => headers.map(h => obj[h] || ''));
        ws.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    return { success: true, count: dataArray.length };
}

// ===== ADD SINGLE ROW =====
function addRow(sheetName, rowData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName(sheetName);
    if (!ws) {
        ws = ss.insertSheet(sheetName);
        const headers = getHeaders(sheetName);
        ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const headers = getHeaders(sheetName);
    const row = headers.map(h => rowData[h] || '');
    ws.appendRow(row);

    return { success: true };
}

// ===== DELETE ROW BY ID =====
function deleteRow(sheetName, id) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(sheetName);
    if (!ws) return { success: false, error: 'Sheet not found' };

    const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
            ws.deleteRow(i + 1);
            return { success: true };
        }
    }
    return { success: false, error: 'ID not found' };
}

// ===== UPDATE ROW BY ID =====
function updateRow(sheetName, id, rowData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName(sheetName);
    if (!ws) return { success: false, error: 'Sheet not found' };

    const data = ws.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(id)) {
            const row = headers.map(h => rowData[h] !== undefined ? rowData[h] : data[i][headers.indexOf(h)]);
            ws.getRange(i + 1, 1, 1, headers.length).setValues([row]);
            return { success: true };
        }
    }
    return { success: false, error: 'ID not found' };
}

// ===== HELPERS =====
function getHeaders(sheetName) {
    const map = {
        'SKUs': ['id', 'name', 'month', 'forecast'],
        'Deliveries': ['id', 'date', 'sku', 'lot', 'qty'],
        'Plans': ['id', 'date', 'sku', 'qty']
    };
    return map[sheetName] || ['id'];
}

function respond(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

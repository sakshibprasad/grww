/**
 * GRW <-> GOOGLE SHEETS BRIDGE (no Firebase, no login required)
 * ---------------------------------------------------------------
 * 1. Open your Sheet.
 * 2. Extensions -> Apps Script
 * 3. Delete any starter code, paste this whole file in.
 * 4. Change API_KEY below to your own secret string (anything you like).
 * 5. Click Deploy -> New deployment -> gear icon -> "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 6. Click Deploy, authorize the permissions it asks for.
 * 7. Copy the "Web app URL" it gives you -> paste into SHEET_WEBAPP_URL
 *    in both index.html and admin.html (and use the SAME API_KEY in index.html).
 * 8. On first run it will auto-create an "Employees" tab with headers.
 *
 * OPTIONAL PIN PROTECTION:
 * If you want a light extra check beyond the API key, add a tab named
 * "Credentials" with two columns: empId | pin. Fill in rows for whichever
 * employees you want to PIN-protect. Anyone whose Employee ID is NOT
 * listed there can sync freely (no PIN needed) - this stays permissive
 * by default so you don't have to set up every employee up front.
 */

const SHEET_NAME = 'Employees';
const CREDENTIALS_SHEET_NAME = 'Credentials';
const API_KEY = 'CHANGE-THIS-SECRET-2026'; // <-- set your own value, must match SHEET_API_KEY in index.html

function doGet(e){
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(row=>{
    const obj = {};
    headers.forEach((h,i)=> obj[h] = row[i]);
    try{ obj.payload = JSON.parse(obj.payload || '{}'); } catch(err){ obj.payload = {}; }
    return obj;
  });
  return jsonOut_({ok:true, employees: rows});
}

function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.apiKey !== API_KEY){
      return jsonOut_({ok:false, error:'unauthorized'});
    }
    if(!body.empId){
      return jsonOut_({ok:false, error:'missing employee id'});
    }
    if(!checkPin_(body.empId, body.pin)){
      return jsonOut_({ok:false, error:'incorrect PIN for this Employee ID'});
    }

    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const empIdCol = headers.indexOf('empId');
    let rowIndex = -1;
    for(let i=1; i<data.length; i++){
      if(data[i][empIdCol] === body.empId){ rowIndex = i+1; break; }
    }
    const rowValues = [
      body.empId,
      body.name || '',
      body.role || '',
      new Date().toISOString(),
      JSON.stringify(body.payload || {})
    ];
    if(rowIndex === -1){
      sheet.appendRow(rowValues);
    } else {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    }
    return jsonOut_({ok:true});
  } catch(err){
    return jsonOut_({ok:false, error:err.message});
  }
}

function checkPin_(empId, pin){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const credSheet = ss.getSheetByName(CREDENTIALS_SHEET_NAME);
  if(!credSheet) return true; // no Credentials tab set up - allow everyone through
  const data = credSheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(empId)){
      // this empId IS listed - PIN must match
      return String(data[i][1]) === String(pin || '');
    }
  }
  return true; // empId not listed in Credentials - no PIN required
}

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['empId','name','role','updatedAt','payload']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

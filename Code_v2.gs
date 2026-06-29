const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  const action = e.parameter.action;
  const cb = e.parameter.callback;

  let result;
  if (action === 'getMenu')       result = getMenu();
  else if (action === 'getStaff') result = getStaff();
  else result = { error: 'Unknown action' };

  const json = JSON.stringify(result);

  // JSONP support untuk browser cross-origin
  if (cb) {
    return ContentService
      .createTextOutput(`${cb}(${json})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  let result;
  if (action === 'updateMenu')       result = updateMenu(data);
  else if (action === 'updateStaff') result = updateStaff(data);
  else if (action === 'logOrder')    result = logOrder(data);
  else result = { error: 'Unknown action' };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMenu() {
  const sheet = SS.getSheetByName('Menu');
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj.varian   = obj.varian ? String(obj.varian).split('|').map(v => v.trim()) : [];
    obj.tersedia = (obj.tersedia === true || String(obj.tersedia).toUpperCase() === 'TRUE');
    return obj;
  });
}

function getStaff() {
  const sheet = SS.getSheetByName('Staff');
  const rows  = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    obj.online = (obj.online === true || String(obj.online).toUpperCase() === 'TRUE');
    return obj;
  });
}

function updateMenu(data) {
  const sheet = SS.getSheetByName('Menu');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 7).setValue(data.tersedia);
      return { success: true };
    }
  }
  return { error: 'Menu not found' };
}

function updateStaff(data) {
  const sheet = SS.getSheetByName('Staff');
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.nama)) {
      sheet.getRange(i + 1, 3).setValue(data.online);
      return { success: true };
    }
  }
  return { error: 'Staff not found' };
}

function logOrder(data) {
  const sheet = SS.getSheetByName('Orders');
  sheet.appendRow([
    new Date(),
    data.nama_user,
    data.staff_tujuan,
    data.total_harga,
    data.detail_pesanan
  ]);
  return { success: true };
}

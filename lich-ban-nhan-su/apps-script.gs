// ============================================================
// LỊCH BẬN NHÂN SỰ — Can Đăng Studio
// Thêm toàn bộ code này vào Google Apps Script hiện tại
// Sau đó deploy lại (Deploy > Manage deployments > Update)
// ============================================================

const LICH_SHEET_ID = '1KF8Iab6Oo2xMQMXR64SxaO5GZCGRxit7JHtd9v3r8sE';

// Gọi hàm này 1 lần để tạo 3 sheet tabs + nhân sự mặc định
// Chạy từ Apps Script editor: Run > initLichBanSheets
function initLichBanSheets() {
  const ss = SpreadsheetApp.openById(LICH_SHEET_ID);

  let staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('Staff');
    staffSheet.appendRow(['id', 'name', 'token', 'role']);
    const defaultStaff = ['Thanh', 'Huy', 'Hiếu', 'Tú', 'Hiền', 'Trung'];
    defaultStaff.forEach(name => {
      staffSheet.appendRow([genId(), name, genToken(), 'staff']);
    });
  }

  if (!ss.getSheetByName('BusyDays')) {
    const s = ss.insertSheet('BusyDays');
    s.appendRow(['staff_id', 'date', 'status']);
  }

  if (!ss.getSheetByName('Projects')) {
    const s = ss.insertSheet('Projects');
    s.appendRow(['id', 'name', 'start_date', 'end_date', 'assigned_staff']);
  }
}

function genId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function genToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = '';
  for (let i = 0; i < 8; i++) t += chars[Math.floor(Math.random() * chars.length)];
  return t;
}

// ── Hook vào doGet / doPost hiện có ──────────────────────────
// Trong doGet(e) của bạn, thêm vào switch hoặc thêm đoạn sau:
//
//   const lichResult = handleLichGet(e.parameter);
//   if (lichResult !== null) return jsonOut(lichResult);
//
// Trong doPost(e) của bạn:
//   let body; try { body = JSON.parse(e.postData.contents); } catch(_) { body = e.parameter; }
//   const lichResult = handleLichPost(body);
//   if (lichResult !== null) return jsonOut(lichResult);

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

const LICH_ACTIONS_GET  = ['getBusyDays','getMyDays','getStaff','getProjects','initSheets'];
const LICH_ACTIONS_POST = ['setBusyDay','addStaff','removeStaff','addProject'];

function handleLichGet(p) {
  if (!LICH_ACTIONS_GET.includes(p.action)) return null;
  try {
    switch (p.action) {
      case 'getBusyDays': return lichGetBusyDays(p.month);
      case 'getMyDays':   return lichGetMyDays(p.token, p.month);
      case 'getStaff':    return lichGetStaff();
      case 'getProjects': return lichGetProjects(p.month);
      case 'initSheets':  initLichBanSheets(); return { ok: true };
    }
  } catch (e) { return { error: e.message }; }
}

function handleLichPost(body) {
  if (!LICH_ACTIONS_POST.includes(body.action)) return null;
  try {
    switch (body.action) {
      case 'setBusyDay':   return lichSetBusyDay(body);
      case 'addStaff':     return lichAddStaff(body);
      case 'removeStaff':  return lichRemoveStaff(body);
      case 'addProject':   return lichAddProject(body);
    }
  } catch (e) { return { error: e.message }; }
}

// ── Helpers ──────────────────────────────────────────────────

function lichGetSS() { return SpreadsheetApp.openById(LICH_SHEET_ID); }

function lichGetStaff() {
  const sheet = lichGetSS().getSheetByName('Staff');
  if (!sheet) return { staff: [] };
  const rows = sheet.getDataRange().getValues().slice(1);
  return { staff: rows.filter(r => r[0]).map(r => ({ id: r[0], name: r[1], token: r[2], role: r[3] })) };
}

function lichGetBusyDays(month) {
  const ss = lichGetSS();
  const busySheet = ss.getSheetByName('BusyDays');
  if (!busySheet) return { days: [] };
  const { staff } = lichGetStaff();
  const staffMap = {};
  staff.forEach(s => staffMap[s.id] = s.name);
  const rows = busySheet.getDataRange().getValues().slice(1);
  const days = rows
    .filter(r => r[0] && r[1] && String(r[1]).startsWith(month))
    .map(r => ({ staff_id: r[0], date: r[1], status: r[2], staff_name: staffMap[r[0]] || '' }));
  return { days };
}

function lichGetMyDays(token, month) {
  const staffSheet = lichGetSS().getSheetByName('Staff');
  if (!staffSheet) return { error: 'Not found', days: [], staff: null };
  const staffRows = staffSheet.getDataRange().getValues().slice(1);
  const row = staffRows.find(r => r[2] === token);
  if (!row) return { error: 'Invalid token', days: [], staff: null };
  const staffId = row[0], staffName = row[1];
  const busySheet = lichGetSS().getSheetByName('BusyDays');
  if (!busySheet) return { days: [], staff: { id: staffId, name: staffName } };
  const rows = busySheet.getDataRange().getValues().slice(1);
  const days = rows
    .filter(r => r[0] === staffId && r[1] && String(r[1]).startsWith(month))
    .map(r => ({ staff_id: r[0], date: r[1], status: r[2] }));
  return { days, staff: { id: staffId, name: staffName } };
}

function lichSetBusyDay(data) {
  const ss = lichGetSS();
  const staffSheet = ss.getSheetByName('Staff');
  const staffRow = staffSheet.getDataRange().getValues().slice(1).find(r => r[2] === data.token);
  if (!staffRow) return { error: 'Invalid token' };
  const staffId = staffRow[0];
  const busySheet = ss.getSheetByName('BusyDays');
  const rows = busySheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === staffId && rows[i][1] === data.date) {
      if (!data.status) busySheet.deleteRow(i + 1);
      else busySheet.getRange(i + 1, 3).setValue(data.status);
      return { ok: true };
    }
  }
  if (data.status) busySheet.appendRow([staffId, data.date, data.status]);
  return { ok: true };
}

function lichAddStaff(data) {
  const sheet = lichGetSS().getSheetByName('Staff');
  const id = genId(), token = genToken();
  sheet.appendRow([id, data.name, token, 'staff']);
  return { ok: true, id, name: data.name, token, role: 'staff' };
}

function lichRemoveStaff(data) {
  const sheet = lichGetSS().getSheetByName('Staff');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { error: 'Not found' };
}

function lichGetProjects(month) {
  const sheet = lichGetSS().getSheetByName('Projects');
  if (!sheet) return { projects: [] };
  const [y, m] = month.split('-').map(Number);
  const monthStart = new Date(y, m - 1, 1);
  const monthEnd   = new Date(y, m, 0);
  const rows = sheet.getDataRange().getValues().slice(1);
  const projects = rows
    .filter(r => r[0] && new Date(r[2]) <= monthEnd && new Date(r[3]) >= monthStart)
    .map(r => ({ id: r[0], name: r[1], start_date: r[2], end_date: r[3], assigned_staff: r[4] ? String(r[4]).split(',') : [] }));
  return { projects };
}

function lichAddProject(data) {
  const sheet = lichGetSS().getSheetByName('Projects');
  const id = genId();
  const assigned = Array.isArray(data.staff) ? data.staff.join(',') : (data.staff || '');
  sheet.appendRow([id, data.name, data.start, data.end, assigned]);
  return { ok: true, id };
}

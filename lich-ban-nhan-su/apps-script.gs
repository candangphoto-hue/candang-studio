// ============================================================
// LỊCH BẬN NHÂN SỰ — Can Đăng Studio · Standalone GAS
// Hướng dẫn deploy:
//   1. Vào script.google.com → New project
//   2. Paste toàn bộ file này vào editor
//   3. Chạy initSheets() một lần (Run > initSheets) để tạo dữ liệu mặc định
//   4. Deploy → New deployment → Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   5. Copy URL → Paste vào trang lich-ban-nhan-su khi được hỏi
// ============================================================

const SHEET_ID = '1KF8Iab6Oo2xMQMXR64SxaO5GZCGRxit7JHtd9v3r8sE';

function getSS() { return SpreadsheetApp.openById(SHEET_ID); }

// Gọi 1 lần để tạo 3 sheet tabs + nhân sự mặc định
function initSheets() {
  const ss = getSS();

  let staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    staffSheet = ss.insertSheet('Staff');
    staffSheet.appendRow(['id', 'name', 'token', 'role']);
    ['Thanh', 'Huy', 'Hiếu', 'Tú', 'Hiền', 'Trung'].forEach(name => {
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
  const c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = '';
  for (let i = 0; i < 8; i++) t += c[Math.floor(Math.random() * c.length)];
  return t;
}

function out(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e.parameter;
  try {
    switch (p.action) {
      case 'getBusyDays': return out(getBusyDays(p.month));
      case 'getMyDays':   return out(getMyDays(p.token, p.month));
      case 'getStaff':    return out(getStaff());
      case 'getProjects': return out(getProjects(p.month));
      case 'initSheets':  initSheets(); return out({ ok: true });
      default: return out({ error: 'Unknown action' });
    }
  } catch (err) { return out({ error: err.message }); }
}

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (_) { p = e.parameter; }
  try {
    switch (p.action) {
      case 'setBusyDay':  return out(setBusyDay(p));
      case 'addStaff':    return out(addStaff(p));
      case 'removeStaff': return out(removeStaff(p));
      case 'addProject':  return out(addProject(p));
      default: return out({ error: 'Unknown action' });
    }
  } catch (err) { return out({ error: err.message }); }
}

// ── Handlers ─────────────────────────────────────────────────

function getStaff() {
  const sheet = getSS().getSheetByName('Staff');
  if (!sheet) return { staff: [] };
  const rows = sheet.getDataRange().getValues().slice(1);
  return { staff: rows.filter(r => r[0]).map(r => ({ id: r[0], name: r[1], token: r[2], role: r[3] })) };
}

function getBusyDays(month) {
  const ss = getSS();
  const sheet = ss.getSheetByName('BusyDays');
  if (!sheet) return { days: [] };
  const { staff } = getStaff();
  const map = {};
  staff.forEach(s => map[s.id] = s.name);
  const rows = sheet.getDataRange().getValues().slice(1);
  return {
    days: rows
      .filter(r => r[0] && r[1] && String(r[1]).startsWith(month))
      .map(r => ({ staff_id: r[0], date: r[1], status: r[2], staff_name: map[r[0]] || '' }))
  };
}

function getMyDays(token, month) {
  const ss = getSS();
  const staffRows = ss.getSheetByName('Staff').getDataRange().getValues().slice(1);
  const row = staffRows.find(r => r[2] === token);
  if (!row) return { error: 'Invalid token', days: [], staff: null };
  const staffId = row[0], staffName = row[1];
  const busySheet = ss.getSheetByName('BusyDays');
  if (!busySheet) return { days: [], staff: { id: staffId, name: staffName } };
  const days = busySheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] === staffId && r[1] && String(r[1]).startsWith(month))
    .map(r => ({ staff_id: r[0], date: r[1], status: r[2] }));
  return { days, staff: { id: staffId, name: staffName } };
}

function setBusyDay(data) {
  const ss = getSS();
  const staffRow = ss.getSheetByName('Staff').getDataRange().getValues().slice(1).find(r => r[2] === data.token);
  if (!staffRow) return { error: 'Invalid token' };
  const staffId = staffRow[0];
  const sheet = ss.getSheetByName('BusyDays');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === staffId && rows[i][1] === data.date) {
      if (!data.status) sheet.deleteRow(i + 1);
      else sheet.getRange(i + 1, 3).setValue(data.status);
      return { ok: true };
    }
  }
  if (data.status) sheet.appendRow([staffId, data.date, data.status]);
  return { ok: true };
}

function addStaff(data) {
  const id = genId(), token = genToken();
  getSS().getSheetByName('Staff').appendRow([id, data.name, token, 'staff']);
  return { ok: true, id, name: data.name, token, role: 'staff' };
}

function removeStaff(data) {
  const sheet = getSS().getSheetByName('Staff');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { error: 'Not found' };
}

function getProjects(month) {
  const sheet = getSS().getSheetByName('Projects');
  if (!sheet) return { projects: [] };
  const [y, m] = month.split('-').map(Number);
  const mStart = new Date(y, m - 1, 1), mEnd = new Date(y, m, 0);
  const rows = sheet.getDataRange().getValues().slice(1);
  return {
    projects: rows
      .filter(r => r[0] && new Date(r[2]) <= mEnd && new Date(r[3]) >= mStart)
      .map(r => ({ id: r[0], name: r[1], start_date: r[2], end_date: r[3], assigned_staff: r[4] ? String(r[4]).split(',') : [] }))
  };
}

function addProject(data) {
  const id = genId();
  const staff = Array.isArray(data.staff) ? data.staff.join(',') : (data.staff || '');
  getSS().getSheetByName('Projects').appendRow([id, data.name, data.start, data.end, staff]);
  return { ok: true, id };
}

// ============================================================
// LỊCH BẬN NHÂN SỰ — Can Đăng Studio
// GAS tự tạo Google Sheet mới, không cần setup thủ công
// Deploy: Web app · Execute as Me · Access: Anyone
// ============================================================

// Token cố định cho 6 nhân sự mặc định
const DEFAULT_STAFF = [
  { name: 'Thanh', token: 'CDThanh1' },
  { name: 'Huy',   token: 'CDHuy001' },
  { name: 'Hiếu',  token: 'CDHieu01' },
  { name: 'Tú',    token: 'CDTu0001' },
  { name: 'Hiền',  token: 'CDHien01' },
  { name: 'Trung', token: 'CDTrung1' },
];

// Tự tạo Sheet nếu chưa có, lưu ID vào ScriptProperties
function getSS() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SS_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (_) {}
  }
  // Tạo mới
  const ss = SpreadsheetApp.create('Lịch Bận Nhân Sự — Can Đăng Studio');
  props.setProperty('SS_ID', ss.getId());
  _setupSheets(ss);
  return ss;
}

function _setupSheets(ss) {
  // Xóa sheet mặc định
  const def = ss.getSheets()[0];

  const staffSheet = ss.insertSheet('Staff');
  staffSheet.appendRow(['id', 'name', 'token', 'role']);
  DEFAULT_STAFF.forEach(s => staffSheet.appendRow([genId(), s.name, s.token, 'staff']));

  const busy = ss.insertSheet('BusyDays');
  busy.appendRow(['staff_id', 'date', 'status']);

  const proj = ss.insertSheet('Projects');
  proj.appendRow(['id', 'name', 'start_date', 'end_date', 'assigned_staff']);

  try { ss.deleteSheet(def); } catch (_) {}
}

// Gọi từ editor nếu muốn reset dữ liệu
function initSheets() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SS_ID');
  getSS(); // recreate
}

function genId() { return Utilities.getUuid().replace(/-/g, '').substring(0, 8); }

function out(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e.parameter;
  try {
    switch (p.action) {
      // Read
      case 'getBusyDays':      return out(getBusyDays(p.month));
      case 'getMyDays':        return out(getMyDays(p.token, p.month));
      case 'getStaff':         return out(getStaff());
      case 'getProjects':      return out(getProjects(p.month));
      case 'initSheets':       getSS(); return out({ ok: true });
      // Write via GET (tránh CORS)
      case 'setBusyDay':       return out(setBusyDay(p));
      case 'adminSetBusyDay':  return out(adminSetBusyDay(p));
      case 'addStaff':         return out(addStaff(p));
      case 'removeStaff':      return out(removeStaff(p));
      case 'addProject':       return out(addProject(p));
      default:                 return out({ error: 'Unknown action' });
    }
  } catch (err) { return out({ error: err.message }); }
}

function doPost(e) {
  let p; try { p = JSON.parse(e.postData.contents); } catch (_) { p = e.parameter; }
  try {
    switch (p.action) {
      case 'setBusyDay':      return out(setBusyDay(p));
      case 'adminSetBusyDay': return out(adminSetBusyDay(p));
      case 'addStaff':        return out(addStaff(p));
      case 'removeStaff':     return out(removeStaff(p));
      case 'addProject':      return out(addProject(p));
      default:                return out({ error: 'Unknown action' });
    }
  } catch (err) { return out({ error: err.message }); }
}

// ── Handlers ─────────────────────────────────────────────────

function getStaff() {
  const sheet = getSS().getSheetByName('Staff');
  if (!sheet) return { staff: [] };
  return { staff: sheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({ id: r[0], name: r[1], token: r[2], role: r[3] })) };
}

function getBusyDays(month) {
  const ss = getSS();
  const sheet = ss.getSheetByName('BusyDays');
  if (!sheet) return { days: [] };
  const { staff } = getStaff(); const map = {}; staff.forEach(s => map[s.id] = s.name);
  return { days: sheet.getDataRange().getValues().slice(1).filter(r => r[0] && r[1] && String(r[1]).startsWith(month)).map(r => ({ staff_id: r[0], date: r[1], status: r[2], staff_name: map[r[0]] || '' })) };
}

function getMyDays(token, month) {
  const staffRows = getSS().getSheetByName('Staff').getDataRange().getValues().slice(1);
  const row = staffRows.find(r => r[2] === token);
  if (!row) return { error: 'Invalid token', days: [], staff: null };
  const busySheet = getSS().getSheetByName('BusyDays');
  const days = busySheet ? busySheet.getDataRange().getValues().slice(1).filter(r => r[0] === row[0] && r[1] && String(r[1]).startsWith(month)).map(r => ({ staff_id: r[0], date: r[1], status: r[2] })) : [];
  return { days, staff: { id: row[0], name: row[1] } };
}

function setBusyDay(data) {
  const staffRow = getSS().getSheetByName('Staff').getDataRange().getValues().slice(1).find(r => r[2] === data.token);
  if (!staffRow) return { error: 'Invalid token' };
  const sheet = getSS().getSheetByName('BusyDays');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === staffRow[0] && rows[i][1] === data.date) {
      if (!data.status) sheet.deleteRow(i + 1); else sheet.getRange(i + 1, 3).setValue(data.status);
      return { ok: true };
    }
  }
  if (data.status) sheet.appendRow([staffRow[0], data.date, data.status]);
  return { ok: true };
}

function adminSetBusyDay(data) {
  const sheet = getSS().getSheetByName('BusyDays');
  if (!sheet) return { error: 'No sheet' };
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.staff_id && rows[i][1] === data.date) {
      if (!data.status) sheet.deleteRow(i + 1); else sheet.getRange(i + 1, 3).setValue(data.status);
      return { ok: true };
    }
  }
  if (data.status) sheet.appendRow([data.staff_id, data.date, data.status]);
  return { ok: true };
}

function addStaff(data) {
  const id = genId(), token = genToken8();
  getSS().getSheetByName('Staff').appendRow([id, data.name, token, 'staff']);
  return { ok: true, id, name: data.name, token, role: 'staff' };
}

function genToken8() {
  const c = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = ''; for (let i = 0; i < 8; i++) t += c[Math.floor(Math.random() * c.length)]; return t;
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
  return { projects: sheet.getDataRange().getValues().slice(1).filter(r => r[0] && new Date(r[2]) <= mEnd && new Date(r[3]) >= mStart).map(r => ({ id: r[0], name: r[1], start_date: r[2], end_date: r[3], assigned_staff: r[4] ? String(r[4]).split(',') : [] })) };
}

function addProject(data) {
  const id = genId();
  getSS().getSheetByName('Projects').appendRow([id, data.name, data.start, data.end, Array.isArray(data.staff) ? data.staff.join(',') : '']);
  return { ok: true, id };
}

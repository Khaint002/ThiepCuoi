/**
 * Backend cho thiệp cưới: lưu xác nhận tham dự + lời chúc, quản lý danh sách khách mời.
 * Cài đặt:
 *  1. Mở Google Sheet → Tiện ích mở rộng → Apps Script → dán toàn bộ file này.
 *  2. Đặt mật khẩu trang quản trị: ⚙️ Cài đặt dự án → Thuộc tính tập lệnh → Thêm thuộc tính
 *     ADMIN_PASSWORD = <mật khẩu của bạn>. Không ghi mật khẩu vào file này vì repo công khai.
 *  3. Triển khai → Ứng dụng web (Thực thi dưới dạng: Tôi · Ai có quyền truy cập: Bất kỳ ai).
 *     Khi sửa file: Triển khai → Quản lý các tùy chọn triển khai → ✏️ → Phiên bản mới (giữ nguyên URL).
 */

const WISHES = { name: 'Lời chúc', headers: ['Thời gian', 'Họ và tên', 'Khách của', 'Tham dự', 'Số người', 'Lời chúc', 'Mã khách'] };
const GUESTS = { name: 'Khách mời', headers: ['Mã', 'Tên hiển thị', 'Khách của', 'Lời mời riêng', 'Ngày tạo', 'Lượt mở', 'Lần mở cuối'] };

function doPost(e) {
  return withLock_(() => {
    const data = JSON.parse(e.postData.contents);
    if (data.action) return admin_(data);
    getSheet_(WISHES).appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.side),
      data.attend === 'yes' ? 'Có' : 'Không',
      clean_(data.guests),
      clean_(data.message),
      clean_(data.g)
    ]);
    return { ok: true };
  });
}

// ?g=<mã khách>: trả về lời mời riêng và đếm lượt mở. Không có tham số: kiểm tra script đang chạy.
function doGet(e) {
  const id = (e.parameter && e.parameter.g) || '';
  if (!id) return json_({ ok: true, message: 'Thiệp cưới: script đang hoạt động' });
  return withLock_(() => {
    const sheet = getSheet_(GUESTS);
    const row = findRow_(sheet, id);
    if (!row) return { ok: false, error: 'Không tìm thấy khách' };
    const v = sheet.getRange(row, 1, 1, GUESTS.headers.length).getValues()[0];
    sheet.getRange(row, 6, 1, 2).setValues([[(Number(v[5]) || 0) + 1, new Date()]]);
    return { ok: true, guest: { name: v[1], side: v[2], invite: v[3] } };
  });
}

/* ---------- Trang quản trị ---------- */
function admin_(data) {
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw) return { ok: false, error: 'Chưa đặt ADMIN_PASSWORD trong Thuộc tính tập lệnh' };
  if (data.password !== pw) { Utilities.sleep(1000); return { ok: false, error: 'Sai mật khẩu' }; }

  const guests = getSheet_(GUESTS);
  const wishes = getSheet_(WISHES);
  switch (data.action) {
    case 'list':
      break;
    case 'addGuests': {
      const used = new Set(readRows_(guests, GUESTS).map((v) => String(v[0])));
      const now = new Date();
      const rows = (data.guests || [])
        .filter((g) => String(g.name || '').trim())
        .map((g) => [newId_(used), clean_(g.name), clean_(g.side), clean_(g.invite), now, 0, '']);
      if (rows.length) guests.getRange(guests.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      break;
    }
    case 'updateGuest': {
      const row = findRow_(guests, data.id);
      if (!row) return { ok: false, error: 'Không tìm thấy khách' };
      guests.getRange(row, 2, 1, 3).setValues([[clean_(data.name), clean_(data.side), clean_(data.invite)]]);
      break;
    }
    case 'deleteGuest': {
      const row = findRow_(guests, data.id);
      if (row) guests.deleteRow(row);
      break;
    }
    case 'deleteWish': {
      // Kiểm tra cả thời gian để không xoá nhầm dòng khi sheet vừa thay đổi
      const row = Number(data.row);
      if (row > 1 && row <= wishes.getLastRow() && iso_(wishes.getRange(row, 1).getValue()) === data.time) wishes.deleteRow(row);
      break;
    }
    default:
      return { ok: false, error: 'Thao tác không hợp lệ' };
  }

  return {
    ok: true,
    guests: readRows_(guests, GUESTS).map((v) => ({
      id: String(v[0]), name: v[1], side: v[2], invite: v[3],
      created: iso_(v[4]), views: Number(v[5]) || 0, lastOpen: iso_(v[6])
    })),
    wishes: readRows_(wishes, WISHES).map((v, i) => ({
      row: i + 2, time: iso_(v[0]), name: v[1], side: v[2], attend: v[3],
      guests: String(v[4]), message: v[5], g: String(v[6] || '')
    }))
  };
}

/* ---------- Tiện ích ---------- */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return json_(fn());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_(def) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(def.name);
  if (!sheet) {
    sheet = ss.insertSheet(def.name);
    sheet.setFrozenRows(1);
  }
  // Tạo hoặc bổ sung dòng tiêu đề (sheet cũ có thể thiếu cột mới)
  const head = sheet.getRange(1, 1, 1, def.headers.length);
  if (head.getValues()[0].join('|') !== def.headers.join('|')) head.setValues([def.headers]).setFontWeight('bold');
  return sheet;
}

function readRows_(sheet, def) {
  const last = sheet.getLastRow();
  return last < 2 ? [] : sheet.getRange(2, 1, last - 1, def.headers.length).getValues();
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().map((v) => String(v[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? 0 : i + 2;
}

// Mã 6 ký tự, bắt đầu bằng chữ để Sheets không tự đổi thành số
function newId_(used) {
  const letters = 'abcdefghjkmnpqrstuvwxyz', chars = letters + '23456789';
  let id;
  do {
    id = letters[Math.floor(Math.random() * letters.length)];
    for (let i = 0; i < 5; i++) id += chars[Math.floor(Math.random() * chars.length)];
  } while (used.has(id));
  used.add(id);
  return id;
}

function iso_(v) {
  return v instanceof Date ? v.toISOString() : (v ? String(v) : '');
}

// Cắt độ dài và chặn chuỗi bắt đầu bằng = + - @ để không bị Sheets hiểu là công thức
function clean_(value) {
  const s = String(value == null ? '' : value).trim().slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

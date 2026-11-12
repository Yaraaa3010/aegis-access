/**
 * AEGIS ACCESS CORE - backend Google Apps Script
 * Database : Google Sheet (sheet "Database" + sheet "Log")
 * Kolom Database : A = Waktu | B = Passcode (hash) | C = Sistem Akhir
 */
const SPREADSHEET_ID = '';        // kosongkan jika script dibuat dari menu Extensions > Apps Script di Sheet
const SHEET_DB = 'Database';
const SHEET_LOG = 'Log';
const MAX_ATTEMPTS = 3;           // salah berturut-turut sebelum dikunci sementara
const LOCKOUT_SECONDS = 60;       // lama kunci sementara
const MIN_LEN = 4, MAX_LEN = 8;   // panjang passcode
const DEFAULT_PASSCODE = '123456';
const SHEET_FACE = 'Wajah';
const FACE_THRESHOLD = 0.5;       // jarak maksimum wajah dianggap cocok (makin kecil makin ketat)
const MAX_SAMPLES = 5;
const PROFILES = { ADMIN: 60, STAF: 30, TAMU: 10 };  // profil : detik sebelum sistem kunci otomatis
const TZ = Session.getScriptTimeZone();

function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('AEGIS // Access Core')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    // Mode GitHub Pages: UI ada di GitHub, Apps Script hanya sebagai API
    return ContentService.createTextOutput('AEGIS API aktif');
  }
}

/** Jalankan SEKALI dari editor: membuat sheet, header, salt, dan passcode awal. */
function setup() {
  const ss = ss_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());
  let db = ss.getSheetByName(SHEET_DB) || ss.insertSheet(SHEET_DB);
  let lg = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  db.getRange(1, 1, 1, 3).setValues([['Waktu', 'Passcode (hash)', 'Sistem Akhir']]).setFontWeight('bold');
  lg.getRange(1, 1, 1, 5).setValues([['Waktu', 'Aksi', 'Hasil', 'Detail', 'Perangkat']]).setFontWeight('bold');
  db.setFrozenRows(1); lg.setFrozenRows(1);
  if (!db.getRange('B2').getValue()) {
    db.getRange(2, 1, 1, 3).setValues([[now_(), hash_(DEFAULT_PASSCODE), 'TERKUNCI']]);
  }
  faceSheet_();
  log_('SETUP', 'OK', 'Sistem diinisialisasi', 'editor');
}

/** Opsional: atur passcode langsung dari editor. Ubah angkanya lalu Run. */
function setPasscodeManual() {
  const code = '654321';
  if (!valid_(code)) throw new Error('Passcode harus ' + MIN_LEN + '-' + MAX_LEN + ' angka');
  const db = db_();
  db.getRange(2, 1, 1, 2).setValues([[now_(), hash_(code)]]);
  log_('UBAH', 'BERHASIL', 'Passcode diubah manual dari editor', 'editor');
}

/* ---------- API untuk UI ---------- */

function getState() {
  const p = PropertiesService.getScriptProperties();
  const left = Math.max(0, Math.ceil((Number(p.getProperty('LOCK_UNTIL') || 0) - Date.now()) / 1000));
  const r = db_().getRange(2, 1, 1, 3).getDisplayValues()[0];
  return { waktu: r[0], status: r[2], remaining: left, minLen: MIN_LEN, maxLen: MAX_LEN };
}

function verifyPasscode(code, device) {
  return locked_(function () {
    const gate = gate_(device, 'AKSES');
    if (gate) return gate;
    code = String(code || '');
    if (hash_(code) === db_().getRange('B2').getValue()) {
      reset_();
      setStatus_('TERBUKA');
      log_('AKSES', 'BERHASIL', 'Passcode benar, sistem dibuka', device);
      return { ok: true, status: 'TERBUKA', message: 'Akses diterima' };
    }
    return fail_('AKSES', device);
  });
}

function lockSystem(device) {
  return locked_(function () {
    setStatus_('TERKUNCI');
    log_('KUNCI', 'BERHASIL', 'Sistem dikunci kembali', device);
    return { ok: true, status: 'TERKUNCI' };
  });
}

function changePasscode(oldCode, newCode, device) {
  return locked_(function () {
    const gate = gate_(device, 'UBAH');
    if (gate) return gate;
    oldCode = String(oldCode || ''); newCode = String(newCode || '');
    if (hash_(oldCode) !== db_().getRange('B2').getValue()) return fail_('UBAH', device);
    if (!valid_(newCode)) return { ok: false, message: 'Passcode baru harus ' + MIN_LEN + '-' + MAX_LEN + ' angka' };
    if (/^(\d)\1+$/.test(newCode)) return { ok: false, message: 'Passcode terlalu mudah ditebak' };
    if (newCode === oldCode) return { ok: false, message: 'Passcode baru harus berbeda' };
    db_().getRange(2, 1, 1, 2).setValues([[now_(), hash_(newCode)]]);
    reset_();
    log_('UBAH', 'BERHASIL', 'Passcode berhasil diubah', device);
    return { ok: true, message: 'Passcode berhasil diubah' };
  });
}

/* ---------- Helper ---------- */

function gate_(device, aksi) {
  const p = PropertiesService.getScriptProperties();
  const left = Math.ceil((Number(p.getProperty('LOCK_UNTIL') || 0) - Date.now()) / 1000);
  if (left > 0) {
    log_(aksi, 'DITOLAK', 'Percobaan saat kunci sementara aktif', device);
    return { ok: false, locked: true, remaining: left, message: 'Sistem terkunci sementara' };
  }
  return null;
}

function fail_(aksi, device) {
  const p = PropertiesService.getScriptProperties();
  const fails = Number(p.getProperty('FAILS') || 0) + 1;
  if (fails >= MAX_ATTEMPTS) {
    p.setProperties({ FAILS: '0', LOCK_UNTIL: String(Date.now() + LOCKOUT_SECONDS * 1000) });
    setStatus_('LOCKOUT');
    log_(aksi, 'LOCKOUT', MAX_ATTEMPTS + 'x salah, kunci ' + LOCKOUT_SECONDS + ' detik', device);
    return { ok: false, locked: true, remaining: LOCKOUT_SECONDS, message: 'Terlalu banyak percobaan salah' };
  }
  p.setProperty('FAILS', String(fails));
  log_(aksi, 'GAGAL', 'Passcode salah (' + fails + '/' + MAX_ATTEMPTS + ')', device);
  return { ok: false, left: MAX_ATTEMPTS - fails, message: 'Passcode salah' };
}

function reset_() { PropertiesService.getScriptProperties().setProperties({ FAILS: '0', LOCK_UNTIL: '0' }); }
function valid_(c) { return new RegExp('^\\d{' + MIN_LEN + ',' + MAX_LEN + '}$').test(c); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function db_() {
  const s = ss_().getSheetByName(SHEET_DB);
  if (!s) throw new Error('Jalankan fungsi setup() terlebih dahulu');
  return s;
}
function setStatus_(st) { db_().getRange(2, 1).setValue(now_()); db_().getRange(2, 3).setValue(st); }

function log_(aksi, hasil, detail, device) {
  const safe = function (v) { v = String(v || '').slice(0, 160); return /^[=+\-@]/.test(v) ? "'" + v : v; };
  ss_().getSheetByName(SHEET_LOG).appendRow([now_(), aksi, hasil, safe(detail), safe(device)]);
}

function hash_(code) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + code, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ================= FACE ID ================= */
/* Sheet "Wajah": ID | Nama | Profil | Descriptor (JSON) | Terdaftar. Descriptor tidak pernah dikirim ke browser. */

function faceSheet_() {
  const ss = ss_();
  let s = ss.getSheetByName(SHEET_FACE);
  if (!s) {
    s = ss.insertSheet(SHEET_FACE);
    s.getRange(1, 1, 1, 5).setValues([['ID', 'Nama', 'Profil', 'Descriptor (JSON)', 'Terdaftar']]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('@');
  }
  return s;
}

function dist_(a, b) { let s = 0; for (let i = 0; i < 128; i++) { const d = a[i] - b[i]; s += d * d; } return Math.sqrt(s); }
function okDesc_(d) { return Array.isArray(d) && d.length === 128 && d.every(function (n) { return typeof n === 'number' && isFinite(n); }); }

/** Pengelolaan wajah wajib passcode. Mengembalikan null jika lolos, atau objek error jika gagal/terkunci. */
function admin_(code, device, aksi) {
  const gate = gate_(device, aksi);
  if (gate) return gate;
  if (hash_(String(code || '')) !== db_().getRange('B2').getValue()) return fail_(aksi, device);
  return null;
}

function verifyFace(desc, device) {
  return locked_(function () {
    const gate = gate_(device, 'WAJAH');
    if (gate) return gate;
    if (!okDesc_(desc)) return { ok: false, message: 'Data wajah tidak valid' };
    let best = null;
    faceSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
      if (!r[3]) return;
      JSON.parse(r[3]).forEach(function (d) {
        const x = dist_(desc, d);
        if (!best || x < best.dist) best = { dist: x, nama: r[1], profil: r[2] };
      });
    });
    if (best && best.dist <= FACE_THRESHOLD) {
      reset_(); setStatus_('TERBUKA');
      log_('WAJAH', 'BERHASIL', best.nama + ' (' + best.profil + ') jarak ' + best.dist.toFixed(2), device);
      return { ok: true, nama: best.nama, profil: best.profil, autoLock: PROFILES[best.profil] || 30 };
    }
    log_('WAJAH', 'DITOLAK', 'Wajah tidak dikenal' + (best ? ' (terdekat ' + best.dist.toFixed(2) + ')' : ''), device);
    return { ok: false, message: 'Wajah tidak dikenal' };
  });
}

function registerFace(code, nama, profil, descs, device) {
  return locked_(function () {
    const bad = admin_(code, device, 'DAFTAR_WAJAH');
    if (bad) return bad;
    nama = String(nama || '').replace(/^[=+\-@\s]+/, '').trim().slice(0, 40);
    profil = String(profil || '');
    if (!nama || !(profil in PROFILES)) return { ok: false, message: 'Nama atau profil tidak valid' };
    if (!Array.isArray(descs) || !descs.length || descs.length > MAX_SAMPLES || !descs.every(okDesc_)) return { ok: false, message: 'Data wajah tidak valid' };
    const sh = faceSheet_(), rows = sh.getDataRange().getValues();
    const row = [Utilities.getUuid().slice(0, 8), nama, profil, JSON.stringify(descs), now_()];
    let idx = -1;
    for (let i = 1; i < rows.length; i++) if (String(rows[i][1]).toLowerCase() === nama.toLowerCase()) idx = i + 1;
    if (idx > 0) sh.getRange(idx, 1, 1, 5).setValues([row]); else sh.appendRow(row);
    log_('DAFTAR_WAJAH', 'BERHASIL', nama + ' (' + profil + ')', device);
    return { ok: true, message: 'Wajah ' + nama + ' terdaftar' };
  });
}

function listFaces(code, device) {
  return locked_(function () {
    const bad = admin_(code, device, 'LIHAT_WAJAH');
    if (bad) return bad;
    const rows = faceSheet_().getDataRange().getValues().slice(1);
    return { ok: true, faces: rows.map(function (r) { return { id: String(r[0]), nama: r[1], profil: r[2] }; }) };
  });
}

function deleteFace(code, id, device) {
  return locked_(function () {
    const bad = admin_(code, device, 'HAPUS_WAJAH');
    if (bad) return bad;
    const sh = faceSheet_(), rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(id)) {
        sh.deleteRow(i + 1);
        log_('HAPUS_WAJAH', 'BERHASIL', rows[i][1], device);
        return { ok: true, message: 'Wajah ' + rows[i][1] + ' dihapus' };
      }
    }
    return { ok: false, message: 'Data tidak ditemukan' };
  });
}

/** Hanya dipakai jika Index.html di-host di luar Apps Script (isi GAS_URL di Index.html). */
function doPost(e) {
  const api = { getState: getState, verifyPasscode: verifyPasscode, lockSystem: lockSystem, changePasscode: changePasscode,
                verifyFace: verifyFace, registerFace: registerFace, listFaces: listFaces, deleteFace: deleteFace };
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!api.hasOwnProperty(req.fn)) throw new Error('Fungsi tidak dikenal');
    out = api[req.fn].apply(null, req.args || []);
  } catch (err) { out = { ok: false, message: 'Permintaan tidak valid' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

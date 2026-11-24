/**
 * ABSENSI WAJAH KIOSK — Backend (Google Apps Script)
 * Pasang: Spreadsheet > Extensions > Apps Script > tempel file ini
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Access: Anyone
 * Salin URL /exec ke konstanta GAS_URL di index.html
 */
const CFG = {
  PIN: '1234',               // PIN admin untuk daftar wajah
  JAM_MASUK: '07:00',        // informasi saja
  BATAS_TERLAMBAT: '07:15',  // lewat jam ini => Terlambat
  TZ: 'Asia/Jakarta'
};
const H_SISWA = ['NIS', 'Nama', 'Kelas', 'Wajah'];
const H_ABSEN = ['Tanggal', 'NIS', 'Nama', 'Kelas', 'Jam', 'Status'];

function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), head.length).setNumberFormat('@'); // cegah auto-konversi tanggal/angka
  }
  return s;
}
const now_ = f => Utilities.formatDate(new Date(), CFG.TZ, f);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  try { return out_(init_()); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'absen') return out_(absen_(String(p.nis)));
    if (p.action === 'daftar') return out_(daftar_(p));
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function init_() {
  const rows = sh_('Siswa', H_SISWA).getDataRange().getValues().slice(1);
  const gallery = [];
  rows.forEach(r => {
    if (!r[3]) return;
    try { gallery.push({ nis: String(r[0]), nama: r[1], kelas: r[2], d: JSON.parse(r[3]) }); } catch (_) {}
  });
  return { ok: true, gallery: gallery, today: today_(), total: rows.length };
}

function today_() {
  const t = now_('yyyy-MM-dd');
  return sh_('Absensi', H_ABSEN).getDataRange().getValues().slice(1)
    .filter(r => String(r[0]) === t)
    .map(r => ({ nis: String(r[1]), nama: r[2], kelas: r[3], jam: String(r[4]), status: r[5] }));
}

function absen_(nis) {
  const siswa = sh_('Siswa', H_SISWA).getDataRange().getValues().slice(1).find(r => String(r[0]) === nis);
  if (!siswa) return { ok: false, error: 'NIS tidak terdaftar' };
  const s = sh_('Absensi', H_ABSEN);
  const t = now_('yyyy-MM-dd');
  const dup = today_().find(r => r.nis === nis);
  if (dup) return { ok: true, dup: true, nama: dup.nama, kelas: dup.kelas, jam: dup.jam, status: dup.status };
  const jam = now_('HH:mm');
  const status = jam > CFG.BATAS_TERLAMBAT ? 'Terlambat' : 'Hadir';
  s.appendRow([t, nis, siswa[1], siswa[2], jam, status]);
  return { ok: true, dup: false, nama: siswa[1], kelas: siswa[2], jam: jam, status: status };
}

function daftar_(p) {
  if (String(p.pin) !== CFG.PIN) return { ok: false, error: 'PIN salah' };
  if (!p.nis || !p.nama || !p.d || !p.d.length) return { ok: false, error: 'Data tidak lengkap' };
  const s = sh_('Siswa', H_SISWA);
  const rows = s.getDataRange().getValues();
  const wajah = JSON.stringify(p.d);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.nis)) {
      s.getRange(i + 1, 1, 1, 4).setValues([[String(p.nis), p.nama, p.kelas || '', wajah]]);
      return { ok: true, updated: true };
    }
  }
  s.appendRow([String(p.nis), p.nama, p.kelas || '', wajah]);
  return { ok: true, updated: false };
}

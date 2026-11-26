/**
 * ABSENSI WAJAH KIOSK — Backend (Google Apps Script)
 * Pasang: Spreadsheet > Extensions > Apps Script > tempel file ini
 * Deploy: Deploy > Manage deployments > Edit > New version (Execute as: Me, Access: Anyone)
 * Salin URL /exec ke konstanta GAS_URL di index.html
 *
 * Status absensi: Hadir | Terlambat | Sakit | Izin | Tanpa Masuk (hanya scan pulang)
 * Sakit/Izin dicatat oleh admin (PIN) lewat tombol "Sakit/Izin" di index.html.
 */
const CFG = {
  PIN: '1234',               // PIN admin (daftar wajah, jadwal, rekap, sakit/izin)
  TZ: 'Asia/Jakarta',
  DEF: { masuk: '07:00', terlambat: '07:15', pulang: '15:00' }
};
const H_SISWA = ['NIS', 'Nama', 'Kelas', 'Wajah'];
const H_ABSEN = ['Tanggal', 'NIS', 'Nama', 'Kelas', 'Jam', 'Status', 'Pulang', 'Catatan'];
const H_JADWAL = ['Hari', 'Masuk', 'Terlambat', 'Pulang'];
const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const IZIN = ['Sakit', 'Izin'];

/* ---------- util ---------- */
function sh_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name) || ss.insertSheet(name);
  if (s.getMaxColumns() < head.length) s.insertColumnsAfter(s.getMaxColumns(), head.length - s.getMaxColumns());
  const cur = s.getRange(1, 1, 1, head.length).getValues()[0];
  if (cur.join('|') !== head.join('|')) {   // sheet baru / kolom baru (mis. Pulang, Catatan) => rapikan header
    s.getRange(1, 1, 1, head.length).setNumberFormat('@').setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange(1, 1, s.getMaxRows(), head.length).setNumberFormat('@');   // cegah auto-konversi tanggal/jam
  }
  return s;
}
const now_ = f => Utilities.formatDate(new Date(), CFG.TZ, f);
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const cell_ = (v, f) => v instanceof Date ? Utilities.formatDate(v, CFG.TZ, f) : String(v == null ? '' : v);
const pinOk_ = p => String(p) === CFG.PIN;

function put_(s, row, arr) {   // tulis 1 baris sebagai teks
  if (row > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), row - s.getMaxRows());
  const rg = s.getRange(row, 1, 1, arr.length);
  rg.setNumberFormat('@'); rg.setValues([arr]);
}
function absRows_() {
  const s = sh_('Absensi', H_ABSEN), v = s.getDataRange().getValues(), rows = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!r[0]) continue;
    rows.push({ i: i + 1, tgl: cell_(r[0], 'yyyy-MM-dd'), nis: String(r[1]), nama: String(r[2]), kelas: String(r[3]),
      jam: cell_(r[4], 'HH:mm'), status: String(r[5]), pulang: cell_(r[6], 'HH:mm'), cat: String(r[7] == null ? '' : r[7]) });
  }
  return { s: s, rows: rows };
}
function siswa_(nis) {
  return sh_('Siswa', H_SISWA).getDataRange().getValues().slice(1).find(r => String(r[0]) === String(nis));
}

/* ---------- jadwal ---------- */
function cfg_() {
  const map = {};
  sh_('Jadwal', H_JADWAL).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0]) map[String(r[0])] = { hari: String(r[0]), masuk: cell_(r[1], 'HH:mm'), terlambat: cell_(r[2], 'HH:mm'), pulang: cell_(r[3], 'HH:mm') };
  });
  return ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'].map(h => map[h] || { hari: h, masuk: CFG.DEF.masuk, terlambat: CFG.DEF.terlambat, pulang: CFG.DEF.pulang });
}
function cfgToday_() {
  const h = HARI[Number(now_('u')) % 7];
  return cfg_().find(x => x.hari === h) || null;   // Sabtu/Minggu => null
}
function setCfg_(p) {
  if (!pinOk_(p.pin)) return { ok: false, error: 'PIN salah' };
  if (!Array.isArray(p.hari) || !p.hari.length) return { ok: false, error: 'Data jadwal kosong' };
  const s = sh_('Jadwal', H_JADWAL);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, H_JADWAL.length).clearContent();
  p.hari.forEach((h, i) => put_(s, i + 2, [String(h.hari), String(h.masuk), String(h.terlambat), String(h.pulang)]));
  return { ok: true };
}

/* ---------- endpoint ---------- */
function doGet() {
  try { return out_(init_()); } catch (err) { return out_({ ok: false, error: String(err) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    switch (p.action) {
      case 'absen':  return out_(absen_(String(p.nis), p.jenis));
      case 'daftar': return out_(daftar_(p));
      case 'izin':   return out_(izin_(p));
      case 'rekap':  return out_(rekap_(p));
      case 'setCfg': return out_(setCfg_(p));
    }
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
  return { ok: true, gallery: gallery, today: today_(), cfg: { hari: cfg_() }, total: rows.length };
}

function today_() {
  const t = now_('yyyy-MM-dd');
  return absRows_().rows.filter(r => r.tgl === t)
    .map(r => ({ nis: r.nis, nama: r.nama, kelas: r.kelas, jam: r.jam, status: r.status, pulang: r.pulang, cat: r.cat }));
}

/* ---------- absen wajah (masuk / pulang) ---------- */
function absen_(nis, jenis) {
  const sw = siswa_(nis);
  if (!sw) return { ok: false, error: 'NIS tidak terdaftar' };
  const c = cfgToday_();
  if (!c) return { ok: false, error: 'Hari ini tidak ada jadwal absensi (Senin–Jumat)' };
  const jn = jenis === 'pulang' ? 'pulang' : 'masuk';
  const t = now_('yyyy-MM-dd'), jam = now_('HH:mm');
  const A = absRows_(), s = A.s, r = A.rows.find(x => x.tgl === t && x.nis === nis);
  const base = { ok: true, jenis: jn, nama: sw[1], kelas: sw[2] };

  // siswa yang tercatat Sakit/Izin hari ini tidak bisa absen hadir/pulang
  if (r && IZIN.indexOf(r.status) >= 0) return Object.assign(base, { dup: true, jam: '', status: r.status, pulang: '', at: '', cat: r.cat });

  if (jn === 'masuk') {
    if (r && r.jam) return Object.assign(base, { dup: true, jam: r.jam, status: r.status, pulang: r.pulang, at: r.jam });
    const status = jam > c.terlambat ? 'Terlambat' : 'Hadir';
    if (r) put_(s, r.i, [t, nis, sw[1], sw[2], jam, status, r.pulang, r.cat]);
    else put_(s, s.getLastRow() + 1, [t, nis, sw[1], sw[2], jam, status, '', '']);
    return Object.assign(base, { dup: false, jam: jam, status: status, pulang: r ? r.pulang : '', at: jam });
  }
  if (r && r.pulang) return Object.assign(base, { dup: true, jam: r.jam, status: r.status, pulang: r.pulang, at: r.pulang });
  if (r) put_(s, r.i, [t, nis, r.nama, r.kelas, r.jam, r.status, jam, r.cat]);
  else put_(s, s.getLastRow() + 1, [t, nis, sw[1], sw[2], '', 'Tanpa Masuk', jam, '']);
  return Object.assign(base, { dup: false, jam: r ? r.jam : '', status: r ? r.status : 'Tanpa Masuk', pulang: jam, at: jam, awal: jam < c.pulang });
}

/* ---------- SAKIT / IZIN (admin) ---------- */
// p: {pin, nis, status: 'Sakit'|'Izin'|'Batal', tgl: 'yyyy-MM-dd' (opsional, default hari ini), cat}
function izin_(p) {
  if (!pinOk_(p.pin)) return { ok: false, error: 'PIN salah' };
  const st = String(p.status);
  if (['Sakit', 'Izin', 'Batal'].indexOf(st) < 0) return { ok: false, error: 'Status harus Sakit atau Izin' };
  const tgl = String(p.tgl || now_('yyyy-MM-dd'));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl)) return { ok: false, error: 'Format tanggal salah' };
  const nis = String(p.nis || '').trim();
  const sw = siswa_(nis);
  if (!sw) return { ok: false, error: 'NIS tidak terdaftar' };

  const A = absRows_(), s = A.s, r = A.rows.find(x => x.tgl === tgl && x.nis === nis);
  if (st === 'Batal') {
    if (!r || IZIN.indexOf(r.status) < 0) return { ok: false, error: 'Tidak ada catatan Sakit/Izin pada tanggal tersebut' };
    s.deleteRow(r.i);
    return { ok: true, removed: true, nama: sw[1], kelas: sw[2], tgl: tgl };
  }
  if (r && r.jam) return { ok: false, error: 'Siswa sudah tercatat ' + r.status + ' (hadir) pada tanggal tersebut' };
  const cat = String(p.cat || '').slice(0, 100);
  const row = [tgl, nis, sw[1], sw[2], '', st, '', cat];
  if (r) put_(s, r.i, row); else put_(s, s.getLastRow() + 1, row);
  return { ok: true, updated: !!r, nama: sw[1], kelas: sw[2], status: st, tgl: tgl, cat: cat };
}

/* ---------- rekap ---------- */
function rekap_(p) {
  if (!pinOk_(p.pin)) return { ok: false, error: 'PIN salah' };
  const dari = String(p.dari), sampai = String(p.sampai);
  const rows = absRows_().rows.filter(r => r.tgl >= dari && r.tgl <= sampai)
    .map(r => ({ tgl: r.tgl, nis: r.nis, nama: r.nama, kelas: r.kelas, jam: r.jam, status: r.status, pulang: r.pulang, cat: r.cat }));
  return { ok: true, rows: rows };
}

/* ---------- daftar wajah ---------- */
function daftar_(p) {
  if (!pinOk_(p.pin)) return { ok: false, error: 'PIN salah' };
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

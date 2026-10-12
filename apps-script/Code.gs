/**
 * SILAB Gizi — Backend Google Apps Script
 * Jurusan Gizi, Poltekkes Kemenkes Makassar
 *
 * CARA PASANG (ringkas, lengkapnya di README):
 * 1. Buat Google Sheet baru → Ekstensi → Apps Script → tempel seluruh file ini.
 * 2. Isi CONFIG.LOGO_URL (opsional, untuk logo di PDF).
 * 3. Jalankan fungsi setup() sekali (izinkan akses).
 * 4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 * 5. Salin URL /exec ke CONFIG.API_URL di index.html.
 */

const CONFIG = {
  APP_NAME: 'SILAB Gizi',
  INSTANSI: 'POLITEKNIK KESEHATAN KEMENKES MAKASSAR',
  JURUSAN: 'JURUSAN GIZI',
  ALAMAT: '',          // alamat/telepon jurusan untuk kop PDF
  LOGO_URL: '',        // mis. https://akun-anda.github.io/silab-gizi/logo.png
  TOKEN_HOURS: 24 * 7, // lama sesi login
  MAKS_SALAH: 5,       // salah login 5x → dikunci
  KUNCI_MENIT: 15
};

/* ---------- Skema sheet (urutan kolom = urutan di Google Sheet) ---------- */
const SCHEMA = {
  Lab: ['id_lab', 'nama_lab', 'plp', 'no_wa', 'mode_jadwal', 'kapasitas', 'link_sop'], // mode_jadwal: sesi | jam
  User: ['id_user', 'nama', 'peran', 'id_lab', 'prodi', 'no_hp', 'pin'],          // staf: dosen, PLP, admin
  Mahasiswa: ['no', 'nim', 'nama', 'jk', 'prodi', 'kelas', 'tgl_lahir'],          // = master jurusan; masuk dengan NIM + tahun lahir
  Alat: ['kode', 'id_lab', 'nama', 'merk', 'satuan', 'jml_total', 'jml_rusak', 'lokasi', 'tahun', 'foto_url', 'status', 'catatan'],
  Bahan: ['kode', 'id_lab', 'nama', 'satuan', 'stok', 'stok_min', 'exp_date', 'b3', 'lokasi', 'status'],
  MutasiBahan: ['id', 'tgl', 'kode_bahan', 'jenis', 'jumlah', 'keperluan', 'oleh'],
  Peminjaman: ['id', 'tgl_ajuan', 'id_user', 'nama', 'prodi', 'no_hp', 'id_lab', 'keperluan', 'mata_kuliah', 'tgl_pinjam', 'tgl_kembali', 'status',
    'tgl_disetujui', 'disetujui_oleh', 'alasan_tolak', 'tgl_diambil', 'tgl_dikembalikan', 'catatan_mhs', 'kondisi_kembali', 'catatan_plp', 'tgl_selesai', 'link_pdf'],
  DetailPinjam: ['id_pinjam', 'kode_alat', 'nama_alat', 'jml', 'jml_rusak'],
  PenggunaanLab: ['id', 'id_lab', 'tanggal', 'sesi', 'jam_mulai', 'jam_selesai', 'kegiatan', 'kelas', 'dosen', 'resep', 'id_user', 'nama', 'jml_peserta', 'status', 'tgl_ajuan', 'catatan']
};
/** Judul kolom yang ditulis di Google Sheet (bila berbeda dari nama kunci) */
const HEADER = { Mahasiswa: ['No.', 'NIM', 'Nama', 'JK', 'Program Studi', 'Kelas Perkuliahan', 'Tanggal Lahir'] };
const NUMERIC = ['jml_total', 'jml_rusak', 'stok', 'stok_min', 'jumlah', 'jml', 'jml_peserta', 'kapasitas'];
/** Sesi untuk lab ber-mode 'sesi' (Lab PM/MSPMI) */
const SESI = { 'Sesi 1': ['08:00', '12:00'], 'Sesi 2': ['13:00', '17:00'], 'Seharian': ['08:00', '17:00'] };
const LABS = [
  ['KONS', 'Lab. Konsultasi Gizi'], ['PMM', 'Lab. Penyelenggaraan Makanan & MSPMI'], ['KIM', 'Lab. Kimia'], ['MIK', 'Lab. Mikrobiologi'],
  ['PSG', 'Lab. Penilaian Status Gizi'], ['ORG', 'Lab. Organoleptik/Citarasa'], ['PGN', 'Lab. Pangan']
];

/* ===================== SETUP & MENU ===================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('SILAB Gizi')
    .addItem('1. Setup awal (sekali saja)', 'setup')
    .addItem('2. Lengkapi kode alat/bahan hasil impor', 'lengkapiKode')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([HEADER[name] || SCHEMA[name]]).setFontWeight('bold').setBackground('#D6F1EE');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, SCHEMA[name].length).setNumberFormat('@'); // simpan sebagai teks agar tanggal tidak berubah
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (table('Lab').length === 0) LABS.forEach(([id, nama]) => append('Lab', { id_lab: id, nama_lab: nama, plp: '(isi nama PLP)', no_wa: '', mode_jadwal: id === 'PMM' ? 'sesi' : 'jam', kapasitas: 50, link_sop: '' }));
  if (table('User').length === 0) append('User', { id_user: 'ADMIN', nama: 'Kepala Laboratorium', peran: 'admin', id_lab: '', prodi: '', no_hp: '', pin: '1234' });

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('FOLDER_PDF')) props.setProperty('FOLDER_PDF', DriveApp.createFolder('SILAB Gizi — Arsip Peminjaman').getId());
  if (!props.getProperty('FOLDER_FOTO')) {
    const f = DriveApp.createFolder('SILAB Gizi — Foto Alat');
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('FOLDER_FOTO', f.getId());
  }
  SpreadsheetApp.getActive().toast('Setup selesai. Login admin: ADMIN / PIN 1234 (segera ganti PIN di sheet User).', 'SILAB Gizi', 10);
}

/** Isi kode & status untuk baris hasil impor template Excel yang kodenya masih kosong */
function lengkapiKode() {
  let n = 0;
  [['Alat', 'ALT'], ['Bahan', 'BHN']].forEach(([name, tag]) => {
    const rows = table(name);
    rows.forEach(r => {
      if (r.kode || !r.id_lab || !r.nama) return;
      r.kode = nextId(rows, `${r.id_lab}-${tag}-`, 'kode');
      if (!r.status) r.status = 'Aktif';
      if (name === 'Alat' && r.jml_rusak === '') r.jml_rusak = 0;
      update(name, r); n++;
    });
  });
  SpreadsheetApp.getActive().toast(`${n} baris dilengkapi kodenya.`, 'SILAB Gizi', 6);
}

/* ===================== WEB APP ===================== */
function doGet() {
  return json({ ok: true, data: { app: CONFIG.APP_NAME, status: 'aktif' } });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const h = HANDLERS[req.action];
    if (!h) throw new Error('Aksi tidak dikenal: ' + req.action);
    let user = null;
    if (req.action !== 'login' && req.action !== 'publik') {
      user = verifyToken(req.token);
      if (!user) return json({ ok: false, code: 'AUTH', error: 'Sesi berakhir, silakan masuk lagi.' });
    }
    const write = ['login', 'bootstrap', 'publik'].indexOf(req.action) < 0;
    const lock = LockService.getScriptLock();
    if (write) lock.waitLock(20000);
    try { return json({ ok: true, data: h(req.payload || {}, user) }); }
    finally { if (write) lock.releaseLock(); }
  } catch (err) {
    return json({ ok: false, error: err.message || String(err) });
  }
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ===================== TOKEN ===================== */
function sign(s) {
  const key = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key));
}
function makeToken(id) {
  const exp = Date.now() + CONFIG.TOKEN_HOURS * 3600e3;
  const body = id + '|' + exp;
  return Utilities.base64EncodeWebSafe(body) + '.' + sign(body);
}
function verifyToken(tok) {
  if (!tok || tok.indexOf('.') < 0) return null;
  const [b, sig] = tok.split('.');
  const body = Utilities.newBlob(Utilities.base64DecodeWebSafe(b)).getDataAsString();
  if (sign(body) !== sig) return null;
  const [id, exp] = body.split('|');
  if (Date.now() > Number(exp)) return null;
  const staf = table('User').find(u => String(u.id_user) === id);
  if (staf) return staf;
  const m = table('Mahasiswa').find(x => normNim(x.nim) === id);
  return m ? asUser(m) : null;
}

/* ===================== AKSES SHEET ===================== */
function sheet(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function table(name) {
  const sh = sheet(name); const cols = SCHEMA[name];
  const last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getDisplayValues().map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c] = NUMERIC.indexOf(c) >= 0 ? (r[j] === '' ? '' : Number(r[j])) : r[j]; });
    return o;
  }).filter(o => cols.some(c => o[c] !== '') && !(name === 'Mahasiswa' && normNim(o.nim) === 'NIM'));
}
function append(name, obj) {
  sheet(name).appendRow(SCHEMA[name].map(c => obj[c] === undefined || obj[c] === null ? '' : String(obj[c])));
}
function update(name, obj) {
  sheet(name).getRange(obj._row, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name].map(c => obj[c] === undefined || obj[c] === null ? '' : String(obj[c]))]);
}
const pad = n => String(n).padStart(2, '0');
function now() { return Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm'); }
function nextId(rows, prefix, field) {
  field = field || 'id'; let max = 0;
  rows.forEach(r => { const v = String(r[field]); const m = v.match(/(\d+)$/); if (m && v.indexOf(prefix) === 0) max = Math.max(max, +m[1]); });
  return prefix + String(max + 1).padStart(3, '0');
}
function fail(m) { throw new Error(m); }
/** Kunci login setelah beberapa kali salah */
function cekKunci(key) {
  const c = CacheService.getScriptCache().get('kunci_' + key);
  if (c) fail('Terlalu banyak percobaan. Coba lagi ' + CONFIG.KUNCI_MENIT + ' menit lagi.');
}
function catatGagal(key, msg) {
  const cache = CacheService.getScriptCache(); const n = Number(cache.get('gagal_' + key) || 0) + 1;
  if (n >= CONFIG.MAKS_SALAH) { cache.put('kunci_' + key, '1', CONFIG.KUNCI_MENIT * 60); cache.remove('gagal_' + key); }
  else cache.put('gagal_' + key, String(n), CONFIG.KUNCI_MENIT * 60);
  fail(msg);
}
/** Samarkan NIM untuk tampilan publik: P07131224001 -> P0713****001 */
function mask(v) { v = String(v || ''); return v.length > 8 ? v.slice(0, 5) + '****' + v.slice(-3) : v.replace(/.(?=.{2})/g, '*'); }
/** NIM dinormalkan: tanpa apostrof/spasi, huruf besar, huruf O = angka 0 (PO713… = P0713…) */
function normNim(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/O/g, '0'); }
function asUser(m) { return { id_user: normNim(m.nim), nama: m.nama, peran: 'mahasiswa', id_lab: '', prodi: m.prodi, kelas: m.kelas, no_hp: '' }; }
/** Tahun lahir dari kolom Tanggal Lahir, format apa pun (17/08/2005, 2005-08-17, 17 Agustus 2005) */
function tahunLahir(t) { const m = String(t || '').match(/(19|20)\d{2}/); return m ? m[0] : ''; }
function publicUser(u) { const o = {}; Object.keys(u).forEach(k => { if (k !== 'pin' && k !== '_row') o[k] = u[k]; }); return o; }
function clean(rows) { return rows.map(r => { const o = Object.assign({}, r); delete o._row; return o; }); }
function terpakai(pinjam, detail, kode, exceptId) {
  const aktif = {}; pinjam.forEach(p => { if (['Disetujui', 'Dipinjam', 'Dikembalikan'].indexOf(p.status) >= 0 && p.id !== exceptId) aktif[p.id] = 1; });
  return detail.filter(d => d.kode_alat === kode && aktif[d.id_pinjam]).reduce((s, d) => s + Number(d.jml), 0);
}

/* ===================== HANDLER (sama dengan mode demo) ===================== */
const HANDLERS = {
  login(p) {
    const key = String(p.id_user || '').trim().toLowerCase();
    cekKunci(key);
    if (p.mode === 'mahasiswa') {
      const m = table('Mahasiswa').find(x => normNim(x.nim) === normNim(key));
      if (!m) catatGagal(key, 'NIM tidak terdaftar. Hubungi admin laboratorium.');
      if (tahunLahir(m.tgl_lahir) !== String(p.pin || '').trim()) catatGagal(key, 'NIM atau tahun lahir tidak cocok.');
      CacheService.getScriptCache().remove('gagal_' + key);
      return { user: asUser(m), token: makeToken(normNim(m.nim)) };
    }
    const u = table('User').find(x => String(x.id_user).toLowerCase() === key);
    if (!u || String(u.pin) !== String(p.pin)) catatGagal(key, 'NIP/ID atau PIN salah.');
    CacheService.getScriptCache().remove('gagal_' + key);
    return { user: publicUser(u), token: makeToken(u.id_user) };
  },

  /** Data halaman depan (tanpa login). NIM disamarkan. */
  publik() {
    const t = now().slice(0, 10);
    return {
      labs: table('Lab').map(l => ({ id_lab: l.id_lab, nama_lab: l.nama_lab, plp: l.plp, mode_jadwal: l.mode_jadwal, kapasitas: l.kapasitas, link_sop: l.link_sop })),
      jadwal: table('PenggunaanLab').filter(b => b.status === 'Disetujui' && b.tanggal >= t)
        .sort((a, b) => (a.tanggal + a.jam_mulai).localeCompare(b.tanggal + b.jam_mulai)).slice(0, 8)
        .map(b => ({ id_lab: b.id_lab, tanggal: b.tanggal, sesi: b.sesi, jam_mulai: b.jam_mulai, jam_selesai: b.jam_selesai, kegiatan: b.kegiatan, nama: b.nama, id_user: mask(b.id_user) })),
      jml_alat: table('Alat').filter(a => a.status !== 'Nonaktif').length
    };
  },

  bootstrap(p, u) {
    const pinjamAll = table('Peminjaman'), detailAll = table('DetailPinjam'), alat = table('Alat');
    const out = { labs: clean(table('Lab')), alat: clean(alat.filter(a => a.status !== 'Nonaktif')), server_time: now() };
    const booking = clean(table('PenggunaanLab'));
    // mahasiswa hanya melihat NIM orang lain dalam bentuk tersamar
    out.booking = u.peran === 'mahasiswa' ? booking.map(b => { if (b.id_user !== u.id_user) b.id_user = mask(b.id_user); return b; }) : booking;
    if (u.peran === 'mahasiswa') { out.pinjam = pinjamAll.filter(x => x.id_user === u.id_user); out.bahan = []; out.mutasi = []; }
    else {
      out.pinjam = u.peran === 'plp' ? pinjamAll.filter(x => x.id_lab === u.id_lab) : pinjamAll;
      out.bahan = clean(table('Bahan').filter(b => b.status !== 'Nonaktif'));
      out.mutasi = clean(table('MutasiBahan'));
    }
    const ids = {}; out.pinjam.forEach(x => ids[x.id] = 1);
    out.pinjam = clean(out.pinjam);
    out.detail = clean(detailAll.filter(d => ids[d.id_pinjam]));
    out.terpakai = {}; alat.forEach(a => out.terpakai[a.kode] = terpakai(pinjamAll, detailAll, a.kode));
    return out;
  },

  ajukanPinjam(p, u) {
    if (!p.items || !p.items.length) fail('Pilih minimal satu alat.');
    const pinjam = table('Peminjaman'), detail = table('DetailPinjam'), alat = table('Alat');
    const hari = now().slice(0, 10);
    if (pinjam.some(x => x.id_user === u.id_user && x.status === 'Dipinjam' && x.tgl_kembali < hari)) fail('Masih ada pinjaman yang terlambat dikembalikan. Kembalikan dulu sebelum mengajukan lagi.');
    if (!p.tgl_pinjam || !p.tgl_kembali || p.tgl_kembali < p.tgl_pinjam) fail('Periksa tanggal pinjam dan tanggal kembali.');
    p.items.forEach(it => {
      const a = alat.find(x => x.kode === it.kode_alat); if (!a) fail('Alat tidak ditemukan: ' + it.kode_alat);
      if (a.id_lab !== p.id_lab) fail('Semua alat harus dari lab yang sama.');
      const sisa = a.jml_total - a.jml_rusak - terpakai(pinjam, detail, a.kode);
      if (Number(it.jml) > sisa) fail('Stok ' + a.nama + ' tidak cukup (tersedia ' + sisa + ').');
    });
    const id = nextId(pinjam, 'PJM-' + Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyMM') + '-');
    append('Peminjaman', {
      id, tgl_ajuan: now(), id_user: u.id_user, nama: u.nama, prodi: u.prodi, no_hp: p.no_hp || u.no_hp, id_lab: p.id_lab,
      keperluan: p.keperluan, mata_kuliah: p.mata_kuliah, tgl_pinjam: p.tgl_pinjam, tgl_kembali: p.tgl_kembali, status: 'Menunggu'
    });
    p.items.forEach(it => { const a = alat.find(x => x.kode === it.kode_alat); append('DetailPinjam', { id_pinjam: id, kode_alat: a.kode, nama_alat: a.nama, jml: Number(it.jml), jml_rusak: 0 }); });
    return { id };
  },

  aksiPinjam(p, u) {
    const pinjam = table('Peminjaman'); const r = pinjam.find(x => x.id === p.id); if (!r) fail('Permohonan tidak ditemukan.');
    const isStaff = u.peran === 'admin' || (u.peran === 'plp' && u.id_lab === r.id_lab);
    const own = r.id_user === u.id_user;
    const need = (c, m) => { if (!c) fail(m); };
    const detail = table('DetailPinjam'); const mine = detail.filter(d => d.id_pinjam === r.id);
    switch (p.aksi) {
      case 'batal': need(own && r.status === 'Menunggu', 'Permohonan tidak bisa dibatalkan.'); r.status = 'Dibatalkan'; break;
      case 'setujui': {
        need(isStaff && r.status === 'Menunggu', 'Aksi tidak diizinkan.');
        const alat = table('Alat');
        mine.forEach(d => { const a = alat.find(x => x.kode === d.kode_alat); const sisa = a.jml_total - a.jml_rusak - terpakai(pinjam, detail, a.kode, r.id); if (d.jml > sisa) fail('Stok ' + a.nama + ' tidak cukup (tersedia ' + sisa + ').'); });
        r.status = 'Disetujui'; r.tgl_disetujui = now(); r.disetujui_oleh = u.nama;
        r.link_pdf = buatPdf(r, mine); break;
      }
      case 'tolak': need(isStaff && r.status === 'Menunggu', 'Aksi tidak diizinkan.'); need(p.alasan, 'Tuliskan alasan penolakan.'); r.status = 'Ditolak'; r.alasan_tolak = p.alasan; r.disetujui_oleh = u.nama; r.tgl_disetujui = now(); break;
      case 'serahkan': need(isStaff && r.status === 'Disetujui', 'Aksi tidak diizinkan.'); r.status = 'Dipinjam'; r.tgl_diambil = now(); break;
      case 'kembalikan': need(own && r.status === 'Dipinjam', 'Aksi tidak diizinkan.'); r.status = 'Dikembalikan'; r.tgl_dikembalikan = now(); r.catatan_mhs = p.catatan || ''; break;
      case 'cek': {
        need(isStaff && ['Dikembalikan', 'Dipinjam'].indexOf(r.status) >= 0, 'Aksi tidak diizinkan.');
        if (!r.tgl_dikembalikan) r.tgl_dikembalikan = now();
        const alat = table('Alat'); let adaRusak = false;
        mine.forEach(d => {
          const rusak = Math.max(0, Math.min(Number((p.rusak || {})[d.kode_alat] || 0), d.jml));
          d.jml_rusak = rusak; update('DetailPinjam', d);
          if (rusak > 0) { adaRusak = true; const a = alat.find(x => x.kode === d.kode_alat); a.jml_rusak = Number(a.jml_rusak || 0) + rusak; a.catatan = rusak + ' rusak saat ' + r.id; update('Alat', a); }
        });
        r.kondisi_kembali = adaRusak ? 'Ada kerusakan' : 'Baik'; r.catatan_plp = p.catatan || ''; r.tgl_selesai = now();
        r.status = adaRusak ? 'Selesai dgn Catatan' : 'Selesai';
        r.link_pdf = buatPdf(r, mine); break;
      }
      default: fail('Aksi tidak dikenal.');
    }
    update('Peminjaman', r);
    return { ok: true, link_pdf: r.link_pdf };
  },

  simpanAlat(p, u) {
    if (u.peran === 'mahasiswa') fail('Tidak diizinkan.');
    const a = p.alat; const lab = u.peran === 'plp' ? u.id_lab : a.id_lab;
    const rows = table('Alat');
    const foto = p.foto ? simpanFoto(p.foto, a.nama) : '';
    if (a.kode) {
      const x = rows.find(r => r.kode === a.kode); if (!x) fail('Alat tidak ditemukan.');
      if (u.peran === 'plp' && x.id_lab !== u.id_lab) fail('Bukan alat lab Anda.');
      ['nama', 'merk', 'satuan', 'jml_total', 'jml_rusak', 'lokasi', 'tahun', 'status', 'catatan'].forEach(k => x[k] = a[k]);
      if (foto) x.foto_url = foto;
      update('Alat', x); return { kode: x.kode };
    }
    if (!a.nama || !lab) fail('Nama alat dan lab wajib diisi.');
    const kode = nextId(rows, lab + '-ALT-', 'kode');
    append('Alat', { kode, id_lab: lab, nama: a.nama, merk: a.merk, satuan: a.satuan || 'unit', jml_total: a.jml_total || 1, jml_rusak: a.jml_rusak || 0, lokasi: a.lokasi, tahun: a.tahun, foto_url: foto, status: 'Aktif', catatan: a.catatan });
    return { kode };
  },

  simpanBahan(p, u) {
    if (u.peran === 'mahasiswa') fail('Tidak diizinkan.');
    const b = p.bahan; const lab = u.peran === 'plp' ? u.id_lab : b.id_lab;
    const rows = table('Bahan');
    if (b.kode) {
      const x = rows.find(r => r.kode === b.kode); if (!x) fail('Bahan tidak ditemukan.');
      if (u.peran === 'plp' && x.id_lab !== u.id_lab) fail('Bukan bahan lab Anda.');
      ['nama', 'satuan', 'stok_min', 'exp_date', 'b3', 'lokasi'].forEach(k => x[k] = b[k]);
      update('Bahan', x); return { kode: x.kode };
    }
    if (!b.nama || !lab) fail('Nama bahan dan lab wajib diisi.');
    const kode = nextId(rows, lab + '-BHN-', 'kode');
    append('Bahan', { kode, id_lab: lab, nama: b.nama, satuan: b.satuan, stok: b.stok || 0, stok_min: b.stok_min || 0, exp_date: b.exp_date, b3: b.b3 || 'Tidak', lokasi: b.lokasi, status: 'Aktif' });
    if (Number(b.stok) > 0) append('MutasiBahan', { id: nextId(table('MutasiBahan'), 'MUT-'), tgl: now(), kode_bahan: kode, jenis: 'Masuk', jumlah: b.stok, keperluan: 'Stok awal', oleh: u.nama });
    return { kode };
  },

  mutasiBahan(p, u) {
    if (u.peran === 'mahasiswa') fail('Tidak diizinkan.');
    const b = table('Bahan').find(x => x.kode === p.kode_bahan); if (!b) fail('Bahan tidak ditemukan.');
    if (u.peran === 'plp' && b.id_lab !== u.id_lab) fail('Bukan bahan lab Anda.');
    const j = Number(p.jumlah); if (!(j > 0)) fail('Jumlah harus lebih dari 0.');
    if (p.jenis === 'Keluar' && j > b.stok) fail('Stok tidak cukup (sisa ' + b.stok + ' ' + b.satuan + ').');
    b.stok = p.jenis === 'Masuk' ? Number(b.stok) + j : Number(b.stok) - j;
    if (p.jenis === 'Masuk' && p.exp_date) b.exp_date = p.exp_date;
    update('Bahan', b);
    append('MutasiBahan', { id: nextId(table('MutasiBahan'), 'MUT-'), tgl: now(), kode_bahan: b.kode, jenis: p.jenis, jumlah: j, keperluan: p.keperluan, oleh: u.nama });
    return { stok: b.stok };
  },

  ajukanBooking(p, u) {
    const lab = table('Lab').find(l => l.id_lab === p.id_lab); if (!lab) fail('Pilih laboratorium.');
    if (!p.tanggal || !p.kegiatan || !p.kelas || !p.dosen) fail('Lengkapi tanggal, mata kuliah, kelas, dan dosen pengampu.');
    let jm = p.jam_mulai, js = p.jam_selesai, sesi = '';
    if (lab.mode_jadwal === 'sesi') { if (!SESI[p.sesi]) fail('Pilih sesi praktikum.'); sesi = p.sesi; jm = SESI[sesi][0]; js = SESI[sesi][1]; }
    else { if (!jm || !js) fail('Isi jam mulai dan selesai.'); if (js <= jm) fail('Jam selesai harus setelah jam mulai.'); }
    const n = Number(p.jml_peserta) || 0, kap = Number(lab.kapasitas) || 50;
    if (n < 1) fail('Isi jumlah peserta.');
    if (n > kap) fail('Kapasitas ' + lab.nama_lab + ' maksimal ' + kap + ' orang.');
    const staff = u.peran !== 'mahasiswa'; const tgls = [p.tanggal];
    if (staff && p.ulang_sampai) { // jadwal reguler: ulangi tiap minggu
      let d = new Date(p.tanggal + 'T00:00:00');
      for (let i = 0; i < 30; i++) { d.setDate(d.getDate() + 7); const t = Utilities.formatDate(d, 'Asia/Makassar', 'yyyy-MM-dd'); if (t > p.ulang_sampai) break; tgls.push(t); }
    }
    const rows = table('PenggunaanLab'); const dibuat = [], dilewati = [];
    tgls.forEach(t => {
      const b = rows.find(x => x.id_lab === p.id_lab && x.tanggal === t && ['Menunggu', 'Disetujui'].indexOf(x.status) >= 0 && jm < x.jam_selesai && x.jam_mulai < js);
      if (b) { if (tgls.length === 1) fail('Bentrok dengan "' + b.kegiatan + '" (' + (b.sesi || b.jam_mulai + '–' + b.jam_selesai) + ').'); dilewati.push(t); return; }
      const rec = { id: nextId(rows, 'BKG-'), id_lab: p.id_lab, tanggal: t, sesi, jam_mulai: jm, jam_selesai: js, kegiatan: p.kegiatan, kelas: p.kelas, dosen: p.dosen,
        resep: lab.mode_jadwal === 'sesi' ? (p.resep || '') : '', id_user: u.id_user, nama: u.nama, jml_peserta: n, status: staff ? 'Disetujui' : 'Menunggu', tgl_ajuan: now(), catatan: '' };
      append('PenggunaanLab', rec); rows.push(rec); dibuat.push(t);
    });
    if (!dibuat.length) fail('Semua tanggal bentrok dengan jadwal lain.');
    return { dibuat, dilewati };
  },

  aksiBooking(p, u) {
    const b = table('PenggunaanLab').find(x => x.id === p.id); if (!b) fail('Jadwal tidak ditemukan.');
    const isStaff = u.peran === 'admin' || (u.peran === 'plp' && u.id_lab === b.id_lab);
    if (p.aksi === 'batal' && b.id_user === u.id_user && b.status === 'Menunggu') b.status = 'Dibatalkan';
    else if (isStaff && p.aksi === 'setujui') b.status = 'Disetujui';
    else if (isStaff && p.aksi === 'tolak') { b.status = 'Ditolak'; b.catatan = p.alasan || ''; }
    else fail('Aksi tidak diizinkan.');
    update('PenggunaanLab', b); return { ok: true };
  }
};

/* ===================== FOTO & PDF (Google Drive) ===================== */
function simpanFoto(dataUrl, nama) {
  const m = String(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/); if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nama + '.jpg');
  const f = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_FOTO')).createFile(blob);
  return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w480';
}

function subFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function logoData() {
  if (!CONFIG.LOGO_URL) return '';
  const cache = CacheService.getScriptCache(); const c = cache.get('logo'); if (c) return c;
  try { const b = UrlFetchApp.fetch(CONFIG.LOGO_URL).getBlob(); const d = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); if (d.length < 95000) cache.put('logo', d, 21600); return d; }
  catch (e) { return ''; }
}

function tgl(s, jam) {
  if (!s) return '—'; const B = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const [d, t] = String(s).split(' '); const [y, m, dd] = d.split('-');
  return Number(dd) + ' ' + B[Number(m) - 1] + ' ' + y + (jam && t ? ', ' + t : '');
}
function e(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

/** Membuat/memperbarui PDF bukti di Drive: Arsip/<Lab>/<yyyy-MM>/Bukti-<id>.pdf */
function buatPdf(r, items) {
  const lab = table('Lab').find(l => l.id_lab === r.id_lab) || { nama_lab: r.id_lab, plp: '' };
  const selesai = !!r.tgl_selesai; const logo = logoData();
  const row = (k, v) => '<tr><td class="k">' + k + '</td><td><b>' + e(v) + '</b></td></tr>';
  const html = '<html><head><style>' +
    'body{font-family:Arial,sans-serif;font-size:11px;color:#111;margin:28px 36px}' +
    '.kop{border-bottom:3px solid #00988F;margin-bottom:0}.kop td{padding:0 0 6px 0;vertical-align:middle}' +
    '.kop img{height:46px}.kop .r{text-align:right}.kop .j{color:#00988F;font-weight:bold}' +
    '.lime{border-top:1.5px solid #B9C91C;margin-top:2px}h1{font-size:14px;text-align:center;margin:16px 0 2px}.no{text-align:center;margin-bottom:12px}' +
    'table{border-collapse:collapse;width:100%;margin-bottom:10px}td,th{padding:4px 6px}.k{color:#555;width:150px}' +
    '.it th{background:#00988F;color:#fff;text-align:left}.it td{border-bottom:1px solid #ddd}' +
    '.ttd{margin-top:26px}.ttd td{text-align:center;width:50%}.el{color:#007A73;font-size:10px;margin:10px 0}' +
    '.foot{margin-top:30px;color:#888;font-size:9px;text-align:center}</style></head><body>' +
    '<table class="kop"><tr><td>' + (logo ? '<img src="' + logo + '">' : '<b>Kemenkes Poltekkes Makassar</b>') +
    '</td><td class="r"><b>' + e(CONFIG.INSTANSI) + '</b><br><span class="j">' + e(CONFIG.JURUSAN) + '</span><br>' + e(lab.nama_lab) + (CONFIG.ALAMAT ? '<br>' + e(CONFIG.ALAMAT) : '') + '</td></tr></table><div class="lime"></div>' +
    '<h1>' + (selesai ? 'BUKTI PEMINJAMAN DAN PENGEMBALIAN ALAT' : 'BUKTI PEMINJAMAN ALAT LABORATORIUM') + '</h1><div class="no">Nomor: ' + e(r.id) + '</div>' +
    '<table>' + row('Nama peminjam', r.nama) + row('NIM / NIP', r.id_user) + row('Program studi', r.prodi || '-') + row('No. HP', r.no_hp) +
    row('Laboratorium', lab.nama_lab) + row('Keperluan', r.keperluan + ' — ' + r.mata_kuliah) + row('Tanggal pinjam', tgl(r.tgl_pinjam)) + row('Batas kembali', tgl(r.tgl_kembali)) + '</table>' +
    '<table class="it"><tr><th>No</th><th>Kode</th><th>Nama Alat</th><th>Jumlah</th>' + (selesai ? '<th>Kondisi Kembali</th>' : '') + '</tr>' +
    items.map((i, n) => '<tr><td>' + (n + 1) + '</td><td>' + e(i.kode_alat) + '</td><td>' + e(i.nama_alat) + '</td><td>' + i.jml + '</td>' + (selesai ? '<td>' + (i.jml_rusak ? i.jml_rusak + ' rusak/hilang' : 'Baik') + '</td>' : '') + '</tr>').join('') + '</table>' +
    '<table>' + row('Disetujui', tgl(r.tgl_disetujui, true) + ' oleh ' + r.disetujui_oleh) + (r.tgl_diambil ? row('Diambil', tgl(r.tgl_diambil, true)) : '') +
    (selesai ? row('Dikembalikan', tgl(r.tgl_dikembalikan, true)) + row('Dicek & selesai', tgl(r.tgl_selesai, true)) + row('Kondisi', r.kondisi_kembali + (r.catatan_plp ? ' — ' + r.catatan_plp : '')) : '') + '</table>' +
    '<table class="ttd"><tr><td>Peminjam,<div class="el">Diajukan elektronik<br>' + tgl(r.tgl_ajuan, true) + '</div><b>' + e(r.nama) + '</b></td>' +
    '<td>PLP / Laboran,<div class="el">Disetujui elektronik<br>' + tgl(r.tgl_disetujui, true) + '</div><b>' + e(r.disetujui_oleh || lab.plp) + '</b></td></tr></table>' +
    '<div class="foot">Dokumen dibuat otomatis oleh ' + e(CONFIG.APP_NAME) + ' — ' + e(CONFIG.JURUSAN) + '. Sah tanpa tanda tangan basah.</div></body></html>';

  const pdf = Utilities.newBlob(html, 'text/html', 'tmp.html').getAs('application/pdf').setName('Bukti-' + r.id + '.pdf');
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_PDF'));
  const folder = subFolder(subFolder(root, lab.nama_lab), String(r.tgl_ajuan).slice(0, 7));
  const old = folder.getFilesByName(pdf.getName()); while (old.hasNext()) old.next().setTrashed(true);
  return folder.createFile(pdf).getUrl();
}

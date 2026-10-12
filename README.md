# SILAB Gizi
Sistem Informasi Laboratorium — Jurusan Gizi, Poltekkes Kemenkes Makassar.
PWA (dihosting di GitHub Pages) + Google Apps Script + Google Sheet sebagai database.

## Isi folder
| File | Fungsi |
|---|---|
| `index.html` | Seluruh aplikasi (tampilan + logika). Berjalan dalam **mode demo** selama `API_URL` kosong. |
| `manifest.json`, `sw.js`, `icons/` | Membuat aplikasi bisa di-*install* di HP (PWA). |
| `logo.png` | Logo untuk kop PDF di Google Drive. |
| `apps-script/Code.gs` | Backend: baca/tulis Google Sheet, buat PDF bukti ke Google Drive. |
| `Template-Impor-SILAB-Gizi.xlsx` | Template untuk memasukkan data mahasiswa (NIM & nama) serta inventaris alat & bahan. |

## Langkah 1 — Upload ke GitHub Pages
1. Buat repository baru di GitHub, mis. `silab-gizi` (Public).
2. Upload semua file di folder ini (kecuali boleh juga ikut: folder `apps-script` dan file template).
3. Buka **Settings → Pages → Branch: main / root → Save**.
4. Tunggu ±1 menit. Aplikasi bisa dibuka di `https://<nama-akun>.github.io/silab-gizi/` (masih mode demo).

## Langkah 2 — Siapkan Google Sheet & Apps Script
1. Pakai akun Google jurusan. Buat Google Sheet baru, beri nama **SILAB Gizi — Database**.
2. Menu **Ekstensi → Apps Script**. Hapus isi `Code.gs`, tempel seluruh isi `apps-script/Code.gs`.
3. Di bagian `CONFIG`, isi:
   - `LOGO_URL: 'https://<nama-akun>.github.io/silab-gizi/logo.png'`
   - `ALAMAT` (opsional) untuk kop PDF.
4. Simpan, pilih fungsi **setup** di toolbar, klik **Run**, lalu izinkan akses (Sheet & Drive).
   Hasilnya: tab-tab database terbentuk, 7 lab terisi, folder Drive **"SILAB Gizi — Arsip Peminjaman"** dan **"SILAB Gizi — Foto Alat"** dibuat.
5. **Deploy → New deployment → Select type: Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Klik Deploy, salin **Web app URL** (berakhiran `/exec`).

## Langkah 3 — Hubungkan aplikasi
Di GitHub, edit `index.html`, cari baris:
```js
API_URL: '',
```
ganti menjadi:
```js
API_URL: 'https://script.google.com/macros/s/XXXX/exec',
```
Commit. Mode demo hilang, aplikasi kini memakai Google Sheet.

## Langkah 4 — Isi data awal
- **Lab** (tab `Lab`): isi nama PLP dan nomor WhatsApp (format `62812…`).
- **Staf** (tab `User`): tambahkan PLP (`peran = plp`, `id_lab = PSG` dsb.), dosen/admin, dan ganti PIN akun `ADMIN`. Staf masuk dengan NIP + PIN.
- **Mahasiswa** (tab `Mahasiswa`): kolomnya sama dengan master jurusan (No., NIM, Nama, JK, Program Studi, Kelas Perkuliahan, Tanggal Lahir), jadi data master bisa langsung ditempel. Mahasiswa masuk dengan **NIM + tahun lahir** (diambil otomatis dari Tanggal Lahir). Salah 5× → dikunci 15 menit. Nomor HP tidak disimpan di master; diisi mahasiswa di setiap form peminjaman.
- **Pengaturan lab** (tab `Lab`): `mode_jadwal` = `sesi` (Lab PM/MSPMI: Sesi 1, Sesi 2, Seharian) atau `jam` (lab lain); `kapasitas` (bawaan 50); `link_sop` = link PDF SOP di Google Drive (bagikan "Siapa saja yang memiliki link"). Selama kosong, aplikasi menampilkan "SOP sedang disusun".
- **Inventaris**: isi `Template-Impor-SILAB-Gizi.xlsx`, tempel ke tab `Alat` dan `Bahan`, lalu klik menu **SILAB Gizi → Lengkapi kode alat/bahan hasil impor**.

## Halaman depan & privasi
Aplikasi selalu dibuka dari **halaman depan** (deskripsi, cara penggunaan, ketentuan, daftar lab, jadwal lab terdekat). Data yang tampil tanpa login hanya berisi nama dan **NIM tersamar** (mis. `P0713****001`). NIM lengkap hanya terlihat oleh pemiliknya dan oleh PLP/admin.

## Jadwal lab
- Satu sesi/jam hanya untuk satu kelas; jadwal bentrok ditolak otomatis, jumlah peserta tidak boleh melebihi kapasitas.
- Isian: mata kuliah/topik, kelas, dosen pengampu, jumlah peserta; khusus Lab PM ditambah resep/topik menu.
- PLP bisa membuat **jadwal reguler** (ulang tiap minggu sampai tanggal tertentu); tanggal yang bentrok dilewati.

## Alur peminjaman
Diajukan → Disetujui PLP (PDF bukti dibuat di Drive) → Alat diambil → Mahasiswa klik **Kembalikan Alat** → PLP **Cek Pengembalian** → Selesai (PDF diperbarui menjadi bukti pengembalian).
Semua tahap tercatat di **satu baris** tab `Peminjaman`; rincian alat di tab `DetailPinjam`.

## Catatan
- Setiap kali `Code.gs` diubah: **Deploy → Manage deployments → Edit → Version: New version** (URL tetap sama).
- PIN disimpan apa adanya di sheet `User`; batasi akses Google Sheet hanya untuk admin.
- Foto alat dari HP otomatis diperkecil sebelum diunggah ke Drive.

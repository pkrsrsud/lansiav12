LANSIA SMART RSUD WONOSARI - GOOGLE EDITION v5

Perbaikan utama:
1. File Word (.docx) tidak perlu lagi memakai "Open with Google Docs" secara manual.
2. Apps Script dapat mengonversi file Word di Drive menjadi Google Docs native.
3. Tombol "Buat Google Docs" menghasilkan URL Google Docs yang benar.
4. Perbaikan format tanggal dashboard tetap termasuk dari v4.

SETUP

1) SPREADSHEET
- Buat Google Spreadsheet.
- Dari URL spreadsheet, salin ID saja (bagian setelah /d/ dan sebelum /edit).
- Isi SPREADSHEET_ID.

2) SECRET / ADMIN
- Isi API_SECRET.
- Isi ADMIN_PASSWORD.

3) TEMPLATE WORD
- Upload "Daftar Hadir 2024 - template Google-compatible.docx" ke Google Drive.
- Ambil ID file Word dari URL Drive.
- Isi DOCX_TEMPLATE_ID dengan ID tersebut.
- Biarkan DOC_TEMPLATE_ID kosong.

4) AKTIFKAN DRIVE API LANJUTAN
- Apps Script -> Services (+) -> pilih "Drive API" -> Add.
- appsscript.json v5 sudah menyertakan dependency Drive API v3.
- Google menyatakan Advanced Drive Service harus diaktifkan sebelum dipakai.

5) KONVERSI TEMPLATE
- Pilih fungsi importDocxTemplate.
- Klik Run.
- Izinkan akses jika diminta.
- Lihat Executions/Logs untuk URL Google Docs hasil konversi.
- DOC_TEMPLATE_ID otomatis disimpan ke Script Properties.

6) SIAPKAN PLACEHOLDER
- Pilih prepareGoogleDocTemplate.
- Run sekali.
- Contoh teks di template akan diganti menjadi {{HARI}}, {{TANGGAL}}, {{WAKTU}}, {{TEMPAT}}, {{ACARA}}.

7) BUAT SHEET OTOMATIS
- Pilih setupSpreadsheet.
- Run.
- Akan dibuat: Peserta, Kegiatan, Kehadiran.

8) DEPLOY
- Deploy -> New deployment -> Web app.
- Execute as: Me.
- Who has access: Anyone.
- Salin URL /exec ke GOOGLE_APPS_SCRIPT_URL.
- GOOGLE_APPS_SCRIPT_SECRET harus sama persis dengan API_SECRET.

CATATAN
- Jangan masukkan URL spreadsheet utuh ke SPREADSHEET_ID.
- Jangan masukkan ID Word ke DOC_TEMPLATE_ID. Gunakan DOCX_TEMPLATE_ID lalu importDocxTemplate.
- Export PDF tetap membuat file PDF di Drive.
- Kolom TANDA TANGAN HADIR dan MEWAKILI dibiarkan kosong sesuai kebutuhan.


IMPORT DATA PESERTA (v8)
- Masuk Admin > Peserta > Import peserta.
- Tempel URL Google Spreadsheet atau Google Docs.
- Spreadsheet: sebaiknya baris pertama berisi kolom Nama dan Alamat. URL dengan #gid=... akan memakai sheet tersebut; tanpa gid memakai sheet aktif.
- Google Docs: sebaiknya gunakan tabel dengan kolom NAMA dan ALAMAT.
- Sistem menambah peserta baru dan mencocokkan peserta lama berdasarkan nama lengkap yang dinormalisasi. Data lama yang tidak ada di sumber tidak dihapus.
- Opsi Perbarui alamat yang sudah ada dan Aktifkan kembali peserta lama dapat dimatikan bila diperlukan.
- Tidak perlu membuat service/API baru untuk import Spreadsheet/Docs. Akses mengikuti akun Google yang menjalankan Apps Script.
- Setelah mengganti Code.gs, deploy ulang Web App ke New version.

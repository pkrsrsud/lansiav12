# V12 CHANGELOG

- Memperbaiki crash dashboard `Cannot read properties of undefined (reading 'find')` dengan validasi/normalisasi respons `adminLoad`.
- Jika Apps Script yang dipakai salah versi atau mengembalikan struktur yang tidak lengkap, dashboard sekarang menampilkan pesan yang jelas alih-alih crash.
- Memperbaiki deklarasi `const sh` ganda pada `saveParticipant_()` di Apps Script.
- Tidak mengubah struktur Spreadsheet maupun alur fitur yang sudah berjalan.

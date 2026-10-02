# LANSIA SMART RSUD WONOSARI - Google Edition v7

Arsitektur: Browser -> Next.js/Vercel -> Next.js API -> Google Apps Script -> Google Spreadsheet / Google Docs.

v7 memperbaiki:
- tanggal Google Sheets yang dapat menyebabkan Runtime RangeError di dashboard
- hasil export yang sebelumnya menyalin DOCX tetapi membentuk URL Google Docs palsu
- proses template Word sekarang dapat dikonversi ke native Google Docs oleh Apps Script

Mulai dari `google-apps-script/README.txt`.

## Import peserta massal (v8)
Pada Admin > Peserta tersedia fitur Import Peserta. Sumber dapat berupa Google Spreadsheet atau Google Docs. Import tidak menghapus data peserta lama. Nama yang sama dicocokkan secara normalisasi; peserta baru ditambahkan, alamat peserta lama dapat diperbarui, dan peserta nonaktif dapat diaktifkan kembali sesuai opsi.


## v11 - Penomoran laporan dinamis
- Export laporan menggunakan blok tabel pertama sebagai master.
- Nomor peserta mengikuti `display_order` dan tidak lagi dibatasi 120 peserta.
- Baris tabel bertambah otomatis sesuai jumlah peserta hadir dan tabel dapat berlanjut ke halaman berikutnya.
- Kop/header dokumen, keterangan Hari/Tanggal/Waktu/Tempat/Acara, serta blok tanda tangan tetap dipertahankan.
- Kolom HADIR dan MEWAKILI tetap kosong.

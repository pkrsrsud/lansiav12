# v10 - Nomor Kehadiran Setelah Konfirmasi

Setelah peserta berhasil menekan **Ya, Saya Hadir**, halaman peserta sekarang menampilkan kartu konfirmasi yang berisi:

- **Nomor kehadiran** sesuai `display_order` dari Google Apps Script/Spreadsheet.
- **Nama peserta** tepat di bawah nomor.

Contoh:

**1.**
**Budi Santoso**

Nomor berasal dari urutan kehadiran pada kegiatan tersebut, sehingga peserta mendapatkan nomor sesuai urutan berhasil dicatat.

Tidak ada perubahan struktur Spreadsheet. Google Apps Script v9/v8 yang mengembalikan `display_order` tetap kompatibel.

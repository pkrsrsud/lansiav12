# Template Export

Gunakan file `docs/Daftar Hadir 2024 - template asli.docx` sebagai referensi.

Di Google Docs template:
- `Selasa` -> `{{HARI}}`
- `29 September 2026` -> `{{TANGGAL}}`
- `15.30 WIB - Selesai` -> `{{WAKTU}}`
- `Ruang Tunggu Ruang Bersalin` -> `{{TEMPAT}}`
- `Kelas Ibu Hamil dan Prenatal Gentle Yoga` -> `{{ACARA}}`

Kode Apps Script sudah memiliki `prepareGoogleDocTemplate()` untuk melakukan penggantian contoh tersebut secara otomatis.

Kolom tabel export:
- NO: otomatis
- NAMA: otomatis
- JABATAN / ALAMAT: otomatis diisi alamat
- HADIR: kosong
- MEWAKILI: kosong


## Field otomatis
Saat laporan diekspor untuk suatu kegiatan, field `{{HARI}}`, `{{TANGGAL}}`, `{{WAKTU}}`, `{{TEMPAT}}`, dan `{{ACARA}}` pada header template diisi otomatis dari data kegiatan yang dipilih di dashboard.

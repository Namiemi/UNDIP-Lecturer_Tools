# 🧰 Dashboard Tools Penilaian

Kumpulan tool berbasis web untuk membantu penilaian mahasiswa — berjalan
sepenuhnya di browser, tanpa server dan tanpa install.

Buka `index.html` di browser untuk masuk ke dashboard.

## Tools

### 📝 Penilaian Nilai Mahasiswa (`penilaian.html`)

Input nilai UTS/UAS (atau 3 komponen bebas) dari file Excel SIAP.

- Upload `.xlsx` / `.xls` / `.csv` (klik / drag & drop)
- Deteksi baris header & baris footer otomatis
- Pemetaan kolom: NIM & Nama otomatis, 3 komponen nilai dipilih manual
  (nama mengikuti kolom yang dipilih, bisa diketik ulang)
- Cari mahasiswa by Nama/NIM, entri cepat full-keyboard
  (`/`, `↑↓`, `Enter`, `Alt+1..3`)
- Edit langsung di tabel, undo (`Ctrl+Z`), sortir, filter, pagination
- Download Excel dengan font, format angka, lebar kolom & proteksi sheet
  file asli tetap terjaga (via ExcelJS)

### 📊 Analisis Penilaian OBE (`analisis.html`)

Analisis nilai akhir dari file Excel (contoh bawaan: SIAP DDKB, 50 mahasiswa).

- Upload Excel — kolom NIM/Nama/komponen + bobot terdeteksi otomatis
- Grafik batang nilai akhir: berdasarkan **skor angka** (bin 0–9 s.d. 90–100)
  atau **grade huruf** (A, AB, B, BC, C, D, E ala SIAP)
- Tabel distribusi grade (Grade, Rentang Skor, Jumlah, Persentase)
- Download Excel 2 sheet: Rekap + Distribusi Grade

## Struktur file

```
Grading Project/
├── index.html        # Dashboard
├── penilaian.html    # Tool Penilaian
├── penilaian.js      # Logika Penilaian
├── analisis.html     # Tool Analisis OBE
├── analisis.js       # Logika Analisis (+ contoh DDKB bawaan)
├── styles.css        # Style bersama (mendukung dark mode)
├── contoh/           # File Excel contoh (SIAP DDKB)
└── README.md
```

## Cara pakai

1. Download / clone repo ini.
2. Buka `index.html` di browser (Chrome/Edge/Firefox).
3. Pilih tool yang ingin dipakai.

> **Catatan:**
>
> - Library SheetJS & ExcelJS dimuat dari CDN, jadi butuh internet saat
>   pertama membuka halaman.
> - Semua data tersimpan lokal di browser (`localStorage`) — tidak dikirim
>   ke mana pun. Menghapus cache browser akan menghilangkan data tersimpan.
> - Format Excel yang didukung merujuk pada export **SIAP** (header berisi
>   `NIM`, `Nama Mahasiswa`, dan kolom `Nilai ... (bobot%)`).

## Format kolom yang dikenali (Analisis)

| Kolom          | Keterangan                                    |
| -------------- | --------------------------------------------- |
| NIM            | Otomatis (kolom berjudul NIM)                 |
| Nama           | Otomatis (kolom berjudul Nama)                |
| Komponen skor  | Kolom `Nilai ...`, kecuali `Akhir/Huruf/Bobot`|
| Bobot          | Dibaca dari judul kolom, cth: `(25%)`         |
| Grade huruf    | A ≥ 85, AB ≥ 80, B ≥ 75, BC ≥ 70, C ≥ 60, D ≥ 40, E di bawahnya |

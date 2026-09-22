---
name: litmas-generator
description: Generate Laporan Penelitian Kemasyarakatan (Litmas) lengkap dari input mentah (transkrip wawancara, PDF, hasil web, dokumen pendukung) menjadi 146-field Excel sesuai struktur Contoh Fix.ods (Bapas Jember). Digunakan oleh Pembimbing Kemasyarakatan untuk Cuti Bersyarat / Pembebasan Bersyarat / Asimilasi.
version: 1.0.0
author: AutoLitmas
license: MIT
tags: [litmas, bapas, pemasyarakatan, corrections, indonesia, legal, document-generation]
---

# Skill: litmas-generator

## Tujuan
Mengubah data mentah klien (transkrip wawancara STT, PDF putusan, KK, hasil web SIPP, dll) menjadi **satu row Excel 146 kolom** sesuai struktur `Contoh Fix.ods` untuk Laporan Penelitian Kemasyarakatan (Litmas) Bapas Jember.

## Konteks
- **Pengguna**: Pembimbing Kemasyarakatan (PK) Bapas Kelas II Jember
- **PK default**: Gema Eka Adi Pamungkas (NIP 19940817 202012 1 001)
- **Output**: File `.xlsx` baru per klien, struktur = persis `Contoh Fix.ods` (146 kolom, header row 1, data row 2)
- **Jenis Litmas**: Cuti Bersyarat (CB) / Pembebasan Bersyarat (PB) / Asimilasi di Rumah (PP 99)

## File Skill
- `litmas_schema.json` — Schema 146 field (tipe, deskripsi, default, enum)
- `templates/extract_prompt.md` — Prompt untuk extract data terstruktur (kolom 1-84)
- `templates/narasi_prompt.md` — Prompt untuk generate narasi paragraf (kolom 85-146)
- `templates/reference_contoh_fix.md` — Sample row 2 (klien Samsul Muarip) sebagai contoh gaya narasi
- `scripts/write_excel.py` — Script openpyxl untuk tulis JSON → `.xlsx` per klien

## Alur Kerja

### 1. Ingest Input
Kumpulkan semua data mentah klien:
- Transkrip wawancara (hasil STT website lain) — teks
- PDF putusan pengadilan / petikan putusan
- Kartu Keluarga (KK) klien & penjamin (OCR/teks)
- Laporan Perkembangan Narapidana dari Lapas
- Hasil web SIPP (Sistem Informasi Penelusuran Perkara)
- Surat permintaan Litmas dari Lapas
- Data perhitungan masa pidana (Telraam)

### 2. Extract Data Terstruktur (Kolom 1-84)
Pakai `templates/extract_prompt.md` + `litmas_schema.json`.
AI baca semua input → tarik 84 field terstruktur (identitas klien, ortu, istri, penjamin, perkara, pidana, masa pidana, meta litmas).
Apply defaults dari schema untuk field repetitif (PK, Bapas, bangsa, WNI).

### 3. Generate Narasi (Kolom 85-146)
Pakai `templates/narasi_prompt.md` + `templates/reference_contoh_fix.md`.
AI susun 62 paragraf narasi mengikuti gaya & struktur sample row 2 Contoh Fix.ods.
Setiap paragraf = 1 sel Excel. Format narasi formal Bapas (Bahasa Indonesia baku, sebut klien sebagai "klien").

### 4. Tulis Excel
Jalankan `scripts/write_excel.py` via `execute_code`:
```bash
python3 scripts/write_excel.py --json output_klien.json --out "Litmas_[NamaKlien].xlsx"
```
Output: file `.xlsx` dengan 146 kolom (header row 1 + data row 2), siap dibuka di Excel/LibreOffice.

## Cara Invoke via Hermes

### CLI
```
hermes run-skill litmas-generator --input ./data_klien/ --nama "Mohammad Zaenal Abidin" --jenis "Cuti Bersyarat"
```

### Chat / Telegram
```
@hermes tolong buatkan Litmas untuk klien Mohammad Zaenal Abidin, jenis Cuti Bersyarat. Data ada di folder ./data_klien/ (transkrip wawancara.txt, putusan.pdf, KK.jpg)
```

Hermes akan:
1. Baca semua file di folder input (OCR PDF/image jika perlu)
2. Extract 84 field terstruktur
3. Generate 62 paragraf narasi
4. Tulis Excel baru `Litmas_Mohammad_Zaenal_Abidin.xlsx`
5. Kirim balik file `.xlsx` ke lo

## Catatan Penting
- **Data sensitif**: Litmas = data pribadi narapidana. Pastikan Hermes jalan di lingkungan yang sesuai compliance (on-prem / VPS private). Jangan kirim data ke API publik tanpa persetujuan.
- **Akurasi**: AI bisa salah baca transkrip STT (noise/noise). Selalu **review output** sebelum submit ke sidang TPP.
- **Field kosong**: Kalau data tidak ada di input, biarkan kosong (jangan halusinasi). AI harus tulis "" bukan ngaruh.
- **Format tanggal**: YYYY-MM-DD (ISO 8601). Contoh: 27 Agustus 1973 → `1973-08-27`.
- **Bahasa narasi**: Bahasa Indonesia baku, sebut subjek sebagai "klien" (bukan nama).

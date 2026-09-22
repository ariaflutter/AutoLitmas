---
name: litmas-generator
description: Generate 146-field Bapas Litmas Excel from raw client data.
version: 1.1.0
author: AutoLitmas
license: MIT
platforms: [linux, macos, windows]
metadata:
  hermes:
    tags: [litmas, bapas, pemasyarakatan, corrections, indonesia, legal, document-generation]
    related_skills: []
---

# Skill: litmas-generator

Mengubah data mentah klien (transkrip wawancara STT, PDF putusan, KK, surat permintaan Lapas, hasil web SIPP, Telraam) menjadi **satu row Excel 146 kolom** sesuai struktur `Contoh Fix.ods` untuk Laporan Penelitian Kemasyarakatan (Litmas) Bapas Jember. Skill ini hanya menulis file `.xlsx`; ia tidak submit ke sidang TPP dan tidak menghitung masa pidana sendiri (angka 1/3, 1/2, 2/3, ekspirasi dibaca dari Telraam/surat Lapas).

## When to Use

- PK minta Litmas untuk klien **Cuti Bersyarat (CB)**, **Pembebasan Bersyarat (PB)**, **Asimilasi di Rumah**, **Pembimbingan**, atau **Permintaan Pemindahan Lapas**.
- Tersedia data mentah per klien: surat permintaan Lapas, surat tugas PK, petikan putusan, dan/atau transkrip wawancara.
- **Don't use for:** menulis narasi Litmas tanpa data mentah, atau administrasi Litmas di luar 146 field (berkas sidang, tanda tangan, arsip).

## Prerequisites

- `openpyxl` di interpreter yang menjalankan script: `python -m pip install "openpyxl>=3.1,<4"`.
- Di Windows tidak ada `python3` — pakai `python`.
- **Data sensitif.** Litmas = data pribadi narapidana. Jangan kirim isi data klien ke API publik tanpa persetujuan; jalankan di lingkungan yang sesuai compliance.

## File Skill

- `litmas_schema.json` — schema 146 field (tipe, enum, default). `templates/litmas_schema.json` duplikat identik; pakai yang di root.
- `templates/extract_prompt.md` — prompt extract data terstruktur (kolom 1-84).
- `templates/narasi_prompt.md` — prompt generate narasi paragraf (kolom 85-146).
- `templates/reference_contoh_fix.json` — 146 header + sample row 2 (klien Samsul Muarip); ini file yang dibaca `write_excel.py` untuk header.
- `templates/reference_contoh_fix.md` — versi human-readable dari file di atas.
- `scripts/write_excel.py` — tulis JSON 146 field → `.xlsx`.

## Alur Kerja

### 1. Ingest Input
Baca **semua** file di folder klien sebelum mengisi field apa pun: transkrip wawancara (STT) · petikan putusan · Kartu Keluarga klien & penjamin · Laporan Perkembangan Narapidana · hasil web SIPP · surat permintaan Litmas dari Lapas · surat tugas PK · perhitungan masa pidana (Telraam).

Selesai bila: daftar file input dan sumber tiap field sudah diketahui.

### 2. Extract Data Terstruktur (Kolom 1-84)
Pakai `templates/extract_prompt.md` + `litmas_schema.json`. Tarik 84 field (meta litmas, identitas klien, ayah, ibu, istri, penjamin) dan apply `defaults` dari schema untuk field repetitif (PK, Bapas, bangsa, WNI).

Selesai bila: 84 key terisi, di mana setiap nilai kosong benar-benar tidak ada di input.

### 3. Generate Narasi (Kolom 85-146)
Pakai `templates/narasi_prompt.md` + `templates/reference_contoh_fix.md`. Baca sample row 2 dulu untuk mimic gaya, lalu susun 62 paragraf (Bahasa Indonesia baku, subjek disebut **"klien"**, tanpa line break di dalam sel).

Selesai bila: 62 paragraf ada, dan paragraf yang datanya tidak tersedia ditandai sesuai konvensi penanda di bawah.

**Konvensi penanda data kosong** — beda per blok, dan tidak boleh dicampur:
- Field terstruktur (1-84) yang tidak ada di input → string kosong `""`.
- Narasi (85-146) yang hanya sebagian datanya tidak ada → `...` di posisi data yang hilang, kalimat template tetap ditulis.
- Narasi yang **seluruhnya** belum bisa diverifikasi dari sumber yang ada → tulis draftnya, lalu akhiri dengan `[PERLU DATA: <apa yang kurang>]` agar mudah ditemukan PK dengan pencarian teks.

### 4. Tulis Excel
Gabung hasil langkah 2 dan 3 jadi satu JSON (146 key), lalu **append ke master**:

```
terminal(command='python "<skill_dir>/scripts/write_excel.py" --json "<workdir>/output_klien.json" --master "<repo>/output/Master Litmas.xlsx"', timeout=120)
```

- `<skill_dir>` = folder skill ini, mis. `~/AppData/Local/hermes/skills/litmas-generator`.
- **`--master` = target default tiap run.** Append row baru ke `Master Litmas.xlsx`; kalau file belum ada, dibuat otomatis (header 146 kolom + row). **Duplikat nama klien di-skip otomatis** (dibandingkan kolom 1 master), jadi run berulang tidak menambah row kembar — tidak perlu cek manual.
- `--out "<path>"` = buat file `.xlsx` baru (1 klien, standalone) — pakai hanya kalau mau file terpisah, bukan masuk master.
- `--append "<path>"` = append ke file yang sudah ada (tanpa auto-create/dedup) — fallback.
- Nama key JSON = nama field di `litmas_schema.json` (mis. `1_Nama_Klien`). Script juga menerima prefix angka saja (`"1"`).

Selesai bila: script mencetak `OK: <path>` dan file `.xlsx` benar-benar ada di disk.

## Cara Invoke via Hermes

Cukup minta di chat / Telegram:

```
@hermes buatkan Litmas untuk klien Mohammad Imron Bin Nawari, jenis Cuti Bersyarat.
Data ada di D:\Github\AutoLitmas\AutoLitmas\data_klien\M.Imron bin Nawari\
```

Hermes akan: (1) `skill_view(name='litmas-generator')`, (2) baca semua file input — `read_file` sudah mengekstrak teks PDF/DOCX, pakai `vision_analyze` untuk scan KK/foto, (3) extract 84 field, (4) generate 62 paragraf narasi, (5) jalankan `write_excel.py` lewat `terminal`, (6) kirim file balik dengan `MEDIA:<path absolut>`.

Tidak ada subcommand `hermes run-skill` — skill dijalankan lewat alur di atas, bukan CLI khusus.

## Pitfalls

- **`python3` tidak ada di Windows.** `python3 scripts/write_excel.py` gagal; pakai `python`.
- **Path relatif gagal.** Script dipanggil dari cwd mana pun — selalu beri path absolut ke `--json` dan `--out`.
- **`openpyxl` tidak terinstall** → script exit 1 dengan `ERROR: openpyxl tidak terinstall.`. Cek interpreter yang benar-benar dipakai, bukan `pip list` global.
- **`python -m pip` bilang `No module named pip`** → venv-nya tidak punya pip (umum di venv yang dibuat uv). Jalankan `python -m ensurepip --upgrade` dulu, baru `python -m pip install "openpyxl>=3.1,<4"`.
- **`Alignment(vertical="middle")` tidak valid di openpyxl** — nilai yang diterima hanya `top`, `center`, `bottom`, `justify`, `distributed`. Ini pernah bikin `--out` crash; kalau mengedit `write_excel.py`, jangan pakai `"middle"`.
- **`--append` ke file yang belum ada** → exit 1. Pakai `--out` untuk file baru.
- **Field hilang hanya WARN, bukan error.** Script mengisi `""` dan mencetak `WARN: N field kosong` ke stderr. Baca peringatan itu sebelum mengirim file ke PK.
- **Format tanggal = ISO 8601 `YYYY-MM-DD`** (standar tunggal di semua output). `write_excel.py` otomatis meng-normalisasi field tanggal (kolom 4, 6, 11, 15, 19-27, 43, 54, 65, 76) dari `M/D/YYYY` atau `D-M-YYYY` ke ISO 8601, jadi aman walau LLM output format lain. Tanggal dalam narasi (85-146) **harus** ISO 8601 — instruksi di `narasi_prompt.md` sudah menegaskan ini. Nomor register (misal `B II A 03/08/2026`) bukan tanggal dan tidak di-normalisasi.
- **Penanda kosong berbeda per blok:** field terstruktur (1-84) kosong = `""`; narasi (85-146) yang sebagian datanya tidak ada = `"..."`; narasi yang seluruhnya belum terverifikasi diakhiri `[PERLU DATA: ...]`.
- **STT noise.** Transkrip wawancara sering salah baca nama/tanggal. Prioritas sumber: putusan > KK > surat Lapas > transkrip > web.
- **Jangan halusinasi.** Field yang tidak ada di input dibiarkan kosong — jangan menebak suku, pendidikan, atau nilai RRI/kriminogenik.

## Verification

1. `python -c "import openpyxl; print(openpyxl.__version__)"` berhasil di interpreter yang dipakai.
2. Script mencetak `OK: <path>`; setiap `WARN: ... field kosong` sudah dijelaskan alasannya.
3. Buka `.xlsx` hasil: row 1 = 146 header, row 2 = 1 row data, kolom 85-146 berisi 62 paragraf.
4. Bandingkan 5 field paling kritis (`1_Nama_Klien`, `21_No_Putusan`, `22_Pidana`, `27_Ekspirasi`, `84_Penjamin_Hubungan`) langsung ke dokumen sumber.
5. Review manual oleh PK sebelum submit ke sidang TPP.

## Catatan Penting

- **PK default**: Gema Eka Adi Pamungkas, S.Sos (NIP 199408172020121001), Bapas Kelas II Jember.
- **Output**: satu file `.xlsx` per klien, struktur persis `Contoh Fix.ods` (146 kolom, header row 1, data row 2).
- **Akurasi**: AI bisa salah baca transkrip STT. Review output wajib sebelum submit ke sidang TPP.
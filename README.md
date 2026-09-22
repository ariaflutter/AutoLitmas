# AutoLitmas

Automate **Laporan Penelitian Kemasyarakatan (Litmas)** generation using AI (Hermes Agent) — dari input mentah (transkrip wawancara, PDF putusan, KK, hasil web) → **146-field Excel** sesuai struktur `Contoh Fix.ods` (Bapas Jember).

## Struktur Repo

```
AutoLitmas/
├── Contoh Fix.ods                          # Template output reference (146 kolom, sample row 2)
├── LitmasV5.xlsm                           # Engine Litmas asli (rumus, RRI, Kriminogenik, Telraam)
├── Cuti_Bersyarat_...docx                  # Contoh Litmas jadi (narasi lengkap)
├── reference_contoh_fix.json               # Extract headers + sample row dari Contoh Fix.ods
├── reference_contoh_fix.md                 # Human-readable reference
├── skills/
│   └── litmas-generator/                   # Hermes Skill
│       ├── SKILL.md                        # Definisi skill (metadata + alur)
│       ├── litmas_schema.json              # Schema 146 field (tipe, default, enum)
│       ├── templates/
│       │   ├── extract_prompt.md           # Prompt extract data terstruktur (kolom 1-84)
│       │   ├── narasi_prompt.md            # Prompt generate narasi (kolom 85-146)
│       │   ├── reference_contoh_fix.json   # Reference headers + sample
│       │   └── reference_contoh_fix.md     # Reference human-readable
│       └── scripts/
│           └── write_excel.py              # Script tulis JSON → .xlsx (openpyxl)
└── README.md                               # File ini
```

## Cara Pakai

### Prasyarat
1. **Hermes Agent** terinstall (lihat https://hermes-agent.nousresearch.com/docs)
2. **Python 3.10+** dengan `openpyxl`:
   ```bash
   pip install openpyxl
   ```
3. **Skill folder** di-load Hermes (lihat dokumentasi Skills System Hermes)

### Alur Singkat
1. Kumpulkan data klien (transkrip wawancara.txt, putusan.pdf, KK.jpg, hasil web SIPP, dll)
2. Invoke skill via Hermes:
   ```
   @hermes buatkan Litmas untuk klien Mohammad Zaenal Abidin, jenis Cuti Bersyarat.
   Data ada di folder ./data_klien/
   ```
3. Hermes akan:
   - Baca semua file input (OCR jika perlu)
   - Extract 84 field terstruktur (identitas, perkara, pidana, masa pidana)
   - Generate 62 paragraf narasi (riwayat → rekomendasi)
   - Tulis Excel baru `Litmas_[NamaKlien].xlsx`
4. Buka file `.xlsx` → review → submit ke sidang TPP

### Manual (tanpa Hermes)
Kalau mau test tanpa Hermes, bisa jalankan manual:

1. Extract data klien ke JSON 146 field (pakai LLM apapun dengan prompt dari `templates/`)
2. Tulis Excel:
   ```bash
   python3 skills/litmas-generator/scripts/write_excel.py \
     --json output_klien.json \
     --out "Litmas_Mohammad_Zaenal_Abidin.xlsx"
   ```

## Output
File `.xlsx` dengan:
- **Row 1**: 146 header (sesuai `Contoh Fix.ods`)
- **Row 2+**: data klien (1 row per klien)
- Kolom 1-84: data terstruktur (identitas, meta litmas)
- Kolom 85-146: 62 paragraf narasi Litmas

## Field Schema
Lihat `skills/litmas-generator/litmas_schema.json` untuk definisi lengkap 146 field (tipe, default, enum, deskripsi).

## Catatan
- **Data sensitif**: Litmas = data pribadi narapidana. Jalankan Hermes di lingkungan yang sesuai compliance.
- **Review wajib**: AI bisa salah baca transkrip STT. Selalu review output sebelum submit.
- **Format tanggal**: YYYY-MM-DD (ISO 8601) (sesuai Excel).
- **PK default**: Gema Eka Adi Pamungkas (NIP 19940817 202012 1 001), Bapas Kelas II Jember.

## Lisensi
MIT

# office-template-batch-pdf

Batch DOC/DOCX/PPT/PPTX → PDF generation with `{{TOKENS}}` replacement from Excel or CSV.  
Includes preflight (tokens vs columns), token filters, dry-run mode, per-row subfolders, and export via **LibreOffice** or **Microsoft Office COM** (on Windows).

![License](https://img.shields.io/badge/license-MIT-blue.svg)

---

## Features
- Replace tokens `{{FIELD}}` in **.docx** and **.pptx** (and legacy .doc/.ppt/.odt/.odp/.rtf after automatic conversion to OOXML).
- **Preflight**: scans templates and warns about missing or unused columns.
- **Token filters & defaults**:  
  `{{Name|trim|upper}}`, `{{Amount|euros}}`, `{{Date|dmy}}`, `{{Field?:N/A}}`
- **Dry-run** mode: simulate replacements without generating PDFs.
- **Row-level SKIP** and per-row **OUTPUT** (subfolder) support.
- **JSON and CSV** reports for every batch run.
- **Export engine**: choose LibreOffice or MS Office (Word / PowerPoint COM on Windows).
- **Progress bar** using `tqdm` (if installed).

---

## Requirements
- **Python 3.9+**
- Libraries:
  ```bash
  pip install pandas python-docx python-pptx
  # optional:
  pip install tqdm pywin32 comtypes
  ```
- **LibreOffice** installed and accessible as `soffice` (or configure `SOFFICE_BIN`).
- (Optional) Microsoft Office (Word/PowerPoint) on Windows for COM export.

---

## Folder structure
```
.
├─ templates/
│  ├─ diploma_base.pptx
│  └─ letter.docx
├─ data.xlsx          # or data.csv
└─ output/            # created automatically
```

---

## Excel / CSV columns
| Column     | Required | Description |
|-------------|-----------|-------------|
| `TEMPLATE`  | ✅ | Template filename (e.g. `diploma_base.pptx`) |
| `SKIP`      | ⛔ | If set to `1`, `true`, `yes`, `si`, etc., the row is skipped |
| `OUTPUT`   | ⛔ | Optional subfolder name for output |
| others...   | — | All other columns are available as tokens like `{{NOMBRE}}`, `{{EMPRESA}}`, etc. |

---

## Token syntax in templates
- Basic: `{{NOMBRE}}`
- With filters: `{{Nombre|trim|upper}}`, `{{Fecha|dmy}}`, `{{Importe|euros}}`
- With default value: `{{Campo?:N/A}}`

**Supported scopes:**
- DOCX → body, tables, headers, and footers  
- PPTX → slides, tables, and master/layouts  

---

## Quick start
```bash
python office-template-batch-pdf.py               # uses default paths (script directory)
python office-template-batch-pdf.py data.xlsx     # specify input data
python office-template-batch-pdf.py data.xlsx output/ templates/
```

---

## CLI Options
```bash
python office-template-batch-pdf.py --help
```

| Option | Description |
|--------|-------------|
| `--sheet` | Excel sheet name or index (ignored for CSV) |
| `--pattern` | Output filename pattern (e.g. `"{NAME}_{COURSE}.pdf"`) |
| `--engine` | `auto` \| `libreoffice` \| `msoffice` |
| `--strict` | Fail if a token has no matching column |
| `--dry-run` | Simulation mode (no PDFs) |
| `--pdf-filter-opts` | LibreOffice PDF filter options |
| `--from` / `--to` | Row range to process |
| `--where` | Pandas query filter, e.g. `"Course == 'A' and SKIP != '1'"` |
| `--verbose` | More detailed logging |
| `--check` | Environment check (LibreOffice, MS Office) |
| `--version` | Show current version |

---

## Examples
```bash
# Simulate replacements only
python office-template-batch-pdf.py --dry-run --verbose

# Process rows 0..49 where Course == 'B'
python office-template-batch-pdf.py data.xlsx output/ templates/ --from 0 --to 49 --where "Course == 'B'"

# Force LibreOffice and custom output name pattern
python office-template-batch-pdf.py --engine libreoffice --pattern "{index:04d}_{Company}.pdf"
```

---

## Output and reports
- PDFs go into `./output` (or your custom output directory).
- Reports generated automatically:
  - `_report.json`
  - `_report.csv`

Each report includes: row index, status (`OK`, `ERROR`, `SKIPPED`, `DRY-RUN`), template name, and output path.

---

## Troubleshooting & tips
- **LibreOffice detection**  
  Run `python office-template-batch-pdf.py --check` to verify.  
  If not found, set the path explicitly:
  ```python
  SOFFICE_BIN = r"C:\Program Files\LibreOffice\program\soffice.exe"
  ```
- **Performance**  
  Disable antivirus scanning in the working folder.  
  For thousands of rows, use `--from/--to` to process in chunks.
- **Fonts**  
  Make sure the fonts used in templates are installed on the system.
- **MS Office COM (Windows)**  
  Requires `pywin32` or `comtypes`.  
  Falls back to LibreOffice automatically if unavailable.

---

## License
**MIT License** – see [LICENSE](LICENSE) for details.

---

## Credits
Created with ❤️ by [Nico](https://github.com/)  
for automating document and diploma generation in batch mode.

---

## Example dataset
You can provide a simple sample in `/examples`:

```
examples/
├─ templates/
│  ├─ demo_docx.docx
│  └─ demo_pptx.pptx
├─ data.xlsx
└─ README.md
```

Each template includes tokens like `{{NAME}}`, `{{COURSE}}`, etc.  
The Excel file contains one row per record.

---

## Changelog
**v0.9.0 (2025-11-02)**  
- Initial public release.  
- Added preflight, token filters, dry-run, JSON/CSV reports, MS Office COM export, and multi-format template support.

---

## Tags
`python` · `pdf` · `docx` · `pptx` · `automation` · `templates` · `libreoffice` · `batch` · `word` · `powerpoint`

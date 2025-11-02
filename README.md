# office-template-batch-pdf

Batch DOC/DOCX/PPT/PPTX → PDF con sustitución de tokens `{{...}}` desde Excel/CSV.
Incluye preflight (tokens vs columnas), filtros de tokens, dry-run, subcarpeta por fila y exportación vía LibreOffice o Microsoft Office (COM en Windows).

![License](https://img.shields.io/badge/license-MIT-blue.svg)

## Características
- Sustitución de tokens `{{CAMPO}}` en **.docx** y **.pptx** (y en .doc/.ppt/.odt/.odp/.rtf tras convertir a OOXML).
- **Preflight**: descubre tokens en plantillas y avisa si faltan columnas o si hay columnas sin uso.
- **Filtros y valores por defecto** en tokens: `{{Nombre|trim|upper}}`, `{{Importe|euros}}`, `{{Fecha|dmy}}`, `{{Campo?:N/A}}`.
- **Dry-run** (simulación) sin generar PDFs.
- **SKIP** por fila y **CARPETA** opcional por fila para salida.
- **Informes** JSON y CSV por lote.
- Motor de exportación **LibreOffice** o **MS Office** (Word/PowerPoint COM en Windows).
- Barra de progreso con `tqdm` (si está instalado).

## Requisitos
- Python 3.9+
- `pandas`, `python-docx`, `python-pptx`
- **LibreOffice** accesible como `soffice` (o variable `SOFFICE_BIN`) para conversiones/exportación.
- Opcional (Windows): `pywin32` o `comtypes` para exportar con MS Office.
- Opcional: `tqdm` para barra de progreso.

```bash
pip install pandas python-docx python-pptx
# opcional
pip install tqdm pywin32 comtypes

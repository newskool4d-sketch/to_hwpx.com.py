# Dependency Matrix

The converter supports Python 3.10 or newer. This matches the current source
syntax, including `str | None` annotations in `hwpx_hierarchy.py`.

| Feature area | Input or command | Required packages/tools | Optional fallback packages | Missing-dependency behavior |
|---|---|---|---|---|
| Core parser dispatch | `.md`, `.txt`, `.csv`, `--list-formats` | Python standard library | none | No third-party dependency expected. |
| Markdown tables | `.md` | Python standard library plus local `markdown_table_parser.py` | none | Parser errors should surface directly. |
| Plain text | `.txt` | Python standard library plus local hierarchy parser | none | Parser errors should surface directly. |
| CSV | `.csv` | Python standard library `csv` | none | Empty CSV returns `[]`. |
| HTML | `.html`, `.htm` | `beautifulsoup4` | none | Raises `RuntimeError` with `pip install beautifulsoup4`. |
| XLSX | `.xlsx` | `openpyxl` | none | Raises `RuntimeError` with `pip install openpyxl`. |
| DOCX | `.docx` | `python-docx` | none | Import failure surfaces from the parser import path. |
| PDF structured extraction | `.pdf` | `opendataloader-pdf` | none | Failure is recorded and PDF text fallbacks are attempted. |
| PDF OCR route | `.pdf` scanned text | `kordoc-ai` available through `--kordoc-home` or `KORDOC_HOME` | none | Unavailable kordoc route returns `None`; fallback extractors are attempted. |
| PDF text fallback | `.pdf` selectable text | one of `pdfplumber`, `PyMuPDF`, or `pypdf` | tries in that order | Raises `RuntimeError` naming each unavailable or failed extractor. |
| HWP COM conversion | normal `to_hwpx_com.py` conversion | Windows, Hancom HWP with COM support, `pywin32` | none | `--preflight` or conversion returns exit code `2` with a Korean startup message. Conversion is serial; do not run parallel conversions against one HWP COM object. |
| Direct HWPX experiment | `hwpx_direct.py` | local `hwpx변환` skill assets and helper scripts | none | Missing helper assets raise `ImportError` or subprocess errors. |
| Development tests | `python -B -m pytest` | `pytest` | optional parser packages for full coverage | Tests skip only when the named optional package is absent. |

If HWP COM startup reports a stale `gen_py` cache, do not delete the cache automatically during conversion. Close HWP and run this explicit repair command in the same Python environment:

```bash
python -c "import win32com.client.gencache as gencache; gencache.Rebuild()"
```

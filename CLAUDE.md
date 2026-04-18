# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project purpose

Python CLI tools that turn CSV exports from an EcoStruxure HMI/SCADA system into professional A4 PDF reports (plus auto-printing helpers). Everything runs locally; the final artifacts are Windows `.exe` files built by CI for non-technical operators on the factory floor. README.md and all user-facing strings are in Italian — keep that convention.

## Commands

```bash
pip install -r requirements.txt           # pandas, reportlab, matplotlib, pyarrow

# Generate a single report (auto-detects the most recent matching CSV in cwd)
python generate_report_alarm.py    [--csv F] [--out F] [--logo F] [--limit-rows N] [--dry-run]
python generate_report_operlog.py  [--csv F] [--out F] [--logo F] [--limit-rows N] [--dry-run]
python generate_report_batch.py    [--csv F] [--out F] [--logo F] [--limit-rows N] [--dry-run] [--chart-first] [--separate-files]

# Run all three against ./data (or another dir)
python run_reports.py --data-dir data [--output-dir PDF] [--type batch|alarm|operlog|all] [--dry-run]

# Print the latest PDF (default printer)
python print_latest_pdf.py
python print_latest_pdf_from_recent_folder.py   # picks latest DDMMYY folder first
```

There is no test suite and no linter configured. For a fast sanity check use `--dry-run --limit-rows 10` on any generator.

## Architecture

Five entry-point scripts share one utilities module; there is no package structure.

- **`report_utils.py`** — shared helpers used by all three generators. When changing report appearance, start here: `get_common_styles`, `get_common_table_style`, `create_logo_header`, `add_page_number`, `convert_date_format` (MM/DD/YYYY → DD/MM/YY), `find_header_row` (skips EcoStruxure metadata preamble by scanning for the row starting with `Date`), `clean_dataframe_columns` / `clean_dataframe_data`, and `get_logo_path` / `get_resource_path` (PyInstaller-aware via `sys._MEIPASS`).
- **`generate_report_{alarm,operlog,batch}.py`** — one script per CSV type. Each auto-detects the most recent matching CSV (`*ALARM*`, `*OPERLOG*`, `*BATCH*`) in its working directory, applies type-specific column whitelisting and word-wrap, and writes a PDF. `batch` additionally renders a matplotlib temperature chart on a landscape A4 page (the rest of the document stays portrait — this is done with a second `SimpleDocTemplate` / page size swap). `--separate-files` splits the batch output into `_chart.pdf` + `_report.pdf`.
- **`operlog` is the odd one out** for paths: it expects to be run from a directory containing `DDMMYY/` subfolders, picks the most recent, and writes both the PDF and its log file *inside* that folder. The other generators write to `./PDF/`.
- **`run_reports.py`** — orchestrator that shells out to each generator via `subprocess`. It does the `BATCH` vs `OPERLOG_BATCH` vs `ALARM_BATCH` disambiguation in `find_csv_files` (filename substring filtering); if you add a new CSV naming scheme, update that function *and* the per-script `find_csv_*` functions — they are separate.
- **`print_latest_pdf*.py`** — Windows-only printing via the default printer; macOS/Linux paths exist but are mostly for dev.

### PyInstaller / Windows build

`.github/workflows/build_windows.yml` builds five `.exe`s with `--onefile --windowed` and bundles `data/logo.png` via `--add-data`. Consequences to keep in mind when editing:

- Use `get_resource_path()` from `report_utils` for any bundled asset — `sys._MEIPASS` only exists in the frozen build.
- Because `--windowed` suppresses stdout/stderr, the frozen binaries redirect them to a log file via `setup_logging_for_pyinstaller()`. Any new script destined for CI build must call it early in `main()`.
- `version_info.txt` and `--uac-admin` on the print scripts exist to reduce Windows antivirus false positives — don't remove without reason.
- The workflow triggers on `main` and the legacy branch `fix-page-and-refactor`; update the branch list if you rename branches.

### CSV input quirks

EcoStruxure CSVs have a variable-length metadata preamble before the real header row. Parsing flow is always: `find_header_row` → `pd.read_csv(skiprows=...)` with auto-detected separator (tab / comma / semicolon) → `clean_dataframe_columns` with the type-specific required column list → `clean_dataframe_data` (strips quotes, collapses `nan` / `'-` placeholders). Missing columns are reported in-PDF as a note rather than aborting.

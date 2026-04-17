# PAH Hospitalist Invoice Generator

Generate Fraser Health hospitalist invoice workbooks from the master Google Sheets schedule.

The app reads the live schedule, matches a physician's shifts for a selected half-month, fills an invoice template workbook, and saves the result into a local `outputs/` folder.

## What It Does

- Matches shifts by schedule name plus optional aliases
- Generates invoices for the first half or second half of a month
- Supports daytime team shifts, ADMIT, VIRTUAL, evening, and overnight work
- Optionally adds a 15-minute clinical admin row after selected shift types
- Avoids adding clinical admin if that extra time overlaps another billed interval
- Saves finished workbooks into `outputs/`

## Shift Rules

- Team/day shifts default to `07:00-17:00`
- If a day-shift name has `*C` after it, such as `Grewal*C`, that shift is treated as on-call daytime coverage and billed as `06:30-17:30`
- `VIRTUAL` shifts are billed as `17:00-21:00` and are `Off-Site`
- Clinical admin is `Off-Site` only when it follows a `VIRTUAL` shift
- Clinical admin is `On-Site` for the other supported shift types

## Requirements

- Windows
- Python 3.12+ recommended
- `openpyxl`
- Access to the master Google Sheet
- An existing invoice template workbook

If you do not already have `openpyxl` installed:

```powershell
python -m pip install openpyxl
```

## Quick Start

### Option 1: Launch the local web app

Double-click:

```text
Start Hospitalist Invoice App.bat
```

If the bundled Python runtime is available, the app starts a local web server and tells you to open:

```text
http://127.0.0.1:8765
```

Then:

1. Fill in physician and billing details.
2. Confirm or change the template workbook path.
3. Confirm or change the Google Sheets master URL.
4. Add schedule aliases if the schedule uses abbreviations or alternate spellings.
5. Choose the month and half-month period.
6. Select which shift types should get a 15-minute clinical admin add-on.
7. Generate the workbook.

The finished file is saved in `outputs/` and can be downloaded from the local app page.

### Option 2: Run from the command line

Start the web server directly:

```powershell
python .\hospitalist_invoice_generator.py --serve
```

Or run a one-off invoice generation from the terminal:

```powershell
python .\hospitalist_invoice_generator.py `
  --physician-name "Dr Jane Smith" `
  --first-name "Jane" `
  --last-name "Smith" `
  --msp "12345" `
  --site "PAH" `
  --schedule-name "Smith" `
  --year 2026 `
  --month 4 `
  --period first `
  --clinical-admin-type scheduled `
  --clinical-admin-type virtual
```

Useful optional flags:

- `--schedule-alias "Smith, J.;J Smith"` for alternate schedule names
- `--clinical-admin-note "doing billings"` to change the clinical admin note
- `--template "C:\path\to\template.xlsx"` to choose a workbook template
- `--master-url "https://docs.google.com/..."` to override the source sheet
- `--output "C:\path\to\invoice.xlsx"` to control the output path
- `--dry-run` to print matched rows without writing a workbook

## Template Path

The script currently has a default template path that points to a local file on the original development machine. If that file does not exist on your computer, update the template path in the web form or pass `--template` on the command line.

## Privacy

Generated invoices in `outputs/` may contain physician-identifying information. This folder is intentionally excluded from Git with `.gitignore`, so those files are not pushed to GitHub unless someone manually overrides that behavior.

## Project Files

- `hospitalist_invoice_generator.py`: main application
- `Start Hospitalist Invoice App.bat`: Windows launcher for local web mode
- `outputs/`: generated invoice workbooks


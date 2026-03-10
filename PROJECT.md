# GLA Calculator — Project Reference

## Overview
A web-based tool for computing GLA (Guaranteed Liability Amount) outputs on insurance policy Excel files. Built for Protiviti audit workflows. Upload an Excel file, the tool reads specific input columns, computes two output columns (TAT and GLA), and returns the file with all other columns untouched.

## Tech Stack
- **Backend**: FastAPI (Python) with SSE streaming for real-time progress
- **Frontend**: Single HTML file (inline CSS + JS), served by FastAPI
- **Excel**: openpyxl for reading/writing .xlsx files
- **Dependencies**: `fastapi`, `uvicorn`, `python-multipart`, `openpyxl`

## File Structure
```
├── app.py                  # FastAPI backend (SSE streaming, file processing)
├── templates/
│   └── index.html          # Frontend (dark luxury theme, all inline)
├── GLA dummy file.xlsx     # Original reference/mapping file from client
├── test_input.xlsx         # Test file with 10 rows across real column positions
├── temp_files/             # Auto-created, stores uploaded/processed files
└── PROJECT.md              # This file
```

## How to Run
```bash
python app.py
# Opens at http://127.0.0.1:8000
```

## Column Mapping (Hardcoded — from dummy file)
These are the REAL column positions in the actual client Excel file:

### Input Columns (read-only, never modified)
| Column Letter | Index | Name         |
|---------------|-------|--------------|
| B             | 2     | Status       |
| Y             | 25    | CNTTYPE      |
| AC            | 29    | PREM TERM    |
| AH            | 34    | PREMCESDTE   |
| AS            | 45    | ORIGINAL_SA  |
| AW            | 49    | NEXT_PAYDATE |

### Output Columns (computed and written)
| Column Letter | Index | Name                      |
|---------------|-------|---------------------------|
| AZ            | 52    | TAT for Payment Due Date  |
| BA            | 53    | Protiviti GLA Calculation |

## Formulas
```
AZ (TAT) = NEXT_PAYDATE - PREMCESDTE  (result in days)

BA (GLA) = IF(Status = "IF" AND CNTTYPE IN {"MSB", "SMB"})
             THEN  1% × ORIGINAL_SA × PREM_TERM
             ELSE  0
```

### Business Logic Notes
- Only policies with Status = **"IF"** (In Force) produce a GLA value
- Only contract types **"MSB"** and **"SMB"** are eligible
- Status "PU" (Paid Up), "LA" (Lapsed), or any other → GLA = 0
- CNTTYPE outside {MSB, SMB} (e.g., "ABC") → GLA = 0
- Rows where Status, CNTTYPE, and ORIGINAL_SA are ALL null → skipped entirely
- All other columns in the file remain completely untouched

## API Endpoints

### `GET /`
Serves the frontend HTML page.

### `POST /api/process`
Accepts multipart form: `file` (Excel) + `header_row` (int, default 1).
Returns **SSE stream** with these events:

```
event: stage
data: {"stage": "reading"}

event: stage
data: {"stage": "processing", "total": 5000}

event: progress  (sent every ~N rows, ~200 updates total)
data: {"current": 150, "total": 5000, "processed": 148, "skipped": 2, "gla": 45230.50, "percent": 3.0}

event: stage
data: {"stage": "saving"}

event: complete
data: {
  "file_id": "uuid-string",
  "output_filename": "original_name_GLA_output.xlsx",
  "total_rows": 5000,
  "processed": 4998,
  "skipped": 2,
  "total_gla": 184225.35,
  "preview": [
    {"row": 2, "status": "IF", "cnttype": "SMB", "prem_term": 10, "premcesdte": "28-Feb-2026", "original_sa": 112650, "next_paydate": "01-Mar-2026", "tat": 1, "gla": 11265.0},
    ...  (up to 25 rows)
  ]
}

event: error  (on failure)
data: {"message": "error description"}
```

### `GET /api/download/{file_id}?filename=output.xlsx`
Downloads the processed file. Files auto-clean after 1 hour.

## Frontend Design
- **Theme**: "Dark Luxury Terminal" — warm gold (#c8a55a) on deep black (#07060c)
- **Fonts**: Instrument Serif (headings/numbers), Outfit (body), JetBrains Mono (data)
- **Distinctive elements**: Film grain overlay, gold corner brackets on upload card, diamond step markers, serif stat numbers
- **Flow**: Upload → Process (SSE streaming with live progress bar, speed, ETA, live stats) → Results (animated stats, staggered table rows) → Download

## Test Data (test_input.xlsx)
10 rows covering all edge cases:

| Row | Status | CNTTYPE | PT | SA      | Expected TAT | Expected GLA |
|-----|--------|---------|-----|---------|-------------|-------------|
| 2   | IF     | SMB     | 10  | 112,650 | 1           | 11,265.00   |
| 3   | PU     | SMB     | 10  | 100,000 | 1           | 0           |
| 4   | IF     | SMB     | 10  | 187,635 | 1           | 18,763.50   |
| 5   | IF     | MSB     | 5   | 257,422 | 1           | 12,871.10   |
| 6   | IF     | MSB     | 5   | 151,515 | 1           | 7,575.75    |
| 7   | IF     | ABC     | 8   | 200,000 | 5           | 0 (bad type)|
| 8   | LA     | SMB     | 12  | 500,000 | 4           | 0 (lapsed)  |
| 9   | IF     | MSB     | 20  | 350,000 | 59          | 70,000.00   |
| 10  | IF     | SMB     | 15  | 425,000 | 9           | 63,750.00   |
| 11  | PU     | MSB     | 7   | 180,000 | 5           | 0 (paid up) |

**Total GLA: 1,84,225.35** — All tests validated and passed.

## Known Issues / Notes
- Windows file locking: `os.remove(input_path)` after `wb.close()` may fail on Windows. Wrapped in try/except, cleanup_old_files() handles it later.
- The column mapping is intentionally hardcoded per client requirements. Not configurable from UI.
- `header_row` parameter lets users specify which row contains headers (data starts on next row).

## Potential Future Improvements
1. Input file preview before processing
2. Row-level validation report (flag bad data per row)
3. Formatted output Excel (bold headers, number formatting, summary sheet)
4. Audit log sheet in output (timestamp, filename, tool version)
5. Batch processing (multiple files → zip download)
6. Comparison/diff mode between two processed files
7. Dashboard charts (GLA distribution by CNTTYPE, status breakdown)

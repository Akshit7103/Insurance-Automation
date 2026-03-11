import os
import csv
import uuid
import json
import time
import shutil
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, numbers
from openpyxl.utils import column_index_from_string, get_column_letter

app = FastAPI(title="GLA Calculator")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp_files")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
os.makedirs(TEMP_DIR, exist_ok=True)

# === Column Mapping (from the GLA dummy/reference file) ===
INPUT_COLUMNS = {
    "Status":       "B",
    "CNTTYPE":      "Y",
    "PREM_TERM":    "AC",
    "PREMCESDTE":   "AH",
    "ORIGINAL_SA":  "AS",
    "NEXT_PAYDATE": "AW",
}
OUTPUT_COLUMNS = {
    "GLA":  "BA",
}

# === Excel Styling ===
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
LIGHT_YELLOW_FILL = PatternFill(start_color="FFFFF0", end_color="FFFFF0", fill_type="solid")
LIGHT_GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="000000")
GLA_HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="000000")
GLA_VALUE_FONT = Font(name="Calibri", size=11, color="1F4E79")
GLA_VALUE_FONT_HIGHLIGHT = Font(name="Calibri", bold=True, size=11, color="006100")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
HEADER_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="medium", color="C8A53A"),
    bottom=Side(style="medium", color="C8A53A"),
)


def csv_to_xlsx(csv_path: str, xlsx_path: str):
    """Convert a CSV file to XLSX so the rest of the pipeline stays unchanged."""
    wb = Workbook()
    ws = wb.active
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(xlsx_path)
    wb.close()


def cleanup_old_files(max_age_seconds: int = 3600):
    now = time.time()
    for fname in os.listdir(TEMP_DIR):
        fpath = os.path.join(TEMP_DIR, fname)
        if os.path.isfile(fpath) and (now - os.path.getmtime(fpath)) > max_age_seconds:
            try:
                os.remove(fpath)
            except OSError:
                pass


def sse(event: str, data: dict) -> str:
    return f"event: {event}\ndata: {json.dumps(data)}\n\n"


HEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
INPUT_HEADER_FONT = Font(name="Calibri", bold=True, size=11, color="1F4E79")
DATA_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
INPUT_HEADER_BORDER = Border(
    left=Side(style="thin", color="8EA9C8"),
    right=Side(style="thin", color="8EA9C8"),
    top=Side(style="medium", color="1F4E79"),
    bottom=Side(style="medium", color="1F4E79"),
)


def format_entire_sheet(ws, header_row: int, data_start: int, max_row: int, gla_col: int):
    """Apply borders and formatting across the entire sheet."""
    max_col = ws.max_column

    # --- Header row: bold, blue fill, bordered ---
    for c in range(1, max_col + 1):
        cell = ws.cell(row=header_row, column=c)
        if c == gla_col:
            continue  # GLA header styled separately
        if cell.value is not None:
            cell.font = INPUT_HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = INPUT_HEADER_BORDER

    # --- Data rows: borders on all cells with data ---
    for row_num in range(data_start, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=row_num, column=c)
            if c == gla_col:
                continue  # GLA data cells styled separately
            cell.border = DATA_BORDER
            cell.alignment = Alignment(vertical="center")

    # --- Auto-fit column widths (approximate) ---
    for c in range(1, max_col + 1):
        if c == gla_col:
            continue  # already set
        col_letter = get_column_letter(c)
        max_len = 0
        for row_num in range(header_row, max_row + 1):
            val = ws.cell(row=row_num, column=c).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 8), 30)


def format_output_column(ws, header_row: int, data_start: int, max_row: int, gla_col: int):
    """Apply beautiful formatting to the GLA output column."""

    # --- Header cell at header_row (BA2) ---
    header_cell = ws.cell(row=header_row, column=gla_col)
    header_cell.value = "Protiviti Output GLA"
    header_cell.fill = YELLOW_FILL
    header_cell.font = GLA_HEADER_FONT
    header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_cell.border = HEADER_BORDER

    # --- Set column width ---
    col_letter = get_column_letter(gla_col)
    ws.column_dimensions[col_letter].width = 22

    # --- Style each data cell ---
    for row_num in range(data_start, max_row + 1):
        cell = ws.cell(row=row_num, column=gla_col)
        val = cell.value

        # Number format with 2 decimal places and comma separator
        cell.number_format = '#,##0.00'
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = THIN_BORDER

        # Green background + bold for non-zero GLA, light yellow for zero
        if val is not None and val != 0:
            cell.fill = LIGHT_GREEN_FILL
            cell.font = GLA_VALUE_FONT_HIGHLIGHT
        else:
            cell.fill = LIGHT_YELLOW_FILL
            cell.font = GLA_VALUE_FONT


def process_stream(input_path: str, output_path: str, header_row: int,
                   file_id: str, original_filename: str):
    """Generator that yields SSE events as it processes the workbook."""
    try:
        col = {}
        for name, letter in {**INPUT_COLUMNS, **OUTPUT_COLUMNS}.items():
            col[name] = column_index_from_string(letter)

        # ── Stage: Reading ──
        yield sse("stage", {"stage": "reading"})

        wb = load_workbook(input_path)
        ws = wb.active
        total_rows = max(0, ws.max_row - header_row)

        # ── Stage: Processing ──
        yield sse("stage", {"stage": "processing", "total": total_rows})

        data_start = header_row + 1
        processed = 0
        skipped = 0
        total_gla = 0.0
        preview = []
        batch = max(1, total_rows // 200)  # ~200 progress updates

        for row_num in range(data_start, ws.max_row + 1):
            status       = ws.cell(row=row_num, column=col["Status"]).value
            cnttype      = ws.cell(row=row_num, column=col["CNTTYPE"]).value
            prem_term    = ws.cell(row=row_num, column=col["PREM_TERM"]).value
            premcesdte   = ws.cell(row=row_num, column=col["PREMCESDTE"]).value
            original_sa  = ws.cell(row=row_num, column=col["ORIGINAL_SA"]).value
            next_paydate = ws.cell(row=row_num, column=col["NEXT_PAYDATE"]).value

            if status is None and cnttype is None and original_sa is None:
                skipped += 1
            else:
                # GLA = IF(AND(Status="IF", OR(CNTTYPE in {MSB,SMB})), 1% * SA * PT, 0)
                gla = 0.0
                status_str  = str(status).strip().upper() if status else ""
                cnttype_str = str(cnttype).strip().upper() if cnttype else ""
                if status_str == "IF" and cnttype_str in ("MSB", "SMB"):
                    try:
                        sa = float(original_sa) if original_sa is not None else 0
                        pt = float(prem_term) if prem_term is not None else 0
                        gla = round(0.01 * sa * pt, 2)
                    except (ValueError, TypeError):
                        gla = 0.0

                ws.cell(row=row_num, column=col["GLA"]).value = gla
                processed += 1
                total_gla += gla

                if len(preview) < 25:
                    def fmt_date(d):
                        if isinstance(d, datetime):
                            return d.strftime("%d-%b-%Y")
                        return str(d) if d else "-"

                    preview.append({
                        "row": row_num,
                        "status": str(status) if status else "-",
                        "cnttype": str(cnttype) if cnttype else "-",
                        "prem_term": prem_term if prem_term is not None else "-",
                        "premcesdte": fmt_date(premcesdte),
                        "original_sa": original_sa if original_sa is not None else "-",
                        "next_paydate": fmt_date(next_paydate),
                        "gla": gla,
                    })

            current = processed + skipped
            if current % batch == 0 or row_num == ws.max_row:
                yield sse("progress", {
                    "current": current,
                    "total": total_rows,
                    "processed": processed,
                    "skipped": skipped,
                    "gla": round(total_gla, 2),
                    "percent": round(current / total_rows * 100, 1) if total_rows > 0 else 100,
                })

        # ── Format the entire sheet + output column ──
        format_entire_sheet(ws, header_row, data_start, ws.max_row, col["GLA"])
        format_output_column(ws, header_row, data_start, ws.max_row, col["GLA"])

        # ── Stage: Saving ──
        yield sse("stage", {"stage": "saving"})
        wb.save(output_path)
        wb.close()

        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except OSError:
            pass  # cleanup_old_files() will handle it later

        # ── Complete ──
        output_name = original_filename.rsplit(".", 1)[0] + "_GLA_output.xlsx" if original_filename else "output.xlsx"
        yield sse("complete", {
            "file_id": file_id,
            "output_filename": output_name,
            "total_rows": total_rows,
            "processed": processed,
            "skipped": skipped,
            "total_gla": round(total_gla, 2),
            "preview": preview,
        })

    except Exception as e:
        for p in [input_path, output_path]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except OSError:
                    pass
        yield sse("error", {"message": str(e)})


# ──────────────────────────── Routes ────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    html_path = os.path.join(TEMPLATES_DIR, "index.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.post("/api/process")
async def api_process(
    file: UploadFile = File(...),
    header_row: int = Form(default=2),
):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls", ".csv")):
        raise HTTPException(status_code=400, detail="Please upload an Excel (.xlsx) or CSV (.csv) file")

    cleanup_old_files()

    file_id = str(uuid.uuid4())
    is_csv = file.filename.lower().endswith(".csv")
    raw_ext = ".csv" if is_csv else ".xlsx"
    raw_path = os.path.join(TEMP_DIR, f"{file_id}_raw{raw_ext}")
    input_path = os.path.join(TEMP_DIR, f"{file_id}_in.xlsx")
    output_path = os.path.join(TEMP_DIR, f"{file_id}_out.xlsx")

    try:
        with open(raw_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {str(e)}")

    if is_csv:
        try:
            csv_to_xlsx(raw_path, input_path)
            os.remove(raw_path)
        except Exception as e:
            for p in [raw_path, input_path]:
                if os.path.exists(p):
                    os.remove(p)
            raise HTTPException(status_code=400, detail=f"Failed to read CSV: {str(e)}")
    else:
        os.rename(raw_path, input_path)

    return StreamingResponse(
        process_stream(input_path, output_path, header_row, file_id, file.filename),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.get("/api/download/{file_id}")
async def api_download(file_id: str, filename: str = "output.xlsx"):
    output_path = os.path.join(TEMP_DIR, f"{file_id}_out.xlsx")
    if not os.path.exists(output_path):
        raise HTTPException(status_code=404, detail="File not found or has expired")
    return FileResponse(
        output_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)

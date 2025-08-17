import logging
from datetime import date
from typing import Optional, List
import io
import os
import re

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from app.services.excel import generate_excel

router = APIRouter()

class DayHours(BaseModel):
    work_date: date
    hours: float

class ExportRequest(BaseModel):
    employee_name: str
    designation: str
    email_primary: str
    email_secondary: Optional[str] = ""

    week_begin: Optional[date] = None
    week_end: Optional[date] = None
    days: Optional[List[DayHours]] = None
    client_name: Optional[str] = None

def _sanitize_name(name: str) -> str:
    s = re.sub(r"\s+", "_", name.strip())  # raw string, replace whitespace with underscore
    s = re.sub(r"[^\w]", "", s)             # remove non-word chars except underscore
    s = re.sub(r"_+", "_", s)               # collapse multiple underscores
    return s

def _ext_and_mime_from_template() -> tuple[str, str]:
    tpl = os.getenv("EXCEL_TEMPLATE_PATH", "template/Gudipati_Phani_Babu_Timesheet_Week_Ending_08152025.xlsx").lower()
    if tpl.endswith(".xlsm"):
        return ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12"
    return ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def _capitalize_preserving_rest(word: str) -> str:
    if not word:
        return word
    return word[0].upper() + word[1:]

def _build_filename(employee_name: str, week_end: Optional[date]) -> str:
    from datetime import date as _d
    end = week_end or _d.today()
    mmddyyyy = end.strftime("%m%d%Y")
    ext, _ = _ext_and_mime_from_template()

    sanitized_name = _sanitize_name(employee_name)
    parts = sanitized_name.split("_")   # underscore literal without backslash
    capitalized_parts = [_capitalize_preserving_rest(p) for p in parts]
    sanitized_name_corrected = "_".join(capitalized_parts)

    suffix_words = ["Timesheet", "Week", "Ending"]
    suffix = "_".join(suffix_words)

    filename = f"{sanitized_name_corrected}_{suffix}_{mmddyyyy}{ext}"
    return filename

@router.post("/weekly/download")
def export_weekly_download(req: ExportRequest):
    if not req.employee_name or not req.designation or not req.email_primary:
        raise HTTPException(status_code=400, detail="Missing required employee information.")

    try:
        xls_bytes = generate_excel(
            employee_name=req.employee_name,
            designation=req.designation,
            email_primary=req.email_primary,
            email_secondary=req.email_secondary,
            client_name=req.client_name,
            week_begin=req.week_begin,
            week_end=req.week_end,
            days=req.days,
        )
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=f"Template not found: {e}")
    except Exception:
        logging.exception("An unexpected error occurred during Excel generation.")
        raise HTTPException(status_code=500, detail="An unexpected export error occurred. See server logs for details.")

    filename = _build_filename(req.employee_name, req.week_end)
    _, content_type = _ext_and_mime_from_template()

    return StreamingResponse(
        io.BytesIO(xls_bytes),
        media_type=content_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )

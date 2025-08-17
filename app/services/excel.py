import logging
from io import BytesIO
from pathlib import Path
from typing import List, Optional, TYPE_CHECKING
from datetime import date
import openpyxl
import os

if TYPE_CHECKING:
    from app.routes.exports import DayHours

logging.basicConfig(level=logging.INFO)

def _resolve_template_path(filename: str) -> Path:
    cwd = Path(os.getcwd())
    candidate = cwd / "template" / filename
    logging.info(f"Checking template path: {candidate}")
    if candidate.exists():
        logging.info("Template found at root template folder.")
        return candidate
    else:
        logging.error(f"Template not found at: {candidate}")
        raise FileNotFoundError(f"Template not found at: {candidate}")

def generate_excel(
    employee_name: str,
    designation: str,
    email_primary: str,
    email_secondary: str,
    client_name: Optional[str] = None,
    week_begin: Optional[date] = None,
    week_end: Optional[date] = None,
    days: Optional[List["DayHours"]] = None,
) -> bytes:
    template_filename = "Gudipati_Phani_Babu_Timesheet_Week_Ending_08152025.xlsx"
    template_path = _resolve_template_path(template_filename)

    wb = openpyxl.load_workbook(template_path)
    ws = wb['Timesheet']  # Change if your sheet name is different

    ws["G2"] = employee_name
    ws["G3"] = designation
    ws["G4"] = email_primary
    ws["G5"] = email_secondary
    ws["B6"] = f"Client : {client_name or 'Burger King'}"

    if week_begin:
        ws["B9"] = week_begin.strftime("%m-%d-%Y")
    if week_end:
        ws["C9"] = week_end.strftime("%m-%d-%Y")

    sorted_days = sorted(days, key=lambda d: d.work_date) if days else []
    total_hours = 0
    for i in range(5):
        col_index = 3 + i  # Columns C, D, E, F, G
        row_date = 11
        row_hours = 12

        cell_date = ws.cell(row=row_date, column=col_index)
        cell_hours = ws.cell(row=row_hours, column=col_index)

        if i < len(sorted_days):
            day = sorted_days[i]
            hours = day.hours if day.hours is not None else 0
            total_hours += hours
            cell_date.value = day.work_date.strftime("%m-%d-%Y")
            cell_hours.value = hours
        else:
            cell_date.value = None
            cell_hours.value = None

    ws["D9"] = total_hours
    ws["E9"] = total_hours
    ws["F9"] = 0

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()

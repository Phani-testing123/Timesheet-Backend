# backend/app/services/excel.py

import os
import uuid
import logging
import xlwings as xw
from pathlib import Path
from typing import List, Optional, TYPE_CHECKING
from datetime import date

if TYPE_CHECKING:
    from app.routes.exports import DayHours

logging.basicConfig(level=logging.INFO)

def _resolve_template_path(filename: str) -> Path:
    """
    Resolves the absolute path to your Excel template file for Render when working directory is /opt/render/project/src.
    """
    tpl_env = os.getenv("EXCEL_TEMPLATE_PATH", filename)
    cwd = Path(os.getcwd())
    candidate = cwd / "backend" / "template" / os.path.basename(tpl_env)
    logging.info(f"Checking for Excel template at: {candidate}")
    if candidate.exists():
        logging.info("Template found in backend/template folder.")
        return candidate
    else:
        logging.error(f"Template not found at path: {candidate}")
        raise FileNotFoundError(f"Template not found at path: {candidate}")

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
    """
    Fill the original Excel template using xlwings to keep all formatting intact.
    Returns the filled Excel file as bytes.
    """
    logging.info(f"Generating Excel with xlwings for: {employee_name}")
    template_filename = "Gudipati_Phani_Babu_Timesheet_Week_Ending_08152025.xlsx"
    template_path = _resolve_template_path(template_filename)

    app = xw.App(visible=False)
    try:
        wb = app.books.open(str(template_path))
        ws = wb.sheets['Timesheet']  # Change if your sheet name is different

        ws.range("G2").value = employee_name
        ws.range("G3").value = designation
        ws.range("G4").value = email_primary
        ws.range("G5").value = email_secondary
        ws.range("B6").value = f"Client : {client_name or 'Burger King'}"

        if week_begin:
            ws.range("B9").value = week_begin.strftime("%m-%d-%Y")
        if week_end:
            ws.range("C9").value = week_end.strftime("%m-%d-%Y")

        sorted_days = sorted(days, key=lambda d: d.work_date) if days else []
        total_hours = 0
        for i in range(5):
            col_index = 3 + i
            date_cell = ws.cells(11, col_index)  
            hours_cell = ws.cells(12, col_index)
            if i < len(sorted_days):
                day = sorted_days[i]
                hours = day.hours if day.hours is not None else 0
                total_hours += hours
                date_cell.value = day.work_date.strftime("%m-%d-%Y")
                hours_cell.value = hours
            else:
                date_cell.value = ""
                hours_cell.value = ""

        ws.range("D9").value = total_hours
        ws.range("E9").value = total_hours
        ws.range("F9").value = 0

        user_desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(user_desktop):
            os.makedirs(user_desktop)
        unique_filename = f"export_{uuid.uuid4().hex}.xlsx"
        export_path = os.path.join(user_desktop, unique_filename)

        wb.save(export_path)
        wb.close()
        app.quit()

        with open(export_path, "rb") as f:
            bytes_data = f.read()

        os.remove(export_path)
        logging.info(f"Successfully generated Excel for: {employee_name}")
        return bytes_data

    except Exception:
        app.quit()
        logging.exception("An error occurred inside generate_excel with xlwings")
        raise

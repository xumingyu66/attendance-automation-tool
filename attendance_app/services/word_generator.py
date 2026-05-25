"""Generates Word attendance report."""

from __future__ import annotations


class WordReportGenerator:
    """Generates a Word document summarizing attendance data."""

    HOURS_PER_DAY = 2

    def __init__(self, docx_wrapper, db_dir: str = "") -> None:
        self._wrapper = docx_wrapper
        self._db_dir = db_dir

    def generate(
        self,
        output_path: str,
        attendance_data: dict[str, "EmployeeAttendance"],
        xlsx_employees: list[str],
        year: int = 2026,
        month: int = 4,
    ) -> None:
        """Generate the Word report with separate tables per position.

        Employees are sorted by employee number; those with 0 hours are excluded.
        """
        from services.attendance_parser import EmployeeAttendance

        # Build position lookup from database
        position_map = self._load_positions()

        report_data: list[dict] = []

        for name in xlsx_employees:
            att = attendance_data.get(name.strip())
            if att is None:
                continue
            if att.total_hours <= 0:
                continue
            report_data.append({
                "name": att.name,
                "employee_number": str(att.employee_id),
                "total_hours": att.total_hours,
                "work_days": len(att.work_days),
                "position": position_map.get(name.strip(), "值班岗"),
            })

        report_data.sort(key=lambda d: int(d["employee_number"]))

        # Split by position
        zhengjia = [d for d in report_data if d["position"] == "整架岗"]
        zhiban = [d for d in report_data if d["position"] == "值班岗"]

        self._wrapper.generate(output_path, year, month, zhengjia, zhiban)

    def _load_positions(self) -> dict[str, str]:
        """Load employee positions from the database."""
        try:
            from wrappers.db_handler import DatabaseHandler
            db = DatabaseHandler(self._db_dir)
            emps = db.get_all_employees()
            return {e.name: e.department for e in emps}
        except Exception:
            return {}

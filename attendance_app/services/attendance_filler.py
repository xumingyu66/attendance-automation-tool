"""Fills attendance hours into 文件考勤.xlsx."""

from __future__ import annotations


class AttendanceFiller:
    """Fills per-day hours into the attendance summary workbook."""

    HOURS_PER_DAY = 2

    def __init__(self, xlsx_handler) -> None:
        self._handler = xlsx_handler

    def fill(
        self,
        attendance_data: dict[str, "EmployeeAttendance"],
        year: int = 2026,
        month: int = 4,
    ) -> tuple[int, int]:
        """Fill attendance data into the xlsx workbook.

        Args:
            attendance_data: Dict of name -> EmployeeAttendance from parser.
            year, month: Target reporting period.

        Returns:
            (matched_count, unmatched_count) of employees.
        """
        from services.attendance_parser import EmployeeAttendance

        self._handler.update_weekdays(year, month)

        template_employees = self._handler.get_employees()
        matched = 0
        unmatched = 0

        for emp_row in template_employees:
            name = emp_row.name.strip()
            if name in attendance_data:
                att = attendance_data[name]
                daily = {
                    d: self.HOURS_PER_DAY if d in att.work_days else None
                    for d in range(1, 32)
                }
                self._handler.fill_hours(name, daily)
                matched += 1
            else:
                # Clear this employee's data since not in attendance report
                daily = {d: None for d in range(1, 32)}
                self._handler.fill_hours(name, daily)
                unmatched += 1

        return matched, unmatched

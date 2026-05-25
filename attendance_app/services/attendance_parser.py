"""Parses 考勤报表.xls and extracts attendance data."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class EmployeeAttendance:
    name: str
    employee_id: int
    work_days: set[int]  # days with attendance
    total_hours: int     # work_days * 2


class AttendanceParser:
    """Parses raw punch data from 考勤报表.xls into structured attendance records."""

    HOURS_PER_DAY = 2

    def __init__(self, xls_reader):
        self._reader = xls_reader

    def parse_all(self) -> dict[str, EmployeeAttendance]:
        """Parse all employee attendance data.

        Returns:
            Dict mapping employee name (stripped) -> EmployeeAttendance.
        """
        result: dict[str, EmployeeAttendance] = {}

        for punch_data in self._reader.iter_employees():
            work_days = {
                day for day, has_punch in punch_data.daily_has_punch.items()
                if has_punch
            }
            name = punch_data.name.strip()
            result[name] = EmployeeAttendance(
                name=name,
                employee_id=punch_data.employee_id,
                work_days=work_days,
                total_hours=len(work_days) * self.HOURS_PER_DAY,
            )

        return result

"""Import employee information from 考勤报表.xls into the database.

Reads the 员工排班表 (first sheet) to extract employee_number, name, and department.
"""

from __future__ import annotations

import xlrd

from wrappers.db_handler import Employee


class EmployeeImporter:
    """Extracts employee data from the first sheet (员工排班表) of 考勤报表.xls."""

    # Column indices in the 员工排班表 sheet
    COL_EMPLOYEE_NUMBER = 0
    COL_NAME = 1
    COL_DEPARTMENT = 2
    DATA_START_ROW = 4  # Rows 0-3 are headers

    def __init__(self, xls_path: str) -> None:
        self._path = xls_path

    def extract(self) -> list[Employee]:
        """Extract all employees from the 员工排班表 sheet.

        Returns:
            List of Employee objects ready for database import.
        """
        wb = xlrd.open_workbook(self._path)
        sheet = wb.sheet_by_index(0)  # First sheet = 员工排班表

        employees: list[Employee] = []

        for row in range(self.DATA_START_ROW, sheet.nrows):
            emp_num = self._cell_str(sheet, row, self.COL_EMPLOYEE_NUMBER)
            name = self._cell_str(sheet, row, self.COL_NAME)
            department = self._cell_str(sheet, row, self.COL_DEPARTMENT)

            if not emp_num or not name:
                continue

            # Skip header-like rows (employee_number should be numeric)
            try:
                int(float(emp_num))
            except ValueError:
                continue

            employees.append(Employee(
                employee_number=emp_num,
                name=name,
                department=department,
            ))

        wb.release_resources()
        return employees

    @staticmethod
    def _cell_str(sheet: xlrd.sheet.Sheet, row: int, col: int) -> str:
        if col >= sheet.ncols:
            return ""
        val = sheet.cell_value(row, col)
        if isinstance(val, float):
            if val == int(val):
                return str(int(val))
            return str(val)
        return str(val).strip()

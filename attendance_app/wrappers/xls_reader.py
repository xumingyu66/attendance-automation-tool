"""Typed wrapper for xlrd — reads .xls attendance punch data."""

from dataclasses import dataclass
from typing import Iterator

import xlrd


@dataclass
class EmployeePunchData:
    name: str
    employee_id: int
    daily_has_punch: dict[int, bool]  # day_of_month -> has_any_punch


class XlsReader:
    """Reads 考勤报表.xls and extracts per-employee daily punch records."""

    def __init__(self, filepath: str) -> None:
        self._path = filepath
        self._workbook = xlrd.open_workbook(filepath)

    @property
    def sheet_names(self) -> list[str]:
        return self._workbook.sheet_names()

    def iter_employees(self) -> Iterator[EmployeePunchData]:
        """Yield EmployeePunchData for each employee in the punch sheets.

        Skips the first 2 sheets (reference/lookup).
        Each data sheet contains up to 3 employees arranged in horizontal blocks.
        """
        for sheet_idx in range(2, self._workbook.nsheets):
            sheet = self._workbook.sheet_by_index(sheet_idx)
            yield from self._parse_sheet_employees(sheet)

    def _parse_sheet_employees(self, sheet: xlrd.sheet.Sheet) -> Iterator[EmployeePunchData]:
        """Parse up to 3 employees from a single punch data sheet.

        Each employee occupies a ~15-column horizontal block.
        Block starts at column 0, 15, 30.
        Employee name at Row 3, Col block_start+9.
        Employee ID at Row 4, Col block_start+9.
        Daily punch rows: R12 through R41 (30 days of April).
        """
        block_starts = [0, 15, 30]

        for bs in block_starts:
            name = self._cell_str(sheet, 3, bs + 9)
            emp_id = self._cell_int(sheet, 4, bs + 9)

            if not name or emp_id == 0:
                continue

            daily: dict[int, bool] = {}

            # Rows 12-41 = days 1-30
            for day_idx in range(30):
                row = 12 + day_idx
                day_num = day_idx + 1
                has_punch = self._block_has_any_data(sheet, row, bs, 15)
                daily[day_num] = has_punch

            yield EmployeePunchData(name=name, employee_id=emp_id, daily_has_punch=daily)

    def _block_has_any_data(self, sheet: xlrd.sheet.Sheet, row: int, col_start: int, width: int) -> bool:
        """Check if any cell in the given block contains punch data.

        Skips the first column of each block (date label like '01 四').
        """
        for c in range(col_start + 1, min(col_start + width, sheet.ncols)):
            val = sheet.cell_value(row, c)
            if val != "" and val != 0.0:
                return True
        return False

    @staticmethod
    def _cell_str(sheet: xlrd.sheet.Sheet, row: int, col: int) -> str:
        if col >= sheet.ncols:
            return ""
        val = sheet.cell_value(row, col)
        if isinstance(val, float) and val == int(val):
            return str(int(val))
        return str(val).strip()

    @staticmethod
    def _cell_int(sheet: xlrd.sheet.Sheet, row: int, col: int) -> int:
        if col >= sheet.ncols:
            return 0
        try:
            return int(sheet.cell_value(row, col))
        except (ValueError, TypeError):
            return 0

    def close(self) -> None:
        self._workbook.release_resources()

"""SQLite database handler for the attendance application.

Manages the employees table and attendance records.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from dataclasses import dataclass
from contextlib import contextmanager


@dataclass
class Employee:
    employee_number: str
    name: str
    department: str


class DatabaseHandler:
    """SQLite database wrapper for employee management."""

    DB_FILENAME = "attendance.db"

    def __init__(self, db_dir: str) -> None:
        self._db_path = Path(db_dir) / self.DB_FILENAME

    def init_db(self) -> None:
        """Create tables if they don't exist."""
        with self._get_conn() as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    employee_number VARCHAR(20) NOT NULL UNIQUE,
                    name VARCHAR(50) NOT NULL,
                    department VARCHAR(100)
                )
            """)
            conn.commit()

    def upsert_employees(self, employees: list[Employee]) -> tuple[int, int, int]:
        """Insert or update employee records.

        Returns:
            (inserted, updated, skipped) counts.
        """
        inserted = 0
        updated = 0
        skipped = 0

        with self._get_conn() as conn:
            for emp in employees:
                existing = conn.execute(
                    "SELECT id, name, department FROM employees WHERE employee_number = ?",
                    (emp.employee_number,),
                ).fetchone()

                if existing is None:
                    conn.execute(
                        "INSERT INTO employees (employee_number, name, department) VALUES (?, ?, ?)",
                        (emp.employee_number, emp.name, emp.department),
                    )
                    inserted += 1
                else:
                    if existing[1] != emp.name or (existing[2] or "") != emp.department:
                        conn.execute(
                            "UPDATE employees SET name = ?, department = ? WHERE employee_number = ?",
                            (emp.name, emp.department, emp.employee_number),
                        )
                        updated += 1
                    else:
                        skipped += 1

            conn.commit()

        return inserted, updated, skipped

    def get_all_employees(self) -> list[Employee]:
        """Return all employees from the database."""
        with self._get_conn() as conn:
            rows = conn.execute(
                "SELECT employee_number, name, department FROM employees ORDER BY CAST(employee_number AS INTEGER)"
            ).fetchall()
            return [
                Employee(employee_number=row[0], name=row[1], department=row[2] or "")
                for row in rows
            ]

    def get_employee_count(self) -> int:
        with self._get_conn() as conn:
            row = conn.execute("SELECT COUNT(*) FROM employees").fetchone()
            return row[0] if row else 0

    @contextmanager
    def _get_conn(self):
        conn = sqlite3.connect(str(self._db_path))
        try:
            yield conn
        finally:
            conn.close()

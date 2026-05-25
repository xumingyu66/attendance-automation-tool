"""Coordinates attendance parsing, filling, and Word report generation.

Emits Qt signals so the UI can display progress.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal

from wrappers.xls_reader import XlsReader
from wrappers.xlsx_handler import XlsxHandler
from wrappers.docx_wrapper import DocxReportGenerator as DocxWrapper
from wrappers.db_handler import DatabaseHandler
from services.attendance_parser import AttendanceParser
from services.attendance_filler import AttendanceFiller
from services.word_generator import WordReportGenerator as WordGenService
from services.employee_importer import EmployeeImporter
from services.template_generator import TemplateGenerator


class AttendanceManager(QObject):
    progress = Signal(str)
    finished = Signal(str, str)  # xlsx_output, docx_output
    import_finished = Signal(int, int, int)  # inserted, updated, skipped
    template_generated = Signal(str)  # output_path
    error = Signal(str)

    def __init__(self) -> None:
        super().__init__()

    def process(
        self,
        attendance_xls_path: str,
        template_xlsx_path: str,
        output_dir: str,
    ) -> None:
        """Run the full attendance processing pipeline.

        Steps:
          1. Parse 考勤报表.xls -> per-employee punch data
          2. Read 文件考勤.xlsx template
          3. Update weekdays + fill attendance hours
          4. Generate Word report
          5. Save outputs
        """
        try:
            # Step 1: Parse attendance punch data
            self.progress.emit("正在读取考勤报表...")
            xls_reader = XlsReader(attendance_xls_path)
            parser = AttendanceParser(xls_reader)
            attendance_data = parser.parse_all()
            year, month = self._read_report_month(xls_reader)
            self.progress.emit(f"考勤报表解析完成，共 {len(attendance_data)} 人，报表年月: {year}年{month}月")

            # Step 2: Read template and get employee list
            self.progress.emit("正在读取考勤模板...")
            xlsx_handler = XlsxHandler(template_xlsx_path)
            template_employees = xlsx_handler.get_employees()
            employee_names = [e.name.strip() for e in template_employees]
            self.progress.emit(f"考勤模板读取完成，在岗 {len(employee_names)} 人")

            # Step 3: Fill attendance data
            self.progress.emit("正在填充考勤数据并更新日期...")
            filler = AttendanceFiller(xlsx_handler)
            matched, unmatched = filler.fill(attendance_data, year, month)
            self.progress.emit(f"数据填充完成 (匹配: {matched}, 未匹配: {unmatched})")

            # Save filled xlsx
            xlsx_output = str(Path(output_dir) / f"{year}年{month}月考勤统计.xlsx")
            xlsx_handler.save(xlsx_output)
            self.progress.emit(f"已保存考勤表: {xlsx_output}")

            # Step 4: Generate Word report
            self.progress.emit("正在生成Word报告...")
            docx_output = str(Path(output_dir) / f"图书馆在岗学生馆员{year}年{month}月工作时长统计表.docx")
            docx_wrapper = DocxWrapper()
            word_gen = WordGenService(docx_wrapper, output_dir)
            word_gen.generate(docx_output, attendance_data, employee_names, year, month)
            self.progress.emit(f"已保存Word报告: {docx_output}")

            # Cleanup
            xls_reader.close()

            self.progress.emit("处理完成！")
            self.finished.emit(xlsx_output, docx_output)

        except Exception as e:
            self.error.emit(f"处理失败: {e}")
            raise

    def import_employees(self, xls_path: str, output_dir: str) -> None:
        """Import employee data from 考勤报表.xls 员工排班表 into the database.

        Args:
            xls_path: Path to 考勤报表.xls.
            output_dir: Directory where the database file will be created.
        """
        try:
            self.progress.emit("正在从员工排班表提取员工信息...")
            importer = EmployeeImporter(xls_path)
            employees = importer.extract()
            self.progress.emit(f"提取到 {len(employees)} 名员工信息")

            self.progress.emit("正在写入数据库...")
            db = DatabaseHandler(output_dir)
            db.init_db()
            inserted, updated, skipped = db.upsert_employees(employees)
            self.progress.emit(
                f"数据库写入完成 (新增: {inserted}, 更新: {updated}, 跳过: {skipped})"
            )

            self.progress.emit("员工信息导入完成！")
            self.import_finished.emit(inserted, updated, skipped)

        except Exception as e:
            self.error.emit(f"导入失败: {e}")
            raise

    def generate_template(self, output_dir: str, year: int, month: int) -> None:
        """Generate a blank attendance template from database employee data.

        Args:
            output_dir: Directory to save the generated template.
            year, month: Target reporting period.
        """
        try:
            db_dir = output_dir
            self.progress.emit("正在从数据库读取员工信息...")
            gen = TemplateGenerator(db_dir)

            output_path = str(Path(output_dir) / f"{year}年{month}月考勤模板.xlsx")
            self.progress.emit("正在生成空白考勤模板...")
            count = gen.generate(output_path, year, month)
            saved = gen.last_saved_path or output_path

            self.progress.emit(f"模板生成完成，共 {count} 名员工，保存至: {saved}")
            self.template_generated.emit(saved)

        except Exception as e:
            self.error.emit(f"模板生成失败: {e}")
            raise

    @staticmethod
    def _read_report_month(xls_reader) -> tuple[int, int]:
        """Extract year and month from the 考勤报表.xls first sheet.

        Parses R1: '统计日期：2026-04-01至2026-04-30'
        """
        import xlrd
        wb = xlrd.open_workbook(xls_reader._path)
        sheet = wb.sheet_by_index(0)
        date_str = str(sheet.cell_value(1, 0))  # R1C0
        # Extract "2026-04-01" pattern
        import re
        match = re.search(r'(\d{4})-(\d{1,2})-\d{1,2}', date_str)
        if match:
            return int(match.group(1)), int(match.group(2))
        return 2026, 4  # fallback

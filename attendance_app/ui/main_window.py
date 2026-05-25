"""Main window for the attendance automation application."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QComboBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from managers.attendance_manager import AttendanceManager


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("考勤自动化处理工具")
        self.setMinimumSize(700, 600)

        self._manager = AttendanceManager()
        self._manager.progress.connect(self._on_progress)
        self._manager.finished.connect(self._on_finished)
        self._manager.import_finished.connect(self._on_import_finished)
        self._manager.template_generated.connect(self._on_template_generated)
        self._manager.error.connect(self._on_error)

        self._output_xlsx = ""
        self._output_docx = ""
        self._busy = False

        self._setup_ui()
        self._load_settings()

    def _setup_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(0)
        layout.setContentsMargins(0, 0, 0, 0)

        tabs = QTabWidget()
        tabs.addTab(self._build_attendance_tab(), "考勤处理")
        tabs.addTab(self._build_import_tab(), "员工信息导入")
        tabs.addTab(self._build_employee_tab(), "员工数据")
        layout.addWidget(tabs)

    # =========================================================================
    # Tab 1: Attendance Processing
    # =========================================================================

    def _build_attendance_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # File Selection
        file_group = QGroupBox("输入文件")
        file_layout = QVBoxLayout(file_group)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("考勤报表 (.xls):"))
        self._xls_edit = QLineEdit()
        self._xls_edit.setPlaceholderText("选择考勤报表.xls文件...")
        self._xls_edit.setReadOnly(True)
        row1.addWidget(self._xls_edit, 1)
        xls_btn = QPushButton("浏览...")
        xls_btn.clicked.connect(self._browse_xls)
        row1.addWidget(xls_btn)
        file_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("考勤模板 (.xlsx):"))
        self._xlsx_edit = QLineEdit()
        self._xlsx_edit.setPlaceholderText("选择文件考勤.xlsx文件...")
        self._xlsx_edit.setReadOnly(True)
        row2.addWidget(self._xlsx_edit, 1)
        xlsx_btn = QPushButton("浏览...")
        xlsx_btn.clicked.connect(self._browse_xlsx)
        row2.addWidget(xlsx_btn)
        file_layout.addLayout(row2)

        layout.addWidget(file_group)

        # Output
        out_group = QGroupBox("输出目录")
        out_layout = QHBoxLayout(out_group)
        self._out_edit = QLineEdit()
        self._out_edit.setPlaceholderText("选择输出目录（默认为模板文件所在目录）...")
        self._out_edit.setReadOnly(True)
        out_layout.addWidget(self._out_edit, 1)
        out_btn = QPushButton("浏览...")
        out_btn.clicked.connect(self._browse_output)
        out_layout.addWidget(out_btn)
        layout.addWidget(out_group)

        # Action
        btn_layout = QHBoxLayout()
        self._process_btn = QPushButton("开始考勤处理")
        self._process_btn.setMinimumHeight(40)
        self._process_btn.setStyleSheet("QPushButton { font-size: 14px; font-weight: bold; }")
        self._process_btn.clicked.connect(self._on_process)
        btn_layout.addWidget(self._process_btn, 1)
        layout.addLayout(btn_layout)

        # Progress
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setVisible(False)
        layout.addWidget(self._progress_bar)

        # Log
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout(log_group)
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setMaximumHeight(150)
        log_layout.addWidget(self._log)
        layout.addWidget(log_group)

        return w

    # =========================================================================
    # Tab 2: Employee Import
    # =========================================================================

    def _build_import_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # File
        file_group = QGroupBox("数据源")
        file_layout = QHBoxLayout(file_group)
        file_layout.addWidget(QLabel("考勤报表 (.xls):"))
        self._import_xls_edit = QLineEdit()
        self._import_xls_edit.setPlaceholderText("选择考勤报表.xls，将从员工排班表提取信息...")
        self._import_xls_edit.setReadOnly(True)
        file_layout.addWidget(self._import_xls_edit, 1)
        import_xls_btn = QPushButton("浏览...")
        import_xls_btn.clicked.connect(self._browse_import_xls)
        file_layout.addWidget(import_xls_btn)
        layout.addWidget(file_group)

        # Output dir (DB location)
        out_group = QGroupBox("数据库位置")
        out_layout = QHBoxLayout(out_group)
        self._import_out_edit = QLineEdit()
        self._import_out_edit.setPlaceholderText("选择数据库保存目录...")
        self._import_out_edit.setReadOnly(True)
        out_layout.addWidget(self._import_out_edit, 1)
        import_out_btn = QPushButton("浏览...")
        import_out_btn.clicked.connect(self._browse_import_output)
        out_layout.addWidget(import_out_btn)
        layout.addWidget(out_group)

        # Info
        info_label = QLabel(
            "将从考勤报表的\"员工排班表\"中提取：员工号、姓名、考勤地区（岗位），"
            "写入 attendance.db 数据库的 employees 表中。\n"
            "已存在的员工号会自动更新姓名和岗位信息。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        # Action
        btn_layout = QHBoxLayout()
        self._import_btn = QPushButton("导入员工信息到数据库")
        self._import_btn.setMinimumHeight(40)
        self._import_btn.setStyleSheet("QPushButton { font-size: 14px; font-weight: bold; }")
        self._import_btn.clicked.connect(self._on_import)
        btn_layout.addWidget(self._import_btn, 1)
        layout.addLayout(btn_layout)

        # Separator
        sep = QLabel("— 从数据库生成考勤模板 —")
        sep.setStyleSheet("color: #666; padding-top: 8px;")
        sep.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(sep)

        # Template generation section
        tmpl_group = QGroupBox("模板生成")
        tmpl_layout = QVBoxLayout(tmpl_group)

        # Month selector
        month_row = QHBoxLayout()
        month_row.addWidget(QLabel("目标年月:"))
        self._tmpl_year_edit = QLineEdit("2026")
        self._tmpl_year_edit.setMaximumWidth(60)
        month_row.addWidget(self._tmpl_year_edit)
        month_row.addWidget(QLabel("年"))
        self._tmpl_month_combo = QComboBox()
        self._tmpl_month_combo.addItems([str(i) for i in range(1, 13)])
        self._tmpl_month_combo.setCurrentIndex(3)  # April
        month_row.addWidget(self._tmpl_month_combo)
        month_row.addWidget(QLabel("月"))
        month_row.addStretch()
        tmpl_layout.addLayout(month_row)

        gen_btn = QPushButton("生成空白考勤模板")
        gen_btn.setMinimumHeight(35)
        gen_btn.clicked.connect(self._on_generate_template)
        tmpl_layout.addWidget(gen_btn)
        layout.addWidget(tmpl_group)

        # Progress
        self._import_progress = QProgressBar()
        self._import_progress.setRange(0, 0)
        self._import_progress.setVisible(False)
        layout.addWidget(self._import_progress)

        # Log
        log_group = QGroupBox("导入日志")
        log_layout = QVBoxLayout(log_group)
        self._import_log = QTextEdit()
        self._import_log.setReadOnly(True)
        self._import_log.setMaximumHeight(150)
        log_layout.addWidget(self._import_log)
        layout.addWidget(log_group)

        return w

    # =========================================================================
    # Tab 3: Employee Data
    # =========================================================================

    def _build_employee_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("数据库中所有员工信息"))
        bar.addStretch()
        refresh_btn = QPushButton("刷新")
        refresh_btn.clicked.connect(self._on_refresh_employees)
        bar.addWidget(refresh_btn)
        layout.addLayout(bar)

        self._emp_table = QTableWidget()
        self._emp_table.setColumnCount(3)
        self._emp_table.setHorizontalHeaderLabels(["工号", "姓名", "岗位"])
        self._emp_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._emp_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._emp_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._emp_table.setAlternatingRowColors(True)
        layout.addWidget(self._emp_table)

        return w

    # =========================================================================
    # Slots - Attendance Tab
    # =========================================================================

    def _browse_xls(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择考勤报表", "", "Excel文件 (*.xls)")
        if path:
            self._xls_edit.setText(path)
            if not self._out_edit.text():
                self._out_edit.setText(str(Path(path).parent))

    def _browse_xlsx(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择考勤模板", "", "Excel文件 (*.xlsx)")
        if path:
            self._xlsx_edit.setText(path)
            if not self._out_edit.text():
                self._out_edit.setText(str(Path(path).parent))

    def _browse_output(self) -> None:
        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if dir_path:
            self._out_edit.setText(dir_path)

    def _on_process(self) -> None:
        xls_path = self._xls_edit.text().strip()
        xlsx_path = self._xlsx_edit.text().strip()
        out_dir = self._out_edit.text().strip()

        if not xls_path:
            QMessageBox.warning(self, "提示", "请先选择考勤报表文件 (.xls)")
            return
        if not xlsx_path:
            QMessageBox.warning(self, "提示", "请先选择考勤模板文件 (.xlsx)")
            return
        if not out_dir:
            out_dir = str(Path(xlsx_path).parent)
            self._out_edit.setText(out_dir)

        self._log.clear()
        self._set_busy(True)
        self._save_settings()
        self._manager.process(xls_path, xlsx_path, out_dir)

    def _on_finished(self, xlsx_output: str, docx_output: str) -> None:
        self._set_busy(False)
        self._output_xlsx = xlsx_output
        self._output_docx = docx_output

        result = (
            f"处理完成！\n\n"
            f"考勤表: {xlsx_output}\n"
            f"Word报告: {docx_output}\n\n"
            f"是否打开输出文件夹？"
        )
        reply = QMessageBox.question(
            self, "处理完成", result,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            import os
            os.startfile(str(Path(xlsx_output).parent))

    # =========================================================================
    # Slots - Import Tab
    # =========================================================================

    def _browse_import_xls(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择考勤报表", "", "Excel文件 (*.xls)")
        if path:
            self._import_xls_edit.setText(path)
            if not self._import_out_edit.text():
                self._import_out_edit.setText(str(Path(path).parent))

    def _browse_import_output(self) -> None:
        dir_path = QFileDialog.getExistingDirectory(self, "选择数据库保存目录")
        if dir_path:
            self._import_out_edit.setText(dir_path)

    def _on_import(self) -> None:
        xls_path = self._import_xls_edit.text().strip()
        out_dir = self._import_out_edit.text().strip()

        if not xls_path:
            QMessageBox.warning(self, "提示", "请先选择考勤报表文件 (.xls)")
            return
        if not out_dir:
            out_dir = str(Path(xls_path).parent)
            self._import_out_edit.setText(out_dir)

        self._import_log.clear()
        self._set_busy(True)
        self._manager.import_employees(xls_path, out_dir)

    def _on_import_finished(self, inserted: int, updated: int, skipped: int) -> None:
        self._set_busy(False)
        self._on_refresh_employees()
        result = (
            f"导入完成！\n\n"
            f"新增: {inserted} 人\n"
            f"更新: {updated} 人\n"
            f"跳过: {skipped} 人\n\n"
            f"数据库文件: attendance.db"
        )
        QMessageBox.information(self, "导入完成", result)

    def _on_generate_template(self) -> None:
        out_dir = self._import_out_edit.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "提示", "请先选择数据库所在目录（数据库位置）")
            return

        try:
            year = int(self._tmpl_year_edit.text().strip())
            month = int(self._tmpl_month_combo.currentText())
        except ValueError:
            QMessageBox.warning(self, "提示", "请输入有效的年份")
            return

        if year < 2020 or year > 2100:
            QMessageBox.warning(self, "提示", "请输入合理的年份（2020-2100）")
            return

        self._import_log.clear()
        self._set_busy(True)
        self._manager.generate_template(out_dir, year, month)

    def _on_refresh_employees(self) -> None:
        """Load employees from the database and display in the table."""
        out_dir = self._import_out_edit.text().strip()
        if not out_dir:
            self._emp_table.setRowCount(0)
            return

        try:
            from wrappers.db_handler import DatabaseHandler
            db = DatabaseHandler(out_dir)
            employees = db.get_all_employees()
        except Exception:
            self._emp_table.setRowCount(0)
            return

        self._emp_table.setRowCount(len(employees))
        for i, emp in enumerate(employees):
            self._emp_table.setItem(i, 0, QTableWidgetItem(emp.employee_number))
            self._emp_table.setItem(i, 1, QTableWidgetItem(emp.name))
            self._emp_table.setItem(i, 2, QTableWidgetItem(emp.department))

    def _on_template_generated(self, output_path: str) -> None:
        self._set_busy(False)
        result = (
            f"模板生成完成！\n\n"
            f"文件: {output_path}\n\n"
            f"是否打开所在文件夹？"
        )
        reply = QMessageBox.question(
            self, "生成完成", result,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            import os
            os.startfile(str(Path(output_path).parent))

    # =========================================================================
    # Shared Slots
    # =========================================================================

    def _on_progress(self, message: str) -> None:
        self._log.append(message)
        self._import_log.append(message)

    def _on_error(self, message: str) -> None:
        self._set_busy(False)
        self._log.append(f"[错误] {message}")
        self._import_log.append(f"[错误] {message}")
        QMessageBox.critical(self, "错误", message)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self._progress_bar.setVisible(busy)
        self._import_progress.setVisible(busy)
        self._process_btn.setEnabled(not busy)
        self._import_btn.setEnabled(not busy)

    # =========================================================================
    # Settings
    # =========================================================================

    def _save_settings(self) -> None:
        from PySide6.QtCore import QSettings
        s = QSettings("AttendanceApp", "AttendanceApp")
        s.setValue("xls_path", self._xls_edit.text())
        s.setValue("xlsx_path", self._xlsx_edit.text())
        s.setValue("output_dir", self._out_edit.text())

    def _load_settings(self) -> None:
        from PySide6.QtCore import QSettings
        s = QSettings("AttendanceApp", "AttendanceApp")
        xls = s.value("xls_path", "")
        xlsx = s.value("xlsx_path", "")
        out = s.value("output_dir", "")
        if xls:
            self._xls_edit.setText(str(xls))
            self._import_xls_edit.setText(str(xls))
        if xlsx:
            self._xlsx_edit.setText(str(xlsx))
        if out:
            self._out_edit.setText(str(out))
            self._import_out_edit.setText(str(out))

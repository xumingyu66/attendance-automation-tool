"""Main window for the attendance automation application."""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractButton,
    QApplication,
    QFileDialog,
    QFrame,
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
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from managers.attendance_manager import AttendanceManager


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("考勤自动化处理工具")
        self.setMinimumSize(700, 600)
        self.setWindowIcon(QIcon(str(Path(__file__).parent / "app_icon.png")))

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
        main_layout = QHBoxLayout(central)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Stacked content pages (created first, referenced by sidebar)
        self._stack = QStackedWidget()
        self._stack.addWidget(self._build_attendance_tab())  # 0
        self._stack.addWidget(self._build_import_tab())       # 1
        self._stack.addWidget(self._build_employee_tab())     # 2

        # Left sidebar
        sidebar = self._build_sidebar()
        main_layout.addWidget(sidebar)

        # Right content area
        right = QWidget()
        right.setObjectName("rightArea")
        right_layout = QVBoxLayout(right)
        right_layout.setSpacing(0)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Header
        header = self._build_header()
        right_layout.addWidget(header)

        # Stacked pages
        right_layout.addWidget(self._stack, 1)

        main_layout.addWidget(right, 1)

        self._apply_styles()

    def _on_nav_clicked(self, btn: QAbstractButton) -> None:
        """Switch content page when a sidebar nav button is clicked."""
        self._stack.setCurrentIndex(self._nav_btns.index(btn))
        if self._stack.currentIndex() == 2:
            self._on_refresh_employees()

    # =========================================================================
    # Sidebar, Header, and Styles
    # =========================================================================

    def _build_sidebar(self) -> QWidget:
        """Build the left sidebar navigation."""
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(220)

        layout = QVBoxLayout(sidebar)
        layout.setSpacing(2)
        layout.setContentsMargins(0, 0, 0, 16)

        # App title area
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(20, 24, 20, 20)

        icon_path = Path(__file__).parent / "app_icon.png"
        icon_label = QLabel()
        icon_label.setPixmap(QIcon(str(icon_path)).pixmap(28, 28))
        icon_label.setFixedSize(28, 28)
        title_layout.addWidget(icon_label)
        title_layout.addSpacing(10)

        title = QLabel("考勤自动化")
        title.setObjectName("sidebarTitle")
        title_layout.addWidget(title)
        title_layout.addStretch()
        layout.addLayout(title_layout)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("background-color: #E2E8F0; max-height: 1px;")
        layout.addWidget(sep)
        layout.addSpacing(8)

        # Nav label
        nav_label = QLabel("  功能导航")
        nav_label.setStyleSheet("color: #94A3B8; font-size: 11px; padding: 8px 16px 4px;")
        layout.addWidget(nav_label)

        # Nav buttons
        self._nav_btns: list[QPushButton] = []
        nav_data = [
            ("📊", "考勤处理"),
            ("📥", "员工导入"),
            ("👥", "员工数据"),
        ]

        for i, (icon, text) in enumerate(nav_data):
            btn = QPushButton(f"  {icon}  {text}")
            btn.setObjectName("navButton")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _checked, idx=i: self._on_nav_clicked(self._nav_btns[idx]))
            self._nav_btns.append(btn)
            layout.addWidget(btn)

        layout.addStretch()

        # Version
        ver_label = QLabel("v1.0.0")
        ver_label.setStyleSheet("color: #CBD5E1; font-size: 11px; padding: 8px 20px;")
        layout.addWidget(ver_label)

        # Default to first nav
        self._nav_btns[0].setChecked(True)

        return sidebar

    def _build_header(self) -> QWidget:
        """Build the top header bar."""
        header = QWidget()
        header.setObjectName("header")
        layout = QVBoxLayout(header)
        layout.setContentsMargins(32, 24, 32, 20)

        title = QLabel("考勤自动化处理工具")
        title.setObjectName("headerTitle")
        layout.addWidget(title)

        subtitle = QLabel("高效处理考勤数据，自动生成统计报表")
        subtitle.setObjectName("headerSubtitle")
        layout.addWidget(subtitle)

        return header

    def _apply_styles(self) -> None:
        """Apply application-wide QSS stylesheet."""
        self.setStyleSheet("""
            * {
                font-family: "Microsoft YaHei", "Segoe UI", "PingFang SC", sans-serif;
                font-size: 13px;
                color: #1F2937;
            }
            QMainWindow, QWidget#rightArea {
                background-color: #FFFFFF;
            }
            QWidget#sidebar {
                background-color: #F8FAFC;
                border-right: 1px solid #E2E8F0;
            }
            QWidget#sidebar QLabel#sidebarTitle {
                font-size: 18px;
                font-weight: bold;
                color: #0F172A;
            }
            QPushButton#navButton {
                background-color: transparent;
                border: none;
                border-radius: 8px;
                text-align: left;
                padding: 10px 16px;
                margin: 2px 12px;
                font-size: 14px;
                color: #475569;
            }
            QPushButton#navButton:hover {
                background-color: #E2E8F0;
                color: #0F172A;
            }
            QPushButton#navButton:checked {
                background-color: #E8F5E9;
                color: #00BFA5;
                font-weight: bold;
            }
            QWidget#header {
                background-color: #FFFFFF;
                border-bottom: 1px solid #E2E8F0;
            }
            QWidget#header QLabel#headerTitle {
                font-size: 22px;
                font-weight: bold;
                color: #0F172A;
            }
            QWidget#header QLabel#headerSubtitle {
                font-size: 13px;
                color: #94A3B8;
                margin-top: 2px;
            }
            QGroupBox {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                margin-top: 20px;
                padding: 20px 20px 16px;
                font-size: 14px;
                font-weight: bold;
                color: #0F172A;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 16px;
                padding: 0 8px;
                background-color: #FFFFFF;
            }
            QPushButton#primaryBtn {
                background-color: #00BFA5;
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 10px 28px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton#primaryBtn:hover {
                background-color: #00A896;
            }
            QPushButton#primaryBtn:pressed {
                background-color: #00897B;
            }
            QPushButton#primaryBtn:disabled {
                background-color: #CBD5E1;
                color: #94A3B8;
            }
            QPushButton#browseBtn {
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px 16px;
                font-size: 13px;
                font-weight: normal;
            }
            QPushButton#browseBtn:hover {
                background-color: #E2E8F0;
            }
            QLineEdit {
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 13px;
                color: #0F172A;
                background-color: #F8FAFC;
            }
            QLineEdit:read-only {
                background-color: #F1F5F9;
            }
            QLineEdit:focus {
                border-color: #00BFA5;
                background-color: #FFFFFF;
            }
            QTableWidget {
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                gridline-color: #F1F5F9;
                background-color: #FFFFFF;
                selection-background-color: #E8F5E9;
                selection-color: #0F172A;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                padding: 10px 12px;
                border: none;
                border-bottom: 2px solid #E2E8F0;
                font-weight: bold;
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 8px 12px;
            }
            QTableWidget::item:selected {
                background-color: #E8F5E9;
                color: #0F172A;
            }
            QComboBox {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 13px;
                color: #0F172A;
                background-color: #FFFFFF;
            }
            QComboBox:hover {
                border-color: #00BFA5;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 8px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                selection-background-color: #E8F5E9;
                selection-color: #0F172A;
            }
            QProgressBar {
                border: none;
                border-radius: 4px;
                background-color: #E2E8F0;
                min-height: 6px;
                max-height: 6px;
                text-align: center;
                font-size: 11px;
                color: #94A3B8;
            }
            QProgressBar::chunk {
                background-color: #00BFA5;
                border-radius: 4px;
            }
            QScrollBar:vertical {
                border: none;
                background: #F1F5F9;
                width: 6px;
                border-radius: 3px;
            }
            QScrollBar::handle:vertical {
                background: #CBD5E1;
                border-radius: 3px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94A3B8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
        """)

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
        xls_btn.setObjectName("browseBtn")
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
        xlsx_btn.setObjectName("browseBtn")
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
        out_btn.setObjectName("browseBtn")
        out_btn.clicked.connect(self._browse_output)
        out_layout.addWidget(out_btn)
        layout.addWidget(out_group)

        # Action
        btn_layout = QHBoxLayout()
        self._process_btn = QPushButton("开始考勤处理")
        self._process_btn.setObjectName("primaryBtn")
        self._process_btn.setMinimumHeight(40)
        self._process_btn.clicked.connect(self._on_process)
        btn_layout.addWidget(self._process_btn, 1)
        layout.addLayout(btn_layout)

        # Progress
        self._progress_bar = QProgressBar()
        self._progress_bar.setRange(0, 0)
        self._progress_bar.setVisible(False)
        layout.addWidget(self._progress_bar)

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
        import_xls_btn.setObjectName("browseBtn")
        import_xls_btn.clicked.connect(self._browse_import_xls)
        file_layout.addWidget(import_xls_btn)
        layout.addWidget(file_group)

        # Info
        info_label = QLabel(
            "将从考勤报表的\"员工排班表\"中提取：员工号、姓名、考勤地区（岗位），"
            "写入 attendance.db 数据库中。\n"
            "已存在的员工号会自动更新姓名和岗位信息。"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)

        # Action
        btn_layout = QHBoxLayout()
        self._import_btn = QPushButton("导入员工信息到数据库")
        self._import_btn.setObjectName("primaryBtn")
        self._import_btn.setMinimumHeight(40)
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
        gen_btn.setObjectName("primaryBtn")
        gen_btn.setMinimumHeight(35)
        gen_btn.clicked.connect(self._on_generate_template)
        tmpl_layout.addWidget(gen_btn)
        layout.addWidget(tmpl_group)

        # Progress
        self._import_progress = QProgressBar()
        self._import_progress.setRange(0, 0)
        self._import_progress.setVisible(False)
        layout.addWidget(self._import_progress)

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
        refresh_btn.setObjectName("primaryBtn")
        refresh_btn.clicked.connect(self._on_refresh_employees)
        bar.addWidget(refresh_btn)
        layout.addLayout(bar)

        self._emp_table = QTableWidget()
        self._emp_table.setColumnCount(4)
        self._emp_table.setHorizontalHeaderLabels(["工号", "姓名", "岗位", "值班岗位"])
        self._emp_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._emp_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._emp_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._emp_table.setAlternatingRowColors(True)
        self._emp_table.verticalHeader().setVisible(False)
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

    def _on_import(self) -> None:
        xls_path = self._import_xls_edit.text().strip()
        if not xls_path:
            QMessageBox.warning(self, "提示", "请先选择考勤报表文件 (.xls)")
            return

        self._set_busy(True)
        self._manager.import_employees(xls_path, self._get_db_dir())

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
        try:
            year = int(self._tmpl_year_edit.text().strip())
            month = int(self._tmpl_month_combo.currentText())
        except ValueError:
            QMessageBox.warning(self, "提示", "请输入有效的年份")
            return

        if year < 2020 or year > 2100:
            QMessageBox.warning(self, "提示", "请输入合理的年份（2020-2100）")
            return

        self._set_busy(True)
        self._manager.generate_template(self._get_db_dir(), year, month)

    def _on_refresh_employees(self) -> None:
        """Load employees from the database and display in the table."""
        try:
            from wrappers.db_handler import DatabaseHandler
            db = DatabaseHandler(self._get_db_dir())
            employees = db.get_all_employees()
        except Exception:
            self._emp_table.setRowCount(0)
            return

        self._emp_table.setRowCount(len(employees))
        for i, emp in enumerate(employees):
            self._emp_table.setItem(i, 0, QTableWidgetItem(emp.employee_number))
            self._emp_table.setItem(i, 1, QTableWidgetItem(emp.name))
            self._emp_table.setItem(i, 2, QTableWidgetItem(emp.department))

            # Position selector (值班岗 / 整架岗)
            combo = QComboBox()
            combo.addItems(["值班岗", "整架岗"])
            combo.setCurrentText(emp.position)
            combo.currentTextChanged.connect(
                lambda text, num=emp.employee_number: self._on_position_changed(num, text)
            )
            self._emp_table.setCellWidget(i, 3, combo)

    def _on_position_changed(self, employee_number: str, position: str) -> None:
        """Save position change to the database."""
        try:
            from wrappers.db_handler import DatabaseHandler
            db = DatabaseHandler(self._get_db_dir())
            db.update_employee_position(employee_number, position)
        except Exception:
            pass

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
        pass

    def _on_error(self, message: str) -> None:
        self._set_busy(False)
        QMessageBox.critical(self, "错误", message)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self._progress_bar.setVisible(busy)
        self._import_progress.setVisible(busy)
        self._process_btn.setEnabled(not busy)
        self._import_btn.setEnabled(not busy)

    @staticmethod
    def _get_db_dir() -> str:
        """Return the path to the database storage folder."""
        if getattr(sys, 'frozen', False):
            root = Path(sys.executable).parent.parent
        else:
            root = Path(__file__).parent.parent.parent
        db_dir = root / "attendance_app" / "output"
        db_dir.mkdir(exist_ok=True)
        return str(db_dir)

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

"""
CSVtoExcel desktop app (PySide6) with SQLite history storage.

Features:
- PySide6 GUI with modern QSS styling
- Browse & load CSV files, preview first 100 rows
- Convert CSV -> Excel (.xlsx)
- Save conversion history to SQLite (file paths, timestamp, rows, cols)
- History viewer: open folder, re-export, delete records
- Full responsive layout with toolbar-style buttons
- Alternating row colors, hover highlight in preview table
"""

import sys 
import os 
import sqlite3 
import traceback 
from datetime import datetime 
from pathlib import Path 

# Third-party 
from PySide6 import QtCore, QtGui, QtWidgets 

from PySide6.QtWidgets import QDialog, QVBoxLayout, QTextEdit

from PySide6.QtGui import QIcon

APP_NAME = "CSVtoExcel"
DB_FILENAME = "csv_to_excel_history.db"

def resource_path(file_name):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, file_name)

# ------------------------- Helper: App data path -------------------------
def get_app_dir() -> Path:
    if sys.platform.startswith("win"):
        base = Path(os.getenv('LOCALAPPDATA') or Path.home() / 'AppData' / 'Local')
    elif sys.platform.startswith("darwin"):
        base = Path.home() / 'Library' / 'Application Support'
    else:
        base = Path(os.getenv('XDG_DATA_HOME') or Path.home() / '.local' / 'share')
    app_dir = base / "csv_to_excel_app"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir

APP_DIR = get_app_dir()
DB_PATH = APP_DIR / DB_FILENAME

# ------------------------- Database utilities -------------------------
def init_db(conn: sqlite3.Connection):
    c = conn.cursor()
    c.execute(
        '''CREATE TABLE IF NOT EXISTS history(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            src_path TEXT NOT NULL,
            dest_path TEXT NOT NULL,
            timestamp TEXT NOT NULL,
            rows INTEGER,
            cols INTEGER,
            has_header INTEGER DEFAULT 1
        )'''
    )
    conn.commit()

# ------------------------- Helper: save to Excel with xlsxwriter -------------------------
def save_to_excel_xlsxwriter(headers, data_rows, out_path, has_header=True):
    """
    headers: list of header strings
    data_rows: list of row lists (each row is a list of cell values)
    out_path: path-like or str to write .xlsx
    has_header: whether to write headers as bold first row
    """
    import xlsxwriter

    # Disable automatic URL conversion (fixes the 65,530 hyperlink warning)
    wb = xlsxwriter.Workbook(
        str(out_path),
        {'strings_to_urls': False}
    )

    ws = wb.add_worksheet()

    # Header format (bold)
    header_format = wb.add_format({'bold': True}) if has_header else None

    row_idx = 0

    if has_header and headers:
        for col_idx, val in enumerate(headers):
            ws.write(row_idx, col_idx, val, header_format)
        row_idx += 1

    for r in data_rows:
        for col_idx, val in enumerate(r):
            ws.write(row_idx, col_idx, val)
        row_idx += 1

    wb.close()

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("CSVtoExcel - Help")
        self.resize(540, 510)
        self.setStyleSheet("""
            QWidget { background-color: #2e2e2e; color: #f0f0f0; font-family: Arial; }
            QTextEdit { background-color: #1e1e1e; border:1px solid #555; border-radius:8px; padding:10px; color:#f0f0f0;}
            h2 { color:#00ccff; }
            h3 { color:#66ccff; }
        """)
        layout = QVBoxLayout(self)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setHtml("""
        <h2>📊 CSVtoExcel - User Guide</h2>
        <p>This tool allows you to quickly convert <b>CSV files</b> to <b>Excel (.xlsx)</b> format with a modern GUI.</p>
        <h3>🛠 Features:</h3>
        <ul>
        <li>📂 Browse and select individual CSV files or entire folders.</li>
        <li>👁 Preview first 100 rows of your CSV before conversion.</li>
        <li>⚙ Choose whether your CSV has a header row.</li>
        <li>📝 Specify encoding (UTF-8, Latin-1, UTF-16, CP1252).</li>
        <li>💾 Convert CSV → Excel with optional destination folder.</li>
        <li>🕑 Keep a conversion history stored in SQLite (source, Excel path, timestamp, rows, cols).</li>
        <li>🔄 Re-export previous conversions with original header settings.</li>
        <li>🗑 Delete history records without removing files.</li>
        <li>🎨 Responsive GUI with alternating row colors and hover highlight.</li>
        </ul>
        <h3>📖 How to use:</h3>
        <ul>
        <li>Click <b>Open CSV</b> or <b>Open Folder</b> to select files.</li>
        <li>Preview the first rows and adjust header or encoding if needed.</li>
        <li>Choose a destination folder or file name for Excel output.</li>
        <li>Click <b>Convert → Excel</b> to start the conversion.</li>
        <li>Use <b>History</b> to re-export, open folders, or delete records.</li>
        <li>You can also drag & drop CSV files or folders into the app window.</li>
        </ul>
        <h3>🏢 About:</h3>
        <ul>
        <li>This tool is built by MateTools.</li>
        <li>Visit our website: https://matetools.gumroad.com</li>
        </ul>
        """)
        layout.addWidget(text_edit)


def get_conn():
    conn = sqlite3.connect(str(DB_PATH))
    init_db(conn)
    return conn

def record_history(src: str, dest: str, rows: int, cols: int, has_header: int):
    try:
        conn = get_conn()
        c = conn.cursor()
        c.execute(
            'INSERT INTO history (src_path, dest_path, timestamp, rows, cols, has_header) VALUES (?,?,?,?,?,?)',
            (src, dest, datetime.utcnow().isoformat(), rows, cols, has_header)
        )
        conn.commit()
        conn.close()
    except Exception:
        print('Failed to write history:', traceback.format_exc())

def fetch_history(limit: int = 200):
    conn = get_conn()
    c = conn.cursor()
    # Include has_header in the SELECT
    c.execute('SELECT id, src_path, dest_path, timestamp, rows, cols, has_header FROM history ORDER BY id DESC LIMIT ?', (limit,))
    rows = c.fetchall()
    conn.close()
    return rows

def delete_history_item(item_id: int):
    conn = get_conn()
    c = conn.cursor()
    c.execute('DELETE FROM history WHERE id = ?', (item_id,))
    conn.commit()
    conn.close()

# ------------------------- UI Components -------------------------
class StyledButton(QtWidgets.QPushButton):
    def __init__(self, text):
        super().__init__(text)
        self.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))

class CsvPreviewModel(QtCore.QAbstractTableModel):
    def __init__(self, headers=None, data=None):
        super().__init__()
        self._headers = headers or []
        self._data = data or []

    def setDataFrame(self, headers, data):
        """headers: list, data: list of lists"""
        self.beginResetModel()
        self._headers = headers or []
        self._data = data or []
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return max(0, len(self._headers))

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == QtCore.Qt.DisplayRole:
            try:
                value = self._data[index.row()][index.column()]
                return "" if value is None else str(value)
            except Exception:
                return ""
        return None

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            try:
                return str(self._headers[section])
            except Exception:
                return None
        else:
            return str(section + 1)

class HistoryTableModel(QtCore.QAbstractTableModel):
    HEADERS = ["ID", "Source CSV", "Excel", "Timestamp (UTC)", "Rows", "Cols", "Has Header"]

    def __init__(self, rows=None):
        super().__init__()
        self._rows = rows or []

    def setRows(self, rows):
        self.beginResetModel()
        self._rows = rows
        self.endResetModel()

    def rowCount(self, parent=None):
        return len(self._rows)

    def columnCount(self, parent=None):
        return len(self.HEADERS)

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == QtCore.Qt.DisplayRole:
            r = self._rows[index.row()]
            col = index.column()
            # map has_header (last column) to Yes/No
            if col == 6:  # Has Header column
                try:
                    return "Yes" if int(r[6]) else "No"
                except Exception:
                    return str(r[col])
            try:
                return str(r[col])
            except Exception:
                return ""
        return None

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole and orientation == QtCore.Qt.Horizontal:
            return self.HEADERS[section]
        return None

# ----------------- Worker signals and QRunnable -----------------
class WorkerSignals(QtCore.QObject):
    progress = QtCore.Signal(int)            # file index / overall progress
    file_processed = QtCore.Signal(str)      # filename processed (for log)
    status = QtCore.Signal(str)              # status text
    finished = QtCore.Signal(list)           # errors list (empty if none)

class ConversionWorker(QtCore.QRunnable):
    """
    QRunnable that converts a list of CSVs to Excel using your helper save_to_excel_xlsxwriter.
    Emits progress/status/finished signals via WorkerSignals.
    """
    def __init__(self, csv_paths, dest_text, has_header, encoding):
        super().__init__()
        self.csv_paths = list(csv_paths)
        self.dest_text = dest_text
        self.has_header = has_header
        self.encoding = encoding
        self.signals = WorkerSignals()

    def run(self):
        import csv
        errors = []
        total = len(self.csv_paths)
        for i, csv_path in enumerate(self.csv_paths, start=1):
            try:
                # Read CSV in streaming fashion to avoid OOM for big files
                rows = []
                with open(csv_path, newline='', encoding=self.encoding, errors='replace') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        rows.append(row)
                if not rows:
                    raise ValueError("CSV file is empty")
                if self.has_header:
                    headers = rows[0]
                    data = rows[1:]
                else:
                    headers = [f"Column {j+1}" for j in range(len(rows[0]))]
                    data = rows

                # determine destination folder
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                if self.dest_text:
                    dest_folder = Path(self.dest_text)
                    if dest_folder.suffix.lower() == '.xlsx':
                        dest_folder = dest_folder.parent
                    dest_folder.mkdir(parents=True, exist_ok=True)
                else:
                    dest_folder = Path(csv_path).parent

                dest_file = dest_folder / f"{Path(csv_path).stem}_{timestamp}.xlsx"

                save_to_excel_xlsxwriter(headers, data, dest_file, has_header=self.has_header)

                # record history (use UTC ISO)
                record_history(str(csv_path), str(dest_file), len(data), len(headers), 1 if self.has_header else 0)

                self.signals.file_processed.emit(Path(csv_path).name)

            except Exception as e:
                errors.append(f"{Path(csv_path).name}: {str(e)}")

            # update progress
            self.signals.progress.emit(i)
            self.signals.status.emit(f"Processing {i}/{total}: {Path(csv_path).name}")

        # finished
        self.signals.finished.emit(errors)

# ------------------------- Main Window -------------------------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon(resource_path("logo.ico")))
        self.setWindowTitle(APP_NAME)
        self.resize(1000, 650)
        self._csv_path = None
   
        # preview storage (pandas-free)
        self._preview_headers = []
        self._preview_data = []

        self._csv_paths = []  
        self.threadpool = QtCore.QThreadPool()

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QVBoxLayout(central)

        # --- Toolbar-style top buttons ---
        toolbar_layout = QtWidgets.QHBoxLayout()
        toolbar_layout.setSpacing(10)

        self.btn_browse = StyledButton(" Open CSV")
        self.btn_browse.setIcon(QIcon.fromTheme("document-open"))
        self.btn_browse.clicked.connect(self.browse_csv)
        toolbar_layout.addWidget(self.btn_browse)

        self.btn_browse_folder = StyledButton(" Open Folder")
        self.btn_browse_folder.setIcon(QIcon.fromTheme("folder"))
        self.btn_browse_folder.clicked.connect(self.browse_folder)
        toolbar_layout.addWidget(self.btn_browse_folder)

        self.lbl_selected = QtWidgets.QLabel("No CSV selected")
        self.lbl_selected.setStyleSheet("font-weight: bold;")
        toolbar_layout.addWidget(self.lbl_selected, 1)  # stretch

        self.btn_convert = StyledButton(" Convert → Excel")
        self.btn_convert.setIcon(QIcon.fromTheme("document-save"))
        self.btn_convert.clicked.connect(self.convert_csv_to_excel)
        self.btn_convert.setEnabled(True)
        toolbar_layout.addWidget(self.btn_convert)

        self.btn_history = StyledButton(" History")
        self.btn_history.setIcon(QIcon.fromTheme("view-history"))
        self.btn_history.clicked.connect(self.open_history)
        toolbar_layout.addWidget(self.btn_history)

        self.btn_help = StyledButton(" Help")
        self.btn_help.setIcon(QIcon.fromTheme("help-browser"))
        self.btn_help.clicked.connect(self.open_help)
        toolbar_layout.addWidget(self.btn_help)

        main_layout.addLayout(toolbar_layout)

        # --- CSV Preview Table ---
        main_layout.addWidget(QtWidgets.QLabel('CSV Preview (first 100 rows)'))
        self.preview_table = QtWidgets.QTableView()
        self.preview_model = CsvPreviewModel()
        self.preview_table.setModel(self.preview_model)
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        self.preview_table.setSortingEnabled(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        main_layout.addWidget(self.preview_table)

        # --- Controls below preview (responsive) ---
        # Row 1: Preview rows, header, encoding
        row1_layout = QtWidgets.QHBoxLayout()
        row1_layout.setSpacing(15)
        self.spin_preview = QtWidgets.QSpinBox()
        self.spin_preview.setRange(5, 10000)
        self.spin_preview.setValue(100)
        row1_layout.addWidget(QtWidgets.QLabel("Preview rows"))
        row1_layout.addWidget(self.spin_preview)
        self.chk_header = QtWidgets.QCheckBox("Has header row")
        self.chk_header.setChecked(True)
        row1_layout.addWidget(self.chk_header)
        self.combo_encoding = QtWidgets.QComboBox()
        self.combo_encoding.addItems(['utf-8', 'latin-1', 'utf-16', 'cp1252'])
        row1_layout.addWidget(QtWidgets.QLabel("Encoding"))
        row1_layout.addWidget(self.combo_encoding)
        row1_layout.addStretch()
        main_layout.addLayout(row1_layout)

        # Row 2: Destination, browse, status
        row2_layout = QtWidgets.QHBoxLayout()
        row2_layout.setSpacing(10)
        row2_layout.addWidget(QtWidgets.QLabel("Destination"))
        self.line_dest = QtWidgets.QLineEdit()
        self.line_dest.setPlaceholderText('Optional: destination folder or file')
        self.line_dest.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        row2_layout.addWidget(self.line_dest)
        self.btn_dest_browse = StyledButton(' Browse')
        self.btn_dest_browse.setIcon(QIcon.fromTheme("folder"))
        self.btn_dest_browse.clicked.connect(self.browse_dest)
        row2_layout.addWidget(self.btn_dest_browse)
        row2_layout.addWidget(QtWidgets.QLabel("Status"))
        self.status_label = QtWidgets.QLabel('Ready')
        self.status_label.setMinimumWidth(120)
        row2_layout.addWidget(self.status_label)
        main_layout.addLayout(row2_layout)

        # Footer: progress and log
        footer = QtWidgets.QHBoxLayout()
        self.progress = QtWidgets.QProgressBar()
        self.progress.setVisible(False)
        self.progress.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        footer.addWidget(self.progress)
        self.log_label = QtWidgets.QLabel('')
        footer.addWidget(self.log_label)
        main_layout.addLayout(footer)

        # Drag & Drop
        self.setAcceptDrops(True)
        self.apply_styles()

        self.spin_preview.valueChanged.connect(self.load_preview)

    # ---------------- Styles ----------------
    def apply_styles(self):
        qss = """
        QWidget { font-family: Inter, 'Segoe UI', Arial; font-size: 12px; }
        QMainWindow { background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f7f9fc, stop:1 #eef2f7); }
        QPushButton { background: #1976d2; color: white; border-radius: 8px; padding: 6px 12px; }
        QPushButton:hover { background: #165ea9; }
        QLineEdit, QComboBox, QSpinBox { background: white; padding: 6px; border: 1px solid #cfd8e3; border-radius: 6px; }
        QTableView { background: #ffffff; alternate-background-color: #f5f7fa; gridline-color: #e0e0e0; selection-background-color: #cce4ff; selection-color: #000000; show-decoration-selected: 1; }
        QTableView::item:hover { background: #e6f0ff; }
        QHeaderView::section { background: #f1f5f9; padding: 6px; border: 1px solid #dfe3e8; font-weight: bold; }
        QProgressBar { border-radius: 8px; height: 14px; text-align: center; }
        QLabel { font-weight: normal; }
        """
        self.setStyleSheet(qss)
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setShowGrid(True)
        self.preview_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.preview_table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.preview_table.verticalHeader().setVisible(False)

    # ---------------- Drag & Drop ----------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        csv_files = []
        invalid_files = []

        for url in urls:
            path = Path(url.toLocalFile())
            if path.is_dir():
                # Add all CSV files in directory recursively
                csv_files.extend(path.rglob("*.csv"))
            elif path.is_file():
                if path.suffix.lower() == ".csv":
                    csv_files.append(path)
                else:
                    invalid_files.append(path.name)

        # Remove duplicates and sort
        csv_files = sorted(set(csv_files))

        if invalid_files:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid files",
                f"The following files were ignored because they are not CSVs:\n" + "\n".join(invalid_files)
            )

        self.set_csv_files(csv_files)

    # ---------------- File Handling ----------------
    def browse_csv(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            "Open CSV file(s)",
            str(Path.home()),
            "CSV Files (*.csv);;All Files (*)"
        )
        if not files:
            return

        csv_files = []
        invalid_files = []
        for f in files:
            path = Path(f)
            if path.suffix.lower() == ".csv":
                csv_files.append(path)
            else:
                invalid_files.append(path.name)

        if invalid_files:
            QtWidgets.QMessageBox.warning(
                self,
                "Invalid files",
                f"The following files were ignored because they are not CSVs:\n" + "\n".join(invalid_files)
            )

        self.set_csv_files(csv_files)


    def browse_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Folder with CSVs", str(Path.home()))
        if not folder:
            return

        folder_path = Path(folder)
        csv_files = list(folder_path.rglob("*.csv"))

        if not csv_files:
            QtWidgets.QMessageBox.warning(self, "No CSV files", "No CSV files were found in the selected folder.")
            return

        self.set_csv_files(csv_files)

    def set_csv_files(self, files: list[Path]):
        files = sorted(set(files))
        if not files:
            self._csv_paths = []
            self._csv_path = None
            self.lbl_selected.setText("No CSV selected")
            self.btn_convert.setEnabled(False)
            return
        self._csv_paths = files
        self._csv_path = files[0]
        self.update_selected_label()
        self.load_preview()

    def update_selected_label(self):
        total_files = len(self._csv_paths)
        if total_files == 0:
            self.lbl_selected.setText("No CSV selected")
        elif total_files == 1:
            self.lbl_selected.setText(self._csv_paths[0].name)
        else:
            preview_names = [f.name for f in self._csv_paths[:3]]
            more = total_files - 3
            self.lbl_selected.setText(
                f"{total_files} files selected: {', '.join(preview_names)}"
                + (f" ... (+{more} more)" if more > 0 else "")
            )

    def load_preview(self):
        if not self._csv_path:
            return
        try:
            rows_to_preview = self.spin_preview.value()
            has_header = self.chk_header.isChecked()
            enc = self.combo_encoding.currentText()

            # Read first N rows using csv
            headers, data = self.preview_csv(self._csv_path, rows_to_preview, has_header, enc)

            # Save into preview storage
            self._preview_headers = headers
            self._preview_data = data

            # Update preview model
            self.preview_model.setDataFrame(headers, data)

            self.btn_convert.setEnabled(True)
            self.status_label.setText(f'Loaded preview: {len(data)} × {len(headers)}')

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Failed', f"Failed to load {self._csv_path}:\n{str(e)}")
            self.btn_convert.setEnabled(False)

    def preview_csv(self, csv_path, n=100, has_header=True, encoding='utf-8'):
        import csv
        rows = []
        # open with 'errors="replace"' to be robust to encoding issues
        with open(csv_path, newline='', encoding=encoding, errors='replace') as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                if i >= n:
                    break
                rows.append(row)

        if not rows:
            return [], []

        if has_header:
            headers = [h if h is not None else "" for h in rows[0]]
            data = rows[1:]
        else:
            headers = [f"Column {i+1}" for i in range(len(rows[0]))]
            data = rows
        return headers, data

    def browse_dest(self):
        if len(self._csv_paths) > 1:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select destination folder", str(Path.home()))
            if folder:
                self.line_dest.setText(folder)
        else:
            default = self._csv_path.with_suffix('.xlsx') if self._csv_path else Path.home() / 'output.xlsx'
            fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Choose destination", str(default), "Excel (*.xlsx)")
            if fn and not fn.lower().endswith('.xlsx'):
                fn += ".xlsx"
            self.line_dest.setText(fn)

    def convert_csv_to_excel(self):
        files = getattr(self, "_csv_paths", None)
        if not files or not any(files):
            self.status_label.setText("⚠ No CSV selected! Please select at least one CSV.")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            return
        else:
            self.status_label.setStyleSheet("color: black; font-weight: normal;")

        dest_text = self.line_dest.text().strip()
        total_files = len(files)

        # Prepare UI progress
        self.progress.setVisible(True)
        self.progress.setRange(0, total_files)
        self.progress.setValue(0)
        self.progress.repaint()
        QtWidgets.QApplication.processEvents()

        # Worker
        has_header = self.chk_header.isChecked()
        enc = self.combo_encoding.currentText()
        worker = ConversionWorker(files, dest_text, has_header, enc)

        # wiring the signals
        def on_progress(val):
            self.progress.setValue(val)

        def on_file_processed(name):
            self.log_label.setText(f"Processed: {name}")

        def on_status(text):
            self.status_label.setText(text)

        def on_finished(errors):
            self.progress.setVisible(False)
            if errors:
                QtWidgets.QMessageBox.warning(self, "Some conversions failed", "\n".join(errors))
            else:
                QtWidgets.QMessageBox.information(self, "Success", f"Converted {total_files} CSV file(s) successfully.")
            self.status_label.setText("Batch conversion complete")
            # update preview/log last file
            last_name = files[-1].name if files else ""
            self.log_label.setText(f"Last: {last_name}")

        worker.signals.progress.connect(on_progress)
        worker.signals.file_processed.connect(on_file_processed)
        worker.signals.status.connect(on_status)
        worker.signals.finished.connect(on_finished)

        # run in threadpool
        self.threadpool.start(worker)

    # ---------------- History ----------------
    def open_history(self):
        dlg = HistoryDialog(self)
        dlg.exec()

    def open_help(self):
        dlg = HelpDialog(self)
        dlg.exec()

# ------------------------- History Dialog -------------------------
class HistoryDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Conversion History')
        self.resize(900, 450)
        layout = QtWidgets.QVBoxLayout(self)

        self.table = QtWidgets.QTableView()
        self.model = HistoryTableModel([])
        self.table.setModel(self.model)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        layout.addWidget(self.table)

        btn_layout = QtWidgets.QHBoxLayout()
        self.btn_refresh = StyledButton('Refresh')
        self.btn_refresh.clicked.connect(self.load_history)
        btn_layout.addWidget(self.btn_refresh)

        self.btn_reopen = StyledButton('Open Folder')
        self.btn_reopen.clicked.connect(self.open_folder_for_selected)
        btn_layout.addWidget(self.btn_reopen)

        self.btn_reexport = StyledButton('Re-export')
        self.btn_reexport.clicked.connect(self.reexport_selected)
        btn_layout.addWidget(self.btn_reexport)

        self.btn_delete = StyledButton('Delete record')
        self.btn_delete.clicked.connect(self.delete_selected)
        btn_layout.addWidget(self.btn_delete)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.load_history()

    def load_history(self):
        rows = fetch_history(500)
        self.model.setRows(rows)

    def selected_row(self):
        idx = self.table.currentIndex()
        if not idx.isValid():
            return None
        return self.model._rows[idx.row()]

    def open_folder_for_selected(self):
        r = self.selected_row()
        if not r:
            QtWidgets.QMessageBox.warning(self, 'No selection', 'Select a history row first')
            return
        dest_path = Path(r[2])
        if not dest_path.exists():
            QtWidgets.QMessageBox.warning(self, 'Missing file', 'The exported Excel file does not exist anymore')
            return
        folder = dest_path.parent
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(folder)))

    def reexport_selected(self):
        import csv

        r = self.selected_row()
        if not r:
            QtWidgets.QMessageBox.warning(self, 'No selection', 'Select a history row first')
            return

        src = Path(r[1])
        if not src.exists():
            QtWidgets.QMessageBox.warning(self, 'Missing CSV', 'The original CSV is missing')
            return

        # Destination file
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Re-export Excel to...",
            str(src.with_suffix('.reexport.xlsx')),
            "Excel Workbook (*.xlsx)"
        )
        if not fn:
            return

        try:
            # --- Get stored has_header from history ---
            has_header = bool(int(r[6])) if len(r) > 6 else True

            # --- Read CSV ---
            with open(src, newline='', encoding='utf-8', errors='replace') as f:
                reader = csv.reader(f)
                all_rows = list(reader)

            if not all_rows:
                raise ValueError("CSV is empty")

            if has_header:
                headers = all_rows[0]
                data = all_rows[1:]
            else:
                headers = [f"Column {i+1}" for i in range(len(all_rows[0]))]
                data = all_rows

            # --- Save to Excel using helper ---
            save_to_excel_xlsxwriter(headers, data, fn, has_header=has_header)

            QtWidgets.QMessageBox.information(self, 'Re-exported', f'Re-exported to {fn}')

            # --- Record re-export in history ---
            record_history(str(src), fn, len(data), len(headers), 1 if has_header else 0)

            self.load_history()

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, 'Failed', f'Error: {e}')

    def delete_selected(self):
        r = self.selected_row()
        if not r:
            QtWidgets.QMessageBox.warning(self, 'No selection', 'Select a history row first')
            return
        item_id = int(r[0])
        ok = QtWidgets.QMessageBox.question(self, 'Confirm', 'Delete this history record? (this does NOT remove files)')
        if ok == QtWidgets.QMessageBox.StandardButton.Yes:
            delete_history_item(item_id)
            self.load_history()

# ------------------------- Entrypoint -------------------------
def main():
    app = QtWidgets.QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()

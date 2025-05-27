import sys
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QFileDialog, QColorDialog, QMessageBox, QMenuBar, QMenu, QComboBox
)
from PyQt6.QtGui import QColor, QFont, QAction
from PyQt6.QtCore import Qt
import pyqtgraph as pg


class GraphWindow(QWidget):
    def __init__(self, x, y):
        super().__init__()
        self.setWindowTitle("Grafik")
        self.resize(600, 400)
        layout = QVBoxLayout()
        self.chartWidget = pg.PlotWidget()
        self.chartWidget.setBackground("#2b2b2b")
        layout.addWidget(self.chartWidget)
        self.setLayout(layout)
        self.chartWidget.plot(x, y, pen="b", symbol="o")


class ExcelLikeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SoudTablo v1.0.3")
        self.resize(1200, 700)

        self.row_count = 100
        self.col_count = 100
        self.formulas = {}
        self.undo_stack = []
        self.redo_stack = []

        self.table = QTableWidget(self.row_count, self.col_count)
        self.table.setFont(QFont("Arial", 10))
        self.table.cellChanged.connect(self.handle_cell_change)
        self.table.cellDoubleClicked.connect(self.uppercase_cell)
        self.table.itemChanged.connect(self.save_undo_state)

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.create_menu()
        self.apply_dark_theme()

        self.selected_currency = "₺"

    def apply_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: white;
            }
            QTableWidget {
                background-color: #3c3f41;
                color: white;
                gridline-color: #555;
            }
            QHeaderView::section {
                background-color: #3c3f41;
                color: white;
            }
            QMenuBar {
                background-color: #2b2b2b;
                color: white;
            }
            QMenu {
                background-color: #3c3f41;
                color: white;
            }
        """)

    def create_menu(self):
        menu = self.menuBar()

        file_menu = menu.addMenu("Dosya")
        save_action = QAction("Kaydet", self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        load_action = QAction("Yükle", self)
        load_action.triggered.connect(self.load_file)
        file_menu.addAction(load_action)

        file_menu.addSeparator()
        undo_action = QAction("Geri Al", self)
        undo_action.triggered.connect(self.undo)
        file_menu.addAction(undo_action)

        redo_action = QAction("Yinele", self)
        redo_action.triggered.connect(self.redo)
        file_menu.addAction(redo_action)

        add_rows_action = QAction("Satır Ekle", self)
        add_rows_action.triggered.connect(self.add_rows)
        file_menu.addAction(add_rows_action)

        add_cols_action = QAction("Sütun Ekle", self)
        add_cols_action.triggered.connect(self.add_columns)
        file_menu.addAction(add_cols_action)

        format_menu = menu.addMenu("Biçim")
        currency_menu = format_menu.addMenu("Para Birimi Seç")

        currencies = ["₺", "$", "€", "£", "¥"]
        for cur in currencies:
            act = QAction(cur, self)
            act.triggered.connect(lambda checked, c=cur: self.set_selected_currency(c))
            currency_menu.addAction(act)

        color_action = QAction("Hücre Rengini Ayarla", self)
        color_action.triggered.connect(self.set_cell_color)
        format_menu.addAction(color_action)

        align_menu = format_menu.addMenu("Hizalama")
        alignments = {
            "Sola Hizala": Qt.AlignmentFlag.AlignLeft,
            "Ortala": Qt.AlignmentFlag.AlignCenter,
            "Sağa Hizala": Qt.AlignmentFlag.AlignRight
        }
        for name, align in alignments.items():
            act = QAction(name, self)
            act.triggered.connect(lambda checked, a=align: self.align_cell(a))
            align_menu.addAction(act)

        condition_action = QAction("Koşullu Biçim (>100 = Yeşil)", self)
        condition_action.triggered.connect(self.apply_conditional_formatting)
        format_menu.addAction(condition_action)

        chart_menu = menu.addMenu("Grafik")
        plot_action = QAction("Grafik Çiz", self)
        plot_action.triggered.connect(self.plot_graph)
        chart_menu.addAction(plot_action)

    def set_selected_currency(self, currency):
        self.selected_currency = currency
        QMessageBox.information(self, "Para Birimi Seçildi", f"Para birimi: {currency}")

    def handle_cell_change(self, row, col):
        item = self.table.item(row, col)
        if not item:
            return
        text = item.text()
        if text.startswith("="):
            self.formulas[(row, col)] = text
            value = self.evaluate_formula(text[1:])
            if value is not None:
                if isinstance(value, (int, float)):
                    try:
                        value_float = float(value)
                        item.setText(f"{value_float:,.2f} {self.selected_currency}")
                    except:
                        item.setText(str(value))
                else:
                    item.setText(str(value))
        elif (row, col) in self.formulas:
            del self.formulas[(row, col)]

    def evaluate_formula(self, formula):
        try:
            refs = re.findall(r"[A-Z]+\d+", formula)
            for ref in refs:
                col = 0
                for i, ch in enumerate(ref):
                    if ch.isdigit():
                        col_str = ref[:i]
                        row_str = ref[i:]
                        break
                else:
                    col_str = ref
                    row_str = ''
                col = 0
                for idx, c in enumerate(reversed(col_str)):
                    col += (ord(c) - ord('A') + 1) * (26 ** idx)
                col -= 1
                row = int(row_str) - 1 if row_str else 0
                ref_item = self.table.item(row, col)
                val = ref_item.text() if ref_item else "0"
                try:
                    val_float = float(val.replace(',', '').replace(self.selected_currency, '').strip())
                except:
                    val_float = 0
                formula = formula.replace(ref, str(val_float))
            return eval(formula)
        except:
            return None

    def save_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Kaydet", "", "ExcelTablo (*.etab)")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    f.write(",".join(row_data) + "\n")

    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Yükle", "", "ExcelTablo (*.etab)")
        if path:
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                self.table.setRowCount(len(lines))
                max_col = 0
                for row, line in enumerate(lines):
                    values = line.strip().split(",")
                    max_col = max(max_col, len(values))
                    for col, val in enumerate(values):
                        self.table.setItem(row, col, QTableWidgetItem(val))
                self.table.setColumnCount(max_col)

    def set_cell_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            row, col = self.table.currentRow(), self.table.currentColumn()
            item = self.table.item(row, col)
            if not item:
                item = QTableWidgetItem()
                self.table.setItem(row, col, item)
            item.setBackground(color)

    def align_cell(self, alignment):
        row, col = self.table.currentRow(), self.table.currentColumn()
        item = self.table.item(row, col)
        if item:
            item.setTextAlignment(alignment)

    def apply_conditional_formatting(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    try:
                        val = float(item.text().replace(',', '').replace(self.selected_currency, '').strip())
                        if val > 100:
                            item.setBackground(QColor("green"))
                    except:
                        continue

    def plot_graph(self):
        x = []
        y = []
        for row in range(self.table.rowCount()):
            x_item = self.table.item(row, 0)
            y_item = self.table.item(row, 1)
            if x_item and y_item:
                try:
                    x_val = float(x_item.text().replace(',', '').replace(self.selected_currency, '').strip())
                    y_val = float(y_item.text().replace(',', '').replace(self.selected_currency, '').strip())
                    x.append(x_val)
                    y.append(y_val)
                except:
                    continue
        if x and y:
            self.graph_window = GraphWindow(x, y)
            self.graph_window.show()
        else:
            QMessageBox.information(self, "Bilgi", "Grafik çizmek için yeterli veri yok.")

    def uppercase_cell(self, row, col):
        item = self.table.item(row, col)
        if item:
            item.setText(item.text().upper())

    def add_rows(self):
        current_rows = self.table.rowCount()
        self.table.setRowCount(current_rows + 10)

    def add_columns(self):
        current_cols = self.table.columnCount()
        self.table.setColumnCount(current_cols + 10)

    def save_undo_state(self):
        state = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            state.append(row_data)
        self.undo_stack.append(state)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def restore_state(self, state):
        self.table.blockSignals(True)
        self.table.setRowCount(len(state))
        self.table.setColumnCount(len(state[0]) if state else 0)
        for row, row_data in enumerate(state):
            for col, val in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(val))
        self.table.blockSignals(False)

    def undo(self):
        if self.undo_stack:
            state = self.undo_stack.pop()
            self.redo_stack.append(state)
            if self.undo_stack:
                self.restore_state(self.undo_stack[-1])

    def redo(self):
        if self.redo_stack:
            state = self.redo_stack.pop()
            self.restore_state(state)
            self.undo_stack.append(state)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelLikeApp()
    window.show()
    sys.exit(app.exec())

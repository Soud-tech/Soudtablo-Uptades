import sys
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QFileDialog, QColorDialog, QMessageBox, QMenuBar, QMenu
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


class UndoRedoStack:
    def __init__(self):
        self.undo_stack = []
        self.redo_stack = []

    def push(self, row, col, old_text, new_text):
        self.undo_stack.append((row, col, old_text, new_text))
        self.redo_stack.clear()

    def undo(self):
        if not self.undo_stack:
            return None
        change = self.undo_stack.pop()
        self.redo_stack.append(change)
        return change

    def redo(self):
        if not self.redo_stack:
            return None
        change = self.redo_stack.pop()
        self.undo_stack.append(change)
        return change


class ExcelLikeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SoudTablo v1.0.3")
        self.resize(1200, 700)

        self.row_count = 100
        self.col_count = 100

        self.table = QTableWidget(self.row_count, self.col_count)
        self.table.setFont(QFont("Arial", 10))
        self.table.cellChanged.connect(self.handle_cell_change)
        self.table.cellDoubleClicked.connect(self.uppercase_cell)

        self.formulas = {}
        self.ignore_cell_change = False  # Döngü önlemek için flag

        self.undo_redo = UndoRedoStack()
        self.is_undo_redo_action = False  # undo/redo esnasında cellChanged çalışmasın diye

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

        add_rows_action = QAction("Satır Ekle", self)
        add_rows_action.triggered.connect(self.add_rows)
        file_menu.addAction(add_rows_action)

        add_cols_action = QAction("Sütun Ekle", self)
        add_cols_action.triggered.connect(self.add_columns)
        file_menu.addAction(add_cols_action)

        edit_menu = menu.addMenu("Düzen")
        undo_action = QAction("Geri Al (Undo)", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.undo)
        edit_menu.addAction(undo_action)

        redo_action = QAction("İleri Al (Redo)", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.redo)
        edit_menu.addAction(redo_action)

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

        chart_menu = menu.addMenu("Grafik")
        plot_action = QAction("Grafik Çiz", self)
        plot_action.triggered.connect(self.plot_graph)
        chart_menu.addAction(plot_action)

    def set_selected_currency(self, currency):
        self.selected_currency = currency
        self.reformat_all_currency_cells()
        QMessageBox.information(self, "Para Birimi Seçildi", f"Para birimi: {currency}")

    def reformat_all_currency_cells(self):
        self.ignore_cell_change = True
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item and item.text():
                    text = item.text()
                    # Sadece sayı içeren hücrelere uygula
                    try:
                        val_float = float(text.replace(',', '').replace(self.selected_currency, '').strip())
                        item.setText(f"{val_float:,.2f} {self.selected_currency}")
                        self.apply_conditional_formatting(row, col, val_float)
                    except:
                        pass
        self.ignore_cell_change = False

    def handle_cell_change(self, row, col):
        if self.ignore_cell_change:
            return
        if self.is_undo_redo_action:
            return

        item = self.table.item(row, col)
        if not item:
            return

        old_text = getattr(item, "_old_text", "")
        new_text = item.text()

        # Undo stack'a eski-yeni değerleri ekle
        self.undo_redo.push(row, col, old_text, new_text)
        item._old_text = new_text  # eski değer olarak kaydet

        text = new_text
        if text.startswith("="):
            self.formulas[(row, col)] = text
            value = self.evaluate_formula(text[1:])
            if value is not None:
                if isinstance(value, (int, float)):
                    try:
                        value_float = float(value)
                        item.setText(f"{value_float:,.2f} {self.selected_currency}")
                        self.apply_conditional_formatting(row, col, value_float)
                    except:
                        item.setText(str(value))
                else:
                    item.setText(str(value))
        elif (row, col) in self.formulas:
            del self.formulas[(row, col)]
            self.apply_conditional_formatting(row, col, None)
        else:
            # Eğer formül yoksa, sayısal değer varsa koşullu renklendirme uygula
            try:
                val_float = float(text.replace(',', '').replace(self.selected_currency, '').strip())
                self.apply_conditional_formatting(row, col, val_float)
            except:
                self.apply_conditional_formatting(row, col, None)

    def apply_conditional_formatting(self, row, col, val):
        item = self.table.item(row, col)
        if not item:
            return
        # Örnek koşul: değer 100'den büyükse arka plan yeşil, değilse normal
        if val is not None and val > 100:
            item.setBackground(QColor(0, 150, 0, 150))  # yarı saydam yeşil
        else:
            item.setBackground(QColor(60, 63, 65))  # koyu gri (uyumlu tema arka plan)

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
                self.ignore_cell_change = True
                self.table.setRowCount(len(lines))
                max_col = 0
                for row, line in enumerate(lines):
                    values = line.strip().split(",")
                    max_col = max(max_col, len(values))
                    for col, val in enumerate(values):
                        self.table.setItem(row, col, QTableWidgetItem(val))
                        # Eski değer kaydet (undo için)
                        item = self.table.item(row, col)
                        if item:
                            item._old_text = val
self.table.setColumnCount(max_col)
self.ignore_cell_change = False 
def set_cell_color(self):
    color = QColorDialog.getColor()
    if color.isValid():
        for item in self.table.selectedItems():
            item.setBackground(color)

def add_rows(self):
    self.table.insertRow(self.table.rowCount())

def add_columns(self):
    self.table.insertColumn(self.table.columnCount())

def uppercase_cell(self, row, col):
    item = self.table.item(row, col)
    if item:
        item.setText(item.text().upper())

def plot_graph(self):
    x, y = [], []
    for item in self.table.selectedItems():
        try:
            val = float(item.text().replace(',', '').replace(self.selected_currency, '').strip())
            y.append(val)
            x.append(len(x) + 1)
        except:
            continue
    if x and y:
        self.graph = GraphWindow(x, y)
        self.graph.show()

def undo(self):
    change = self.undo_redo.undo()
    if change:
        row, col, old_text, new_text = change
        self.is_undo_redo_action = True
        self.table.setItem(row, col, QTableWidgetItem(old_text))
        self.is_undo_redo_action = False

def redo(self):
    change = self.undo_redo.redo()
    if change:
        row, col, old_text, new_text = change
        self.is_undo_redo_action = True
        self.table.setItem(row, col, QTableWidgetItem(new_text))
        self.is_undo_redo_action = False
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelLikeApp()
    window.show()
    sys.exit(app.exec())
            

import sys
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QMenuBar, QFileDialog, QColorDialog, QInputDialog, QPushButton,
    QLabel, QHBoxLayout, QComboBox, QMessageBox
)
from PyQt6.QtGui import QColor, QFont
from PyQt6.QtCore import Qt
import pyqtgraph as pg

class ExcelLikeApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SoudTablo v1.0.1")
        self.resize(1000, 600)

        self.table = QTableWidget(20, 10)
        self.table.setFont(QFont("Arial", 12))
        self.table.cellChanged.connect(self.handle_cell_change)
        self.table.cellDoubleClicked.connect(self.uppercase_cell)

        self.formulas = {}

        self.chartWidget = pg.PlotWidget()
        self.chartWidget.setBackground("w")

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addWidget(self.chartWidget)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.create_menu()
        self.apply_dark_theme()

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
        """)

    def create_menu(self):
        menu = self.menuBar()
        file_menu = menu.addMenu("Dosya")

        save_action = file_menu.addAction("Kaydet")
        save_action.triggered.connect(self.save_file)

        load_action = file_menu.addAction("Yükle")
        load_action.triggered.connect(self.load_file)

        format_menu = menu.addMenu("Biçim")
        currency_action = format_menu.addAction("Para Birimi Olarak Biçimlendir")
        currency_action.triggered.connect(self.format_currency)

        color_action = format_menu.addAction("Hücre Rengini Ayarla")
        color_action.triggered.connect(self.set_cell_color)

        chart_menu = menu.addMenu("Grafik")
        plot_action = chart_menu.addAction("Grafik Çiz")
        plot_action.triggered.connect(self.plot_graph)

    def handle_cell_change(self, row, col):
        item = self.table.item(row, col)
        if not item:
            return
        text = item.text()
        if text.startswith("="):
            self.formulas[(row, col)] = text
            value = self.evaluate_formula(text[1:])
            if value is not None:
                item.setText(str(value))
        elif (row, col) in self.formulas:
            del self.formulas[(row, col)]

    def evaluate_formula(self, formula):
        try:
            refs = re.findall(r"[A-Z]+\d+", formula)
            for ref in refs:
                col = ord(ref[0]) - ord("A")
                row = int(ref[1:]) - 1
                ref_item = self.table.item(row, col)
                val = ref_item.text() if ref_item else "0"
                formula = formula.replace(ref, val)
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
                for row, line in enumerate(lines):
                    values = line.strip().split(",")
                    self.table.setColumnCount(max(self.table.columnCount(), len(values)))
                    for col, val in enumerate(values):
                        self.table.setItem(row, col, QTableWidgetItem(val))

    def format_currency(self):
        row, col = self.table.currentRow(), self.table.currentColumn()
        item = self.table.item(row, col)
        if item:
            try:
                number = float(item.text())
                item.setText(f"{number:,.2f} ₺")
            except ValueError:
                QMessageBox.warning(self, "Hata", "Bu hücrede geçerli bir sayı yok.")

    def set_cell_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            row, col = self.table.currentRow(), self.table.currentColumn()
            item = self.table.item(row, col)
            if not item:
                item = QTableWidgetItem()
                self.table.setItem(row, col, item)
            item.setBackground(color)

    def plot_graph(self):
        x = []
        y = []
        for row in range(self.table.rowCount()):
            x_item = self.table.item(row, 0)
            y_item = self.table.item(row, 1)
            if x_item and y_item:
                try:
                    x_val = float(x_item.text())
                    y_val = float(y_item.text())
                    x.append(x_val)
                    y.append(y_val)
                except:
                    continue
        if x and y:
            self.chartWidget.clear()
            self.chartWidget.plot(x, y, pen="b", symbol="o")

    def uppercase_cell(self, row, col):
        item = self.table.item(row, col)
        if item:
            item.setText(item.text().upper())

if __name__ == "__main__":
    print("Yeni sürüm yüklendi!")
    app = QApplication(sys.argv)
    window = ExcelLikeApp()
    window.show()
    sys.exit(app.exec())

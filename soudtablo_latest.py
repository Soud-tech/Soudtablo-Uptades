import sys
import locale
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QPushButton, QHBoxLayout, QMenu, QDialog, QLabel
)
from PyQt6.QtGui import QAction, QColor
from PyQt6.QtCore import Qt
import pyqtgraph as pg

print("Yeni sürüm yüklendi!")


locale.setlocale(locale.LC_ALL, '')

class ChartWindow(QDialog):
    def __init__(self, data):
        super().__init__()
        self.setWindowTitle("Grafik")
        self.resize(500, 400)
        layout = QVBoxLayout()
        self.plot_widget = pg.PlotWidget()
        self.plot_widget.plot(data, pen='c', symbol='o')
        layout.addWidget(self.plot_widget)
        self.setLayout(layout)

class ExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SoudTablo - Gelişmiş Sürüm")
        self.resize(1000, 600)
        self.apply_dark_theme()

        self.table = QTableWidget(10, 10)
        self.table.cellChanged.connect(self.evaluate_cell)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.open_context_menu)

        add_row_btn = QPushButton("Satır Ekle")
        add_row_btn.clicked.connect(self.add_row)

        add_col_btn = QPushButton("Sütun Ekle")
        add_col_btn.clicked.connect(self.add_column)

        chart_btn = QPushButton("Grafik Çiz")
        chart_btn.clicked.connect(self.plot_graph)

        layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        button_layout.addWidget(add_row_btn)
        button_layout.addWidget(add_col_btn)
        button_layout.addWidget(chart_btn)

        layout.addLayout(button_layout)
        layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def apply_dark_theme(self):
        dark_stylesheet = """
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QTableWidget {
                background-color: #3c3f41;
                color: #ffffff;
                gridline-color: #6c6c6c;
            }
            QPushButton {
                background-color: #4e4e4e;
                color: white;
                border: 1px solid #5a5a5a;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #646464;
            }
        """
        self.setStyleSheet(dark_stylesheet)

    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def add_column(self):
        self.table.insertColumn(self.table.columnCount())

    def evaluate_cell(self, row, col):
        item = self.table.item(row, col)
        if not item:
            return
        text = item.text()
        if text.startswith('='):
            try:
                result = self.evaluate_formula(text[1:])
                self.table.blockSignals(True)
                self.table.setItem(row, col, QTableWidgetItem(str(result)))
                self.table.blockSignals(False)
            except Exception:
                pass  # Hatalı formül
        else:
            self.apply_conditional_format(item)

    def apply_conditional_format(self, item):
        try:
            value = float(item.text().replace(',', '').replace('₺', '').replace('$', ''))
            if value > 100:
                item.setBackground(QColor('#4caf50'))  # yeşil
            elif value < 0:
                item.setBackground(QColor('#f44336'))  # kırmızı
            else:
                item.setBackground(QColor('#3c3f41'))  # nötr koyu
        except:
            item.setBackground(QColor('#3c3f41'))  # yazıysa koyu zemin

    def evaluate_formula(self, formula):
        formula = formula.upper()
        for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            for row_num in range(1, self.table.rowCount()+1):
                cell_id = f"{col_letter}{row_num}"
                col_idx = ord(col_letter) - 65
                row_idx = row_num - 1
                if 0 <= row_idx < self.table.rowCount() and 0 <= col_idx < self.table.columnCount():
                    item = self.table.item(row_idx, col_idx)
                    value = item.text() if item else '0'
                    formula = formula.replace(cell_id, value)
        return eval(formula)

    def open_context_menu(self, pos):
        item = self.table.itemAt(pos)
        if not item:
            return
        menu = QMenu()
        currency_action = QAction("Para Biçimine Dönüştür", self)
        currency_action.triggered.connect(lambda: self.format_currency(item))
        menu.addAction(currency_action)
        menu.exec(self.table.mapToGlobal(pos))

    def format_currency(self, item):
        try:
            value = float(item.text().replace(',', '').replace('₺', '').replace('$', ''))
            formatted = locale.currency(value, grouping=True)
            item.setText(formatted)
            self.apply_conditional_format(item)
        except:
            pass

    def plot_graph(self):
        data = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)  # İlk sütundaki veriyi çiziyoruz
            if item:
                try:
                    value = float(item.text().replace(',', '').replace('₺', '').replace('$', ''))
                    data.append(value)
                except:
                    continue
        if data:
            self.chart_window = ChartWindow(data)
            self.chart_window.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelApp()
    window.show()
    sys.exit(app.exec())

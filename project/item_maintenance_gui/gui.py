from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidgetItem,
    QVBoxLayout, QPushButton, QWidget, QLabel, QTableWidget,
    QHBoxLayout, QLineEdit, QMessageBox, QDialog, QListWidget, QListWidgetItem
)
from PyQt5.QtGui import QKeySequence, QFont, QIntValidator, QColor
from PyQt5.QtCore import Qt
import pandas as pd
import sys

class FitmentDialog(QDialog):
    def __init__(self, parent=None, selected_items=None):
        super().__init__(parent)
        self.setWindowTitle("Select Reason for Maintenance")
        self.layout = QVBoxLayout(self)
        self.search_field = QLineEdit(self)
        # ... (remaining code)

class CustomTable(QTableWidget):
    def __init__(self, rows, cols, columns, df, excel_file, parent=None):
        super().__init__(rows, cols, parent)
        # ... (remaining code)

    def keyPressEvent(self, event):
        # ... (remaining code)

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setWindowTitle("Simple data entry form")
        self.EXCEL_FILE = "Data_Entry.xlsx"
        self.df = pd.read_excel(self.EXCEL_FILE)
        self.sections = {"Item Maint": [
            # ... (remaining code)
        ]}
        self.columns = [item for sublist in self.sections.values() for item in sublist]
        self.main_widget = QWidget()
        # ... (remaining code)

        self.layout.addWidget(button)

    def open_fitment_dialog(self):
        dialog = FitmentDialog(self)
        result = dialog.exec_()
        if result == QDialog.Accepted:
            selected_items = dialog.selected_items()
            # Do something with selected_items

    def add_section(
        self, title, table, clear_button, clear_button_duplicate, start_row_input, end_row_input
    ):
        header_layout = QHBoxLayout()
        label = QLabel(title)
        font = QFont()
        font.setPointSize(24)
        label.setFont(font)
        header_layout.addWidget(label)
        header_layout.addWidget(clear_button)
        header_layout.addWidget(clear_button_duplicate)
        # ... (remaining code)

    def clear_rows(self):
        start_row_text = self.start_row_input.text()
        end_row_text = self.end_row_input.text()
        # ... (remaining code)

    def create_table(self, columns):
        table = CustomTable(50, len(columns), columns, self.df, self.EXCEL_FILE)
        # ... (remaining code)
        return table

    def submit(self):
        self.table.submit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec_())
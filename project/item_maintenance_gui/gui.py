from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QTableWidgetItem,
    QVBoxLayout,
    QPushButton,
    QWidget,
    QLabel,
    QTableWidget,
    QHBoxLayout,
    QLineEdit,
    QMessageBox,
    QDialog,
    QListWidget,
    QListWidgetItem,
)
from PyQt5.QtGui import QKeySequence, QFont, QIntValidator, QColor
from PyQt5.QtCore import Qt
import pandas as pd
import sys
from PyQt5.QtWidgets import QLineEdit

from PyQt5.QtWidgets import QDialog, QListWidget, QVBoxLayout, QPushButton

import os
from PyQt5.QtWidgets import QComboBox


from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QLineEdit,
)
from PyQt5.QtCore import Qt

# todo Add OEM Part Number to the list

class FitmentDialog(QDialog):
    def __init__(self, parent=None, selected_items=None):
        super().__init__(parent)
        self.setWindowTitle("Select Reason for Maintenance")
        self.layout = QVBoxLayout(self)

        # Add search field
        self.search_field = QLineEdit(self)
        self.search_field.setPlaceholderText("Search...")
        self.search_field.textChanged.connect(self.filter_items)
        self.layout.addWidget(self.search_field)

        self.list_widget = QListWidget(self)
        self.layout.addWidget(self.list_widget)

        self.clear_button = QPushButton("Clear Selections", self)
        self.clear_button.setToolTip("Clear all selections")
        self.clear_button.clicked.connect(self.clear_selections)
        self.layout.addWidget(self.clear_button)

        self.ok_button = QPushButton("Apply", self)
        self.ok_button.setToolTip("Apply the selected fitment to the cell")
        self.ok_button.clicked.connect(self.accept)
        self.layout.addWidget(self.ok_button)

        # Replace the options with the provided list
        self.populate_list_widget(
            [
                "Add Replacement #",
                "Bring Prod Back to CUR",
                "Bring Prod Back to NEW",
                "Bring Prod to TUA",
                "Delete Part numbers",
                "NLP Per Vendor",
                "NLP Per Yourself",
                "Remove Description 2 (30)",
                "Remove EAN Code",
                "Remove OEM No.",
                "Remove Replacement Number",
                "Remove UPC #",
                "Supersede Part",
                "Update Amazon Item#",
                "Update Apparel Type",
                "Update Apparel Cat (Y/N)",
                "Update ATV/UTV Cat (Y/N)",
                "Update Bicycle Cat (Y/N)",
                "Update Brand",
                "Update Carb To Prohibited State Lockout",
                "Update Carb To Waiver State Carb Restriction",
                "Update Catalog Price Factor",
                "Update Color",
                "Update Commodity #",
                "Update Cost",
                "Update Country of Origin",
                "Update Dealer Map (Y/N)",
                "Update Description 1 (30)",
                "Update Description 2 (30)",
                "Update Description 3 (30)",
                "Update Dirt Cat (Y/N)",
                "Update Distributor Whse Y or Blank",
                "Update Division, Class, Sub-Class, Sub-Sub-Class",
                "Update Duty Rate",
                "Update EAN Code",
                "Update Expiration Date (Go Live Date)",
                "Update Fly Cat (Y/N)",
                "Update Freight Category 1",
                "Update Freight Category 2",
                "Update Gender",
                "Update Hazmat Code",
                "Update Height (Inches)",
                "Update HTS CODE (Tariff)",
                "Update Item Category",
                "Update Item Taxable",
                "Update Landed Cost Code",
                "Update Length (Inches)",
                "Update Lithium",
                "Update Lot Control",
                "Update Material",
                "Update Model",
                "Update MSDS Req",
                "Update OEM No.",
                "Update Dealer A & Dealer B & Retail",
                "Update Price 2 Closeout",
                "Update Price 3",
                "Update Price 5 Distributor USA",
                "Update Price 6 Distributor Asia",
                "Update Price 7",
                "Update Price 8",
                "Update Price 9 Retail Map",
                "Update Prop65 to Cancerous",
                "Update Prop65 to Cancerous Developmental",
                "Update Prop65 to Developmental",
                "Update Retail Map (Y/N)",
                "Update Round Retail (Y/N)",
                "Update Schedule B#",
                "Update Segment",
                "Update Serial Item (Y/N)",
                "Update Size",
                "Update Snow Cat (Y/N)",
                "Update State Lockout",
                "Update Stock Item",
                "Update Street Cat (Y/N)",
                "Update Style",
                "Update Sub-Segment",
                "Update Tires Cat (Y/N)",
                "Update UN#",
                "Update Unit of Measure",
                "Update UPC #",
                "Update Vendor #",
                "Update Vendor/Manufacturer Part # (20)",
                "Update V-Twin Cat (Y/N)",
                "Update Watercraft Cat (Y/N)",
                "Update Weight (LBS)",
                "Update Weight (LBS) and Dims",
                "Update Width (Inches)",
                "Update Year Design",
            ],
            selected_items,
        )

    def clear_selections(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setCheckState(Qt.Unchecked)

    def filter_items(self, text):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if text.lower() in item.text().lower():
                item.setHidden(False)
            else:
                item.setHidden(True)

        self.ok_button = QPushButton("Apply", self)
        self.ok_button.setToolTip("Apply the selected fitment to the cell")
        self.ok_button.clicked.connect(self.accept)
        self.layout.addWidget(self.list_widget)
        self.layout.addWidget(self.ok_button)

    def populate_list_widget(self, vehicles, selected_items):
        self.list_widget.clear()
        for vehicle in vehicles:
            item = QListWidgetItem(vehicle)
            item.setFlags(
                item.flags() | Qt.ItemIsUserCheckable
            )  # Make the item checkable
            item.setCheckState(Qt.Unchecked)  # Set initial check state to unchecked
            if selected_items and vehicle in selected_items:
                item.setCheckState(
                    Qt.Checked
                )  # Pre-select the item if it was previously selected
            self.list_widget.addItem(item)

    def select_all(self):
        # Iterate over all items in the QListWidget and set their check state to Qt.Checked
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setCheckState(Qt.Checked)

    def selected_items(self):
        selected_items = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                selected_items.append(item.text())
        return selected_items

    def clear_all(self):
        # Iterate over all items in the QListWidget and set their check state to Qt.Unchecked
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setCheckState(Qt.Unchecked)

class CustomTable(QTableWidget):
    def __init__(self, rows, cols, columns, df, excel_file, parent=None):
        super().__init__(rows, cols, parent)
        self.columns = columns
        self.df = df
        self.EXCEL_FILE = excel_file
        self.cellClicked.connect(self.cell_clicked)  # Connect the cellClicked signal

    def keyPressEvent(self, event):
        if event.key() in {Qt.Key_Backspace, Qt.Key_Delete}:
            self.clear_selected_cells()
        elif event.matches(QKeySequence.Paste):
            self.paste_clipboard()
        else:
            super().keyPressEvent(event)

    def paste_clipboard(self):
        clipboard = QApplication.clipboard()
        if clipboard.mimeData().hasText():
            str_clipboard = clipboard.text()
            for i, row in enumerate(str_clipboard.split("\n")):
                for j, column in enumerate(row.split("\t")):
                    new_item = QTableWidgetItem(column)
                    self.setItem(
                        self.currentRow() + i, self.currentColumn() + j, new_item
                    )

    def clear_selected_cells(self):
        for item in self.selectedItems():
            item.setText("")

    def submit(self):
        for row in range(self.rowCount()):
            data = {}
            for col, col_name in enumerate(self.columns):
                item = self.item(row, col)
                if item is not None and item.text().strip():
                    data[col_name] = item.text()
                else:
                    if col == 0:
                        break
                    if item is None:
                        item = QTableWidgetItem()
                        self.setItem(row, col, item)
                    item.setBackground(QColor(255, 0, 0))
                    QMessageBox.warning(
                        self,
                        "Submission Error",
                        f"All cells in a row must have values. Please fill in the cell at row {row+1}, column {col+1}.",
                    )
                    return

            if all(value for value in data.values()):
                self.df = self.df._append(data, ignore_index=True)

        self.df.to_excel(self.EXCEL_FILE, index=False)
        QMessageBox.information(
            self, "Submission Success", "Your submission was successful!"
        )
        self.clearContents()

    def clear_selected_cells(self):
        for item in self.selectedItems():
            item.setText("")

    def cell_clicked(self, row, column):
        if self.horizontalHeaderItem(column).text() == "Select Reason For Maintenance":
            item = self.item(row, column)
            selected_items = item.text().split(", ") if item else None
            dialog = FitmentDialog(self, selected_items)
            if dialog.exec_() == QDialog.Accepted:
                selected_items = ", ".join(dialog.selected_items())
                if item is None:
                    item = QTableWidgetItem()
                    self.setItem(row, column, item)
                item.setText(selected_items)


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setWindowTitle("Simple data entry form")
        self.EXCEL_FILE = "Data_Entry.xlsx"
        self.df = pd.read_excel(self.EXCEL_FILE)
        self.sections = {
            "Item Maint": [
                "Name",
                "Date",
                "Select Reason For Maintenance",
                "Item Status",
                "Item Category",
                "Item Life Cycle",
                "Vendor #",
                "WPS Item #",
                "Description 1 (30)",
                "Description 2 (30)",
                "Brand",
                "Serial Item (Y/N)",
                "Replacement #",
                "Vendor/Manufacturer Part # (20)",
                "UPC #",
                "EAN Code",
                "OEM No.",
                "Expiration Date ( Go Live Date)",
                "Division",
                "Class",
                "Sub-Class",
                "Sub-Sub-Class",
                "Segment",
                "Sub-Segment",
                "Year Design",
                "Model",
                "Style",
                "Size",
                "Color",
                "Apparel Type",
                "Material",
                "Gender",
                "Cost",
                "Dealer A",
                "STD Dealer B",
                "Round Retail (Y/N)",
                "List Price 3",
                "Price 7",
                "Price 8",
                "Closeout",
                "Distributor USA #5",
                "Distributor Asia #6",
                "List Price",
                "Retail Map (Y/N)",
                "Retail Map ($)",
                "Dealer Map (Y/N)",
                "Street Cat (Y/N)",
                "Dirt Cat (Y/N)",
                "Snow Cat (Y/N)",
                "ATV/UTV Cat (Y/N)",
                "Watercraft Cat (Y/N)",
                "Bicycle Cat (Y/N)",
                "Fly Cat (Y/N)",
                "V-Twin Cat (Y/N)",
                "Apparel Cat (Y/N)",
                "Tires Cat (Y/N)",
                "Unit of Measure",
                "UoM Description",
                "Catalog Price Factor",
                "Weight (LBS)",
                "Length (Inches)",
                "Width (Inches)",
                "Height (Inches)",
                "Country of Origin",
                "Freight Category 1",
                "Freight Category 2",
                "Commodity #",
                "HTS CODE (Tariff)",
                "Schedule B#",
                "MSDS Req",
                "UN#",
                "Lithium",
                "Hazmat Code",
                "Landed Cost Code",
                "Duty Rate",
                "Resale / Consumable",
                "Stock Item",
                "Item Taxable",
                "Lot Control",
                "Item Type (Must be F, S, or leave Blank)",
                "Assembly (Must be Y, N, or leave Blank)",
                "Pre-Assembly",
                "Price by Components",
                "Description 3",
                "Distributor Whse Y or Blank",
                "Comments",
                "Prop65 Flag",
                "Prop65 Type",
                "Prop Descrip1",
                "Prop Descrip 2",
                "State Lockout",
                "Carb Restriction",
                "Carb Flag",
                "Carb Type",
                "Amazon Item#",
            ],
        }
        self.columns = [item for sublist in self.sections.values() for item in sublist]
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.layout = QVBoxLayout(self.main_widget)
        self.table = self.create_table(self.columns)
        self.clear_button = QPushButton("Clear All")
        self.clear_button.clicked.connect(self.table.clearContents)
        self.clear_button_duplicate = QPushButton("Clear All (Duplicate)")
        self.clear_button_duplicate.clicked.connect(self.table.clearContents)

        self.start_row_input = QLineEdit()
        self.start_row_input.setValidator(QIntValidator(0, 999))
        self.start_row_input.setMaxLength(3)
        self.start_row_input.setFixedWidth(50)
        self.end_row_input = QLineEdit()
        self.end_row_input.setValidator(QIntValidator(0, 999))
        self.end_row_input.setMaxLength(3)
        self.end_row_input.setFixedWidth(50)
        self.clear_button_duplicate = QPushButton("Clear Row(s)")
        self.clear_button_duplicate.clicked.connect(self.clear_rows)

        self.add_section(
            "Part Number Data Template",
            self.table,
            self.clear_button,
            self.clear_button_duplicate,
            self.start_row_input,
            self.end_row_input,
        )
        for name, func in [("Submit", self.submit), ("Exit", self.close)]:
            button = QPushButton(name)
            button.clicked.connect(func)
            self.layout.addWidget(button)

    def open_fitment_dialog(self):
        dialog = FitmentDialog(self)
        result = dialog.exec_()
        if result == QDialog.Accepted:
            selected_items = dialog.selected_items()
            # Do something with selected_items

    def add_section(
        self,
        title,
        table,
        clear_button,
        clear_button_duplicate,
        start_row_input,
        end_row_input,
    ):
        header_layout = QHBoxLayout()
        label = QLabel(title)
        font = QFont()
        font.setPointSize(24)
        label.setFont(font)
        header_layout.addWidget(label)
        header_layout.addWidget(clear_button)
        header_layout.addWidget(clear_button_duplicate)
        header_layout.addWidget(QLabel("Start Row:"))
        header_layout.addWidget(start_row_input)
        header_layout.addWidget(QLabel("End Row:"))
        header_layout.addWidget(end_row_input)
        header_layout.addStretch()

        fitment_button = QPushButton("Open Fitment Dialog")
        fitment_button.clicked.connect(self.open_fitment_dialog)
        header_layout.addWidget(fitment_button)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        self.layout.addWidget(header_widget)
        self.layout.addWidget(table)

    def clear_rows(self):
        start_row_text = self.start_row_input.text()
        end_row_text = self.end_row_input.text()

        if not start_row_text or not end_row_text:
            return

        start_row = int(start_row_text) - 1
        end_row = int(end_row_text)
        for row in range(start_row, end_row):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item is not None:
                    item.setText("")


    def create_table(self, columns):
        table = CustomTable(50, len(columns), columns, self.df, self.EXCEL_FILE)
        table.setHorizontalHeaderLabels(columns)
        for i in range(len(columns)):
            table.setColumnWidth(i, 200)
        table.setVerticalHeaderLabels([str(i + 1) for i in range(50)])
        table.setSelectionMode(QTableWidget.ExtendedSelection)
        table.setDragDropMode(QTableWidget.InternalMove)
        return table

    def submit(self):
        self.table.submit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec_())

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
    QListWidgetItem
)
from PyQt5.QtGui import QKeySequence, QFont, QIntValidator, QColor
from PyQt5.QtCore import Qt
import pandas as pd
import sys
from PyQt5.QtWidgets import QLineEdit

from PyQt5.QtWidgets import QDialog, QListWidget, QVBoxLayout, QPushButton


from PyQt5.QtWidgets import QComboBox

class FitmentDialog(QDialog):
    def __init__(self, parent=None, selected_items=None):
        super().__init__(parent)
        self.setWindowTitle("Select Fitment") 
        self.layout = QVBoxLayout(self)

        # Read the vehicles from the Excel file
        self.df = pd.read_excel("FitmentVehicles.xlsx")
        self.vehicles = self.df["agg"].tolist()

        # Create combo boxes for Year and Make, and a search box for Model
        self.year_combo = QComboBox(self)
        self.make_combo = QComboBox(self)
        self.model_search = QLineEdit(self)

        # Populate the combo boxes with the unique values from the respective columns
        self.year_combo.addItem("Select Year")
        self.year_combo.addItems([str(year) for year in self.df["year"].unique().tolist()])
        self.make_combo.addItem("Select Make")
        self.make_combo.addItems([str(make) for make in self.df["make"].unique().tolist()])

        # Connect the currentIndexChanged signal of the combo boxes to the filter_vehicles method
        self.year_combo.currentIndexChanged.connect(self.filter_vehicles)
        self.make_combo.currentIndexChanged.connect(self.filter_vehicles)
        self.model_search.textChanged.connect(self.filter_vehicles)

        self.layout.addWidget(self.year_combo)
        self.layout.addWidget(self.make_combo)
        self.layout.addWidget(self.model_search)

        self.list_widget = QListWidget(self)
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)
        self.populate_list_widget(self.vehicles, selected_items)

        self.clear_all_button = QPushButton("Clear All", self)
        self.clear_all_button.setToolTip("Clear all models currently visible") 
        self.clear_all_button.clicked.connect(self.clear_all)
        self.layout.addWidget(self.clear_all_button)

        self.select_all_button = QPushButton("Select All", self)
        self.select_all_button.setToolTip("Select all models currently visible")
        self.select_all_button.clicked.connect(self.select_all)
        self.layout.addWidget(self.select_all_button)

        self.ok_button = QPushButton("Apply", self)
        self.ok_button.setToolTip("Apply the selected fitment to the cell")  # Set tooltip after defining the button
        self.ok_button.clicked.connect(self.accept)
        self.layout.addWidget(self.list_widget)
        self.layout.addWidget(self.ok_button)

    def populate_list_widget(self, vehicles, selected_items):
        self.list_widget.clear()
        for vehicle in vehicles:
            item = QListWidgetItem(vehicle)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)  # Make the item checkable
            item.setCheckState(Qt.Unchecked)  # Set initial check state to unchecked
            if selected_items and vehicle in selected_items:
                item.setCheckState(Qt.Checked)  # Pre-select the item if it was previously selected
            self.list_widget.addItem(item)

    def select_all(self):
        # Iterate over all items in the QListWidget and set their check state to Qt.Checked
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            item.setCheckState(Qt.Checked)

    def filter_vehicles(self):
        year = int(self.year_combo.currentText()) if self.year_combo.currentText().isdigit() else None
        make = self.make_combo.currentText() if self.make_combo.currentText() != "Select Make" else None
        model = self.model_search.text()

        filtered_df = self.df
        if year:
            filtered_df = filtered_df[filtered_df["year"] == year]
        if make:
            filtered_df = filtered_df[filtered_df["make"] == make]
        if model:
            filtered_df = filtered_df[filtered_df["model"].str.contains(model, case=False)]

        self.update_combo_boxes(filtered_df)
        self.populate_list_widget(filtered_df["agg"].tolist(), None)

    def update_combo_boxes(self, df):
        self.year_combo.currentIndexChanged.disconnect()
        self.make_combo.currentIndexChanged.disconnect()

        current_make = self.make_combo.currentText()

        # Filter the DataFrame by the selected year only
        year_df = self.df[self.df["year"] == int(self.year_combo.currentText())] if self.year_combo.currentText().isdigit() else self.df

        self.make_combo.clear()
        self.make_combo.addItems([str(make) for make in year_df["make"].unique().tolist()])

        self.make_combo.setCurrentText(current_make)

        self.year_combo.currentIndexChanged.connect(self.filter_vehicles)
        self.make_combo.currentIndexChanged.connect(self.filter_vehicles)

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
                self.df = self.df.append(data, ignore_index=True)

        self.df.to_excel(self.EXCEL_FILE, index=False)
        QMessageBox.information(
            self, "Submission Success", "Your submission was successful!"
        )
        self.clearContents()

    def clear_selected_cells(self):
        for item in self.selectedItems():
            item.setText("")

    def cell_clicked(self, row, column):
        if self.horizontalHeaderItem(column).text() == "Fitment":
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
            "Item Info": [
                "Part Number",
                "Description 1",
                "Description 2",
                "Brand Name",
                "UPC #",
                "Supersede",
                "Fitment",
            ],
            "Pricing": [
                "Distributor Cost",
                "Dealer Price",
                "Retail/MSRP",
                "Dealer MAP/UPP Price",
                "Retail MAP Price",
            ],
            "Restrictions": ["State Lockout", "Carb Restriction"],
            "Logistics": [
                "Unit of Measure",
                "Minimum Order Qty",
                "Case Qty",
                "Weight (LBS)",
                "Length (Inches)",
                "Width (Inches)",
                "Height (Inches)",
                "Country of Origin",
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

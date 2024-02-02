class ItemCategoryDialog(QDialog):
    def __init__(self, parent=None, selected_items=None):
        super().__init__(parent)
        self.setWindowTitle("Select Item Category")
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
        self.ok_button.setToolTip("Apply the selected category to the cell")
        self.ok_button.clicked.connect(self.accept)
        self.layout.addWidget(self.ok_button)

        # Replace the options with the provided list
        self.populate_list_widget(
            [
                "Category 1",
                "Category 2",
                "Category 3",
                # Add more categories as needed
            ],
            selected_items,
        )

    # The rest of the methods (clear_selections, filter_items, populate_list_widget, select_all, selected_items, clear_all)
    # remain the same as in FitmentDialog
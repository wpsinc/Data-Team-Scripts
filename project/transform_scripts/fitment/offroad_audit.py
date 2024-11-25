import pandas as pd
import os
import tkinter as tk
import numpy as np
import threading

# Thread class to process vehicle data in parallel
class ThreadedVehicleDataProcessor(threading.Thread):
    def __init__(self, file, output_dir, alias):
        threading.Thread.__init__(self)
        self.file = file  # Excel file to process
        self.output_dir = output_dir  # Where to save the processed file
        self.alias = alias  # Alias lookup table

    def run(self):
        table = pd.read_excel(self.file)  # Read Excel file
        alias_merged = self.merge_with_alias(table)  # Merge with alias data
        self.save_output(alias_merged, self.file)  # Save the result

    # Merge the table with alias data
    def merge_with_alias(self, table):
        table["make_New"] = np.nan  # Create a new column for make
        table.loc[table['*'] == '*', 'make_New'] = table.loc[table['*'] == '*', 'Application']  # Fill make_New
        table['make_New'] = table['make_New'].ffill()  # Forward fill make_New
        table["CMS all"] = table["Application"].ffill()  # Forward fill CMS all
        table["alias lookup"] = (table["make_New"].astype(str) + " " + table["CMS all"].astype(str)).str.upper()  # Create alias lookup column
        alias_merged = pd.merge(table, self.alias, on="alias lookup", how="left")  # Merge with alias
        return alias_merged

    # Save the merged table to a CSV file
    def save_output(self, table, file):
        output_file = os.path.splitext(os.path.basename(file))[0] + '.csv'  # Create output filename
        output_path = os.path.join(self.output_dir, output_file)  # Get full path
        table.to_csv(output_path, index=False)  # Save to CSV

# Main class to handle multiple Excel files
class VehicleDataProcessor:
    def __init__(self, source_dir='C:\\Users\\LucyHaskew\\OneDrive\\Templates\\Code Projects\\2.16.24', output_dir='output'):
        self.source_dir = source_dir  # Directory with Excel files
        self.output_dir = output_dir  # Directory for output files

    # Process all Excel files in the source directory
    def process_files(self):
        alias = pd.read_csv("alias.csv", encoding='latin-1')  # Load alias lookup table
        xlsx_files = [os.path.join(self.source_dir, f) for f in os.listdir(self.source_dir) if os.path.isfile(os.path.join(self.source_dir, f)) and f.endswith('.xlsx')]  # Find all Excel files
        os.makedirs(self.output_dir, exist_ok=True)  # Make output directory if it doesn't exist

        threads = []
        for file in xlsx_files:
            processor = ThreadedVehicleDataProcessor(file, self.output_dir, alias)  # Create thread for file
            threads.append(processor)
            processor.start()  # Start thread

        for thread in threads:
            thread.join()  # Wait for all threads to finish

# Class for a simple GUI to select items from a list
class SelectionWindow:
    def __init__(self, df, column_name, title, selectmode):
        self.df = df  # DataFrame to display
        self.column_name = column_name  # Column to list items from
        self.title = title  # Window title
        self.selectmode = selectmode  # Single or multiple selection
        self.selected_items = []  # Store selected items

    # Display selection window and get selected items
    def get_selection(self):
        root = tk.Tk()
        listbox = tk.Listbox(root, selectmode=self.selectmode)  # Create listbox
        listbox.pack()

        for item in sorted(set(self.df[self.column_name].astype(str))):  # Add items to listbox
            listbox.insert(tk.END, item)

        def print_selection():
            selection = listbox.curselection()  # Get selected items
            for i in selection:
                self.selected_items.append(listbox.get(i))
            root.destroy()  # Close the window

        button = tk.Button(root, text="Get Selection", command=print_selection)  # Button to confirm selection
        button.pack()

        root.mainloop()  # Show the window

        return self.selected_items  # Return selected items

# Class for handling year conversions
class YearConverter:
    # Convert year formats to a uniform format
    @staticmethod
    def convert_year(x):
        if isinstance(x, str):
            x = x.split(", ")
            result = []
            for y in x:
                if len(y) == 2:
                    y = "20" + y  # Convert short year to full year
                elif len(y) == 5:
                    y = y.split("-")
                    y = "20" + y[0] + "-" + "20" + y[1]  # Convert range to full years
                result.append(y)
            return ", ".join(result)
        else:
            return x

    # Fill year ranges as lists
    @staticmethod
    def fill_year(x):
        if isinstance(x, str):
            if len(x) == 4:
                return x  # Return single year as is
            elif len(x) == 9:
                x = x.split("-")
                return list(range(int(x[0]), int(x[1]) + 1))  # Return range as list
        return x

# Main function to run the processor
def main():
    processor = VehicleDataProcessor()  # Create processor
    processor.process_files()  # Process all files

if __name__ == "__main__":
    main()

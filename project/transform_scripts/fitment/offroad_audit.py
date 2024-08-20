import pandas as pd
import os
import tkinter as tk
import numpy as np
import threading

class ThreadedVehicleDataProcessor(threading.Thread):
    def __init__(self, file, output_dir, alias):
        threading.Thread.__init__(self)
        self.file = file
        self.output_dir = output_dir
        self.alias = alias

    def run(self):
        table = pd.read_excel(self.file)
        alias_merged = self.merge_with_alias(table)
        self.save_output(alias_merged, self.file)

    def merge_with_alias(self, table):
        table["make_New"] = np.nan
        table.loc[table['*'] == '*', 'make_New'] = table.loc[table['*'] == '*', 'Application']
        table['make_New'] = table['make_New'].ffill()
        table["CMS all"] = table["Application"].ffill()
        table["alias lookup"] = (table["make_New"].astype(str) + " " + table["CMS all"].astype(str)).str.upper()
        alias_merged = pd.merge(table, self.alias, on="alias lookup", how="left")
        return alias_merged

    def save_output(self, table, file):
        output_file = os.path.splitext(os.path.basename(file))[0] + '.csv'
        output_path = os.path.join(self.output_dir, output_file)
        table.to_csv(output_path, index=False)

class VehicleDataProcessor:
    def __init__(self, source_dir='C:\\Users\\LucyHaskew\\OneDrive\\Templates\\Code Projects\\2.16.24', output_dir='output'):
        self.source_dir = source_dir
        self.output_dir = output_dir

    def process_files(self):
        alias = pd.read_csv("alias.csv", encoding='latin-1')
        xlsx_files = [os.path.join(self.source_dir, f) for f in os.listdir(self.source_dir) if os.path.isfile(os.path.join(self.source_dir, f)) and f.endswith('.xlsx')]
        os.makedirs(self.output_dir, exist_ok=True)

        threads = []
        for file in xlsx_files:
            processor = ThreadedVehicleDataProcessor(file, self.output_dir, alias)
            threads.append(processor)
            processor.start()

        for thread in threads:
            thread.join()  # Wait for all threads to finish

class SelectionWindow:
    def __init__(self, df, column_name, title, selectmode):
        self.df = df
        self.column_name = column_name
        self.title = title
        self.selectmode = selectmode
        self.selected_items = []

    def get_selection(self):
        root = tk.Tk()
        listbox = tk.Listbox(root, selectmode=self.selectmode)
        listbox.pack()

        for item in sorted(set(self.df[self.column_name].astype(str))):
            listbox.insert(tk.END, item)

        def print_selection():
            selection = listbox.curselection()
            for i in selection:
                self.selected_items.append(listbox.get(i))
            root.destroy()

        button = tk.Button(root, text="Get Selection", command=print_selection)
        button.pack()

        root.mainloop()

        return self.selected_items

class YearConverter:
    @staticmethod
    def convert_year(x):
        if isinstance(x, str):
            x = x.split(", ")
            result = []
            for y in x:
                if len(y) == 2:
                    y = "20" + y
                elif len(y) == 5:
                    y = y.split("-")
                    y = "20" + y[0] + "-" + "20" + y[1]
                result.append(y)
            return ", ".join(result)
        else:
            return x

    @staticmethod
    def fill_year(x):
        if isinstance(x, str):
            if len(x) == 4:
                return x
            elif len(x) == 9:
                x = x.split("-")
                return list(range(int(x[0]), int(x[1]) + 1))
        return x

def main():
    processor = VehicleDataProcessor()
    processor.process_files()

if __name__ == "__main__":
    main()
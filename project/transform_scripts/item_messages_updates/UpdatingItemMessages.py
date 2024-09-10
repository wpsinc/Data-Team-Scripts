import pandas as pd
import os
import time
import chardet
from openpyxl import load_workbook
from halo import Halo
import warnings
from concurrent.futures import ThreadPoolExecutor
from flask import Flask, render_template, request

warnings.simplefilter("ignore")
app = Flask(__name__)
# Get the current user's home directory
home_dir = os.path.expanduser("~")

# Replace hardcoded paths
base_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech"
)

# Replace hardcoded paths
message_path = os.path.join(
    home_dir, "OneDrive - Arrowhead EP/Data Tech/How To Templates/Updating Item Messages"
)
# Replace hardcoded paths
Warranty_path = os.path.join(
    base_path, "WARRANTY NOTES/Warranty Master File"
)
ItemLookup_path = os.path.join(
    message_path, "ItemMessagesLookup.csv"
)
# Replace hardcoded paths
UN_path = os.path.join(
    base_path, "End Of Month Templates/Linked Reports"
)
class FileHandler:
    @staticmethod
    # Function to detect encoding
    def detect_encoding(file_path):
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']
        return encoding
    @staticmethod
    # Function to read csv
    def read_csv(file_path, encoding, sep = ",", header =0):
        return pd.read_csv(file_path, encoding=encoding, sep=sep, header=0)

    @staticmethod
    def read_excel(file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name=sheet_name)
    
class DataFrameOperations:
    def __init__ (self, dataframe):
        self.dataframe = dataframe
    # Function to rename columns    
    def rename_columns(self, columns_dict):
        self.dataframe.rename(columns=columns_dict, inplace = True)

    # Function to concatenate columns
    def concat_columns(self, new_col_name, col_list, sep="-"):
        try:
            # Convert columns to integers (removing decimals) and then to strings
            self.dataframe[new_col_name] = self.dataframe[col_list[0]].apply(lambda x: str(int(x)) if pd.notna(x) else "")
            for col in col_list[1:]:
                self.dataframe[new_col_name] += sep + self.dataframe[col].apply(lambda x: str(int(x)) if pd.notna(x) else "")
        except KeyError as e:
            print(f"KeyError: {e} - One of the columns {col_list} does not exist in the DataFrame.")
    def trim_columns(self, columns_to_keep):
        self.dataframe = self.dataframe[columns_to_keep]

class DataProcessor:
    @staticmethod
    def clean_text_column(df, column_name):
        df[column_name] = df[column_name].str.strip().str.lower()
        return df

    @staticmethod
    def merge_dataframes(df1, df2, left_on, right_on, how='left', indicator=False):
        return pd.merge(df1, df2, left_on=left_on, right_on=right_on, how=how, indicator=indicator)

    @staticmethod
    def filter_non_matching_rows(df, indicator_column='_merge', non_matching_value='left_only'):
        return df[df[indicator_column] == non_matching_value]

    
item_messages_lookup_file = os.path.join(message_path, "ItemMessagesLookup.csv")
updating_item_messages_file = os.path.join(message_path, "Item Messages Download.csv")

# Detect encoding and read CSV files
encoding_type = FileHandler.detect_encoding(item_messages_lookup_file)
ItemLookup = FileHandler.read_csv(item_messages_lookup_file, encoding=encoding_type, sep=",")
# Initialize DataFrameOperations with the ItemLookup DataFrame
df_ops = DataFrameOperations(ItemLookup)

# Rename columns
df_ops.rename_columns({
    "ItemLookup_ItemNumber": "ItemNumber",
    "Division_DivisionCode": "DivisionCode",
    "Class_ClassCode": "ClassCode",
    "SubClass_SubClass": "SubClassCode",
    "SubSubClass_SubSubClass": "SubSubClassCode",
    "Brand_Lookup_Brand": "Brand",
    "Item_Detail_ItemCategory": "ItemCategory",
    "Item_Detail_Hazmat": "Hazmat",
    "Item_Flags_CarbRestriction" : "CarbRestriction"
})

# List of new columns to create and their respective columns to concatenate
new_columns = {
    "Div-Sub": ["DivisionCode", "SubClassCode"],
    "Div-Sub-Sub": ["DivisionCode", "SubClassCode", "SubSubClassCode"],
    "Div-Class": ["DivisionCode", "ClassCode"]
}

# Loop to create new columns
for new_col, cols in new_columns.items():
    df_ops.concat_columns(new_col, cols)

# Trim Columns (example: keeping only specific columns)
df_ops.trim_columns(["ItemNumber", "DivisionCode", "ClassCode", "SubClassCode", "SubSubClassCode", "Brand", "ItemCategory", "Hazmat", "Div-Sub", "Div-Sub-Sub", "Div-Class", "CarbRestriction"])

def main(updating_item_messages_file, warranty_path):
    # Detect encoding and read the CSV file
    encoding_type = FileHandler.detect_encoding(updating_item_messages_file)
    messages = FileHandler.read_csv(updating_item_messages_file, encoding=encoding_type, sep="\t")
    
    # Clean the 'Explanation Text' column
    messages = DataProcessor.clean_text_column(messages, "Explanation Text")
    
    # Read the warranty file
    warranty_file = os.path.join(warranty_path, "Warranty Notes File.xlsx")
    warranty = FileHandler.read_excel(warranty_file, sheet_name="For Upload")
    warranty_dup = warranty.copy()
    
    # Clean the 'Explanation' column
    warranty = DataProcessor.clean_text_column(warranty, "Explanation")
    warranty = warranty[["Explanation"]]
    
    # Merge the dataframes
    messagestokeep = DataProcessor.merge_dataframes(messages, warranty, left_on="Explanation Text", right_on="Explanation", indicator=True)
    
    # Filter non-matching rows
    non_matching_messages = DataProcessor.filter_non_matching_rows(messagestokeep)
    
    return non_matching_messages

# Run the main function to get non_matching_messages
updating_item_messages_file = os.path.join(message_path, "Item Messages Download.csv")
warranty_path = Warranty_path
non_matching_messages = main(updating_item_messages_file, warranty_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        selected_messages = request.form.getlist('messages')
        filtered_df = non_matching_messages[~non_matching_messages['Message'].isin(selected_messages)]
        return render_template('index.html', tables=[filtered_df.to_html(classes='data')], titles=filtered_df.columns.values)
    return render_template('index.html', tables=[non_matching_messages.to_html(classes='data')], titles=non_matching_messages.columns.values)

if __name__ == '__main__':
    app.run(debug=True)


#Bring in Explanation
#Add column Merge Item Code : Sequence
#Clean / Trim Columns
#Sort Rows by Item then Sequence
#Store this Directory
#Duplicate Directory
#Filter Explanation column from Warranty File to be empty
#Remove Merged Column: Merge Item Code : Sequence
#Merge messages with UN Codes by item Left outer
#Bring in UN#
#Merge Querries with ItemLookup
#"Expanded Item Lookup" = Table.ExpandTableColumn(#"Merged Queries", "Item Lookup", {"Brand", "ItemCategory", "HazmatCode", "Div-Sub", "Div-Sub-Sub", "CarbRestriction", "Div-Class"}, {"Brand", "ItemCategory", "HazmatCode", "Div-Sub", "Div-Sub-Sub", "CarbRestriction", "Div-Class"}),
    #VendorCat = Table.AddColumn(#"Expanded Item Lookup", "VendorCategory", each Text.Combine({[#"Vendor Code"], [ItemCategory]}, "-"), type text),
   # VendorBrand = Table.AddColumn(VendorCat, "VendorBrand", each Text.Combine({[#"Vendor Code"], [Brand]}, "-"), type text),
    #"Trimmed Text" = Table.TransformColumns(VendorBrand,{{"VendorBrand", Text.Trim, type text}}),
    #VendorDivClass = Table.AddColumn(#"Trimmed Text", "VendorDivClass", each Text.Combine({[#"Vendor Code"], [#"Div-Class"]}, "-"), type text),
    #VendorDivSub = Table.AddColumn(VendorDivClass, "VendorDivSub", each Text.Combine({[#"Vendor Code"], [#"Div-Sub"]}, "-"), type text),
   # VendorDivSubSub = Table.AddColumn(VendorDivSub, "VendorDivSubSub", each Text.Combine({[#"Vendor Code"], [#"Div-Sub-Sub"]}, "-"), type text),
   # VendorBrandDivClass = Table.AddColumn(VendorDivSubSub, "VendorBrandDivClass", each Text.Combine({[VendorBrand], [#"Div-Class"]}, "-"), type text),
   # VendorBrandDivSub = Table.AddColumn(VendorBrandDivClass, "VendorBrandDivSub", each Text.Combine({[VendorBrand], [#"Div-Sub"]}, "-"), type text),
    #SpecialVendorMessages = Table.AddColumn(VendorBrandDivSub, "SpecialVendorMessages", each if [VendorBrand] = "6271-4618" then "6271SCORP" else if [VendorBrand] = "5924-840" then "5924BURLY" else if [VendorBrand] = "5436-1875" then "5436-1875" else if [VendorDivSub] = "5716-136-18" then "5716TRAILER" else if [VendorDivClass] = "6514-136-9" then "6514WHEEL" else if [VendorDivClass] = "6514-136-5" then "6514TIRE" else if [VendorDivSubSub] = "5601-134-9" then "5601SUS" else if [VendorBrand] = "6084-4288" then "6084RACING" else if [VendorCategory] = "5849-DIR" then "5849DIR" else if [VendorCategory] = "5333-DIR" then 6 else if [VendorDivClass] = "6371-110-19" then "6371RACEFUEL" else if [VendorDivSub] = "5864-112-4" then "5864Axle" else if [Vendor Code] = "5864" then 5864 else if [VendorBrandDivClass] = "5157-3740-138-3" then "5157OPEN" else if [VendorBrandDivClass] = "5157-1510-102-5" then "5157COVERSLUG" else if [VendorBrandDivClass] = "5157-5930-102-5" then "5157COVERSLUG" else if [VendorBrandDivClass] = "5157-1510-106-19" then "5157COVERSLUG" else if [VendorBrandDivClass] = "5157-5930-106-19" then "5157COVERSLUG" else if Text.StartsWith([VendorBrandDivClass], "5157-1510-104") then "5157COVERSLUG" else if Text.StartsWith([VendorBrandDivClass], "5157-1510-130") then "5157COVERSLUG" else if Text.StartsWith([VendorBrandDivClass], "5157-5930-130") then "5157COVERSLUG" else if [VendorBrandDivSub] = "5157-3740-106-6" then "5157OPEN" else if [#"Div-Sub-Sub"] = "110-1" then "3DRUMDIR" else null),
    #CarbMessage = Table.AddColumn(SpecialVendorMessages, "CarbMessage", each if [CarbRestriction] = "P" then 1 else if [CarbRestriction] = "W" then 2 else null),
    #DrumMessage = Table.AddColumn(CarbMessage, "DrumMessage", each if [#"Div-Sub-Sub"] = "110-1" then "3DRUM" else null),
    #VendorMessages = Table.AddColumn(DrumMessage, "VendorMessages", each if [SpecialVendorMessages] is null then [Vendor Code] else null),
    #DIRMessage = Table.AddColumn(VendorMessages, "DIR Message", each if Text.Contains([VendorCategory], "5849") then null else if [VendorDivClass] = "6371-110-19" then null else if Text.Contains([VendorCategory], "DIR") then 4 else null),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(DIRMessage, {{"SpecialVendorMessages", type text}, {"DrumMessage", type text}, {"CarbMessage", type text}, {"UN#", type text}, {"DIR Message", type text}}, "en-US"),{"SpecialVendorMessages", "DrumMessage", "CarbMessage", "VendorMessages", "UN#", "DIR Message"},Combiner.CombineTextByDelimiter(",", QuoteStyle.None),"Messages"),
    #"Split Column by Delimiter" = Table.ExpandListColumn(Table.TransformColumns(#"Merged Columns", {{"Messages", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}), "Messages"),
    #"Removed Other Columns" = Table.SelectColumns(#"Split Column by Delimiter",{"Item Code", "Messages"}),
    #"Filtered Rows3" = Table.SelectRows(#"Removed Other Columns", each [Messages] <> null and [Messages] <> ""),
    #"Removed Duplicates" = Table.Distinct(#"Filtered Rows3"),
    #"Merged Queries1" = Table.NestedJoin(#"Removed Duplicates", {"Messages"}, WarrantyFile, {"Source.Name.1"}, "WarrantyFile", JoinKind.LeftOuter),
    #"Expanded WarrantyFile" = Table.ExpandTableColumn(#"Merged Queries1", "WarrantyFile", {"Sequence", "Explanation", "Pick Ticket Program", "Invoice Program", "Labels", "R/A", "O/E", "P.O. Entry", "P.O. Print", "M.O. Entry", "M.O. Print", "P.O. Receiving", "WEB", "Expiration Date"}, {"Sequence", "Explanation", "Pick Ticket Program", "Invoice Program", "Labels.1", "R/A", "O/E", "P.O. Entry", "P.O. Print", "M.O. Entry", "M.O. Print", "P.O. Receiving", "WEB", "Expiration Date"}),
    #"Filtered Rows1" = Table.SelectRows(#"Expanded WarrantyFile", each [Explanation] <> null and [Explanation] <> ""),
    #"Added Custom3" = Table.AddColumn(#"Filtered Rows1", "Messages.1", each if Text.Start([Messages],2)="UN" then "1-"&[Messages] else [Messages]),
    #"Removed Columns1" = Table.RemoveColumns(#"Added Custom3",{"Messages"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Removed Columns1",{{"Messages.1", "Messages"}}),
    #"Sorted Rows" = Table.Sort(#"Renamed Columns2",{{"Item Code", Order.Ascending}, {"Messages", Order.Ascending}, {"Sequence", Order.Ascending}}),
    #"Added Index" = Table.AddIndexColumn(#"Sorted Rows", "Index", 0, 1, Int64.Type),
    #"Added Index1" = Table.AddIndexColumn(#"Added Index", "Index.1", 1, 1, Int64.Type),
    #"Merged Queries2" = Table.NestedJoin(#"Added Index1", {"Index"}, #"Added Index1", {"Index.1"}, "Added Index1", JoinKind.LeftOuter),
    #"Expanded Added Index1" = Table.ExpandTableColumn(#"Merged Queries2", "Added Index1", {"Item Code"}, {"Added Index1.Item Code"}),
    #"Sorted Rows1" = Table.Sort(#"Expanded Added Index1",{{"Index", Order.Ascending}}),
    #"Added Custom" = Table.AddColumn(#"Sorted Rows1", "New Seq", each if [Item Code] <> [Added Index1.Item Code] then [Index] else null),
    #"Filled Down" = Table.FillDown(#"Added Custom",{"New Seq"}),
    #"Inserted Subtraction" = Table.AddColumn(#"Filled Down", "Subtraction", each [Index.1] - [New Seq], type number),
    #"Removed Columns" = Table.RemoveColumns(#"Inserted Subtraction",{"Messages", "Sequence", "Index", "Index.1", "Added Index1.Item Code", "New Seq"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns",{"Subtraction", "Item Code", "Explanation", "Pick Ticket Program", "Invoice Program", "Labels.1", "R/A", "O/E", "P.O. Entry", "P.O. Print", "M.O. Entry", "M.O. Print", "P.O. Receiving", "WEB", "Expiration Date"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns",{{"Subtraction", "Sequence"}, {"Labels.1", "Labels"}}),
    #"Added Custom1" = Table.AddColumn(#"Renamed Columns", "COMPANY #", each 1),
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "ADD / DEL", each "A"),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Custom2",{"COMPANY #", "ADD / DEL", "Sequence", "Item Code", "Explanation", "Pick Ticket Program", "Invoice Program", "Labels", "R/A", "O/E", "P.O. Entry", "P.O. Print", "M.O. Entry", "M.O. Print", "P.O. Receiving", "WEB", "Expiration Date"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns1",{{"Item Code", "Item"}, {"Expiration Date", "Message Expire Date"}}),
    #"Inserted Merged Column" = Table.AddColumn(#"Renamed Columns1", "Merged", each Text.Combine({[Item], Text.From([Sequence], "en-US")}, ":"), type text),
    #"Filtered Rows2" = Table.SelectRows(#"Inserted Merged Column", each [Item] <> null and [Item] <> "")

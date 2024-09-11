import os
import pandas as pd
import chardet
import warnings
from flask import Flask, request, render_template

warnings.simplefilter("ignore")
app = Flask(__name__)

# Get the current user's home directory
class PathManager:
    def __init__(self):
        self.home_dir = os.path.expanduser("~")
        self.base_path = os.path.join(self.home_dir, "OneDrive - Arrowhead EP/Data Tech")
        self.message_path = os.path.join(self.base_path, "How To Templates/Updating Item Messages")
        self.item_lookup_path = os.path.join(self.message_path, "ItemMessagesLookup.csv")
        self.messages_cur_path = os.path.join(self.message_path, "Item Messages Download.csv")
        self.warranty_path = os.path.join(self.base_path, "WARRANTY NOTES/Warranty Master File/Warranty Notes File.xlsx")
        self.un_path = os.path.join(self.base_path, "End Of Month Templates/Linked Reports/UNCodes.xlsx")

class FileHandler:
    @staticmethod
    def detect_encoding(file_path):
        """Detect the encoding of a file."""
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
        return result['encoding']

    @staticmethod
    def read_csv(file_path, encoding, sep=",", header=0):
        """Read a CSV file with the specified encoding."""
        return pd.read_csv(file_path, encoding=encoding, sep=sep, header=header)

    @staticmethod
    def read_excel(file_path, sheet_name):
        """Read an Excel file."""
        return pd.read_excel(file_path, sheet_name=sheet_name)

class DataframeOperations:
    def __init__(self, df):
        self.df = df

    def rename_columns(self, columns_dict):
        """Rename columns in the dataframe."""
        self.df.rename(columns=columns_dict, inplace=True)

    def concat_columns(self, new_col_name, col_list, sep="-"):
        """Concatenate columns with a specified separator."""
        try:
            self.df[new_col_name] = self.df[col_list[0]].apply(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x) if pd.notna(x) else "")
            for col in col_list[1:]:
                self.df[new_col_name] += self.df[col].apply(lambda x: sep + str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else sep + str(x) if pd.notna(x) else "")
            # Remove trailing separator if present
            self.df[new_col_name] = self.df[new_col_name].str.rstrip(sep)
        except KeyError as e:
            print(f"KeyError: {e} - One of the columns {col_list} does not exist in the df.")
    def concat_colon(self, new_col_name, col_list, sep=":"):
        """Concatenate columns with a colon separator."""
        try:
            self.df[new_col_name] = self.df[col_list[0]].apply(lambda x: str(x) if pd.notna(x) else "")
            for col in col_list[1:]:
                self.df[new_col_name] += sep + self.df[col].apply(lambda x: str(int(x)) if pd.notna(x) else "")
        except KeyError as e:
            print(f"KeyError: {e} - One of the columns {col_list} does not exist in the df.")

    def trim_columns(self, columns_to_keep):
        """Keep only specified columns in the dataframe."""
        self.df = self.df[columns_to_keep]

    def get_dataframe(self):
        """Return the dataframe."""
        return self.df

class DataProcessor:
    @staticmethod
    def clean_text_column(df, column_name):
        """Clean text in a specified column."""
        df[column_name] = df[column_name].str.strip().str.lower()
        return df

    @staticmethod
    def merge_dataframes(df1, df2, left_on, right_on, how='left', indicator=False):
        """Merge two dataframes."""
        return pd.merge(df1, df2, left_on=left_on, right_on=right_on, how=how, indicator=indicator)

    @staticmethod
    def filter_non_matching_rows(df, indicator_column='_merge', non_matching_value='left_only'):
        """Filter rows that do not match in a merge."""
        return df[df[indicator_column] == non_matching_value]

# Initialize PathManager
paths = PathManager()

# Initialize objects and files
item_messages_lookup_file = paths.item_lookup_path
updating_item_messages_file = paths.messages_cur_path
warranty_path = FileHandler.read_excel(paths.warranty_path, sheet_name="For Upload")
warranty_dup = warranty_path.copy()

# Detect encoding and read CSV files
encoding_type = FileHandler.detect_encoding(item_messages_lookup_file)
ItemLookup = FileHandler.read_csv(item_messages_lookup_file, encoding=encoding_type, sep=",")
encoding_type = FileHandler.detect_encoding(updating_item_messages_file)
messages = FileHandler.read_csv(updating_item_messages_file, encoding=encoding_type, sep="\t")
messages_dup = messages.copy()

# Read the UNCodes file
UN = FileHandler.read_excel(paths.un_path, sheet_name="page")

# Initialize DataframeOperations with the ItemLookup df
df_ops = DataframeOperations(ItemLookup)

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
    "Item_Flags_CarbRestriction": "CarbRestriction",
    "VendorDetail_VendorNumber": "Vendor"
})

# List of new columns to create and their respective columns to concatenate
new_columns = {
    "Div-Sub": ["DivisionCode", "SubClassCode"],
    "Div-Sub-Sub": ["DivisionCode", "SubSubClassCode"],
    "Div-Class": ["DivisionCode", "ClassCode"],
    "Vendor-Cat": ["Vendor", "ItemCategory"],
    "Vendor-Brand": ["Vendor", "Brand"],
    "Vendor-Div-Class": ["Vendor", "DivisionCode", "ClassCode"],
    "Vendor-Div-Sub" : ["Vendor", "DivisionCode", "SubClassCode"],
    "Vendor-Div-Sub-Sub" : ["Vendor", "DivisionCode", "SubSubClassCode"],
    "Vendor-Brand-Div-Class" : ["Vendor", "Brand", "DivisionCode", "ClassCode"],
    "Vendor-Brand-Div-Sub" : ["Vendor", "Brand", "DivisionCode", "SubClassCode"]
}

# Loop to create new columns
for new_col, cols in new_columns.items():
    df_ops.concat_columns(new_col, cols)

message_current = DataframeOperations(messages_dup)
new_columns = {
    "Item-Seq": ["Item Code", "Sequence"]
}

# Loop to create new columns
for new_col, cols in new_columns.items():
    message_current.concat_colon(new_col, cols)

# Access the underlying df
message_current_df = message_current.get_dataframe()
message_current_unique = message_current_df["Item Code"].unique()
# Convert message_current_unique to DataFrame
message_current_unique_df = pd.DataFrame(message_current_unique, columns=["Item Code"])

# Merge the dataframes
MessageCreation = DataProcessor.merge_dataframes(message_current_unique_df, ItemLookup, left_on="Item Code", right_on="ItemNumber")

# Bring in UN#
MessageCreation_merge = DataProcessor.merge_dataframes(MessageCreation, UN, left_on="Item Code", right_on="Item Code")

# Initialize DataframeOperations with the merged dataframe
MessageCreation_merge = DataframeOperations(MessageCreation_merge)

# Trim Columns (example: keeping only specific columns)
MessageCreation_merge.trim_columns(["ItemNumber", "Vendor", "Div-Sub", "Div-Sub-Sub", "Div-Class", "Vendor-Cat", "Vendor-Brand", "Vendor-Div-Class", "Vendor-Div-Sub", "Vendor-Div-Sub-Sub", "Vendor-Brand-Div-Class", "Vendor-Brand-Div-Sub", "Hazmat", "CarbRestriction", "UN#"])

# Access the underlying dataframe if needed
MessageCreation_merge_df = MessageCreation_merge.get_dataframe()

def process_special_messages(df):
    # Apply the if-then-else logic for SpecialVendorMessages
    df['SpecialVendorMessages'] = df.apply(
        lambda row: "6271SCORP" if row['Vendor-Brand'] == "6271-4618" else
                    "5924BURLY" if row['Vendor-Brand'] == "5924-840" else
                    "5436-1875" if row['Vendor-Brand'] == "5436-1875" else
                    "5716TRAILER" if row['Vendor-Div-Sub'] == "5716-136-18" else
                    "6514WHEEL" if row['Vendor-Div-Class'] == "6514-136-9" else
                    "6514TIRE" if row['Vendor-Div-Class'] == "6514-136-5" else
                    "5601SUS" if row['Vendor-Div-Sub-Sub'] == "5601-134-9" else
                    "6084RACING" if row['Vendor-Brand'] == "6084-4288" else
                    "5849DIR" if row['Vendor-Cat'] == "5849-DIR" else
                    6 if row['Vendor-Cat'] == "5333-DIR" else
                    "6371RACEFUEL" if row['Vendor-Div-Class'] == "6371-110-19" else
                    "5864Axle" if row['Vendor-Div-Sub'] == "5864-112-4" else
                    5864 if row['Vendor'] == "5864" else
                    "5157OPEN" if row['Vendor-Brand-Div-Class'] == "5157-3740-138-3" else
                    "5157COVERSLUG" if row['Vendor-Brand-Div-Class'] in ["5157-1510-102-5", "5157-5930-102-5", "5157-1510-106-19", "5157-5930-106-19"] else
                    "5157COVERSLUG" if row['Vendor-Brand-Div-Class'].startswith("5157-1510-104") or row['Vendor-Brand-Div-Class'].startswith("5157-1510-130") or row['Vendor-Brand-Div-Class'].startswith("5157-5930-130") else
                    "5157OPEN" if row['Vendor-Brand-Div-Sub'] == "5157-3740-106-6" else
                    "3DRUMDIR" if row['Div-Sub-Sub'] == "110-1" else
                    None,
        axis=1
    )

    # Apply the if-then-else logic for CarbMessage
    df['CarbMessage'] = df.apply(
        lambda row: 1 if row['CarbRestriction'] == "P" else
                    2 if row['CarbRestriction'] == "W" else
                    None,
        axis=1
    )

    # Apply the if-then-else logic for DrumMessage
    df['DrumMessage'] = df.apply(
        lambda row: "3DRUM" if row['Div-Sub-Sub'] == "110-1" else None,
        axis=1
    )

    # Apply the if-then-else logic for VendorMessages
    df['VendorMessages'] = df.apply(
        lambda row: row['Vendor'] if pd.isna(row['SpecialVendorMessages']) else None,
        axis=1
    )

    # Apply the if-then-else logic for DIRMessage
    df['DIR Message'] = df.apply(
        lambda row: None if "5849" in row['Vendor-Cat'] or row['Vendor-Div-Class'] == "6371-110-19" else
                    4 if "DIR" in row['Vendor-Cat'] else
                    None,
        axis=1
    )

    # Combine columns into a single 'Messages' column
    df['Messages'] = df[['SpecialVendorMessages', 'DrumMessage', 'CarbMessage', 'VendorMessages', 'UN#', 'DIR Message']].apply(
        lambda x: ','.join(x.dropna().astype(str)), axis=1
    )

    # Split the 'Messages' column into lists
    df['Messages'] = df['Messages'].str.split(',')

    # Explode the 'Messages' column to create new rows
    df_exploded = df.explode('Messages')

    # Convert the values in the 'Messages' column to strings, handling both numeric and non-numeric values
    df_exploded['Messages'] = df_exploded['Messages'].apply(lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else str(x))

    # Trim columns (example: keeping only specific columns)
    df_exploded = df_exploded[['ItemNumber', 'Messages']]

    

    return df_exploded

# Example usage
message_creation_merge_df = MessageCreation_merge.get_dataframe()
processed_df = process_special_messages(message_creation_merge_df)
processed_df.to_excel('WarrantyMessageCreation.xlsx', index=False)

# Your function definition
def create_messages(processed_df, warranty_dup):
    # Check if inputs are DataFrames
    if not isinstance(processed_df, pd.DataFrame):
        raise TypeError("processed_df is not a pandas DataFrame")
    if not isinstance(warranty_dup, pd.DataFrame):
        raise TypeError("warranty_dup is not a pandas DataFrame")
    
    # Print column names to debug
    print("Columns in processed_df:", processed_df.columns)
    print("Columns in warranty_dup:", warranty_dup.columns)
    
    # Merge dataframes
    processed_df = DataProcessor.merge_dataframes(processed_df, warranty_dup, left_on='Messages', right_on='Source.Name.1')
    
    # List of columns to keep
    columns_to_keep = [
        "ItemNumber"
        "Sequence",
        "Explanation",
        "Pick Ticket Program",
        "Invoice Program",
        "Labels",
        "R/A",
        "O/E",
        "P.O. Entry",
        "P.O. Print",
        "M.O. Entry",
        "M.O. Print",
        "P.O. Receiving",
        "WEB",
        "Expiration Date"
    ]
    
    # Initialize DataframeOperations with the merged dataframe
    df_ops = DataframeOperations(processed_df)
    
    # Trim columns
    df_ops.trim_columns(columns_to_keep)
    
    # Get the trimmed dataframe
    processed_df = df_ops.get_dataframe()
    processed_df["Sequence"] = processed_df["sequence"].apply(lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else str(x))
    return processed_df

# Example usage
message_creation_merge_df = MessageCreation_merge.get_dataframe()
processed_df = process_special_messages(message_creation_merge_df)
result_df = create_messages(processed_df, warranty_dup)

# Now result_df contains the merged dataframe with the specified columns
print(result_df)

def main(messages, warranty_path):
    # Clean the 'Explanation Text' column
    messages = DataProcessor.clean_text_column(messages, "Explanation Text")

    # Clean the 'Explanation' column
    warranty = DataProcessor.clean_text_column(warranty_path, "Explanation")
    warranty = warranty[["Explanation"]]

    # Merge the dfs
    messagestokeep = DataProcessor.merge_dataframes(messages, warranty, left_on="Explanation Text", right_on="Explanation", indicator=True)

    # Filter non-matching rows
    non_matching_messages = DataProcessor.filter_non_matching_rows(messagestokeep)

    return non_matching_messages, messagestokeep

# Run the main function to get non_matching_messages and messagestokeep
non_matching_messages, messagestokeep = main(messages, warranty_path)
def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        print('Not running with the Werkzeug Server')
    else:
        func()
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Rename the column
        messagestokeep.rename(columns={'Item Code': 'ItemNumber'}, inplace=True)
        selected_explanations = request.form.getlist('explanations')
        filtered_df = messagestokeep[~messagestokeep['Explanation Text'].isin(selected_explanations)]
        filtered_df.to_csv('filtered_messages.csv', index=False)  # Save the filtered dataset

        # Render the shutdown page
        response = render_template('shutdown.html')

        # Set a custom flag to shutdown the server
        request.environ['shutdown_flag'] = True

        return response
    explanations = non_matching_messages['Explanation Text'].unique()
    return render_template('index.html', explanations=explanations)

@app.teardown_request
def teardown_request(exception):
    if request.environ.get('shutdown_flag'):
        shutdown_server()

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)
print("Thanks for your submissions")



# Store messages to keep
filtered_messages = pd.read_csv("filtered_messages.csv")
# List of new columns to create and their respective columns to concatenate
messages_toKeep = DataframeOperations(filtered_messages)
new_columns = {
"Item-Seq": ["Item Code", "Sequence"]

}

# Loop to create new columns
for new_col, cols in new_columns.items():
    messages_toKeep.concat_colon(new_col, cols)


    
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

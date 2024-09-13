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
        self.output_path = os.path.join(self.home_dir, "OneDrive - Arrowhead EP/Templates")

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
    
    def drop_column(self, column_name):
        self.df = self.df.drop(columns=[column_name])
        return self.df
    
    def add_columns(self):
        self.df['ADD/DEL'] = self.df.apply(
            lambda row: 'U' if pd.notna(row['Company Code_dup']) and pd.notna(row['Company Code_merge']) else 
                                'D' if pd.notna(row['Company Code_dup']) else
                                'A' if pd.notna(row['Company Code_merge']) else
                                None,
            axis=1
        )

        columns_to_add = {
            'COMPANY #': ['Company Code_dup', 'Company Code_merge'],
            'WPS Item #': ['ItemNumber_dup', 'ItemNumber_merge'],
            'Sequence': ['Sequence_dup', 'Sequence_merge'],
            'Explanation/Message (CAPS, 60)': ['Explanation_dup', 'Explanation_merge'],
            'PT/Pick Ticket Program': ['Pick Ticket Program_dup', 'Pick Ticket Program_merge'],
            'Invoice Program': ['Invoice Program_dup', 'Invoice Program_merge'],
            'Labels': ['Labels_dup', 'Labels_merge'],
            'Return Auth/Warranty (X or blank)': ['R/A_dup', 'R/A_merge'],
            'Order Entry/Customer Service (X or blank)': ['O/E_dup', 'O/E_merge'],
            'PO Entry': ['P.O. Entry_dup', 'P.O. Entry_merge'],
            'PO Form': ['P.O. Print_dup', 'P.O. Print_merge'],
            'MO Entry': ['M.O. Entry_dup', 'M.O. Entry_merge'],
            'MO Form': ['M.O. Print_dup', 'M.O. Print_merge'],
            'PO Receiving': ['P.O. Receiving_dup', 'P.O. Receiving_merge'],
            'WEB': ['WEB_dup', 'WEB_merge'],
            'Expiration Date': ['Expiration Date_dup', 'Expiration Date_merge']
        }

        for new_col, (dup_col, merge_col) in columns_to_add.items():
            self.df[new_col] = self.df.apply(
                lambda row: row[dup_col] if pd.notna(row[dup_col]) and pd.notna(row[merge_col]) else 
                                    row[dup_col] if pd.notna(row[merge_col]) else
                                    row[merge_col] if pd.notna(row[dup_col]) else
                                    None,
                axis=1
            )
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
output_path = paths.output_path

# Detect encoding and read CSV files
encoding_type = FileHandler.detect_encoding(item_messages_lookup_file)
ItemLookup = FileHandler.read_csv(item_messages_lookup_file, encoding=encoding_type, sep=",")
encoding_type = FileHandler.detect_encoding(updating_item_messages_file)
messages = FileHandler.read_csv(updating_item_messages_file, encoding=encoding_type, sep="\t")
messages_dup = messages.copy()

messages_dup = DataframeOperations(messages_dup)
new_columns = {
    "Item-Seq": ["Item Code", "Sequence"]
}

for new_col, cols in new_columns.items():
    messages_dup.concat_colon(new_col, cols)
messages_dup = messages_dup.get_dataframe()
messages_dup['Company Code'] = 1
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
def create_messages(processed_df, warranty_dup):
    # Check if inputs are DataFrames
    if not isinstance(processed_df, pd.DataFrame):
        raise TypeError("processed_df is not a pandas DataFrame")
    if not isinstance(warranty_dup, pd.DataFrame):
        raise TypeError("warranty_dup is not a pandas DataFrame")
    
    # Remove duplicates from each DataFrame
    processed_df = processed_df.drop_duplicates(ignore_index=True)
    warranty_dup = warranty_dup.drop_duplicates(ignore_index=True)
    
    # Merge processed_df with warranty_dup by the Message and Source.Name.1 column
    merged_df = pd.merge(processed_df, warranty_dup, left_on='Messages', right_on='Source.Name.1', how='left')
    merged_df['Company Code'] = 1
    # List of columns to keep
    columns_to_keep = [
        "ItemNumber",
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
    
    # Trim columns
    merged_df = merged_df[columns_to_keep]
    
    # Drop rows with NaN in Explanation
    merged_df = merged_df.dropna(subset=['Explanation'])
    merged_df = merged_df.dropna(subset = ['Sequence'])
    #Drop Duplicates
    merged_df = merged_df.drop_duplicates(ignore_index=True)
    merged_df['Company Code'] = 1
  
    return merged_df

def append_messages(merged_df, filtered_df):
    # Concatenate the DataFrames
    appended = pd.concat([merged_df, filtered_df])
    
    # Create new column Item-Seq
    appended.drop_duplicates(ignore_index=True)
    appended = appended.dropna(subset=['Expiration Date'])
    appended = appended.sort_values(by=['ItemNumber', 'Company Code', 'Sequence'])
    # Renumber Sequence based on ItemNumber
    appended['Sequence'] = appended.groupby('ItemNumber').cumcount() + 1
    appended['Item-Seq'] = appended['ItemNumber'].astype(str) + ':' + appended['Sequence'].astype(str)
    return appended

def main(messages, warranty_path):
    # Create an instance of the class with the messages DataFrame
    df_operations = DataframeOperations(messages)
    messages = df_operations.get_dataframe()
    # Drop the 'Explanation' column
    
    messages = messages.dropna(subset=['Explanation Text'])
    messages = DataProcessor.clean_text_column(messages, "Explanation Text")

    # Clean the 'Explanation' column
    warranty = DataProcessor.clean_text_column(warranty_path, "Explanation")
    warranty = warranty[["Explanation"]]

    # Merge the dfs
    messagestokeep = DataProcessor.merge_dataframes(messages, warranty, left_on="Explanation Text", right_on="Explanation", indicator=True)
    messagestokeep = messagestokeep.drop(columns=["Explanation"])

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
        selected_explanations = request.form.getlist('explanations')
        print("Selected explanations:", selected_explanations)  # Debug: Print selected explanations

        if not selected_explanations:
            print("No explanations selected")
        pass

        left_only_df = messagestokeep[messagestokeep['_merge'] == 'left_only']
 
        # Replace 'nan' strings with actual NaN values
        left_only_df['Explanation Text'] = left_only_df['Explanation Text'].replace('nan', pd.NA)

        # Ensure data types match and strip spaces
        left_only_df['Explanation Text'] = left_only_df['Explanation Text'].astype(str).str.strip()
        selected_explanations = [str(explanation).strip() for explanation in selected_explanations]

        # Filter out the selected explanations
        filtered_df = left_only_df[~left_only_df['Explanation Text'].isin(selected_explanations)]
        filtered_df.rename(columns={
            "Item Code": "ItemNumber",
            "Explanation Text": "Explanation",
            "Message Expire Date": "Expiration Date"
        }, inplace=True)
        
        filtered_df = filtered_df.drop_duplicates()
        filtered_df['Company Code'] = 0

        filtered_df.to_csv('filtered_messages.csv', index=False)
        
        response = render_template('shutdown.html')
        request.environ['shutdown_flag'] = True
        return response

    if 'Explanation Text' in non_matching_messages.columns:
        explanations = non_matching_messages['Explanation Text'].unique()
        
    else:
        explanations = []
        print("Column 'Explanation Text' not found in non_matching_messages")
    return render_template('index.html', explanations=explanations)

@app.teardown_request
def teardown_request(exception):
    if request.environ.get('shutdown_flag'):
        shutdown_server()

if __name__ == '__main__':
    app.run(debug=True, use_reloader=False)

# Example usage
message_creation_merge_df = MessageCreation_merge.get_dataframe()
processed_df = process_special_messages(message_creation_merge_df)
filtered_df = pd.read_csv("filtered_messages.csv")  # Assuming filtered_df is read from a CSV file
result_df = create_messages(processed_df, warranty_dup)
appendmessages = append_messages(result_df,filtered_df)

# Rename columns in messages_dup
messages_dup = messages_dup.rename(columns={'Item Code': 'ItemNumber', 'Explanation Text': 'Explanation', 'Web': 'WEB', 'Message Expire Date': 'Expiration Date'})

# Remove 'Vendor Code' column from both dataframes
messages_dup = messages_dup.drop(columns=['Vendor Code'])

# Add 'Company Code' column with value 1 to both dataframes
appendmessages['Company Code'] = 1

# Ensure column names are correct
assert 'Item-Seq' in messages_dup.columns, "Column 'Item-Seq' not found in messages_dup"
assert 'Item-Seq' in appendmessages.columns, "Column 'Item-Seq' not found in result_df"

# Check for overlapping 'Item-Seq' values
common_item_seq = set(messages_dup['Item-Seq']).intersection(set(appendmessages['Item-Seq']))
print(f"Number of common 'Item-Seq' values: {len(common_item_seq)}")

# Merge dataframes
merged_df = pd.merge(messages_dup, appendmessages, on='Item-Seq', how='outer', suffixes=('_dup', '_merge'))
merged_df.drop_duplicates(ignore_index=True)

def check_sequence(row, col_dup, col_merge):
    if pd.notna(row['Sequence_dup']) and pd.notna(row['Sequence_merge']):
        return row[col_dup] if pd.notna(row[col_dup]) else row[col_merge]
    elif pd.notna(row['Sequence_dup']):
        return row[col_dup]
    elif pd.notna(row['Sequence_merge']):
        return row[col_merge]
    else:
        return None
# Apply the function to each column
merged_df['ADD/DEL'] = merged_df.apply(
    lambda row: 'U' if pd.notna(row['Sequence_dup']) and pd.notna(row['Sequence_merge']) else 
                        'D' if pd.notna(row['Sequence_dup']) else
                        'A' if pd.notna(row['Sequence_merge']) else
                        None,
    axis=1
)

merged_df['COMPANY #'] = merged_df.apply(
    lambda row: check_sequence(row, 'Company Code_dup', 'Company Code_merge'),
    axis=1
)

merged_df['WPS Item #'] = merged_df.apply(
    lambda row: check_sequence(row, 'ItemNumber_dup', 'ItemNumber_merge'),
    axis=1
)

merged_df['Sequence'] = merged_df.apply(
    lambda row: check_sequence(row, 'Sequence_dup', 'Sequence_merge'),
    axis=1
)

merged_df['Explanation/Message (CAPS, 60)'] = merged_df.apply(
    lambda row: check_sequence(row, 'Explanation_dup', 'Explanation_merge'),
    axis=1
)

merged_df['PT/Pick Ticket Program'] = merged_df.apply(
    lambda row: check_sequence(row, 'Pick Ticket Program_dup', 'Pick Ticket Program_merge'),
    axis=1
)

merged_df['Invoice Program'] = merged_df.apply(
    lambda row: check_sequence(row, 'Invoice Program_dup', 'Invoice Program_merge'),
    axis=1
)

merged_df['Labels'] = merged_df.apply(
    lambda row: check_sequence(row, 'Labels_dup', 'Labels_merge'),
    axis=1
)

merged_df['Return Auth/Warranty (X or blank)'] = merged_df.apply(
    lambda row: check_sequence(row, 'R/A_dup', 'R/A_merge'),
    axis=1
)

merged_df['Order Entry/Customer Service (X or blank)'] = merged_df.apply(
    lambda row: check_sequence(row, 'O/E_dup', 'O/E_merge'),
    axis=1
)

merged_df['PO Entry'] = merged_df.apply(
    lambda row: check_sequence(row, 'P.O. Entry_dup', 'P.O. Entry_merge'),
    axis=1
)

merged_df['PO Form'] = merged_df.apply(
    lambda row: check_sequence(row, 'P.O. Print_dup', 'P.O. Print_merge'),
    axis=1
)

merged_df['MO Entry'] = merged_df.apply(
    lambda row: check_sequence(row, 'M.O. Entry_dup', 'M.O. Entry_merge'),
    axis=1
)

merged_df['MO Form'] = merged_df.apply(
    lambda row: check_sequence(row, 'M.O. Print_dup', 'M.O. Print_merge'),
    axis=1
)

merged_df['PO Receiving'] = merged_df.apply(
    lambda row: check_sequence(row, 'P.O. Receiving_dup', 'P.O. Receiving_merge'),
    axis=1
)

merged_df['WEB'] = merged_df.apply(
    lambda row: check_sequence(row, 'WEB_dup', 'WEB_merge'),
    axis=1
)

merged_df['Expiration Date'] = merged_df.apply(
    lambda row: check_sequence(row, 'Expiration Date_dup', 'Expiration Date_merge'),
    axis=1
)
merged_df = DataframeOperations(merged_df)
merged_df.trim_columns(['ADD/DEL', 'COMPANY #', 'WPS Item #', 'Sequence', 'Explanation/Message (CAPS, 60)', 
    'PT/Pick Ticket Program', 'Invoice Program', 'Labels', 
    'Return Auth/Warranty (X or blank)', 'Order Entry/Customer Service (X or blank)', 
    'PO Entry', 'PO Form', 'MO Entry', 'MO Form', 'PO Receiving', 'WEB', 'Expiration Date'
])
# Access the DataFrame stored in the 'df' attribute
data = merged_df.df

# Ensure 'data' is a pandas DataFrame
if isinstance(data, pd.DataFrame):
    data.to_csv('ItemUpload.csv', index=False)
    data.to_csv(os.path.join(output_path, 'ItemMessages Template.csv'), index=False)
else:
    print("The 'df' attribute is not a pandas DataFrame.")
    
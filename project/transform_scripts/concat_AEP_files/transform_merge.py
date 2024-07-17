import pandas as pd

temp_output_file = 'C:/Users/LucyHaskew/Downloads/WPS_Harley/output.xlsx'
transformed_output_file = 'C:/Users/LucyHaskew/Downloads/WPS_Harley/transformed_output.xlsx'

align_columns = ['Type','Make','CC','Model','Years','Value']

df = pd.read_excel(temp_output_file, skiprows=2)

condensed_data = {col: [] for col in align_columns}

for col in df.columns:
    if col in align_columns:
        pd.concat(condensed_data[col])

    else:
        if df[col].dtype != ['Type','Make','CC','Model','Years']:
            pattern = r'\d{2}-\d+'
            if df[col].str.match(pattern).all():
                condensed_data['Value'].extend(df[col].dropna().tolist())

condensed_df = pd.DataFrame({col: pd.Series(condensed_data[col]) for col in align_columns})
condensed_df.to_excel(transformed_output_file, index=False)

print(f'Combined data saved to: {transformed_output_file}')

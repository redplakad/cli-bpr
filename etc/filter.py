import pandas as pd

# Path to your CSV file (assuming voucher.csv is in the same directory as your script)
csv_file = 'voucher.csv'

# Load CSV into DataFrame with delimiter '|'
df = pd.read_csv(csv_file, sep='|')

# Print current columns to inspect the actual headers
print("Current columns:")
print(df.columns)

# Update column names to match the expected headers
expected_columns = ['NO REK', 'NO KUPON', 'NAMA', 'AREA', 'PLAFON', 'KELIPATAN PLAFON',
                    'KELIPATAN TOPUP', 'JUMLAH KUPON', 'TGL BUKA', 'TGL JT', 'KAB/KOTA']
df.columns = [col.replace(' ', '_').replace('.', '_') for col in df.columns]

# Function to fill missing values based on 'NAMA' column
def fill_missing_values(df):
    filled_df = df.copy()
    for col in filled_df.columns:
        filled_df[col] = filled_df.groupby('NAMA')[col].fillna(method='ffill')
    return filled_df

# Fill missing values
df_filled = fill_missing_values(df)

# Output the DataFrame to CSV
output_csv = 'output_filled.csv'
df_filled.to_csv(output_csv, index=False, sep='|')

print(f"Processed data with filled values saved to {output_csv}")
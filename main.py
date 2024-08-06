import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side

# Load the new data Excel file
new_data_file_path = 'Kilkenny tankering. 2023 (5).xlsx'
df = pd.read_excel(new_data_file_path, sheet_name='June', header=2)

# Rename 'Location' to 'Source'
df.rename(columns={'Location': 'Source'}, inplace=True)

# Define the columns to extract and convert to numeric
columns = ['Source', '4000g', '3500g', '2600g', 'Hours']
data_to_extract = df[columns]
df['4000g'] = pd.to_numeric(df['4000g'], errors='coerce')
df['3500g'] = pd.to_numeric(df['3500g'], errors='coerce')
df['2600g'] = pd.to_numeric(df['2600g'], errors='coerce')
df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')

# Group by 'Source' and sum
grouped_df = data_to_extract.groupby('Source').sum().reset_index()

# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Add the custom header and make A1 bold with font size 14
ws['A1'] = 'Enva'  # Custom header above the data
ws['A1'].font = Font(bold=True, size=14)  # Make cell A1 bold with font size 14
ws['A2'] = ''  # Empty row

# Define border style
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Append the DataFrame to the worksheet starting from the fourth row
for r_idx, row in enumerate(dataframe_to_rows(grouped_df, index=False, header=True), start=3):
    for c_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=r_idx, column=c_idx, value=value)
        cell.border = thin_border  # Apply border to each cell
        cell.font = Font(bold=True, size=14)  # Make each cell bold with font size 14

# Apply borders and bold font with size 14 to header manually
for cell in ws[3]:
    cell.font = Font(bold=True, size=14)
    cell.border = thin_border

# Apply borders and bold font with size 14 to every cell manually
for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
    for cell in row:
        cell.font = Font(bold=True, size=14)
        cell.border = thin_border

# Save the workbook
output_file_path = 'output.xlsx'
wb.save(output_file_path)

print(f"Data has been saved to {output_file_path} with the desired formatting.")

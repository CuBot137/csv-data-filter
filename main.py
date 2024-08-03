import pandas as pd


file_path = 'Kilkenny tankering. 2023 (5).xlsx'

excel_data = pd.ExcelFile(file_path)

sheet_name = excel_data.sheet_names[8]
print(sheet_name)
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Border, Side
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import os.path

# Global file path variable
file_path = None

def detect_header_and_extract_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # Attempt to find the header row by looking for specific column names
    header_row = None
    for i, row in df.iterrows():
        if 'Source' in row.values:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Header row with 'Source' not found")
    
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    # Ensure that all required columns are present
    required_columns = ['Source', '4500G', '4000g', '3500g', '2600g', 'Hours']
    if not all(col in df.columns for col in required_columns):
        raise ValueError(f"One or more required columns are missing in the sheet {sheet_name}")

    return df

def extract_data_create_new_excel(file_path, sheet_name):
    try:
        df = detect_header_and_extract_data(file_path, sheet_name)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process the Excel file\n{e}")
        return

    df.rename(columns={'Location': 'Source'}, inplace=True)

    columns = ['Source', '4500G', '4000g', '3500g', '2600g', 'Hours']
    if not all(col in df.columns for col in columns):
        messagebox.showerror("Error", f"Expected columns {columns} not found in the sheet {sheet_name}")
        return
##############################################################################################
    data_to_extract = df[columns]
    # Ensure numeric conversion
    df['4500G'] = pd.to_numeric(df['4500G'], errors='coerce')
    df['4000g'] = pd.to_numeric(df['4000g'], errors='coerce')
    df['3500g'] = pd.to_numeric(df['3500g'], errors='coerce')
    df['2600g'] = pd.to_numeric(df['2600g'], errors='coerce')
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')

    # Clean 'Source' values to avoid duplicates
    df['Source'] = df['Source'].astype(str).str.strip().str.title()

    # Regenerate `data_to_extract` to include cleaned Source column
    data_to_extract = df[columns]

    grouped_df = data_to_extract.groupby('Source').sum().reset_index()
#################################################################################
    wb = Workbook()
    ws = wb.active

    # Add custom header
    ws['A1'] = 'Enva'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = ''

    # Define border style
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # Append data starting from row 3
    for r_idx, row in enumerate(dataframe_to_rows(grouped_df, index=False, header=True), start=3):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = thin_border
            cell.font = Font(bold=True, size=14)

    for cell in ws[3]:
        cell.font = Font(bold=True, size=14)
        cell.border = thin_border

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = Font(bold=True, size=14)
            cell.border = thin_border

    try:
        desktop_path = os.path.join(os.path.expanduser("~"), "Downloads")
        output_file_path = os.path.join(desktop_path, "Enva Monthly.xlsx")
        wb.save(output_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save the Excel file\n{e}")
        return

    messagebox.showinfo("Success", f"Data has been saved to:\n{output_file_path}")

def select_file():
    global file_path
    file_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        try:
            sheet_names = pd.ExcelFile(file_path).sheet_names
            sheet_name_combobox['values'] = sheet_names
            sheet_name_combobox.set('')  # Clear previous selection
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel sheet names\n{e}")

def import_file():
    global file_path
    if not file_path:
        messagebox.showwarning("Input Required", "Please select an Excel file first.")
        return

    sheet_name = sheet_name_combobox.get()
    if not sheet_name:
        messagebox.showwarning("Input Required", "Please select a sheet name.")
        return

    extract_data_create_new_excel(file_path, sheet_name)

# GUI setup
root = Tk()
root.title("Enva Monthly Data Extractor")
root.geometry("400x300")
root.configure(bg="#2e3f4f")

label_font = ("Helvetica", 12, "bold")
entry_font = ("Helvetica", 12)
button_font = ("Helvetica", 12, "bold")

frame = Frame(root, bg="#2e3f4f")
frame.pack(pady=30, padx=30, fill="both", expand=True)

sheet_name_label = Label(frame, text="Sheet Name:", font=label_font, bg="#2e3f4f", fg="#ffffff")
sheet_name_label.pack(pady=5)

sheet_name_combobox = ttk.Combobox(frame, font=entry_font, state="readonly")
sheet_name_combobox.pack(pady=10)

select_file_button = Button(frame, text="Select Excel File", font=button_font, bg="#2196f3", fg="#ffffff",
                            padx=10, pady=5, bd=0, relief="ridge", highlightthickness=0,
                            activebackground="#1976d2", cursor="hand2", command=select_file)
select_file_button.pack(pady=10)

import_button = Button(frame, text="Import File", font=button_font, bg="#4caf50", fg="#ffffff",
                       padx=10, pady=5, bd=0, relief="ridge", highlightthickness=0,
                       activebackground="#45a049", cursor="hand2", command=import_file)
import_button.pack(pady=20)

def style_button(button):
    button.config(
        borderwidth=0,
        relief="flat",
        overrelief="flat",
        highlightthickness=0
    )

style_button(select_file_button)
style_button(import_button)

root.mainloop()

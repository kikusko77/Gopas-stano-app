import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def scrape_data():
    # Open a file dialog to select the first Excel file
    file_path1 = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("OpenOffice files", "*.ods"), ("Any file", "*")))
    
    # Read the first Excel file
    df_export = pd.read_excel(file_path1)

    # Open a file dialog to select the second Excel file
    file_path2 = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"),("OpenOffice files", "*.ods"), ("Any file", "*")))
    
    # Read the second Excel file
    df_ucast = pd.read_excel(file_path2)

    # Get the list of columns from the second Excel file
    columns = list(df_ucast.columns)
    print(columns)
    # Check if the "Course sequence" column is in the second Excel file
    if "Kód kurzu" not in columns:
        messagebox.showerror("Error", "The column 'Kód kurzu' is not present in the second Excel file.")
        return

    # Create a list of matched rows
    matched_rows = []
    for i, row_export in df_export.iterrows():
       for j, row_ucast in df_ucast.iterrows():
          if row_export["Kód kurzu"] == row_ucast["Kód kurzu"]:
              matched_rows.append(row_ucast)

    # Create a new dataframe from the matched rows
    df_filtered = pd.DataFrame(matched_rows)
    
    # Open a file dialog to save the filtered data as an Excel file
    file_path3 = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    # Save the filtered data as an Excel file
    df_filtered.to_excel(file_path3, index=False)

root = tk.Tk()
root.title("Excel Scraper")
root.geometry("600x400")

button = tk.Button(root, text="Scrape Data", font=("Helvetica", 20), height=2, width=20, command=scrape_data)
button.pack(pady=100)

root.update_idletasks()
w = button.winfo_width()
h = button.winfo_height()
x = (root.winfo_width() // 2) - (w // 2)
y = (root.winfo_width() // 2) - (h // 2)
root.mainloop()
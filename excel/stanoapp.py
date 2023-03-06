import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox

import pandas as pd
import tkinter as tk
from tkinter import filedialog

def scrape_data():
    # Open a file dialog to select the first Excel file
    ucast = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")))

    # Read the first Excel file
    ucast = pd.read_excel(ucast)

    # Open a file dialog to select the second Excel file
    export = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")))

    # Read the second Excel file
    export = pd.read_excel(export)
    export.columns=['Authorization', 'ProductManager', 'CourseCode', 'CourseSequence','Alternative', 'Related']

    # Rename columns in the first Excel file
    valid_cols=['KodKurzu', 'DatumOd', 'DatumDo', 'Lektor', 'LectorOdkud',
       'Ucastnik', 'Email', 'FirmaUcastnika', 'CisloObjednavky',
       'Pobocka', 'Garance', 'Poznamka', 'Podtyp', 'Oblast',
       'ZpusobPlatby', 'Status1', 'Status2', 'Status3', 'CisloDD',
       'CisloPF', 'Mena', 'Castka']
    ucast.columns=valid_cols

    # Merge the two Excel files and extract relevant columns
    tmp=ucast.merge(export,how='left',left_on='KodKurzu',right_on='CourseCode')
    tmp=tmp[tmp.Ucastnik.notna()]
    tmp=tmp[['Ucastnik','Email','FirmaUcastnika','KodKurzu','CourseSequence','Alternative','Related']]
    tmp.drop_duplicates(inplace=True)
    
    save_all = tk.messagebox.askyesno(title="Save All Data", message="Do you want to save all scraped data?")

    if save_all:
        # Save all data to an Excel file
        file_path = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")), defaultextension='.xlsx')
        tmp.to_excel(file_path, index=False)
    else:
        # Ask the user to enter a KodKurzu to filter the data
        kod_kurzu = tk.simpledialog.askstring(title="Enter KodKurzu", prompt="Enter a KodKurzu to filter the data:")
        filtered_tmp = tmp[tmp['KodKurzu'] == kod_kurzu]
        file_path = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")), defaultextension='.xlsx')
        filtered_tmp.to_excel(file_path, index=False)
    
        


    

root = tk.Tk()
root.title("Excel Scraper")
root.geometry("600x400")

button = tk.Button(root, text="Scrape Data", font=("Helvetica", 20), height=2, width=20, command=scrape_data)
button.pack(pady=100)

root.update_idletasks()
w = button.winfo_width()
h = button.winfo_height()
x = (root.winfo_width() // 2) - (w // 2)
y = (root.winfo_height() //2) - (h // 2)

root.mainloop()


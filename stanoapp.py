import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox

def scrape_data():
    # Open a file dialog to select the first Excel file
    file_path1 = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")))
    print(file_path1)
    # Read the first Excel file
    df1 = pd.read_excel(file_path1)

    # Open a file dialog to select the second Excel file
    file_path2 = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx"), ("Excel files", "*.xlsb"), ("Any file", "*")))
    print(file_path2)
    # Read the second Excel file
    df2 = pd.read_excel(file_path2)

    # Get the list of columns from the second Excel file
    columns = list(df2.columns)

    # Ask the user to choose the column with conditions
    column_name = tk.simpledialog.askstring("Column Name", "Enter the name of the column with conditions:", parent=root, initialvalue=columns[0])

    # Check if the user entered a valid column name
    if column_name not in columns:
        messagebox.showerror("Error", "The entered column name is invalid. Please enter a valid column name.")
        return

    # Get the conditions from the first Excel file
    conditions = df1['Course Sequence'].tolist()

    # Filter the data in the second Excel file based on the conditions
    df2 = df2[df2[column_name].isin(conditions)]
    df2 = df2[df2[column_name].str.contains("->") & ~df2[column_name].str.contains("vyrazen")]

    # Get the filtered data
    filtered_data = []
    for i, row in df2.iterrows():
        condition = row[column_name]
        course_sequence = condition.split("->")[-1].strip()
        if "Vyrazen" not in course_sequence:
            filtered_data.append(row.to_dict())
            print(filtered_data)
    # Open a file dialog to save the filtered data as an Excel file
    file_path3 = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

    # Save the filtered data as an Excel file
    pd.DataFrame(filtered_data).to_excel(file_path3, index=False)

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
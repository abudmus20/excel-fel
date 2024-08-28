
#import os
#print(os.getcwd())
#import pandas as pd
# CSV dosyasını oku

#df = pd.read_excel('C:/Users/user/Desktop/ExcelFiles/New Microsoft Excel Worksheet (5).xlsx', engine='openpyxl')

#df['Time [s]']
#start_time = df[1,0]
#df['Time [s]'] = df['Time [s]'] - start_time
#İlk satırı yazdır

#print(df.head(1))

import os
import tkinter as tk
from tkinter import Listbox, Button, Entry, messagebox
import pandas as pd


def load_columns():
    try:
        file = listbox_files.get(tk.ACTIVE)
        global df  # Make df a global variable to access it in save_columns
        df = pd.read_excel(os.path.join(folder_path, file))
        listbox_columns.delete(0, tk.END)
        for col in df.columns:
            listbox_columns.insert(tk.END, col)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load columns.\n{e}")


def save_columns():
    try:
        selected_cols = [listbox_columns.get(i) for i in listbox_columns.curselection()]
        if not selected_cols:
            messagebox.showerror("Error", "No columns selected.")
            return

        # Check if df is defined
        if 'df' not in globals():
            messagebox.showerror("Error", "No file loaded.")
            return

        # Check if the first column is a time column
        first_column = df.iloc[:, 0]
        if not pd.to_datetime(first_column, errors='coerce').notna().all():
            # Generate a series of seconds if the first column is not time
            num_rows = df.shape[0]
            seconds_series = [f'{i} ' for i in range(num_rows)]  # Generate seconds series
            df.iloc[:, 0] = seconds_series  # Replace first column with seconds series

        # Save the modified file with selected columns
        save_path = os.path.join(folder_path, save_entry.get() + ".xlsx")
        df[selected_cols].to_excel(save_path, index=False)
        messagebox.showinfo("Success", "File saved successfully")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save file.\n{e}")


root = tk.Tk()
root.title("Excel Column Selector")
folder_path = "C:/Users/user/Desktop/ExcelFiles"

listbox_files = Listbox(root, width=40, height=15)
listbox_files.grid(row=0, column=0, padx=10, pady=10)
listbox_columns = Listbox(root, width=40, height=15, selectmode=tk.MULTIPLE)
listbox_columns.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Load Columns", command=load_columns).grid(row=1, column=0, padx=10, pady=10)
Button(root, text="Save Selected Columns", command=save_columns).grid(row=1, column=1, padx=10, pady=10)
save_entry = Entry(root)
save_entry.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        listbox_files.insert(tk.END, file)

root.mainloop()

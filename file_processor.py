import os
import pandas as pd
import gdown
import requests
import re
import tkinter as tk
from tkinter import filedialog

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

def download_file(url, destination):
    gdown.download(url, destination, quiet=True)

def eng_letter(n):
    if 1 <= n <= 26:
        return chr(n + 64)
    else:
        return "Invalid input"

def process_excel(file_path):
    err_cells = []
    df = pd.read_excel(file_path)
    name_column = int(entry_name_column.get())
    link_columns = [int(column.strip()) for column in entry_link_columns.get().split(',')]

    for index, row in df.iterrows():
        folder_name = str(row[name_column])
        folder_path = os.path.join(os.getcwd(), folder_name)

        create_folder(folder_path)
        for col_num in link_columns:
            attachment_url = str(row[col_num])

            if attachment_url:  # Check if the cell is not empty
                # Extract file ID from the Drive link
                file_id = attachment_url.split('id=')[1]

                # Construct direct download link
                direct_download_url = f'https://drive.google.com/uc?id={file_id}'
                url_2 = f'https://drive.usercontent.google.com/download?id={file_id}&export=download&authuser=0'
                
                response = requests.head(url_2)
                file_details = None
                if response.status_code == 200:
                    file_name = re.search(r'filename="([^"]+)"',  response.headers.get('Content-Disposition')).group(1) 
                    destination_file = os.path.join(folder_path, os.path.basename(file_name))

                    download_file(direct_download_url, destination_file)

                    result_label.config(text="Script executed successfully!'\n'")

                else:
                    cell_addr =  f'{eng_letter(col_num+1)}{index + 2}'
                    err_cells.append(cell_addr)

                if err_cells != []:
                    result_label.config(text=f"Script executed successfully!\n\nError: Invalid URL found in cells {err_cells}")

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")])
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)
        file_name = os.path.basename(file_path)
        selected_file_label.config(text=f"Selected File: {file_name}")
    else:
        selected_file_label.config(text="No file selected")

def run_script():
    try:
        excel_file_path = entry_path.get()
        name_column = int(entry_name_column.get())
        link_columns = [int(column.strip()) for column in entry_link_columns.get().split(',')]
        if excel_file_path and name_column and link_columns:
            process_excel(excel_file_path)
            # result_label.config(text="Script executed successfully!")
        else:
            result_label.config(text="Please fill all the settings.")
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("File Processor")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label_path = tk.Label(frame, text="Excel File Path:")
label_path.grid(row=0, column=0, padx=5, pady=5)

entry_path = tk.Entry(frame, width=50)
entry_path.grid(row=0, column=1, padx=5, pady=5)

button_browse = tk.Button(frame, text="Browse", command=browse_excel_file)
button_browse.grid(row=0, column=2, padx=5, pady=5)

selected_file_label = tk.Label(frame, text="Selected File: None")
selected_file_label.grid(row=1, column=0, columnspan=3, pady=5)

label_name_column = tk.Label(frame, text="Name Column Number:")
label_name_column.grid(row=2, column=0, padx=5, pady=5)

entry_name_column = tk.Entry(frame, width=15)
entry_name_column.grid(row=2, column=1, padx=5, pady=5)

label_link_columns = tk.Label(frame, text="Link Columns (Comma-separated):")
label_link_columns.grid(row=3, column=0, padx=5, pady=5)

entry_link_columns = tk.Entry(frame, width=15)
entry_link_columns.grid(row=3, column=1, padx=5, pady=5)

button_run = tk.Button(frame, text="Run Script", command=run_script)
button_run.grid(row=4, column=0, columnspan=3, pady=10)

result_label = tk.Label(frame, text="")
result_label.grid(row=5, column=0, columnspan=3)

root.mainloop()


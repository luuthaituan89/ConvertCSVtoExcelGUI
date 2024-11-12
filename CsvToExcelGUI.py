import pandas as pd
import csv
import os
from tkinter import Label, Entry, Button, filedialog, messagebox, StringVar, Checkbutton, BooleanVar, PhotoImage
from tkinter.ttk import Progressbar, Frame, Style
from ttkthemes import ThemedTk
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import threading

def select_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    csv_file_path.set(file_path)

def select_output_dir():
    dir_path = filedialog.askdirectory()
    output_dir_path.set(dir_path)

def adjust_column_widths(ws):
    for col in ws.columns:
        max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max_length + 2

def convert_csv_to_excel():
    status_label_var.set("Converting...")
    root.update_idletasks()

    file_path = csv_file_path.get()
    output_dir = output_dir_path.get()
    output_file_name = name_entry.get()
    output_file = os.path.join(output_dir, f"{output_file_name}.xlsx")
    split_into_sheets = split_checkbox_var.get()

    if not file_path or not output_dir or not output_file_name:
        messagebox.showerror("Error", "Please complete all fields.")
        return

    if not file_path.lower().endswith('.csv'):
        messagebox.showerror("Error", "The selected file is not in CSV format.")
        return

    try:
        with open(file_path, mode='r', encoding='utf-8') as file:
            dialect = csv.Sniffer().sniff(file.read(1024))
            file.seek(0)

            chunk_size = 10000
            total_lines = sum(1 for _ in open(file_path, 'r', encoding='utf-8')) - 1
            num_chunks = (total_lines // chunk_size) + 1

            progress_bar['value'] = 0
            root.update_idletasks()

            sheet_index = 1
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                if not split_into_sheets:
                    chunk = pd.read_csv(file, delimiter=dialect.delimiter)
                    chunk.to_excel(writer, index=False, sheet_name='Sheet1')
                    progress_bar['value'] = 100
                    root.update_idletasks()
                else:
                    for i, chunk in enumerate(pd.read_csv(file, delimiter=dialect.delimiter, chunksize=chunk_size)):
                        startrow = 0 if writer.sheets.get(f'Sheet{sheet_index}') is None else writer.sheets[f'Sheet{sheet_index}'].max_row

                        if startrow + len(chunk) > 1048576:
                            sheet_index += 1
                            startrow = 0

                        sheet_name = f'Sheet{sheet_index}'
                        chunk.to_excel(writer, index=False, header=writer.sheets.get(sheet_name) is None, startrow=startrow, sheet_name=sheet_name)
                        progress_bar['value'] = (i + 1) * 50 / num_chunks
                        root.update_idletasks()

            status_label_var.set("Adjusting columns...")
            root.update_idletasks()

            wb = load_workbook(output_file)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                adjust_column_widths(ws)
                progress_bar['value'] += 50 / len(wb.sheetnames)
                root.update_idletasks()

            wb.save(output_file)
            status_label_var.set("Completed")
            progress_bar['value'] = 100
            messagebox.showinfo("Success", f"Conversion complete. File saved as {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label_var.set("Error")
        progress_bar['value'] = 0

def start_conversion_thread():
    threading.Thread(target=convert_csv_to_excel, daemon=True).start()

def on_closing():
    root.quit()
    root.destroy()

# Setup GUI
root = ThemedTk(theme="winnative")
root.title("CSV to Excel Converter")

# Set kích thước và vị trí cửa sổ
window_width = 650
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")
root.resizable(False, False)

# Icon cho ứng dụng
try:
    root.iconphoto(False, PhotoImage(file="icon.png"))
except:
    pass  # Nếu không có icon, bỏ qua lỗi

# Style hiện đại
style = Style()
style.configure("TButton", font=("Arial", 12))
style.configure("TLabel", font=("Arial", 10))
style.configure("TEntry", font=("Arial", 10))

# Variables
csv_file_path = StringVar()
output_dir_path = StringVar()
status_label_var = StringVar()
split_checkbox_var = BooleanVar()

# Frames để tổ chức các thành phần
main_frame = Frame(root, padding=(20, 20))
main_frame.pack(fill="both", expand=True)

# Các thành phần giao diện
Label(main_frame, text="Select CSV file:").grid(row=0, column=0, sticky="w", pady=5)
csv_entry = Entry(main_frame, textvariable=csv_file_path, width=50)
csv_entry.grid(row=0, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_csv_file).grid(row=0, column=2, padx=10)

Label(main_frame, text="Select output directory:").grid(row=1, column=0, sticky="w", pady=5)
dir_entry = Entry(main_frame, textvariable=output_dir_path, width=50)
dir_entry.grid(row=1, column=1, padx=10, pady=5)
Button(main_frame, text="Browse", command=select_output_dir).grid(row=1, column=2, padx=10)

Label(main_frame, text="Enter Excel file name (without extension):").grid(row=2, column=0, sticky="w", pady=5)
name_entry = Entry(main_frame, width=50)
name_entry.grid(row=2, column=1, padx=10, pady=5)

split_checkbox = Checkbutton(main_frame, text="Split CSV into multiple sheets (if needed)", variable=split_checkbox_var)
split_checkbox.grid(row=3, column=1, pady=10, sticky="w")

Button(main_frame, text="Convert to Excel", command=start_conversion_thread).grid(row=4, column=1, pady=20)

status_label = Label(main_frame, textvariable=status_label_var, font=("Arial", 10))
status_label.grid(row=5, column=1, pady=5)

progress_bar = Progressbar(main_frame, orient="horizontal", length=300, mode="determinate")
progress_bar.grid(row=6, column=1, pady=10)

root.protocol("WM_DELETE_WINDOW", on_closing)

# Run ứng dụng
root.mainloop()

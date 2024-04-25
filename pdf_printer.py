import tkinter as tk
from tkinter import filedialog, messagebox
import os
import win32print
import win32api

def print_pdf(pdf_file, printer_name):
    try:
        win32api.ShellExecute(0, "print", pdf_file, f'/d:"{printer_name}"', None, 0)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to print {pdf_file}. Error: {str(e)}")

def print_files():
    folder = folder_path.get()
    printer = printer_var.get()
    if not folder or not printer:
        messagebox.showwarning("Warning", "Please select a folder and a printer.")
        return
    
    for f in os.listdir(folder):
        if f.lower().endswith('.pdf'):
            print_pdf(os.path.join(folder, f), printer)
    messagebox.showinfo("Success", "All files have been sent to the printer.")

def browse_folder():
    folder = filedialog.askdirectory()
    folder_path.set(folder)

app = tk.Tk()
app.title("PDF Batch Printer")

folder_path = tk.StringVar()
printer_var = tk.StringVar()

tk.Label(app, text="Select Folder:").pack()
folder_entry = tk.Entry(app, textvariable=folder_path, width=50)
folder_entry.pack()
browse_button = tk.Button(app, text="Browse", command=browse_folder)
browse_button.pack()

tk.Label(app, text="Select Printer:").pack()
all_printers = [printer[2] for printer in win32print.EnumPrinters(2)]
printer_var.set(all_printers[0])  # default to the first printer
printer_dropdown = tk.OptionMenu(app, printer_var, *all_printers)
printer_dropdown.pack()

print_button = tk.Button(app, text="Print PDFs", command=print_files)
print_button.pack()

app.mainloop()

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import customtkinter as ctk
from customtkinter import *
from docx2pdf import convert
from PIL import Image
import win32com.client as win32
import os
import time
import threading

def select_directory():
    directory_path = filedialog.askdirectory()
    if directory_path:
        treeview.delete(*treeview.get_children())
        directory_entry.configure(state='normal')
        directory_entry.delete(0, tk.END)
        directory_entry.insert(0, directory_path)
        directory_entry.configure(state='readonly')

        populate_treeview(treeview, directory_path)

def populate_treeview(treeView, main_directory, parent=""):
    for file in os.listdir(main_directory):
        full_path = os.path.join(main_directory, file)

        if os.path.isdir(full_path):
            file_directory = treeView.insert(parent, "end", text=os.path.basename(full_path), open=False)
            populate_treeview(treeView, full_path, file_directory)

        elif full_path.endswith(".docx") or full_path.endswith(".xlsx") or full_path.endswith(".xlsm") or full_path.endswith(".xls"):
            treeView.insert(parent, "end", text=file)

    parent_directory = os.path.basename(os.path.dirname(main_directory))
    treeView.heading("#0", text=parent_directory)

def get_excel_and_word_files(directory):
    files_list = []

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".docx") or file.endswith(".xlsx") or file.endswith(".xlsm") or file.endswith(".xls"):
                files_list.append(os.path.join(root, file))

    return files_list

def convert_docs_to_pdf(file_path):
    input_file = r"{}".format(file_path.replace('/', '\\')) # Gets path of file and corrects mixed '/' and '\' with only '\'
    output_file = r"{}.pdf".format(os.path.splitext(input_file)[0]) # Uses file path and replaces Office extension (.xlsx/.docx) with .pdf

    if file_path.endswith(".docx"):
        convert(input_file, output_file) # This is done with the "convert" method from "docx2pdf" library

    elif file_path.endswith('.xlsx') or file_path.endswith('.xls') or file_path.endswith('.xlsm'):
        convert_xlsx_to_pdf(input_file, output_file)

# This function uses the library win32com for exporting the Excel file to a PDF, using Windows's COM objects
def convert_xlsx_to_pdf(input_file, output_file):
    excel = win32.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(input_file)
    worksheet = workbook.Worksheets[0]
    worksheet.ExportAsFixedFormat(0, output_file)
    workbook.Close(False)
    excel.Quit()

def convert_docs(root, files_list):
    progress_popup = ctk.CTkToplevel(root)
    progress_popup.title("Progresso da Conversão")
    progress_popup.configure(bg_color='#FFF', fg_color='#FFF')
    progress_popup.attributes("-toolwindow", 1) # Removing windows buttons
    progress_popup.resizable(False, False)
    progress_popup.geometry("300x100")
    progress_label = ctk.CTkLabel(progress_popup, text="Progresso da Conversão:", font=("Fira Sans", 16), fg_color='#FFF')
    progress_label.pack(pady=5)
    progress_bar = ctk.CTkProgressBar(progress_popup, orientation="horizontal", height=25, mode="determinate", fg_color='#F5EEE6', progress_color='#756AB6')
    progress_bar.set(0)
    progress_bar.pack(pady=(0,5))
    status_label = ctk.CTkLabel(progress_popup, text="0 / {}".format(X), font=("Fira Sans", 16), fg_color='#FFF')
    status_label.pack(pady=(0,5))
    center_window(progress_popup, 300, 100)
    progress_popup.lift()

    totalFiles = len(files_list)
    def convert_and_update_ui(current_file):
        progress_popup.lift()

        if current_file < totalFiles:
            file = files_list[current_file]
            convert_docs_to_pdf(file)
            totalProgress = (current_file + 1) / totalFiles # Adding 1 because the index of the list starts at 0, wich would make the progress bar 1 step behind
            progress_bar.set(totalProgress)
            status_label.configure(text="{:d} / {:d}".format(current_file + 1, totalFiles))
            root.after(100, convert_and_update_ui, current_file + 1)
        else:
            time.sleep(1)
            progress_popup.destroy()

    convert_and_update_ui(0) # Starting the function at current_file = 0

def main_function():
    main_directory = directory_entry.get()
    files_list = get_excel_and_word_files(main_directory)
    convert_docs(root, files_list)

# CREATING INTERFACE WITH CUSTOMTKINTER

# Interface functions
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (width / 2))
    y_coordinate = int((screen_height / 2) - (height / 2))
    window.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

# Interface creation
root = ctk.CTk()
root.title("Fábrica de PDFs")
root.config(bg='#FFF')
window_width = 600
window_height = 700
root.resizable(False, False)
center_window(root, 600, 700)

title_label = ctk.CTkLabel(root, text="Fábrica de PDFs", font=("Fira Sans", 24), text_color='#FFF', fg_color="#756AB6", width=600, height=60)
title_label.pack()

directory_frame = ctk.CTkFrame(root)
directory_frame.configure(fg_color='#FFF')
directory_frame.pack(fill="x", padx=10, pady=(5,10))

directory_label = ctk.CTkLabel(directory_frame, text="Selecione o diretório principal:", font=("Fira Sans", 16))
directory_label.grid(row=0, column=0, sticky='w')

directory_entry = ctk.CTkEntry(directory_frame, font=("Fira Sans", 16), width=531, height=42, border_width=0, fg_color='#F5EEE6', state='readonly')
directory_entry.grid(row=1, column=0)

def on_enter(e):
    openedFolder = Image.open('C:\Teste\Opened Folder.png')
    select_button.configure(fg_color='#62518A', image=CTkImage(openedFolder))

folder = Image.open('C:\Teste\Folder.png')
def on_leave(e):
    select_button.configure(fg_color='#756AB6', image=CTkImage(folder))

select_button = ctk.CTkButton(directory_frame, text='', font=("Fira Sans", 16), command=select_directory, 
                              width=42, height=42, image=CTkImage(folder), fg_color='#756AB6', border_width=0)
select_button.grid(row=1, column=2, padx=6)
select_button.bind("<Enter>", on_enter)
select_button.bind("<Leave>", on_leave)

# Custom TKinter TreeView Style
treeViewStyle = ttk.Style()
treeViewStyle.theme_use("default")
treeViewStyle.configure("Treeview", background="#F5EEE6", font=("Fira Sans", 12), foreground="black", borderwidth=0, rowheight=25, fieldbackground="#F5EEE6")
treeViewStyle.configure("Treeview.Heading", background="#F5EEE6", foreground="black", borderwidth=0, font=("Fira Sans", 12))
treeViewStyle.map("Treeview", background=[('selected', '#756AB6')])

treeview = ttk.Treeview(root)
treeview.column("#0")
treeview.pack(expand=True, fill="both", padx=10)

function_button = ctk.CTkButton(root, text="Fabricar PDFs", font=("Fira Sans", 18), height=42, fg_color='#756AB6', hover_color='#62518A', command=main_function)
function_button.pack(pady=10)

root.mainloop()
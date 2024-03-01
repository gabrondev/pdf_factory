'''
File name: pdf_factory.py
Author: Gabriel Almeida Oliveira
LinkedIn: https://www.linkedin.com/in/gabrondev/
Creation date: 13/02/2024
Update date: 17/02/2024
Description: Simple app for converting specific file formats to PDF format
'''

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import customtkinter as ctk
from customtkinter import *
from PIL import Image, ImageTk
from win32com import client as win32
import os
import threading
import pygetwindow as gw
import json
import time

# GLOBAL VARIABLES
# These variables are global so that they can be used on differents threads
# Widgets:
progress_popup = None
progress_bar = None
status_label = None
file_details_text = None
msgbox_yes_no = None
# Threads:
convert_thread = None
update_progress_ui_thread = None
cancel_event = threading.Event()
pause_event = threading.Event()
done_event = threading.Event()
# General variables:
files_list = []
current_file = 0
total_files = 0

# APP SETTINGS
def load_settings():
    try:
        with open('resources/settings.json', 'r') as file:
            return json.load(file)
    except FileNotFoundError:
        return {}
    
settings = load_settings()
docx_enabled = settings['docx_enabled']
doc_enabled = settings['doc_enabled']
xlsx_enabled = settings['xlsx_enabled']
xls_enabled = settings['xls_enabled']
xlsm_enabled = settings['xlsm_enabled']
open_folder_enabled = settings['open_folder_after_conversion']

def save_settings(settings):
    with open('resources/settings.json', 'w') as file:
        json.dump(settings, file)

def update_settings(docx_checkbox_state, doc_checkbox_state, xlsx_checkbox_state, 
                    xls_checkbox_state, open_folder_switch_state, xlsm_checkbox_state):
    settings['docx_enabled'] = docx_checkbox_state
    settings['doc_enabled'] = doc_checkbox_state
    settings['xlsx_enabled'] = xlsx_checkbox_state
    settings['xls_enabled'] = xls_checkbox_state
    settings['xlsm_enabled'] = xlsm_checkbox_state
    settings['open_folder_after_conversion'] = open_folder_switch_state
    save_settings(settings)

def configure_settings_popup(docx_checkbox, doc_checkbox, xlsx_checkbox, 
                             xls_checkbox, xlsm_checkbox, open_folder_switch):
    if docx_enabled:
        docx_checkbox.select()
    if doc_enabled:
        doc_checkbox.select()
    if xlsx_enabled:
        xlsx_checkbox.select()
    if xls_enabled:
        xls_checkbox.select()
    if xlsm_enabled:
        xlsm_checkbox.select()
    if open_folder_enabled:
        open_folder_switch.select()

# APP LOGIC
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
    files_list = get_excel_and_word_files(main_directory)
    
    if main_directory != "" and len(files_list) > 0:
        for file in os.listdir(main_directory):
            full_path = os.path.join(main_directory, file)

            if os.path.isdir(full_path):
                file_directory = treeView.insert(parent, "end", text=os.path.basename(full_path), open=False, tags="bold")
                populate_treeview(treeView, full_path, file_directory)
            elif file.startswith('~$'):
                pass
            elif docx_enabled and full_path.endswith(".docx"):
                treeView.insert(parent, "end", text=file)
            elif doc_enabled and full_path.endswith(".doc"):
                treeView.insert(parent, "end", text=file)
            elif xlsx_enabled and full_path.endswith(".xlsx"):
                treeView.insert(parent, "end", text=file)
            elif xls_enabled and full_path.endswith(".xls"):
                treeView.insert(parent, "end", text=file)
            elif xlsm_enabled and full_path.endswith(".xlsm"):
                treeView.insert(parent, "end", text=file)

        parent_directory = os.path.basename(os.path.dirname(main_directory))
        heading_text = "Diretório Principal: " + '"' + parent_directory + '"'
        treeView.heading("#0", text=heading_text)

def get_excel_and_word_files(directory):
    files_list = []
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.startswith('~$'):
                pass
            elif docx_enabled and file.endswith(".docx"): 
                files_list.append(os.path.join(root, file))
            elif doc_enabled and file.endswith(".doc"): 
                files_list.append(os.path.join(root, file))
            elif xlsx_enabled and file.endswith(".xlsx"):
                files_list.append(os.path.join(root, file))
            elif xls_enabled and file.endswith(".xls"):
                files_list.append(os.path.join(root, file))
            elif xlsm_enabled and file.endswith(".xlsm"):
                files_list.append(os.path.join(root, file))

    return files_list

# FILE CONVERSION FUNCTIONS
def convert_docs(files_list):
    global current_file, total_files, excel_application, word_application, file_details_text, open_folder_enabled

    time.sleep(2) # Wait time so the progress UI can be fully created before starting

    has_word_files = any(file_path.endswith('.docx') for file_path in files_list)
    has_excel_files = any(file_path.endswith('.xlsx') or file_path.endswith('.xlsm') or file_path.endswith('.xls') for file_path in files_list)
    
    if has_excel_files:
        excel_application = win32.Dispatch("Excel.Application")

    if has_word_files:
        word_application = win32.Dispatch("Word.Application")

    current_file = 1
    for file in files_list:
        if pause_event.is_set():
            while pause_event.is_set(): time.sleep(0.1)
        if cancel_event.is_set():
            progress_popup.destroy()
            break
        else:
            input_file = "{}".format(file.replace('/', '\\')) # Gets path of file and corrects mixed '/' and '\' with only '\'
            output_file = "{}.pdf".format(os.path.splitext(input_file)[0]) # Uses file path and replaces Office extension (.xlsx/.docx) with .pdf

            if file.endswith(".docx") or file.endswith(".doc"):
                convert_docx_to_pdf(input_file, output_file)
            elif file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.xlsm'):
                convert_xlsx_to_pdf(input_file, output_file)

            current_file += 1

    done_event.set()
    
    if open_folder_enabled:
        main_directory = str(directory_entry.get())
        os.startfile(main_directory)

def convert_xlsx_to_pdf(input_file, output_file):
    global file_details_text

    file_name = os.path.basename(input_file)
    file_details_text.configure(state='normal')
    file_details_text.insert("end", "\n" + file_name + " - INICIADO")
    file_details_text.configure(state='disabled')
    file_details_text.see("end")

    if is_file_open(input_file):
        file_details_text.configure(state='normal')
        file_details_text.insert("end", "\n" + file_name + " - ERRO: ARQUIVO JÁ ABERTO", "error")
        file_details_text.configure(state='disabled')
        file_details_text.see("end")
    else:
        try:
            workbook = excel_application.Workbooks.Open(input_file)
            worksheet = workbook.Worksheets[0]
            worksheet.ExportAsFixedFormat(0, output_file)
            if workbook is not None: workbook.Close(False)

            file_details_text.configure(state='normal')
            file_details_text.insert("end", "\n" + file_name + " - OK", "sucess")
            file_details_text.configure(state='disabled')
            file_details_text.see("end")
        except Exception as e:
            file_details_text.configure(state='normal')
            file_details_text.insert("end", "\n" + file_name + " - ERRO", "error")
            file_details_text.configure(state='disabled')
            file_details_text.see("end")

def convert_docx_to_pdf(input_file, output_file):
    global file_details_text

    file_name = os.path.basename(input_file)
    file_details_text.configure(state='normal')
    file_details_text.insert("end", "\n" + file_name + " - INICIADO")
    file_details_text.configure(state='disabled')
    file_details_text.see("end")
    
    if is_file_open(input_file):
        file_details_text.configure(state='normal')
        file_details_text.insert("end", "\n" + file_name + " - ERRO: ARQUIVO JÁ ABERTO", "error")
        file_details_text.configure(state='disabled')
        file_details_text.see("end")
    else:
        try:
            document = word_application.Documents.Open(input_file)
            document.ExportAsFixedFormat(output_file, 17)
            if document is not None: document.Close(False)

            file_details_text.configure(state='normal')
            file_details_text.insert("end", "\n" + file_name + " - OK", "sucess")
            file_details_text.configure(state='disabled')
            file_details_text.see("end")
        except Exception as e:
            file_details_text.configure(state='normal')
            file_details_text.insert("end", "\n" + file_name + " - ERRO", "error")
            file_details_text.configure(state='disabled')
            file_details_text.see("end")

def is_file_open(filename):
    filename = str(os.path.basename(filename))
    windows = gw.getWindowsWithTitle(filename)
    return len(windows) > 0

def update_progress_ui():
    global current_file, total_files, progress_popup, progress_bar, status_label, file_details_text, files_list

    def toggle_file_details():
        if file_details_frame.winfo_ismapped():
            center_popup(progress_popup, 300, 150)
            file_details_frame.pack_forget()
        else:
            center_popup(progress_popup, 300, 250)
            file_details_frame.pack(fill="both", expand=True, padx=10, pady=5)

    def cancel_conversion():
        pause_event.clear()
        cancel_event.set()
        progress_popup.destroy()

    def keep_conversion():
        progress_popup.grab_set()
        progress_popup.lift()
        pause_event.clear()
        root.update_idletasks()

    def stop_conversion_confirmation():
        def on_close():
            progress_popup.grab_set()
            progress_popup.lift()
            pause_event.clear()

        if not done_event.is_set():
            pause_event.set()
            message_box_yes_no("Cancelar Conversão", "Tem certeza que deseja cancelar a conversão?", cancel_conversion, keep_conversion)
        else:
            progress_popup.destroy()

        if pause_event.is_set():
            grabbed_widget = root.grab_current()
            grabbed_widget.bind("<Destroy>", lambda event: on_close())

    total_files = len(files_list)
    def update_progress():
        if not done_event.is_set() or current_file <= total_files:
            totalProgress = current_file / total_files if current_file > 0 else 0
            progress_bar.set(totalProgress)
            status_label.configure(text="Total: {:d} / {:d}".format(current_file, total_files))
            progress_popup.after(100, update_progress)
        else:
            center_popup(progress_popup, 300, 250)
            progress_bar.destroy()
            status_label.destroy()
            toggle_button.destroy()
            progress_bar_frame.destroy()
            cancel_button.destroy()

            file_details_text.see("end")
            file_details_text.configure(state='normal')
            file_details_text.insert("end", "\n" + "Processo finalizado!")
            file_details_text.configure(state='disabled')
            popup_title_label.configure(text="Conversões concluídas!")
            file_details_frame.pack(fill="both", expand=True, padx=10, pady=5)

    #PROGRESS POPUP CREATION
    progress_popup = ctk.CTkToplevel(root)
    center_popup(progress_popup, 300, 150)
    progress_popup.attributes("-toolwindow", 1)
    progress_popup.title("Progresso da Conversão")
    progress_popup.configure(bg_color='#FFF', fg_color='#FFF')
    progress_popup.resizable(False, False)
    
    popup_title_label = ctk.CTkLabel(progress_popup, text="Progresso da Conversão", font=normal, fg_color='#FFF')
    popup_title_label.pack(pady=5)

    # PROGRESS SECTION WIDGETS
    progress_bar_frame = ctk.CTkFrame(progress_popup)
    progress_bar_frame.configure(fg_color='#FFF')
    progress_bar_frame.pack(pady=(0,5))
    progress_bar = ctk.CTkProgressBar(progress_bar_frame, orientation="horizontal", height=25, 
                                      mode="determinate", fg_color='#F1F1F1', progress_color='#E70000')
    progress_bar.set(0)
    progress_bar.grid(row=0, column=0, padx=(0, 5))
    details_icon = Image.open('resources/images/terminal_icon.png')
    toggle_button = ctk.CTkButton(progress_bar_frame, text="", command=toggle_file_details,
                                  fg_color='#E70000', hover_color='#C50000',
                                  width=32, height=32, image=CTkImage(details_icon))
    toggle_button.grid(row=0, column=1)
    status_label = ctk.CTkLabel(progress_popup, text="0 / {}".format(total_files), font=normal, fg_color='#FFF')
    status_label.pack(pady=(0,5))

    # DETAILS SECTION WIDGETS
    file_details_frame = ctk.CTkFrame(progress_popup)
    file_details_frame.configure(bg_color='#FFF', fg_color='#FFF')
    file_details_frame.pack_propagate(False)

    file_details_text = ctk.CTkTextbox(file_details_frame, font=log, text_color="#FFF", wrap="none", 
                                       height=5, bg_color='#FFF', fg_color='#000')
    file_details_text.pack(fill="both", expand=True)
    file_details_text.tag_config("sucess", foreground="green")
    file_details_text.tag_config("error", foreground="red")
    file_details_text.insert("0.0", "Processo iniciado...")
    file_details_text.configure(state='disabled')
    cancel_button = ctk.CTkButton(progress_popup, text="Cancelar Conversão", font=button, text_color="#FFF", command=stop_conversion_confirmation,
                                  fg_color='#E70000', bg_color='#FFF',  hover_color='#C50000')
    cancel_button.pack(pady=5)

    progress_popup.protocol("WM_DELETE_WINDOW", stop_conversion_confirmation)

    progress_popup.grab_set()
    progress_popup.lift()
    update_progress()

def reset_stats():
    global progress_popup, progress_bar, status_label, convert_thread, update_progress_ui_thread, cancel_event, files_list, current_file, total_files 

    progress_popup = None
    progress_bar = None
    status_label = None
    convert_thread = None
    update_progress_ui_thread = None
    cancel_event = threading.Event()
    files_list = []
    current_file = 0
    total_files = 0

def main_function():
    global convert_thread, update_progress_ui_thread, files_list

    reset_stats()
    done_event.clear()

    main_directory = directory_entry.get()
    files_list = get_excel_and_word_files(main_directory)

    def start_conversion():
        convert_thread = threading.Thread(target=convert_docs, args=(files_list,))
        update_progress_ui_thread = threading.Thread(target=update_progress_ui)
        update_progress_ui_thread.start()
        convert_thread.start()

    def ask_to_start_conversion():
        if len(files_list) > 0:
            message_box_yes_no("Iniciar Conversão", "Deseja iniciar a conversão dos arquivos?", start_conversion)
        else:
            message_box_ok_only("Nenhum arquivo encontrado", "Nenhum arquivo encontrado para converter.")
            treeview.delete(*treeview.get_children())
            treeview.heading("#0", text='')
            directory_entry.configure(state='normal')
            directory_entry.delete(0, tk.END)
            directory_entry.configure(state='readonly')
    
    ask_to_start_conversion()

# CREATING INTERFACE WITH CUSTOMTKINTER
ctk.set_appearance_mode('light')

# Interface functions
def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (width / 2))
    y_coordinate = int((screen_height / 2) - (height / 2))
    window.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

def center_popup(window, width, height):
    root_width = root.winfo_width()
    root_height = root.winfo_height()
    root_x = root.winfo_rootx()
    root_y = root.winfo_rooty()

    x_coordinate = root_x + (root_width // 2) - (width // 2)
    y_coordinate = root_y + (root_height // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x_coordinate}+{y_coordinate}")

def message_box_ok_only(title, message):
    msgbox_ok_only = ctk.CTkToplevel(root)
    msgbox_ok_only.resizable(False, False)
    msgbox_ok_only.attributes("-toolwindow", 1)
    msgbox_ok_only.title(title)
    msgbox_ok_only.configure(bg='#FFF')
    msgbox_ok_only.config(bg='#FFF')
    msgbox_ok_only.lift()
    msgbox_ok_only.grab_set()
    center_popup(msgbox_ok_only, 250, 150)

    message_label = ctk.CTkLabel(msgbox_ok_only, text=message, font=normal, fg_color='#FFF', wraplength=220)
    message_label.pack(pady=(30, 20))

    ok_button = ctk.CTkButton(msgbox_ok_only, text="OK", font=button, text_color="#FFF", command=msgbox_ok_only.destroy,
                               fg_color='#E70000', hover_color='#C50000', bg_color='#FFF')
    ok_button.pack(pady=(0,5))

def message_box_yes_no(title, message, callback_yes, callback_no=None):
    global msgbox_yes_no

    def on_yes():
        callback_yes()
        msgbox_yes_no.destroy()

    def on_no():
        if callback_no == None: 
            pass 
        else:
            callback_no()
        msgbox_yes_no.destroy()

    msgbox_yes_no = ctk.CTkToplevel(root)
    msgbox_yes_no.attributes("-toolwindow", 1)
    msgbox_yes_no.title(title)
    msgbox_yes_no.config(bg='#FFF')
    msgbox_yes_no.resizable(False, False)
    center_popup(msgbox_yes_no, 250, 150)

    message_label = ctk.CTkLabel(msgbox_yes_no, text=message, font=normal, fg_color='#FFF', wraplength=220)
    message_label.pack(pady=(30, 20))

    buttons_frame = ctk.CTkFrame(msgbox_yes_no, bg_color='#FFF', fg_color='#FFF')
    buttons_frame.pack(pady=(0,5))

    yes_button = ctk.CTkButton(buttons_frame, text="Sim", font=button, text_color="#FFF", command=on_yes,
                               fg_color='#E70000', hover_color='#C50000', bg_color='#FFF', width=100)
    yes_button.grid(row=0, column=0, padx=(0,5), pady=(0,5))

    no_button = ctk.CTkButton(buttons_frame, text="Não", font=button, text_color="#FFF", command=on_no,
                               fg_color='#E70000', hover_color='#C50000', bg_color='#FFF', width=100)
    no_button.grid(row=0, column=1, padx=(0,5), pady=(0,5))

    msgbox_yes_no.grab_set()
    msgbox_yes_no.lift()

def open_help_popup():
    with open('resources/help.txt', 'r', encoding='utf-8') as file:
        help_text = file.read()

    help_popup = ctk.CTkToplevel(root)
    help_popup.resizable(False, False)
    help_popup.attributes("-toolwindow", 1)
    help_popup.configure(fg_color='#FFF')
    help_popup.title("Ajuda")
    
    tabview = ctk.CTkTabview(help_popup)
    tabview.configure(segmented_button_selected_color='#E70000',
                      segmented_button_selected_hover_color='#C50000', 
                      segmented_button_unselected_hover_color='#C50000')
    tabview.pack(padx=10, pady=(0,10), fill='both', expand=True)
    
    help_tab = tabview.add('Ajuda')
    tabview._segmented_button.configure(font=normal, text_color='#FFF')
    
    help_textbox = ctk.CTkTextbox(help_tab, wrap='word', font=text, width=280, fg_color='#F1F1F1')
    help_textbox.tag_config("center", justify="center")
    help_textbox.insert('end', help_text, "center")
    help_textbox.configure(state='disabled')
    help_textbox.pack(padx=10, pady=10, fill='both', expand=True)

    credits_tab = tabview.add('Créditos')
    app_name_label = ctk.CTkLabel(credits_tab, text='Fábrica de PDFs', font=title)
    app_name_label.pack()
    app_version_label = ctk.CTkLabel(credits_tab, text='Versão: v1.0', font=subtitle)
    app_version_label.pack()

    with open('resources/credits.txt', 'r', encoding='utf-8') as file:
        credits_text = file.read()
    credits_textbox = ctk.CTkTextbox(credits_tab, wrap='word', font=text, width=280, fg_color='#F1F1F1')
    credits_textbox.tag_config("center", justify="center")
    credits_textbox.insert('end', credits_text, "center")
    credits_textbox.configure(state='disabled')
    credits_textbox.pack(padx=10, pady=10, fill='both', expand=True)

    # Going to the end of textbox, then back to the start, 
    # so the scrollbar works properly when dragged
    help_textbox.see('end')
    help_textbox.see('1.0')

    center_popup(help_popup, 400, 350)
    help_popup.lift()
    help_popup.grab_set()

def open_settings_popup():
    settings_popup = ctk.CTkToplevel(root)
    settings_popup.config(bg='#FFF')
    settings_popup.attributes("-toolwindow", 1)
    settings_popup.title("Configurações")
    settings_popup.configure(bg='#FFF')
    settings_popup.resizable(False, False)
    settings_popup.lift()
    settings_popup.grab_set()
    center_popup(settings_popup, 400, 300)

    settings_frame = ctk.CTkFrame(settings_popup, corner_radius=10)
    settings_frame.configure(fg_color="#F1F1F1", bg_color='#FFF')
    settings_frame.pack(anchor='w', pady=8, padx=8, fill='both', expand=True)

    # FORMATS SETTINGS
    formats_label = ctk.CTkLabel(settings_frame, text='Formatos:', font=subtitle, fg_color='#F1F1F1', width=80, anchor='e')
    formats_label.grid(row=0, column=0, pady=(10,0), padx=(4,4))

    docx_checkbox = CTkCheckBox(settings_frame, text=".docx", font=normal, fg_color='#E70000')
    docx_checkbox.configure(hover_color='#E70000')
    docx_checkbox.grid(row=0, column=1, pady=(10, 0), sticky='w')

    doc_checkbox = CTkCheckBox(settings_frame, text=".doc", font=normal, fg_color='#E70000')
    doc_checkbox.configure(hover_color='#E70000')
    doc_checkbox.grid(row=1, column=1, sticky='w', pady=(5,0))

    xlsx_checkbox = CTkCheckBox(settings_frame, text=".xlsx", font=normal, fg_color='#E70000')
    xlsx_checkbox.configure(hover_color='#E70000')
    xlsx_checkbox.grid(row=2, column=1, sticky='w', pady=(5,0))

    xls_checkbox = CTkCheckBox(settings_frame, text=".xls", font=normal, fg_color='#E70000')
    xls_checkbox.configure(hover_color='#E70000')
    xls_checkbox.grid(row=0, column=2, pady=(10,0))

    xlsm_checkbox = CTkCheckBox(settings_frame, text=".xlsm", font=normal, fg_color='#E70000')
    xlsm_checkbox.configure(hover_color='#E70000')
    xlsm_checkbox.grid(row=1, column=2, pady=(5,0))

    dwg_checkbox = CTkCheckBox(settings_frame, text=".dwg", font=normal, state='disabled', border_color='#9999A3')
    dwg_checkbox.grid(row=2, column=2, pady=(5,0))

    # GENERAL SETTINGS
    geral_frame = CTkFrame(settings_frame, corner_radius=0)
    geral_frame.configure(fg_color="#F1F1F1", bg_color='#FFF')
    geral_frame.grid(row=3, column=0, columnspan=4, pady=(20,0), sticky='w')

    general_settings_label = CTkLabel(geral_frame, text='Geral:', font=subtitle, fg_color='#F1F1F1', width=80, anchor='e')
    general_settings_label.grid(row=0, column=0, padx=(4,4))

    open_folder_switch = CTkSwitch(geral_frame, text="Abrir pasta após finalizar", font=normal)
    open_folder_switch.configure(progress_color='#E70000')
    open_folder_switch.grid(row=0, column=1, sticky='w', padx=0)

    subdirectories_switch = CTkSwitch(geral_frame, text="Converter arquivos em subpastas", font=normal)
    subdirectories_switch.select()
    subdirectories_switch.configure(progress_color='#939BA2', state='disabled')
    subdirectories_switch.grid(row=1, column=1, sticky='w', padx=0)

    subdirectories_switch = CTkSwitch(geral_frame, text="Ignorar se arquivo já tiver PDF", font=normal)
    subdirectories_switch.configure(state='disabled')
    subdirectories_switch.grid(row=2, column=1, sticky='w', padx=0)

    settings = load_settings()
    configure_settings_popup(docx_checkbox, doc_checkbox, xlsx_checkbox, xls_checkbox, xlsm_checkbox, open_folder_switch)

    def on_popup_close():
        global docx_enabled, doc_enabled, xlsx_enabled, xls_enabled, xlsm_enabled, open_folder_enabled

        docx_enabled = docx_checkbox.get() == 1
        doc_enabled = doc_checkbox.get() == 1
        xlsx_enabled = xlsx_checkbox.get() == 1
        xls_enabled = xls_checkbox.get() == 1
        xlsm_enabled = xlsm_checkbox.get() == 1
        open_folder_enabled = open_folder_switch.get() == 1

        update_settings(docx_enabled, doc_enabled, xlsx_enabled, 
                        xls_enabled, xlsm_enabled, open_folder_enabled)

        treeview.delete(*treeview.get_children())
        populate_treeview(treeview, directory_entry.get())
        settings_popup.destroy()
        
    settings_popup.protocol("WM_DELETE_WINDOW", on_popup_close)

# MAIN WINDOW
root = ctk.CTk()
root.iconbitmap('resources/images/pdf_factory_icon.ico')
root.title("Fábrica de PDFs")
root.config(bg='#FFF')
root.resizable(False, False)
center_window(root, 600, 700)

# CUSTOM FONTS
ctk.FontManager.load_font("resources/fonts/Fira Sans Bold.ttf")
ctk.FontManager.load_font("resources/fonts/Fira Sans Regular.ttf")
title = ctk.CTkFont(family="Fira Sans Condensed", size=24, weight="bold")
subtitle = ctk.CTkFont(family="Fira Sans Condensed", size=18, weight="bold")
normal = ctk.CTkFont(family="Fira Sans Condensed", size=16, weight="bold")
text = ctk.CTkFont(family="Fira Sans Condensed", size=16)
button = ctk.CTkFont(family="Fira Sans Condensed", size=18, weight="bold")
log = ctk.CTkFont(family="Fira Sans Condensed", size=14)

# TITLE SECTION
background_image = ImageTk.PhotoImage(Image.open('resources/images/title_background.png'))
title_label = ctk.CTkLabel(root, text="Fábrica de PDFs", font=title, text_color='#FFF',
                           width=600, height=60, image=background_image)
title_label.pack()

help_button = ctk.CTkButton(root, text='?', font=button, text_color="#000", command=open_help_popup,
                            width=32, height=32, border_width=0,
                            fg_color='#F1F1F1', hover_color="#C0C0C0", bg_color='#C50000')
help_button.place(x=594, y=7, anchor='ne')
help_button.lift()

# DIRECTORY SECTION
directory_frame = ctk.CTkFrame(root)
directory_frame.configure(fg_color='#FFF')
directory_frame.pack(fill="x", padx=10, pady=(5,10))

directory_label = ctk.CTkLabel(directory_frame, text="Selecione o diretório principal:", font=subtitle)
directory_label.grid(row=0, column=0, sticky='w')

directory_entry = ctk.CTkEntry(directory_frame, font=normal, state='readonly',
                               width=531, height=42, border_width=0, 
                               fg_color='#F1F1F1')
directory_entry.grid(row=1, column=0)

folder = Image.open('resources/images/folder.png')
def on_enter(e):
    openedFolder = Image.open('resources/images/opened_folder.png')
    select_button.configure(fg_color='#C50000', image=CTkImage(openedFolder))
def on_leave(e):
    select_button.configure(fg_color='#E70000', image=CTkImage(folder))

select_button = ctk.CTkButton(directory_frame, text='', command=select_directory, 
                              width=42, height=42, image=CTkImage(folder), 
                              fg_color='#E70000', border_width=0)
select_button.grid(row=1, column=2, padx=6)
select_button.bind("<Enter>", on_enter)
select_button.bind("<Leave>", on_leave)

# TREE VIEW SECTION
treeViewStyle = ttk.Style()
treeViewStyle.theme_use("default")
treeViewStyle.configure("Treeview", background="#F1F1F1", font=("Fira Sans Condensed", 12), foreground="black", borderwidth=0, rowheight=25, fieldbackground="#F1F1F1")
treeViewStyle.configure("Treeview.Heading", background="#F1F1F1", foreground="black", borderwidth=0, font=("Fira Sans Condensed", 12, 'bold'))
treeViewStyle.map("Treeview", background=[('selected', '#DD4A48')])
treeview = ttk.Treeview(root)
treeview.column("#0")
treeview.pack(expand=True, fill="both", padx=10)
treeview.tag_configure("bold", font=("Fira Sans Condensed", 12, "bold"))

# MAIN FUNCTION SECTION
main_function_frame = ctk.CTkFrame(root)
main_function_frame.configure(fg_color='#FFF')
main_function_frame.pack(pady=8)
function_button = ctk.CTkButton(main_function_frame, text="Fabricar PDFs", font=button, text_color="#FFF", command=main_function,
                                height=42, fg_color='#E70000', hover_color='#C50000')
function_button.grid(row=0, column=0, padx=6)
settings_icon = Image.open("resources/images/settings_icon.png")
settings_button = ctk.CTkButton(main_function_frame, text='', command=open_settings_popup,
                                width=42, height=42, image=CTkImage(settings_icon),
                                fg_color='#E70000', hover_color='#C50000', border_width=0)
settings_button.grid(row=0, column=1)

root.mainloop()
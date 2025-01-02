# from asyncio import subprocess
# from operator import truediv
# from pydoc import Helper
from re import I
from threading import Thread
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from service import writeValue
# import os
# import subprocess
# from subprocess import Popen, PIPE
# from threading import Thread
# import helper
# import time

root = Tk()
root.title("TTIMESHEET GENERATOR")
root.iconbitmap("engineering.ico")
frm = ttk.Frame(root, padding=50)
status = StringVar()

def browse_button_pdf():
    global pdf_file_path
    filename = filedialog.askopenfilename()
    if filename == "":
        pdf_file_path.set("Pilih file .pdf")
    else:
        pdf_file_path.set(filename)

def browse_button_excel():
    global excel_file_path
    filename = filedialog.askopenfilename()
    if filename == "":
        excel_file_path.set("Pilih file .excel")
    else:
        excel_file_path.set(filename)
        

def start_writting():
    btnGenerate.config(state="disabled")
    th = Thread(target = write_value)
    th.start()
    
def write_value():
    status.set(writeValue(13, pdf_file_path.get(), excel_file_path.get(), btnGenerate))
        
# stop = False
# th = Thread(target = Convert)
def entry_point():
    frm.grid()

    pdf_file_path = StringVar()
    pdf_file_path.set("Pilih file .pdf")

    excel_file_path = StringVar()
    excel_file_path.set("Pilih file .excel")

    ttk.Label(frm, text="== TTIMESHEET GENERATOR ==").grid(column=1, columnspan=2, row=10, pady=(0, 25))

    ttk.Label(frm, textvariable=excel_file_path).grid(column=1, row=50)
    btnBrowseExcel = ttk.Button(frm, text="Browse Excel..", command=browse_button_excel)
    btnBrowseExcel.grid(column=2, row=50)

    ttk.Label(frm, textvariable=pdf_file_path).grid(column=1, row=51)
    btnBrowsePdf = ttk.Button(frm, text="Browse Pdf..", command=browse_button_pdf)
    btnBrowsePdf.grid(column=2, row=51)

    labelStatus = ttk.Label(frm, textvariable=status)
    labelStatus.grid(column=1, columnspan=2, row=99, pady=(50, 5))

    btnGenerate = ttk.Button(frm, text="GENERATE", command=start_writting)
    btnGenerate.grid(column=1, columnspan=2, row=100)

    # btnStop = ttk.Button(frm, text="Stop", command=StopConvert)
    # btnStop.grid(column=2, row=100)
    # helper.SwitchButton(btnStop, "disabled")

    root.mainloop()
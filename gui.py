from threading import Thread
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from service import writeValue

root = Tk()
root.title("TIMESHEET GENERATOR")
root.iconbitmap("engineering.ico")
frm = ttk.Frame(root, padding=50)
status = StringVar()
excel_file_path = StringVar()
excel_file_path.set("Pilih file .excel")
pdf_file_path = StringVar()
pdf_file_path.set("Pilih file .pdf")

def browse_button_pdf():
    filename = filedialog.askopenfilename()
    if filename == "":
        pdf_file_path.set("Pilih file .pdf")
    else:
        pdf_file_path.set(filename)

def browse_button_excel():
    filename = filedialog.askopenfilename()
    if filename == "":
        excel_file_path.set("Pilih file .excel")
    else:
        excel_file_path.set(filename)
        

def start_writting():
    if pdf_file_path.get() == "Pilih file .pdf" or not pdf_file_path:
        status.set("file pdf belum dipilih!")
    elif excel_file_path.get() == "Pilih file .excel" or not excel_file_path:
        status.set("file excel belum dipilih!")
    else:
        btnGenerate.config(state="disabled")
        status.set("processing...")
        th = Thread(target = write_value)
        th.start()
    
def write_value():
    status.set(writeValue(13, pdf_file_path.get(), excel_file_path.get(), btnGenerate))

btnGenerate = ttk.Button(frm, text="GENERATE", command=start_writting)
btnGenerate.grid(column=1, columnspan=2, row=100)

def entry_point():
    frm.grid()

    ttk.Label(frm, text="== TTIMESHEET GENERATOR ==").grid(column=1, columnspan=2, row=10, pady=(0, 25))

    ttk.Label(frm, textvariable=excel_file_path).grid(column=1, row=50)
    btnBrowseExcel = ttk.Button(frm, text="Browse Excel..", command=browse_button_excel)
    btnBrowseExcel.grid(column=2, row=50)

    ttk.Label(frm, textvariable=pdf_file_path).grid(column=1, row=51)
    btnBrowsePdf = ttk.Button(frm, text="Browse Pdf..", command=browse_button_pdf)
    btnBrowsePdf.grid(column=2, row=51)
    
    labelStatus = ttk.Label(frm, textvariable=status)
    labelStatus.grid(column=1, columnspan=2, row=99, pady=(50, 5))

    root.mainloop()
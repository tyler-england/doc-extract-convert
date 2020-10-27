import tkinter as tk, xlwings, os
from tika import parser
from xlsxwriter.workbook import Workbook
from tkinter import filedialog

root = tk.Tk()  # tkinter to find file
root.withdraw()
file_path = filedialog.askopenfilename()  # dialog box to navigate to file
file_ext = file_path[file_path.rfind(".") + 1:]  # find the filename extension
i=1
while i >0:
   j=i
   i=file_path.rfind("\\")

file_name=file_path[j + 1:file_path.rfind(".")]
if file_ext.lower().find("pdf") == 0:  # if extension starts with "pdf"
    doc_type = "pdf"
elif file_ext.find("doc") == 0:  # if extension starts with "doc"
    doc_type = "doc"
else:  # neither a pdf nor a doc file
    doc_type = "xxx"
    print("Unrecognized document format. Program designed for PDF or DOCX.")
    quit()

raw = parser.from_file(file_path)  # raw content from the file (stored as a list)
nice_text = str(raw)  # convert that raw content into one long string


import xlwings, tkinter as tk
from tika import parser
from tkinter import filedialog

root=tk.Tk() #tkinter to find file
file_path = filedialog.askopenfilename() #dialog box to navigate to file

file_ext = file_path[file_path.rfind(".") + 1:]  # find the filename extension
if file_ext.find("pdf") == 0:  # if extension starts with "pdf"
    doc_type = "pdf"
elif file_ext.find("doc") == 0:  # if extension starts with "doc"
    doc_type = "doc"
else:  # neither a pdf nor a doc file
    doc_type = "xxx"
    print("Unrecognized document format. Please use PDF or DOCX.")
    quit()

# TODO: see if numbering is alphabetical/numerical/etc

raw = parser.from_file(file_path)  # raw content from the file (stored as a list)
nice_text = str(raw)  # convert that raw content into one long string

lineitems_array = [""] * len(nice_text)  # create an array/list to store the line items
k = 0  # current position in lineitems array/list
for j in range(len(nice_text)):  # for each letter position in the long string of content
    lineitems_array[k] = lineitems_array[k] + nice_text[j]  # add the letter/character to current line item
    if j < (len(nice_text) - 3):  # required to stay in the index range (looking 3 letters ahead for PDFs)
        if nice_text[j + 1] == "\\":  # if next letter is a backslash
            if nice_text[j + 2] == "n" or nice_text[j + 2] == "t":  # \n = new line ... \t = new table entry
                if doc_type == "pdf":  # because every line on a PDF registers as a new line "\n"
                    if nice_text[j + 3].isnumeric():  # if the new line is followed by a # (likely to be intentional)
                        k += 1  # new array/list item
                        j += 2  # skip past the new line ("\n")
                elif doc_type == "doc":  # only actual new lines register as new lines "\n" (number not required)
                    k += 1  # new array/list item
                    j += 2  # skip past the new line ("\n")
        elif nice_text[j + 1] == "\u2022":  # next character is a bullet point
            k += 1  # new array/list item

for i in range(len(lineitems_array) - 1):  # for every line item
    lineitems_array[i] = str(lineitems_array[i]).replace("\\n", "").strip()  # remove line breaks
    lineitems_array[i] = str(lineitems_array[i]).replace("\\t", "").strip()  # remove table breaks

lineitems_array = list(filter(lambda x: len(x) > 3, lineitems_array))  # remove line items with 3 or fewer char's

wkbk = xlwings.Book()  # open empty workbook
wsheet = wkbk.sheets["Sheet1"]  # declare sheet as sheet1

for i in range(len(lineitems_array)):  # for all line items
    print(str(round(100 * i / len(lineitems_array), ndigits=2)) + " %")  # progress notification
    cell = "A" + str(i)  # use i to identify the cell in Excel to receive the line item
    wsheet.range(i + 1, 1).value = lineitems_array[i]  # put the line item content into the cell

print("100.0 %")  # success notification

#TODO: Create columns for section numbers/letters?

# #TODO: Make wkbk top window

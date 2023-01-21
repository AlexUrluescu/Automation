from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import openpyxl

# ----------------------------------------------------------

root = Tk()

root.title("Send students grades")

root.resizable(0,0)

root.geometry("500x500")

# ------------------- conectare excel -------------------------

book = openpyxl.load_workbook("registru.xlsx", data_only=True)

hoja = book.sheetnames

# --------------------------------------------------------

# --------------- Functions ---------------------------

def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.config(text=root.filename)
    path = root.filename
    # print(path)
    return path


def sheet_select_btn_clb():
    valoare = combobox.get()
    print(valoare)


def getPath():
    print (path_label.cget("text"))


# -----------------------------------------------------------------


# ------------------- GUI -------------------------------------------

path_to_file = Label(root, text="Path")
path_to_file.place(x=50, y=20)

path_label = Label(root, text="")
path_label.place(x=100, y=20)

email_label = Label(root, text="Email")
email_label.place(x=50, y=170)

input_email = Entry(root)
input_email.place(x=90, y=170)
input_email.config(width=35)

password_label = Label(root, text="Password")
password_label.place(x=50, y=200)

input_password = Entry(root)
input_password.place(x=130, y=200)

combobox = Combobox(root, value = hoja)
combobox.place(x=50, y=50)

# ------- buttons ----------------

file_select_btn = Button(root, text="Select file", command=file_select_btn_clb)
file_select_btn.place(x=100, y=240)

butonCoord = Button(root, text="Get coordinates")
butonCoord.place(x=200, y=240)

sheet_select = Button(root, text="Get value", command=sheet_select_btn_clb);
sheet_select.place(x=50, y=100)

butonTrimite = Button(root, text="Get the path", command=getPath)
butonTrimite.place(x=50, y=300)

# -------------------------------------------------------------------------------


root.mainloop()
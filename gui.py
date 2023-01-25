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



# hoja = book.sheetnames

# --------------------------------------------------------

# --------------- Functions ---------------------------

def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.config(text=root.filename)
    path = root.filename
    # print(path)
    return path


def parameters_selected_btn_clb():
    sheet = combobox.get()
    path = (path_label.cget("text"))
    email = input_email.get()
    password = input_password.get()
    return path


def load_sheets_clb():
    path = parameters_selected_btn_clb();
    book = open_xlsx(path)
    combobox.config(values=book.sheetnames)


def open_xlsx(path):
    book = openpyxl.load_workbook(path, data_only=True)
    # combobox.config(values=book.sheetnames)
    return book


def email_emisor():
    email = input_email.get()
    return email


def password_emisor():
    password = input_password.get()
    return password


def get_sheet():
    sheet = combobox.get()
    return sheet


def get_table_data(book):
    active_sheet = book.active # stocam in ative_sheet, sheet-ul activ
    print(type(active_sheet)) # ne returneaza tipul
    return active_sheet
    # data_cells = active_sheet["A4":"D5"]
    # for row in data_cells:
    #     data_array.append([cell_data.value for cell_data in row])

    # return data_array


def print_data(data):
    print(data)


def print_sheet(sheet):
    print(sheet)

def main():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    sheet = get_sheet()
    email = email_emisor()
    password = password_emisor()
    data = get_table_data(book)
    print_data(data)
    

# -----------------------------------------------------------------


# ------------------- GUI -------------------------------------------

path_to_file = Label(root, text="Path")
path_to_file.place(x=50, y=20)

path_label = Label(root, text="")
path_label.place(x=85, y=22)
path_label.config(bg="gray", width=40)

sheet_label = Label(root, text="Sheet")
sheet_label.place(x=50, y=70)

combobox = Combobox(root)
combobox.place(x=90, y=70)
combobox.set("SheetName")


email_label = Label(root, text="Email")
email_label.place(x=50, y=170)

input_email = Entry(root)
input_email.place(x=90, y=170)
input_email.config(width=35)

password_label = Label(root, text="Password")
password_label.place(x=50, y=200)

input_password = Entry(root)
input_password.place(x=130, y=200)


# ------- buttons ----------------

file_select_btn = Button(root, text="Select file", command=file_select_btn_clb)
file_select_btn.place(x=400, y=20)

load_sheets = Button(root, text="Load Sheets", command=load_sheets_clb)
load_sheets.place(x=250, y=65)

parameters_selected = Button(root, text="Get the parameters", command=main)
parameters_selected.place(x=50, y=300)

# -------------------------------------------------------------------------------

root.mainloop()



if __name__ == "__main__":
    main()
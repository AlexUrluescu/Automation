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
    # print(sheet)
    # print(path)
    # print(email)
    # print(password)
    return path


def open_xlsx(path):
    book = openpyxl.load_workbook(path, data_only=True)
    return book

def main():
    path = parameters_selected_btn_clb()
    print(path)
    book = open_xlsx(path)
    print(book)
    combobox.config(values=book.sheetnames)

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

parameters_selected = Button(root, text="Get the parameters", command=main)
parameters_selected.place(x=50, y=300)

# -------------------------------------------------------------------------------

root.mainloop()



if __name__ == "__main__":
    main()
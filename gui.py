from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Combobox
import openpyxl
from email.message import EmailMessage
import ssl
import smtplib

# ----------------------------------------------------------

root = Tk()

root.title("Send students grades")

root.resizable(0,0)

root.geometry("500x500")

# --------------------------------------------------------

# --------------- Functions ------------------------------

def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.config(text=root.filename)
    path = root.filename

    return path


def parameters_selected_btn_clb():
    path = (path_label.cget("text"))

    return path


def load_sheets_clb():
    path = parameters_selected_btn_clb();
    book = open_xlsx(path)
    combobox.config(values=book.sheetnames)


def open_xlsx(path):
    book = openpyxl.load_workbook(path, data_only=True)
  
    return book


def email_emisor():
    email = input_email.get()
    return email


def password_emisor():
    password = input_password.get()
    return password


def get_sheet(book):
    list_data = []
    sheet = combobox.get()
    book.active = book[sheet]
    book_active = book.active
    print(book_active.tables.items())
    paramas_table = book_active.tables.items()
    for data in paramas_table:
        list_data.append(data[1])

    list_data.append(book_active)
    
    return list_data


def get_table_data(sheet, range):
    data_array = []

    active_sheet = sheet
    print(active_sheet) 

    data_cells = active_sheet[range]
    for row in data_cells:
        data_array.append([cell_data.value for cell_data in row])
   
    print(data_array)
    return data_array


def send_emails(data, email, password):
    print(f"from send_emails_data -> {data}")
    print(email)

    array_students = []

    for student in data:
        if(student[0] == None):
            break

        array_students.append(student)
        
    array_students.pop(0)
    print(array_students)

    for student in array_students:
        student_firstName = student[0]
        student_lastName = student[1]
        student_grade = student[2]
        student_email = student[3]

        email_emisor = email
        email_password = password
        email_receptor = student_email

        email_subject = 'Nota examen Fizica'
        email_body = 'Buna ziua ' + student_firstName + " " + student_lastName + ', ai luat nota ' + str(student_grade) + "."

        em = EmailMessage()
        em['From'] = email_emisor
        em['To'] = email_receptor
        em['Subject'] = email_subject
        em.set_content(email_body)

        contexto = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=contexto) as smtp:
            smtp.login(email_emisor, email_password)
            smtp.sendmail(email_emisor, email_receptor, em.as_string())

        print('s-a terminat')


def main():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    range,sheet = get_sheet(book)
    email = email_emisor()
    password = password_emisor()
    data = get_table_data(sheet, range)
    send_emails(data, email, password)
    

# -----------------------------------------------------------------


# ------------------- GUI -------------------------------------------

path_to_file = Label(root, text="Path")
path_to_file.place(x=50, y=20)

path_label = Label(root, text="")
path_label.place(x=85, y=22)
path_label.config(bg="black", width=40, fg="white")

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
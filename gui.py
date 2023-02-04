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

drop_down_sheet_list = StringVar()
drop_down_sheet_list.set("SheetName")

# --------------------------------------------------------

contor_students = 0

# --------------- Functions ------------------------------

def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.config(text=root.filename)
    path = root.filename
    combobox.configure(state=NORMAL)

    return path


def parameters_selected_btn_clb():
    path = (path_label.cget("text"))

    return path


def clear_widget(contor_students):
    info_students = f"S-au incarcat: {contor_students} studenti"
    info_students_data.delete(1.0, END)
    info_students_data.insert(1.0, info_students)
    info_students_data.configure(state=DISABLED)


def load_sheets_clb():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    combobox.config(values=book.sheetnames)
    


def open_xlsx(path):
    book = openpyxl.load_workbook(path, data_only=True)
    root.update()
    return book


def email_emisor():
    email = input_email.get()
    return email


def password_emisor():
    password = input_password.get()
    return password


def get_sheet(book):
    list_data = []
    sheet = drop_down_sheet_list.get()
    book.active = book[sheet]
    book_active = book.active
    print(book_active.tables.items())
    paramas_table = book_active.tables.items()
    for data in paramas_table:
        list_data.append(data[1])

    list_data.append(book_active)
    
    return list_data


def get_table_data(sheet, range):
    global contor_students
    data_array = []
    array_students = []

    active_sheet = sheet
    print(active_sheet) 

    data_cells = active_sheet[range]
    for row in data_cells:
        data_array.append([cell_data.value for cell_data in row])
   

    for student in data_array:
        if(student[0] == None):
            break

        array_students.append(student)

    array_students.pop(0)

    print(array_students)

    
    contor_students = len(array_students)
    clear_widget(contor_students)
    print(contor_students)
    parameters_selected.configure(state=NORMAL)
    return array_students


def send_emails(data, email, password):

    print(f"from send_emails_data -> {data}")
    print(email)

    for student in data:

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


def get_data_btn_clb():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    range,sheet = get_sheet(book)
    email = email_emisor()
    print(email)
    password = password_emisor()
    print(password)
    get_table_data(sheet, range)
    


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

combobox = Combobox(root, textvariable= drop_down_sheet_list, postcommand=load_sheets_clb)
combobox.place(x=90, y=70)
combobox.configure(state=DISABLED)
#combobox.set("SheetName")


email_label = Label(root, text="Email")
email_label.place(x=50, y=170)

input_email = Entry(root)
input_email.place(x=90, y=170)
input_email.config(width=35)

password_label = Label(root, text="Password")
password_label.place(x=50, y=200)

input_password = Entry(root)
input_password.place(x=130, y=200)

details_author_label = Label(root, text="Made by Alexandre Urluescu, contact: alexurluescu23@gmail.com")
details_author_label.place(x=50, y=470)

text_afisare = "Sunt un text";

info_students_data = Text(root, width=50, height=5);
info_students_data.insert(1.0, "")

info_students_data.place(x=50, y=360)

# ------- buttons ----------------

file_select_btn = Button(root, text="Select file", command=file_select_btn_clb)
file_select_btn.place(x=400, y=20)


parameters_selected = Button(root, text="Send", width=15, command=main)
parameters_selected.place(x=250, y=270)
parameters_selected.configure(state=DISABLED)

get_data_btn = Button(root, text="Get students data", command=get_data_btn_clb)
get_data_btn.place(x=50, y=270)

# -------------------------------------------------------------------------------

root.mainloop()


if __name__ == "__main__":
    main()
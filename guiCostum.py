
from email.message import EmailMessage
import ssl
import smtplib
from customtkinter import CTk, CTkButton, CTkCheckBox, CTkComboBox,CTkEntry, CTkLabel, CTkFrame, CTkTextbox, CTkProgressBar
from tkinter import *
from tkinter import filedialog
import openpyxl
import time
from tkinter import ttk
import logging
import customtkinter

# customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

# customtkinter.set_appearance_mode("dark")
# customtkinter.set_appearance_mode("light")

logging.basicConfig(level=logging.DEBUG, format="%(levelname)s: %(time)s %(message)s")


root = CTk()
root.geometry("630x750")
root.title("Send students grades")
root.resizable(0,0)

drop_down_sheet_list = StringVar()
drop_down_sheet_list.set("SheetName")
contor_students = 0

grades_var = IntVar()
absente_var = IntVar()
extra_text_var = IntVar()
restante_var = IntVar()

numar = 0
value = 0

costumize_var = True

# -------------- fecth email -----------------------------------------------

date_frame = CTkFrame(root, height=70)
date_frame.grid(row = 2, column=0, sticky="we")

input_email = CTkEntry(date_frame, width=250, border_color="#C0C0C0")
input_email.grid(column=1, row=0, padx=10, pady=30, sticky="w")

with open("email.txt", "r") as file_object:
    email = file_object.read()
    input_email.delete(0,END)
    input_email.insert(0,email)
    logging.debug(email)

# ----------------------------------------------------------------------

# ----- functii ---------------

def costumize():
    global costumize_var

    if(costumize_var == True):
        logging.debug("A intrat in true")

        customtkinter.set_appearance_mode("light")
        
        mode_button.configure(text="Dark", fg_color="black", hover_color="black", text_color="white")
        path_to_file.configure(text_color="black")
        sheet_label.configure(text_color="black")
        email_label.configure(text_color="black")
        subject_label.configure(text_color="black")
        grades_checkbutton.configure(text_color="black")
        absente_checkbutton.configure(text_color="black")
        extraText_checkbutton.configure(text_color="black")
        restante_checkbutton.configure(text_color="black")
        date_label.configure(text_color="black")
        room_label.configure(text_color="black")
        details_author_label.configure(text_color="black")
        progress_bar_label.configure(text_color="black")

        costumize_var = False
    
    else:
        logging.debug("A intrat in false")

        customtkinter.set_appearance_mode("dark")

        mode_button.configure(text="Light", fg_color="white", hover_color="white", text_color="black")
        path_to_file.configure(text_color="white")
        sheet_label.configure(text_color="white")
        email_label.configure(text_color="white")
        subject_label.configure(text_color="white")
        grades_checkbutton.configure(text_color="white")
        absente_checkbutton.configure(text_color="white")
        extraText_checkbutton.configure(text_color="white")
        restante_checkbutton.configure(text_color="white")
        date_label.configure(text_color="white")
        room_label.configure(text_color="white")
        details_author_label.configure(text_color="white")
        progress_bar_label.configure(text_color="white")
        
        costumize_var = True

        

def fetchFileTxt():
    vector = []
    email = input_email.get()
    vector.append(email)
    file = open('email.txt', 'w') # ne conectam cu fisierul 'date.txt' si vrem sa scriem date in el, deci punem 'w' (write)
    # pentru fiecare valoare din vector vreau sa mi-l scrii in fisier
    for i in vector:
        file.write(str(i) + "\n")
    
    file.close() #inchidem conexiunea cu fisierul 'date.txt'


def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.configure(text=root.filename)
    path = root.filename
    sheets_btn.configure(state=NORMAL)

    logging.debug(path)
    return path


def parameters_selected_btn_clb():
    path = (path_label.cget("text"))

    logging.debug(path)
    return path


def open_xlsx(path):
    book = openpyxl.load_workbook(path, data_only=True)
    root.update()

    logging.debug(book)
    return book


def get_sheet(book):
    list_data = []
    sheet = drop_down_sheet_list.get()
    book.active = book[sheet]
    book_active = book.active
    logging.debug(book_active.tables.items())
    paramas_table = book_active.tables.items()
    for data in paramas_table:
        list_data.append(data[1])

    list_data.append(book_active)
    
    return list_data


def email_emisor():
    email = input_email.get().strip()
    return email


def password_emisor():
    password = input_password.get()
    return password


def get_subject():
    subject = input_subject.get().capitalize()
    return subject


def load_sheets_clb():
    combobox.configure(state=NORMAL)
    path = parameters_selected_btn_clb()
    logging.debug(path)
    book = open_xlsx(path)
    logging.debug(book)
    combobox.configure(values=book.sheetnames)
    logging.debug(book.sheetnames)


def extraText_callback():
    if(extra_text_var.get() == 1):
# info_students_data.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        text_proba.grid(row=1, column=0, padx=10, pady=10, sticky="w")

    if(extra_text_var.get() == 0):
        text_proba.place(x=1000, y=1000)


def restante_callback():
    if(restante_var.get() == 1):
        restante_frame.grid(row=4, column=0, sticky="we")
        input_date.configure(state=NORMAL)
        input_room.configure(state=NORMAL)

        absente_checkbutton.configure(state=DISABLED)
        grades_checkbutton.configure(state=DISABLED)
        grades_checkbutton.deselect()
        absente_checkbutton.deselect()
        

    if(restante_var.get() == 0):
        restante_frame.place(x=1000, y=1000)
        
        absente_checkbutton.configure(state=NORMAL)
        grades_checkbutton.configure(state=NORMAL)
        grades_checkbutton.select()


def clear_widget(contor_students):
    info_students = f"S-au incarcat: {contor_students} studenti"
    info_students_data.delete(1.0, END)
    info_students_data.insert(1.0, info_students)
    info_students_data.configure(state=DISABLED)
    

def get_table_data(sheet, range):
    global contor_students
    data_array = []
    array_students = []

    active_sheet = sheet
    logging.debug(active_sheet)

    data_cells = active_sheet[range]
    for row in data_cells:
        data_array.append([cell_data.value for cell_data in row])
   

    for student in data_array:
        if(student[0] == None):
            break

        array_students.append(student)

    array_students.pop(0)

    logging.debug(array_students)

    
    contor_students = len(array_students)
    clear_widget(contor_students)
    logging.debug(contor_students)
    parameters_selected.configure(state=NORMAL)
    return array_students


def get_data_btn_clb():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    range,sheet = get_sheet(book)
    logging.debug(f"Range {range}, Sheet {sheet}")
    fetchFileTxt()
    email = email_emisor()
    logging.debug(email)
    password = password_emisor()
    logging.debug(password)
    get_table_data(sheet, range)
    

def get_room():
    room = input_room.get()
    return room


def get_date():
    date = input_date.get()
    return date


def send_emails(data, email, password, subject, room, date):
    global numar
    global value

    logging.debug(f"from send_emails_data -> {data}")

    # progress bar placement in the root ----------------
    progress_bar_label.pack(padx=10)
    progress_bar.pack(padx=10, pady=10) 
    

    range_progress = len(data) # variable for the dynamic increment
   
    # ------------------------------------------------------

    for student in data:

        student_firstName = student[0]
        student_lastName = student[1]
        student_email = student[2]
        student_grade = student[3]
        student_abs = student[4]
        student_restanta = student[5]

        email_emisor = email
        email_password = password
        email_receptor = student_email

        if(grades_var.get() == 1 and absente_var.get() == 0):
            email_subject = 'Nota examen la ' + subject
            email_body = 'Buna ziua ' + student_firstName + " " + student_lastName + ', ai luat nota ' + str(student_grade) + "\n" + text_proba.get(1.0, END) + "."

        elif(absente_var.get() == 1 and grades_var.get() == 0):
            email_subject = 'Numarul de absente la ' + subject
            email_body = 'Buna ziua ' + student_firstName + " " + student_lastName + ', ai in total ' + str(student_abs) + " absente. \n" + text_proba.get(1.0, END) + "."

        elif(absente_var.get() == 1 and grades_var.get() == 1):
            email_subject = 'Notele si absentele la ' + subject
            email_body = 'Buna ziua ' + student_firstName + " " + student_lastName + ', ai nota ' + str(student_grade) + ' si ai in total ' + str(student_abs) + " absente. \n" + text_proba.get(1.0, END) + "."

        if(restante_var.get() == 1):
            if(int(student_grade) < 5):
                email_subject = "Informatii despre restanta la " + subject
                email_body = "Buna ziua " + student_firstName + " " + student_lastName + ", ai restanta in data de " + date +" in clasa " + room + ".\n" + text_proba.get(1.0, END) + "."

            if(int(student_grade) > 5):
                email_subject = "Informatii despre restanta la " + subject
                email_body = "Buna ziua " + student_firstName + " " + student_lastName + ", esti integralist, felicitari!.\n" + text_proba.get(1.0, END) + "."

        em = EmailMessage()
        em['From'] = email_emisor
        em['To'] = email_receptor
        em['Subject'] = email_subject
        em.set_content(email_body)

        contexto = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=contexto) as smtp:
            smtp.login(email_emisor, email_password)
            smtp.sendmail(email_emisor, email_receptor, em.as_string())

        #progress bar dynamic increment ----------------------------------
       
       
        progress_bar['value'] += 100/range_progress
        progress_bar_label.configure(text= str(round(progress_bar['value'], 1))+"%")
        root.update_idletasks()
        time.sleep(1)

        # --------------------------------------------------------------------

        logging.debug('s-a terminat')

    
def main():
    path = parameters_selected_btn_clb()
    book = open_xlsx(path)
    range,sheet = get_sheet(book)
    email = email_emisor()
    password = password_emisor()
    subject = get_subject()
    room = get_room()
    date = get_date()
    data = get_table_data(sheet, range)
    send_emails(data, email, password, subject, room, date)

# ---------- Frames --------------------------------------------------

path_frame = CTkFrame(root, height=70)
path_frame.grid(row = 0, column = 0, sticky="we")

sheet_frame = CTkFrame(root, height=70)
sheet_frame.grid(row = 1, column=0, sticky="we")

checkbuttons_frame = CTkFrame(root, height=70)
checkbuttons_frame.grid(row = 3, column=0, sticky="we")

restante_frame = CTkFrame(root)

buttons_frame = CTkFrame(root)
buttons_frame.grid(row=5, column=0, sticky="we")

textBox_Frame = CTkFrame(root)
textBox_Frame.grid(row=6, column=0, sticky="we")

progressBar_frame = CTkFrame(root, height=70)
progressBar_frame.grid(row=7, column=0, sticky="we")

contact_details_frame = CTkFrame(root, height=20)
contact_details_frame.place(x=100, y=715)

# ------------------------------------------------------------------------

# ------------- Content Path Frame ------------------------------------------------

path_to_file = CTkLabel(path_frame, text="Path", font=('Comic Sans MS', 18), text_color="white")
path_to_file.grid(column = 0, row = 0, padx=10, pady=10)

path_label = CTkLabel(path_frame, text="", bg_color="black", width=340)
path_label.grid(column=1, row=0, padx=10, pady=10)


file_select_btn = CTkButton(path_frame, text="Select file",height=30, width=120, command=file_select_btn_clb, font=('Comic Sans MS', 15), text_color="white", corner_radius=15)
file_select_btn.grid(column=2, row=0, padx=10, pady=10)

mode_button = CTkButton(path_frame, text="Light", width=40, height=40, fg_color="white", hover_color="white", text_color="black", command=costumize)
mode_button.grid(row=0, column=3, padx=10, pady=10)

# ---------------------------------------------------------------------------------

# ------------ Content Sheet Frame -------------------------------------------------------

sheet_label = CTkLabel(sheet_frame, text="Sheet", font=('Comic Sans MS', 18), text_color="white")
sheet_label.grid(column=0, row=0, padx=10, pady=10)

combobox = CTkComboBox(sheet_frame, variable=drop_down_sheet_list)
combobox.grid(column=1, row=0, padx=10, pady=10)
combobox.configure(state=DISABLED)

sheets_btn = CTkButton(sheet_frame, text="Load Sheets", command=load_sheets_clb, font=('Comic Sans MS', 15), text_color="white", corner_radius=15)
sheets_btn.grid(row=0, column=2, padx=30, pady=10)
sheets_btn.configure(state=DISABLED)

# -----------------------------------------------------------------------------------

# ------------ Content Date Frame ----------------------------------------------

email_label = CTkLabel(date_frame, text="Email", font=('Comic Sans MS', 18), text_color="white")
email_label.grid(column=0, row=0, padx=0, pady=30)


password_label = CTkLabel(date_frame, text="Password", font=('Comic Sans MS', 18))
password_label.grid(column=2, row=0, padx=10, pady=10)

input_password = CTkEntry(date_frame, border_color="white")
input_password.grid(column=3, row=0, padx=10, pady=10, sticky="w")

subject_label = CTkLabel(date_frame, text="Subject", font=('Comic Sans MS', 18), text_color="white")
subject_label.grid(row=2, column=0, padx=10, pady=10)

input_subject = CTkEntry(date_frame)
input_subject.grid(row=2, column=1, padx=10, pady=10, sticky="w")

# ------------------------------------------------------------------------------

# ---------- Content Checkbutton Frame ------------------------------------------

grades_checkbutton = CTkCheckBox(checkbuttons_frame, text="Grades", variable=grades_var, onvalue=1, offvalue=0, font=('Comic Sans MS', 16), text_color="white")
grades_checkbutton.grid(row=0, column=0, padx=10, pady=30)
grades_checkbutton.select()

absente_checkbutton = CTkCheckBox(checkbuttons_frame, text="Absente", variable=absente_var, onvalue=1, offvalue=0, font=('Comic Sans MS', 16), text_color="white")
absente_checkbutton.grid(row=0, column=1, padx=10, pady=30)

extraText_checkbutton = CTkCheckBox(checkbuttons_frame, text="Extra text", variable=extra_text_var, onvalue=1, offvalue=0, command=extraText_callback, font=('Comic Sans MS', 16), text_color="white")
extraText_checkbutton.grid(row=0, column=2, padx=10, pady=30)

restante_checkbutton = CTkCheckBox(checkbuttons_frame, text="Restante", variable=restante_var, onvalue=1, offvalue=0, command=restante_callback, font=('Comic Sans MS', 16), text_color="white")
restante_checkbutton.grid(row=0, column=3, padx=10, pady=30)

# -------------------------------------------------------------------------------

# ------------- Content Restante_Frame--------------------------------------------

room_label = CTkLabel(restante_frame, text="Room", font=('Comic Sans MS', 16), text_color="white")
room_label.grid(row=0, column=0, padx=10, pady=10)

input_room = CTkEntry(restante_frame, state=DISABLED)
input_room.grid(row=0, column=1, padx=10, pady=10)

date_label = CTkLabel(restante_frame, text="Exam date", font=('Comic Sans MS', 16), text_color="white")
date_label.grid(row=0, column=2, padx=10, pady=10)

input_date = CTkEntry(restante_frame, state=DISABLED)
input_date.grid(row=0, column=3, padx=10, pady=10)

# --------------------------------------------------------------------------------

# -------------------- Content Buttons Frame -------------------------------------

get_data_btn = CTkButton(buttons_frame, text="Get students data", command=get_data_btn_clb, font=('Comic Sans MS', 15), text_color="white", corner_radius=15)
get_data_btn.grid(row=0, column=0, padx=10, pady=10)

parameters_selected = CTkButton(buttons_frame, text="Send", width=100, command=main, font=('Comic Sans MS', 15), text_color="white", corner_radius=15)
parameters_selected.grid(row=0, column=1, padx=10, pady=10)
parameters_selected.configure(state=DISABLED)

# ------------------------------------------------------------------------------

# --------------------- Content TextBox Frame ------------------------------------

info_students_data = CTkTextbox(textBox_Frame, height=70, width=350);
info_students_data.insert(1.0, "")
info_students_data.grid(row=0, column=0, padx=10, pady=10, sticky="w")

text_proba = CTkTextbox(textBox_Frame, height=100, width=350);
text_proba.insert(1.0, "")
info_students_data.grid(row=0, column=0, padx=10, pady=10, sticky="w")

# ------------------------------------------------------------------------------

# --------------------- Content progressBar Frame ------------------------------------

progress_bar = ttk.Progressbar(progressBar_frame, orient=HORIZONTAL, length=400, mode="determinate")

progress_bar_label = CTkLabel(progressBar_frame, text='', font=('Comic Sans MS', 15), text_color="white")

# ---------------------------------------------------------------------------------

# ------------ Content Contact details Frame --------------------------------------------

details_author_label = CTkLabel(contact_details_frame, text="Made by Alexandre Urluescu, contact: alexurluescu23@gmail.com", font=('Comic Sans MS', 15), text_color="white")
details_author_label.pack()

root.mainloop()

if __name__ == "__main__":
    main()
    


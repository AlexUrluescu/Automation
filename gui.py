from tkinter import *
from tkinter import filedialog

root = Tk()

root.title("Send students grades")

root.resizable(0,0)

root.geometry("500x500")

list = []


# --------------- Functions ---------------------------

def file_select_btn_clb():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    path_label.config(text=root.filename)
    path = root.filename
    print(path)
    return path

# -----------------------------------------------------------------


# ------------------- GUI -------------------------------------------

path_to_file = Label(root, text="Path")
path_to_file.place(x=50, y=50)

path_label = Label(root, text="")
path_label.place(x=100, y=50)

email_label = Label(root, text="Email")
email_label.place(x=50, y=170)

input_email = Entry(root)
input_email.place(x=90, y=170)
input_email.config(width=35)

# ------- buttons ----------------

file_select_btn = Button(root, text="Select file", command=file_select_btn_clb)
file_select_btn.place(x=100, y=240)

butonCoord = Button(root, text="Get coordinates", command=getData)
butonCoord.place(x=200, y=240)

# -------------------------------------------------------------------------------
def main():
    book = excel()

root.mainloop()
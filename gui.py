from tkinter import *
from tkinter import filedialog

root = Tk()

root.title("Automation")

root.resizable(0,0)

root.geometry("500x500")

# --------------- Functions ---------------------------

def getExcelFile():
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes = (("Excel files", "*.xlsx"),("All files", "*.*")))
    print(root.filename)
    labelExcel.config(text=root.filename)
    # return root.filename

def getParameters():
    valueCoordX = inputCoordX.get()
    valueCoordY = inputCoordY.get()
    valueEmail = inputEmail.get()
    print(f"CoordX: {valueCoordX}, CoordY: {valueCoordY}, Email: {valueEmail}")

# -----------------------------------------------------------------

# ------------------- GUI -------------------------------------------

pathExcel = Label(root, text="Path")
pathExcel.place(x=50, y=50)

labelExcel = Label(root, text="")
labelExcel.place(x=100, y=50)

coordX = Label(root, text="Coordinate X")
coordX.place(x=50, y=90)

inputCoordX = Entry(root)
inputCoordX.place(x=140, y=90)

coordY = Label(root, text="Coordinate Y")
coordY.place(x=50, y=120)

inputCoordY = Entry(root)
inputCoordY.place(x=140, y=120)

emailLabel = Label(root, text="Email")
emailLabel.place(x=50, y=170)

inputEmail = Entry(root)
inputEmail.place(x=90, y=170)
inputEmail.config(width=35)

# ------- buttons ----------------

buton = Button(root, text="Get excel file", command=getExcelFile)
buton.place(x=100, y=240)

butonCoord = Button(root, text="Get coordinates", command=getParameters)
butonCoord.place(x=200, y=240)

# -------------------------------------------------------------------------------


root.mainloop()
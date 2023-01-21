from tkinter import *

window = Tk()

window.title("Prima noastra pagina")

window.geometry("500x500")

window.resizable(1,1)

numele_meu_label = Label(window, text="")
numele_meu_label.place(x=100, y=100)
   
# ----------- elements -------------------

nume_label = Label(window, text="Nume:")
nume_label.place(x=50, y=100)

nume_entry = Entry(window)
nume_entry.place(x=50, y=150)


def getName():
   nume = nume_entry.get()
   print("trece")
   numele_meu_label.config(text=str(nume))

buton_nume = Button(window, text="Get name", command=getName)
buton_nume.place(x=50, y=200)
buton_nume.config(bg="gray", width="20")

nr1_label=Label(window, text="primul nr")
nr1_label.place(x=70, y=250)

nr1_entry= Entry(window)
nr1_entry.place(x=170, y=250)

nr2_label=Label(window, text="primul nr")
nr2_label.place(x=70, y=300)

nr2_entry= Entry(window)
nr2_entry.place(x=170, y=300)

sum_label=Label(window, text="Suma=")
sum_label.place(x=70, y=340)

sum_label=Label(window, text="")
sum_label.place(x=150, y=340)

def getsum():
    sum=int(nr1_entry.get())+int(nr2_entry.get())
    print (sum)
    sum_label.config(text=sum)


butonas=Button(window, text="sum", command=getsum )
butonas.place(x=70, y=400)


window.mainloop()


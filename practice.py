from tkinter import *

window= Tk()

window.title("operatii elementare")

window.geometry("500x500")

window.resizable(0,0)

nr1_label=Label(window, text="nr1")
nr1_label.place(x=20,y=20)

nr1_entry=Entry(window)
nr1_entry.place(x=50, y=20)

nr2_label=Label(window, text="nr2")
nr2_label.place(x=20,y=50)

nr2_entry=Entry(window)
nr2_entry.place(x=50, y=50)

sum_label=Label(window, text="suma:")
sum_label.place(x=20,y=80)

sum_label=Label(window, text="")
sum_label.place(x=55,y=80)

def suma():
    sumanr=int(nr1_entry.get())+int(nr2_entry.get())
    sum_label.config(text=sumanr)

buton_sum = Button(window, text="suma", command=suma)
buton_sum.place(x=30, y=400)
buton_sum.config(bg="gray", width="15")

diferenta_label=Label(window, text="diferenta:")
diferenta_label.place(x=20,y=100)

diferenta_label=Label(window, text="")
diferenta_label.place(x=75,y=100)

def diferenta():
    difer=int(nr1_entry.get())-int(nr2_entry.get())
    diferenta_label.config(text=difer)

buton_dif=Button(window, text="diferenta", command=diferenta)
buton_dif.place(x=30,y=430)
buton_dif.config(bg="gray", width="15")

produs_label=Label(window, text="produsul:")
produs_label.place(x=20,y=120)

produs_label=Label(window, text="")
produs_label.place(x=75,y=120)

def produs():
    prod=int(nr1_entry.get())*int(nr2_entry.get())
    produs_label.config(text=prod)

buton_produs=Button(window, text="produs", command=produs)
buton_produs.place(x=200,y=400)
buton_produs.config(bg="gray", width="15")

impartire_label=Label(window, text="impartire:")
impartire_label.place(x=20,y=140)

impartire_label=Label(window, text="")
impartire_label.place(x=75,y=140)


def impartire():
    imp=int(nr1_entry.get())/int(nr2_entry.get())
    impartire_label.config(text=imp)

buton_impartire=Button(window, text="impartire", command=impartire)
buton_impartire.place(x=200,y=430)
buton_impartire.config(bg="gray", width="15")






window.mainloop()
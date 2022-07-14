from tkinter import *
from tkinter import ttk
from tkinter import messagebox as msgbox
import saPosMerge_selectFile as sp


win = Tk()
win.title("Raspberry Pi UI")
win.geometry('200x100+200+200')
def ok():
        print("ck3")

    #str = 'nothing selected'

        print(radVar.get())

        if radVar.get() == 1:
                str = "order_merge"
        elif radVar.get() == 2:
                str = "delivery_merge"
        elif radVar.get() == 3:
                str = "tax_merge"
        msgbox.showinfo("Button Clickec", str)

        #win.destroy()
radVar = IntVar()
r1=ttk.Radiobutton(win, text="Radio 1", variable=radVar,  value=1)
r1.grid(column=0, row=0)
r2=ttk.Radiobutton(win, text="Radio 2", variable=radVar, value=2)
r2.grid(column=0, row=1)
r2=ttk.Radiobutton(win, text="Radio 3", variable=radVar, value=3)
r2.grid(column=0, row=2)


action = ttk.Button(win, text = "Click Me", command = ok)
action.grid(column=0, row=4)


win.mainloop()
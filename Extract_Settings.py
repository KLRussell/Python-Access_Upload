from Global import grabobjs
from Global import ShelfHandle
from tkinter import *
from tkinter import messagebox

import os
import random
import pandas as pd

CurrDir = os.path.dirname(os.path.abspath(__file__))
PreserveDir = os.path.join(CurrDir, '04_Preserve')
Global_Objs = grabobjs(CurrDir)


class MainGUI:
    listbox = None
    obj = None

    def __init__(self):
        self.root = Tk()
        self.var = StringVar()

    def buildgui(self):
        self.root.geometry('325x270+500+300')
        self.root.title('Extract Settings')

        top_frame = Frame(self.root)
        middle_frame = Frame(self.root)
        bottom_frame = Frame(self.root)
        top_frame.pack(side=TOP)
        middle_frame.pack()
        bottom_frame.pack(side=BOTTOM, fill=BOTH, expand=True)

        text = Message(self.root, textvariable=self.var, width=250, justify=CENTER, pady=5)
        self.var.set('Please choose an access table to edit settings:')
        text.pack(in_=top_frame)

        scrollbar = Scrollbar(self.root, orient="vertical")
        scrollbar2 = Scrollbar(self.root, orient="horizontal")
        self.listbox = Listbox(self.root, selectmode=SINGLE, width=40, yscrollcommand=scrollbar.set,
                               xscrollcommand=scrollbar2.set)
        self.populatebox()
        scrollbar.config(command=self.listbox.yview)
        scrollbar2.config(command=self.listbox.xview)
        scrollbar.pack(in_=middle_frame, side="right", fill="y")
        scrollbar2.pack(in_=middle_frame, side="bottom", fill="x")
        self.listbox.pack(in_=middle_frame, pady=5)

        btn = Button(self.root, text="Change Settings", width=15, command=self.change_settings)
        btn2 = Button(self.root, text="Cancel", width=15, command=self.cancel)
        btn.pack(in_=bottom_frame, side=LEFT, padx=10)
        btn2.pack(in_=bottom_frame, side=RIGHT, padx=10)

        if self.listbox.size() > 0:
            self.listbox.select_set(0)
        else:
            btn.config(state=DISABLED)

    def showgui(self):
        self.root.mainloop()

    def populatebox(self):
        configs = Global_Objs['Local_Settings'].grab_item('Accdb_Configs')

        for i in configs:
            self.listbox.insert("end", i[0])

    def change_settings(self):
        if self.listbox.curselection():
            print('hi')
        else:
            messagebox.showerror('Selection Error!', 'No shelf date was selected. Please select a valid shelf item')

    def cancel(self):
        if self.obj:
            self.obj.close()

        self.root.destroy()


if __name__ == '__main__':
    myobj = MainGUI()
    myobj.buildgui()
    myobj.showgui()

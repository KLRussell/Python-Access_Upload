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
    selection = 0
    listbox = None
    obj = None

    def __init__(self):
        self.root = Tk()
        self.var = StringVar()

    def buildgui(self):
        self.root.geometry('300x270+500+300')
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
        self.listbox.bind("<Down>", self.on_list_down)
        self.listbox.bind("<Up>", self.on_list_up)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)

        if self.listbox and self.listbox.curselection():
            self.listbox.select_set(self.selection)

        btn = Button(self.root, text="Change Settings", width=15, command=self.change_settings)
        btn2 = Button(self.root, text="Cancel", width=15, command=self.cancel)
        btn.pack(in_=bottom_frame, side=LEFT, padx=10)
        btn2.pack(in_=bottom_frame, side=RIGHT, padx=10)

        if self.listbox.size() > 0:
            self.listbox.select_set(0)
        else:
            btn.config(state=DISABLED)

    def on_select(self, event):
        if self.listbox and self.listbox.curselection() and -1 < self.selection < self.listbox.size() - 1:
            self.selection = self.listbox.curselection()[0]

    def on_list_down(self, event):
        if self.selection < self.listbox.size()-1:
            self.listbox.select_clear(self.selection)
            self.selection += 1
            self.listbox.select_set(self.selection)

    def on_list_up(self, event):
        if self.selection > 0:
            self.listbox.select_clear(self.selection)
            self.selection -= 1
            self.listbox.select_set(self.selection)

    def showgui(self):
        self.root.mainloop()

    def populatebox(self):
        configs = Global_Objs['Local_Settings'].grab_item('Accdb_Configs')

        for i in configs:
            self.listbox.insert("end", i[0])

    def change_settings(self):
        if self.listbox.size() > 0 and self.listbox.curselection():
            if self.obj:
                self.obj.close()

            self.obj = SettingsGUI(self.root, self.listbox.get(self.listbox.curselection()))
            self.obj.buildgui()
        elif self.listbox.size() > 0:
            messagebox.showerror('Settings Empty Error!', 'No settings have been saved. Please run Extract_Upload.py')
        else:
            messagebox.showerror('Selection Error!', 'No table settings was selected. Please select from list')

    def cancel(self):
        if self.obj:
            self.obj.close()

        self.root.destroy()


class SettingsGUI:
    entry1 = None
    entry2 = None
    entry3 = None
    listbox = None
    selection = 0

    def __init__(self, root, table):
        assert root
        assert table
        self.dialog = Toplevel(root)
        self.rvar = IntVar()
        self.evar = StringVar()
        self.table = table
        configs = Global_Objs['Local_Settings'].grab_item('Accdb_Configs')

        for config in configs:
            if config[0] == table:
                self.config = config
                break

        assert self.config

    def buildgui(self):
        self.dialog.geometry('700x295+500+300')
        self.dialog.title('Update TBL Settings')

        top_frame = Frame(self.dialog)
        middle_frame_main = Frame(self.dialog)
        middle_frame_right_main = Frame(self.dialog)
        middle_frame_right_submain = Frame(self.dialog)
        middle_frame = Frame(self.dialog)
        middle_frame2 = Frame(self.dialog)
        middle_frame3 = Frame(self.dialog)
        data_frame1 = Frame(self.dialog)
        data_frame2 = Frame(self.dialog)
        data_frame3 = Frame(self.dialog)
        bottom_frame = Frame(self.dialog)
        top_frame.pack()
        middle_frame.pack()
        middle_frame_main.pack()
        middle_frame2.pack(in_=middle_frame_main, side=LEFT)
        middle_frame_right_main.pack(in_=middle_frame_main, side=RIGHT)
        middle_frame3.pack(in_=middle_frame_right_main, side=LEFT)
        middle_frame_right_submain.pack(in_=middle_frame_right_main, side=RIGHT)
        data_frame1.pack(in_=middle_frame_right_submain)
        data_frame2.pack(in_=middle_frame_right_submain)
        data_frame3.pack(in_=middle_frame_right_submain)
        bottom_frame.pack(fill=BOTH)

        header = Message(self.dialog, text='Please input custom settings for access table {0}:'.format(self.table),
                         width=675)
        header.pack(in_=top_frame)

        label1 = Label(self.dialog, text='SQL Server TBL: ', pady=7)
        self.entry1 = Entry(self.dialog, width=40)
        label1.pack(in_=middle_frame, side=LEFT)
        self.entry1.pack(in_=middle_frame, side=RIGHT)

        scrollbar = Scrollbar(self.dialog, orient="vertical")
        scrollbar2 = Scrollbar(self.dialog, orient="horizontal")
        self.listbox = Listbox(self.dialog, selectmode=SINGLE, width=60, yscrollcommand=scrollbar.set,
                               xscrollcommand=scrollbar2.set)

        self.populatefields()

        scrollbar.config(command=self.listbox.yview)
        scrollbar2.config(command=self.listbox.xview)
        scrollbar.pack(in_=middle_frame2, side="right", fill="y")
        scrollbar2.pack(in_=middle_frame2, side="bottom", fill="x")
        self.listbox.pack(in_=middle_frame2, side="left", pady=5)
        self.listbox.bind("<Down>", self.on_list_down)
        self.listbox.bind("<Up>", self.on_list_up)
        self.listbox.bind('<<ListboxSelect>>', self.on_select)
        self.listbox.select_set(self.selection)

        edit_but = Button(self.dialog, text='Edit', width=6, command=self.edit)
        edit_but.pack(in_=middle_frame3, pady=5, padx=5)
        del_but = Button(self.dialog, text='Delete', width=6, command=self.delete)
        del_but.pack(in_=middle_frame3, pady=5, padx=5)

        label2 = Label(self.dialog, text='Access TBL Col: ', padx=10, pady=7)
        self.entry2 = Entry(self.dialog)
        label2.pack(in_=data_frame1, side=LEFT)
        self.entry2.pack(in_=data_frame1, side=RIGHT)
        label3 = Label(self.dialog, text='SQL TBL Col: ', padx=16, pady=7)
        self.entry3 = Entry(self.dialog)
        label3.pack(in_=data_frame2, side=LEFT)
        self.entry3.pack(in_=data_frame2, side=RIGHT)

        add_but = Button(self.dialog, text='Add', width=15, command=self.add)
        add_but.pack(in_=data_frame3, padx=5, pady=5)

        btn = Button(self.dialog, text="Save Settings", width=20, command=self.save_settings)
        btn2 = Button(self.dialog, text="Cancel", width=20, command=self.close)
        btn3 = Button(self.dialog, text="Delete Settings", width=20, command=self.del_settings)
        btn.pack(in_=bottom_frame, side=LEFT, padx=10, pady=10)
        btn2.pack(in_=bottom_frame, side=RIGHT, padx=10, pady=10)
        btn3.pack(in_=bottom_frame, side=TOP, pady=10)

        self.dialog.mainloop()

    def add(self):
        print('adding item')

    def edit(self):
        print('editing item')

    def delete(self):
        print('deleting item')

    def populatefields(self):
        self.entry1.insert(0, self.config[2])
        for from_col, to_col in zip(self.config[1], self.config[3]):
            self.listbox.insert("end", '{0} => {1}'.format(from_col, to_col))

    def on_select(self, event):
        if self.listbox and self.listbox.curselection() and -1 < self.selection < self.listbox.size() - 1:
            self.selection = self.listbox.curselection()[0]

    def on_list_down(self, event):
        if self.selection < self.listbox.size()-1:
            self.listbox.select_clear(self.selection)
            self.selection += 1
            self.listbox.select_set(self.selection)

    def on_list_up(self, event):
        if self.selection > 0:
            self.listbox.select_clear(self.selection)
            self.selection -= 1
            self.listbox.select_set(self.selection)

    def del_settings(self):
        print('deleting setting')

    def save_settings(self):
        if self.entry1.get() == "General_Settings_Path":
            messagebox.showerror('Table Error', 'Unable to have General_Settings_Path as a table name')
        elif self.entry1.get() == 'Server':
            messagebox.showerror('Table Error', 'Unable to have Server as a table name')
        elif self.entry1.get() == 'Database':
            messagebox.showerror('Table Error', 'Unable to have Database as a table name')
        else:
            if self.rvar.get() == 1:
                myitems = [True, self.entry2.get()]
            else:
                myitems = [False, self.entry2.get()]

            if self.entry1.get() in self.local_settings:
                Global_Objs['Local_Settings'].del_item(self.entry1.get())

            if self.rvar.get() != 1 or self.entry2.get() != '14':
                Global_Objs['Local_Settings'].add_item(self.entry1.get(), myitems)

            self.dialog.destroy()

    def close(self):
        self.dialog.destroy()


if __name__ == '__main__':
    myobj = MainGUI()
    myobj.buildgui()
    myobj.showgui()

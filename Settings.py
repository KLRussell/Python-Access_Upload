# Global Module import
from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle
from Global import ShelfHandle

import os

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(__file__))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir)


class SettingsGUI:
    insert = False
    atc_list_sel = 0
    atcs_list_sel = 0
    stc_list_sel = 0
    stcs_list_sel = 0
    save_button = None
    atc_list_box = None
    atcs_list_box = None
    acc_right_button = None
    acc_all_right_button = None
    acc_left_button = None
    acc_all_left_button = None
    stc_list_box = None
    stcs_list_box = None
    sql_right_button = None
    sql_all_right_button = None
    sql_left_button = None
    sql_all_left_button = None
    sql_tbl_name_txtbox = None

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to Access DB Upload Settings!\nSettings can be changed below.\nPress save when finished'
        self.asql = global_objs['SQL']
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.acc_tbl_name = StringVar()
        self.sql_tbl_name = StringVar()

        # GUI Bind On Destruction event
        self.main.bind('<Destroy>', self.gui_destroy)

    # Function executes when GUI is destroyed
    def gui_destroy(self, event):
        self.asql.close()

    # Static function to fill textbox in GUI
    @staticmethod
    def fill_textbox(setting_list, val, key):
        assert (key and val and setting_list)
        item = global_objs[setting_list].grab_item(key)

        if isinstance(item, CryptHandle):
            val.set(item.decrypt_text())

    # static function to add setting to Local_Settings shelf files
    @staticmethod
    def add_setting(setting_list, val, key, encrypt=True):
        assert (key and val and setting_list)

        global_objs[setting_list].del_item(key)
        global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)

    # Function to build GUI for settings
    def build_gui(self, header=None, insert=False):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        self.insert = insert

        # Set GUI Geometry and GUI Title
        self.main.geometry('497x645+500+90')
        self.main.title('Access DB Upload Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=444, height=70)
        add_upload_frame = LabelFrame(self.main, text='Add Upload Settings', width=444, height=210)
        acc_tbl_frame = LabelFrame(add_upload_frame, text='Access Table', width=444, height=105)
        acc_tbl_top_frame = Frame(acc_tbl_frame)
        acc_tbl_bottom_frame = Frame(acc_tbl_frame)
        sql_tbl_frame = LabelFrame(add_upload_frame, text='SQL Table', width=444, height=105)
        sql_tbl_top_frame = Frame(sql_tbl_frame)
        sql_tbl_bottom_frame = Frame(sql_tbl_frame)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        add_upload_frame.pack(fill="both")
        acc_tbl_frame.grid(row=0)
        acc_tbl_top_frame.grid(row=0, sticky=W+E)
        acc_tbl_bottom_frame.grid(row=1, sticky=W+E)
        sql_tbl_frame.grid(row=1)
        sql_tbl_top_frame.grid(row=0, sticky=W + E)
        sql_tbl_bottom_frame.grid(row=1, sticky=W + E)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)
        server_txtbox.bind('<KeyRelease>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Widgets to the Acc_TBL_Frame
        #     Access Table Name Input Box
        acc_tbl_name_label = Label(acc_tbl_top_frame, text='TBL Name:')
        acc_tbl_name_txtbox = Entry(acc_tbl_top_frame, textvariable=self.acc_tbl_name, width=64)
        acc_tbl_name_label.grid(row=0, column=0, padx=8, pady=5)
        acc_tbl_name_txtbox.grid(row=0, column=1, padx=5, pady=5)
        acc_tbl_name_txtbox.configure(state=DISABLED)

        #     Access Table Column List
        atc_xscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='horizontal')
        atc_yscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='vertical')
        self.atc_list_box = Listbox(acc_tbl_bottom_frame, selectmode=SINGLE, width=30, yscrollcommand=atc_yscrollbar,
                                    xscrollcommand=atc_xscrollbar)
        atc_xscrollbar.config(command=self.atc_list_box.xview)
        atc_yscrollbar.config(command=self.atc_list_box.yview)
        self.atc_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        atc_xscrollbar.grid(row=4, column=0, sticky=W+E)
        atc_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N+S)
        self.atc_list_box.bind("<Down>", self.atc_list_down)
        self.atc_list_box.bind("<Up>", self.atc_list_up)
        self.atc_list_box.bind('<<ListboxSelect>>', self.atc_select)

        #     Access Column Migration Buttons
        self.acc_right_button = Button(acc_tbl_bottom_frame, text='>', width=5, command=self.acc_right_migrate)
        self.acc_all_right_button = Button(acc_tbl_bottom_frame, text='>>', width=5, command=self.acc_all_right_migrate)
        self.acc_left_button = Button(acc_tbl_bottom_frame, text='<', width=5, command=self.acc_left_migrate)
        self.acc_all_left_button = Button(acc_tbl_bottom_frame, text='<<', width=5, command=self.acc_all_left_migrate)
        self.acc_right_button.grid(row=0, column=2, padx=5)
        self.acc_all_right_button.grid(row=1, column=2, padx=5)
        self.acc_left_button.grid(row=2, column=2, padx=5)
        self.acc_all_left_button.grid(row=3, column=2, padx=5)

        #     Access Table Column Selection List
        atcs_xscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='horizontal')
        atcs_yscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='vertical')
        self.atcs_list_box = Listbox(acc_tbl_bottom_frame, selectmode=SINGLE, width=30, yscrollcommand=atcs_yscrollbar,
                                     xscrollcommand=atcs_xscrollbar)
        atcs_xscrollbar.config(command=self.atcs_list_box.xview)
        atcs_yscrollbar.config(command=self.atcs_list_box.yview)
        self.atcs_list_box.grid(row=0, column=3, rowspan=4, padx=8, pady=5)
        atcs_xscrollbar.grid(row=4, column=3, sticky=W + E)
        atcs_yscrollbar.grid(row=0, column=4, rowspan=4, sticky=N + S)
        self.atcs_list_box.bind("<Down>", self.atcs_list_down)
        self.atcs_list_box.bind("<Up>", self.atcs_list_up)
        self.atcs_list_box.bind('<<ListboxSelect>>', self.atcs_select)

        # Apply Widgets to the SQL_TBL_Frame
        #     SQL Table Name Input Box
        sql_tbl_name_label = Label(sql_tbl_top_frame, text='TBL Name:')
        self.sql_tbl_name_txtbox = Entry(sql_tbl_top_frame, textvariable=self.sql_tbl_name, width=64)
        sql_tbl_name_label.grid(row=0, column=0, padx=8, pady=5)
        self.sql_tbl_name_txtbox.grid(row=0, column=1, padx=5, pady=5)

        #     SQL Table Column List
        stc_xscrollbar = Scrollbar(sql_tbl_bottom_frame, orient='horizontal')
        stc_yscrollbar = Scrollbar(sql_tbl_bottom_frame, orient='vertical')
        self.stc_list_box = Listbox(sql_tbl_bottom_frame, selectmode=SINGLE, width=30, yscrollcommand=stc_yscrollbar,
                                    xscrollcommand=stc_xscrollbar)
        stc_xscrollbar.config(command=self.stc_list_box.xview)
        stc_yscrollbar.config(command=self.stc_list_box.yview)
        self.stc_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        stc_xscrollbar.grid(row=4, column=0, sticky=W + E)
        stc_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N + S)
        self.stc_list_box.bind("<Down>", self.stc_list_down)
        self.stc_list_box.bind("<Up>", self.stc_list_up)
        self.stc_list_box.bind('<<ListboxSelect>>', self.stc_select)

        #     SQL Column Migration Buttons
        self.sql_right_button = Button(sql_tbl_bottom_frame, text='>', width=5, command=self.sql_right_migrate)
        self.sql_all_right_button = Button(sql_tbl_bottom_frame, text='>>', width=5, command=self.sql_all_right_migrate)
        self.sql_left_button = Button(sql_tbl_bottom_frame, text='<', width=5, command=self.sql_left_migrate)
        self.sql_all_left_button = Button(sql_tbl_bottom_frame, text='<<', width=5, command=self.sql_all_left_migrate)
        self.sql_right_button.grid(row=0, column=2, padx=5)
        self.sql_all_right_button.grid(row=1, column=2, padx=5)
        self.sql_left_button.grid(row=2, column=2, padx=5)
        self.sql_all_left_button.grid(row=3, column=2, padx=5)

        #     SQL Table Column Selection List
        stcs_xscrollbar = Scrollbar(sql_tbl_bottom_frame, orient='horizontal')
        stcs_yscrollbar = Scrollbar(sql_tbl_bottom_frame, orient='vertical')
        self.stcs_list_box = Listbox(sql_tbl_bottom_frame, selectmode=SINGLE, width=30, yscrollcommand=stcs_yscrollbar,
                                     xscrollcommand=atcs_xscrollbar)
        stcs_xscrollbar.config(command=self.stcs_list_box.xview)
        stcs_yscrollbar.config(command=self.stcs_list_box.yview)
        self.stcs_list_box.grid(row=0, column=3, rowspan=4, padx=8, pady=5)
        stcs_xscrollbar.grid(row=4, column=3, sticky=W + E)
        stcs_yscrollbar.grid(row=0, column=4, rowspan=4, sticky=N + S)
        self.stcs_list_box.bind("<Down>", self.stcs_list_down)
        self.stcs_list_box.bind("<Up>", self.stcs_list_up)
        self.stcs_list_box.bind('<<ListboxSelect>>', self.stcs_select)

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text='Save Settings', width=15, command=self.save_settings)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        extract_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        extract_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        #     Extract Shelf Button
        extract_button = Button(self.main, text='Upload Settings', width=15, command=self.upload_settings)
        extract_button.pack(in_=button_frame, side=TOP, padx=10, pady=5)

        # Fill Textboxes with settings
        self.fill_gui()

        # Show GUI Dialog
        self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')

        if not self.server.get() or not self.database.get() or not self.asql.test_conn('alch'):
            self.save_button.configure(state=DISABLED)
        else:
            self.asql.connect('alch')

        if not self.insert:
            self.atc_list_box.configure(state=DISABLED)
            self.atcs_list_box.configure(state=DISABLED)
            self.acc_right_button.configure(state=DISABLED)
            self.acc_all_right_button.configure(state=DISABLED)
            self.acc_left_button.configure(state=DISABLED)
            self.acc_all_left_button.configure(state=DISABLED)
            self.sql_tbl_name_txtbox.configure(state=DISABLED)

        self.stc_list_box.configure(state=DISABLED)
        self.stcs_list_box.configure(state=DISABLED)
        self.sql_right_button.configure(state=DISABLED)
        self.sql_all_right_button.configure(state=DISABLED)
        self.sql_left_button.configure(state=DISABLED)
        self.sql_all_left_button.configure(state=DISABLED)

    # Function to check network settings if populated
    def check_network(self, event):
        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.asql.test_conn('alch'):
                self.save_button.configure(state=NORMAL)
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')
                self.asql.connect('alch')

    # Function to validate whether a SQL table exists in SQL server
    def check_table(self, table):
        table2 = table.split('.')

        if len(table2) == 2:
            myresults = self.asql.query('''
                SELECT
                    1
                FROM information_schema.tables
                WHERE
                    Table_Schema = '{0}'
                        AND
                    Table_Name = '{1}'
            '''.format(table2[0], table2[1]))

            if myresults.empty:
                return False
            else:
                return True
        else:
            return False

    # Function to connect to SQL connection for this class
    def sql_connect(self):
        if self.asql.test_conn('alch'):
            self.asql.connect('alch')
            return True
        else:
            return False

    # Function to close SQL connection for this class
    def sql_close(self):
        self.asql.close()

    # Function adjusts selection of item when user clicks item (ATC List)
    def atc_select(self, event):
        if self.atc_list_box and self.atc_list_box.curselection() \
                and -1 < self.atc_list_sel < self.atc_list_box.size() - 1:
            self.atc_list_sel = self.atc_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (ATC List)
    def atc_list_down(self, event):
        if self.atc_list_sel < self.atc_list_box.size() - 1:
            self.atc_list_box.select_clear(self.atc_list_sel)
            self.atc_list_sel += 1
            self.atc_list_box.select_set(self.atc_list_sel)

    # Function adjusts selection of item when user presses up key (ATC List)
    def atc_list_up(self, event):
        if self.atc_list_sel > 0:
            self.atc_list_box.select_clear(self.atc_list_sel)
            self.atc_list_sel -= 1
            self.atc_list_box.select_set(self.atc_list_sel)

    # Function adjusts selection of item when user clicks item (ATCS List)
    def atcs_select(self, event):
        if self.atcs_list_box and self.atcs_list_box.curselection() \
                and -1 < self.atcs_list_sel < self.atcs_list_box.size() - 1:
            self.atcs_list_sel = self.atcs_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (ATCS List)
    def atcs_list_down(self, event):
        if self.atcs_list_sel < self.atcs_list_box.size() - 1:
            self.atcs_list_box.select_clear(self.atcs_list_sel)
            self.atcs_list_sel += 1
            self.atcs_list_box.select_set(self.atcs_list_sel)

    # Function adjusts selection of item when user presses up key (ATCS List)
    def atcs_list_up(self, event):
        if self.atcs_list_sel > 0:
            self.atcs_list_box.select_clear(self.atcs_list_sel)
            self.atcs_list_sel -= 1
            self.atcs_list_box.select_set(self.atcs_list_sel)

    # Function adjusts selection of item when user clicks item (STC List)
    def stc_select(self, event):
        if self.stc_list_box and self.stc_list_box.curselection() \
                and -1 < self.stc_list_sel < self.stc_list_box.size() - 1:
            self.stc_list_sel = self.stc_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (STC List)
    def stc_list_down(self, event):
        if self.stc_list_sel < self.stc_list_box.size() - 1:
            self.stc_list_box.select_clear(self.stc_list_sel)
            self.stc_list_sel += 1
            self.stc_list_box.select_set(self.stc_list_sel)

    # Function adjusts selection of item when user presses up key (STC List)
    def stc_list_up(self, event):
        if self.stc_list_sel > 0:
            self.stc_list_box.select_clear(self.stc_list_sel)
            self.stc_list_sel -= 1
            self.stc_list_box.select_set(self.stc_list_sel)

    # Function adjusts selection of item when user clicks item (STCS List)
    def stcs_select(self, event):
        if self.stcs_list_box and self.stcs_list_box.curselection() \
                and -1 < self.stcs_list_sel < self.stcs_list_box.size() - 1:
            self.stcs_list_sel = self.stcs_list_box.curselection()[0]

    # Function adjusts selection of item when user presses down key (STCS List)
    def stcs_list_down(self, event):
        if self.stcs_list_sel < self.stcs_list_box.size() - 1:
            self.stcs_list_box.select_clear(self.stcs_list_sel)
            self.stcs_list_sel += 1
            self.stcs_list_box.select_set(self.stcs_list_sel)

    # Function adjusts selection of item when user presses up key (STCS List)
    def stcs_list_up(self, event):
        if self.stcs_list_sel > 0:
            self.stcs_list_box.select_clear(self.stcs_list_sel)
            self.stcs_list_sel -= 1
            self.stcs_list_box.select_set(self.stcs_list_sel)

    # Button to migrate single record to right list (Access TBL Section)
    def acc_right_migrate(self):
        print('migrate to right')

    # Button to migrate single record to right list (Access TBL Section)
    def acc_all_right_migrate(self):
        print('migrate all to right')

    # Button to migrate single record to right list (Access TBL Section)
    def acc_left_migrate(self):
        print('migrate to left')

    # Button to migrate single record to right list (Access TBL Section)
    def acc_all_left_migrate(self):
        print('migrate all to left')

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_right_migrate(self):
        print('migrate to right')

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_all_right_migrate(self):
        print('migrate all to right')

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_left_migrate(self):
        print('migrate to left')

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_all_left_migrate(self):
        print('migrate all to left')

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if self.server.get() and self.database.get():
            if not self.sql_tbl.get():
                messagebox.showerror('SQL TBL Empty Error!', 'No value has been inputed for SQL TBL',
                                     parent=self.main)
            elif not self.shelf_life.get():
                messagebox.showerror('Shelf Life Empty Error!', 'No value has been inputed for Shelf Life',
                                     parent=self.main)
            elif self.shelf_life.get() <= 0:
                messagebox.showerror('Invalid Shelf Life Error!', 'Shelf Life <= 0',
                                     parent=self.main)
            else:
                if not self.check_table(self.sql_tbl.get()):
                    messagebox.showerror('Invalid SQL TBL!',
                                         'SQL TBL does not exist in sql server',
                                         parent=self.main)
                else:
                    if self.autofill.get() == 1:
                        myitems = [True, self.shelf_life.get()]
                    else:
                        myitems = [False, self.shelf_life.get()]

                    if self.autofill.get() != 1 or self.shelf_life.get() != 14:
                        self.add_setting('Local_Settings', myitems, self.sql_tbl.get(), False)
                    else:
                        global_objs['Local_Settings'].del_item(self.sql_tbl.get())

                    self.main.destroy()

    # Function to load extract Shelf GUI
    def upload_settings(self):
        if self.shelf_obj:
            self.shelf_obj.cancel()

        self.shelf_obj = ExtractShelf(self.main)
        self.shelf_obj.build_gui()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    obj = SettingsGUI()

    try:
        obj.build_gui()
    finally:
        obj.sql_close()
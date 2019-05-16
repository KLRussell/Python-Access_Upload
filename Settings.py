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
    save_button = None
    atc_list_box = None

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to Access DB Upload Settings!\nSettings can be changed below.\nPress save when finished'
        self.asql = global_objs['SQL']
        self.main = Tk()

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.tbl_name = StringVar()

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
    def build_gui(self, header=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('444x357+500+160')
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
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        add_upload_frame.pack(fill="both")
        acc_tbl_frame.grid(row=0)
        acc_tbl_top_frame.grid(row=0)
        acc_tbl_bottom_frame.grid(row=1)
        sql_tbl_frame.grid(row=1)
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
        tbl_name_label = Label(acc_tbl_top_frame, text='TBL Name:')
        tbl_name_txtbox = Entry(acc_tbl_top_frame, textvariable=self.tbl_name, width=55)
        tbl_name_label.grid(row=0, column=0, padx=8, pady=5)
        tbl_name_txtbox.grid(row=0, column=0, padx=5, pady=5)
        tbl_name_txtbox.configure(state=DISABLED)

        #     Access Table Column List
        atc_xscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='horizontal')
        atc_yscrollbar = Scrollbar(acc_tbl_bottom_frame, orient='vertical')
        self.atc_list_box = Listbox(acc_tbl_bottom_frame, selectmode=SINGLE, width=35, yscrollcommand=atc_yscrollbar,
                                    xscrollcommand=atc_xscrollbar)
        atc_xscrollbar.config(command=self.atc_list_box.xview)
        atc_yscrollbar.config(command=self.atc_list_box.yview)
        self.atc_list_box.grid(row=0, column=0, rowspan=4, padx=8, pady=5)
        atc_xscrollbar.grid(row=5, column=0, sticky=W+E)
        atc_yscrollbar.grid(row=0, column=1, rowspan=4, sticky=N+S)

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text='Save Settings', width=15, command=self.save_settings)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        extract_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        extract_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        #     Extract Shelf Button
        extract_button = Button(self.main, text='Extract Shelf', width=15, command=self.extract_shelf)
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
    def extract_shelf(self):
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

# Global Module import
from tkinter import *
from tkinter import messagebox
from Global import grabobjs
from Global import CryptHandle
from Global import SQLHandle
from dateutil import relativedelta
from datetime import datetime

import os
import smtplib
import pandas as pd
import ftplib
import sys
import pathlib as pl
import portalocker

if getattr(sys, 'frozen', False):
    application_path = sys.executable
    ico_dir = sys._MEIPASS
else:
    application_path = __file__
    ico_dir = os.path.dirname(__file__)

# Global Variable declaration
curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
global_objs = grabobjs(main_dir, 'STC_Upload')
ProcessedDir = os.path.join(main_dir, "03_Processed")
icon_path = os.path.join(ico_dir, '%s.ico' % os.path.splitext(os.path.basename(application_path))[0])
SQLDir = os.path.join(main_dir, "05_SQL")


class AccdbHandle:
    config = None
    accdb_cols = None
    sql_cols = None
    upload_df = pd.DataFrame()

    def __init__(self, file, batch):
        self.file = file
        self.batch = batch
        self.log = global_objs['Event_Log']
        self.settings = global_objs['Local_Settings']
        self.settings.read_shelf()
        self.configs = self.settings.grab_item('Accdb_Configs')
        self.accdb = SQLHandle(logobj=self.log, accdb_file=file)
        self.asql = SQLHandle(logobj=self.log, settingsobj=global_objs['Settings'])
        self.asql.connect(conn_type='alch')
        self.accdb.connect(conn_type='accdb')
        self.sql_tables = self.asql.tables()
        self.accdb_tables = self.accdb.tables()

    def get_accdb_tables(self):
        return self.accdb_tables

    @staticmethod
    def validate_cols(cols, true_cols):
        if cols:
            for col in cols:
                found = False

                for true_col in true_cols:
                    if col.lower() == true_col.lower():
                        found = True
                        break

                if not found:
                    return False

            return True

    def process_err(self, table, header, insert):
        config_review = self.settings.grab_item('Config_Review')

        if config_review:
            for config in config_review:
                if config[1].lower() == table.lower():
                    return
        else:
            config_review = []

        config_review.append([self.file, table, self.accdb_cols, header, insert, self.config])
        self.settings.add_item('Config_Review', config_review)
        self.settings.write_shelf()

    def get_config(self, table):
        self.config = None

        if table and self.configs:
            for config in self.configs:
                if config[0].lower() == table.lower():
                    self.config = config
                    break

    def get_accdb_cols(self, table):
        myresults = self.accdb.execute('''
            SELECT TOP 1 *
            FROM [{0}]'''.format(table))

        if len(myresults) < 1:
            raise Exception('File {0} has no valid columns in table {1}'.format(
                os.path.basename(self.file), table))

        self.accdb_cols = myresults.columns.tolist()

    def get_sql_cols(self, table):
        myresults = self.asql.execute('''
            SELECT
                Column_Name

            FROM INFORMATION_SCHEMA.COLUMNS

            WHERE
                TABLE_SCHEMA = '{0}'
                    AND
                TABLE_NAME = '{1}'
        '''.format(table.split('.')[0], table.split('.')[1]))

        if len(myresults) < 1:
            raise Exception('SQL Table {} has no valid tables in the access database file'.format(table))

        self.sql_cols = myresults['Column_Name'].tolist()

    def validate_sql_table(self, table):
        if table:
            splittable = table.split('.')

            if len(splittable) == 2:
                for table in self.sql_tables:
                    if table[0] == 'TABLE' and table[1][1].lower() == splittable[0].lower()\
                            and table[1][2].lower() == splittable[1].lower():
                        return True

    def validate(self, table):
        header = None
        insert = False
        self.get_config(table)
        self.get_accdb_cols(table)

        if not self.config:
            header = ['Welcome to STC Upload!', 'There is no configuration for table.',
                      'Please add configuration setting below:']
            insert = True
        elif self.config and not self.validate_sql_table(self.config[3]):
            header = ['Welcome to STC Upload!', 'SQL Server TBL does not exist.',
                      'Please fix configuration in Upload Settings:']
        elif self.config and not self.validate_cols(self.config[2], self.accdb_cols):
            header = ['Welcome to STC Upload!', 'One or more column does not exist.',
                      'Please redo config for access table columns:']
        else:
            self.get_sql_cols(self.config[3])

            if not self.validate_cols(self.config[4], self.sql_cols):
                header = ['Welcome to STC Upload!', 'One or more column does not exist.',
                          'Please redo config for sql table columns:']

            if not self.validate_cols(['Source_File'], self.sql_cols):
                header = ['Welcome to STC Upload!', 'SQL Table does not have a Source_File column:']

        if header:
            self.log.write_log('Access table {0} failed validation [{1}]. Appending to GUI for fix'.format(
                table, '.'.join(header)), 'warning')
            self.process_err(table, '\n'.join(header), insert)
            return False
        elif self.config:
            if self.updates_not_exists():
                return True
            else:
                self.log.write_log('Updates already exist for SQL table %s' % self.config[3], 'warning')
                return False

    def updates_not_exists(self):
        return self.asql.execute('''
            SELECT TOP 1 1
            FROM {0}
            WHERE
                Source_File = 'Updated Records {1}'
        '''.format(self.config[3], self.batch)).empty

    '''
    def switch_config(self):
        if self.configs and self.config:
            for config in self.configs:
                if config == self.config:
                    self.configs.remove(config)
                    break

            self.configs.append(self.config)
            global_objs['Local_Settings'].del_item('Accdb_Configs')
            global_objs['Local_Settings'].add_item('Accdb_Configs', self.configs)

    def config_gui(self, table, header_text, insert=True):
        if insert:
            obj = AccSettingsGUI()
            obj.build_gui(header_text, table, self.accdb_cols)
            self.get_config(table)
        else:
            old_config = self.config
            self.switch_config()
            obj = AccSettingsGUI(config=old_config)

            obj.build_gui(header_text)
            self.get_config(table)

            if old_config == self.config:
                self.config = None
    '''

    def process(self, table):
        self.log.write_log('Reading data from accdb table [%s]' % table)

        self.upload_df = self.accdb.execute('''
            SELECT
                [{0}],
                'Updated Records {2}' As Source_File

            FROM [{1}]
        '''.format('], ['.join(self.config[2]), table, self.batch))

        if not self.upload_df.empty:
            self.upload_df.columns = self.config[4] + ('Source_File',)

            if self.config[5]:
                self.log.write_log('Truncating table [%s]' % self.config[3])
                self.asql.execute('truncate table %s' % self.config[3], execute=True)

            self.log.write_log('Uploading data to sql table %s' % self.config[3])
            self.asql.upload_df(self.upload_df, self.config[3], index=False, index_label=None)
            self.log.write_log('Data successfully uploaded from table [{0}] to sql table {1}'.format(
                table, self.config[3]))
            return True
        else:
            self.log.write_log('Failed to grab data from access table [{0}]. No update made'.format(table), 'error')
            return False

    def apply_features(self, table):
        self.log.write_log('Appending features to %s if they exist' % self.config[3])

        if self.validate_cols(['Edit_DT'], self.sql_cols):
            self.asql.execute('''
                UPDATE %s
                SET
                    Edit_DT = GETDATE()
            ''' % self.config[3], execute=True)
        elif self.validate_cols(['Edit_Date'], self.sql_cols):
            self.asql.execute('''
                UPDATE %s
                SET
                    Edit_Date = GETDATE()
            ''' % self.config[3], execute=True)

        path = os.path.join(SQLDir, table)

        if os.path.exists(path):
            for file in list(pl.Path(path).glob('*.sql')):
                with portalocker.Lock(str(file), 'r') as f:
                    query = f.read()

                self.asql.execute(query, execute=True)

    def close_conn(self):
        self.accdb.close_conn()
        self.asql.close_conn()

    def get_success_stats(self, table):
        return [table, self.config[3], len(self.upload_df)]


class SettingsGUI:
    change_upload_settings_obj = None
    save_button = None
    cancel_button = None
    upload_button = None
    fserver_txtbox = None
    fuser_txtbox = None
    fpass_txtbox = None
    eserver_txtbox = None
    eport_txtbox = None
    euname_txtbox = None
    eupass_txtbox = None
    eto_txtbox = None
    efrom_txtbox = None
    ecc_txtbox = None

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self):
        self.header_text = 'Welcome to STC Upload Settings!\nSettings can be changed below.\nPress save when finished'

        self.fpass_obj = global_objs['Local_Settings'].grab_item('FTP Password')
        self.epass_obj = global_objs['Settings'].grab_item('Email_Pass')
        self.asql = global_objs['SQL']
        self.main = Tk()
        self.main.iconbitmap(icon_path)

        # GUI Variables
        self.server = StringVar()
        self.database = StringVar()
        self.fserver = StringVar()
        self.fuser = StringVar()
        self.fpass = StringVar()
        self.email_server = StringVar()
        self.email_port = StringVar()
        self.email_user_name = StringVar()
        self.email_user_pass = StringVar()
        self.email_from = StringVar()
        self.email_to = StringVar()
        self.email_cc = StringVar()

        try:
            self.asql.connect('alch')
            self.sql_tables = self.asql.tables()
        except:
            self.sql_tables = None
        finally:
            self.asql.close_conn()

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
        assert (key and setting_list)

        global_objs[setting_list].del_item(key)

        if val:
            global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)
            global_objs[setting_list].write_shelf()

    # Function to connect to SQL connection for this class
    def test_sql_conn(self):
        if self.asql.connect('alch', test_conn=True):
            return True
        else:
            return False

    # Function to build GUI for settings
    def build_gui(self, header=None, acc_table=None, acc_cols=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        # Set GUI Geometry and GUI Title
        self.main.geometry('540x367+500+150')
        self.main.title('STC Upload Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        network_frame = LabelFrame(self.main, text='Network Settings', width=444, height=70)
        ftp_frame = LabelFrame(self.main, text="FTP Settings", width=444, height=70)
        email_frame = LabelFrame(self.main, text='E-mail', width=444, height=170)
        econn_frame = LabelFrame(email_frame, text='Settings', width=222, height=170)
        emsg_frame = LabelFrame(email_frame, text='Message', width=223, height=170)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        network_frame.pack(fill="both")
        ftp_frame.pack(fill="both")
        email_frame.pack(fill="both")
        econn_frame.grid(row=0, column=0, ipady=5)
        emsg_frame.grid(row=0, column=1, ipady=20)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=510, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Network Labels & Input boxes to the Network_Frame
        #     SQL Server Input Box
        server_label = Label(self.main, text='Server:', padx=15, pady=7)
        server_txtbox = Entry(self.main, textvariable=self.server, width=25)
        server_label.pack(in_=network_frame, side=LEFT)
        server_txtbox.pack(in_=network_frame, side=LEFT)
        server_txtbox.bind('<FocusOut>', self.check_network)

        #     Server Database Input Box
        database_label = Label(self.main, text='Database:')
        database_txtbox = Entry(self.main, textvariable=self.database, width=25)
        database_txtbox.pack(in_=network_frame, side=RIGHT, pady=7, padx=15)
        database_label.pack(in_=network_frame, side=RIGHT)
        database_txtbox.bind('<KeyRelease>', self.check_network)

        # Apply Widgets to the FTP_Frame
        #     Host Server Input Box
        fserver_label = Label(ftp_frame, text='Server:')
        self.fserver_txtbox = Entry(ftp_frame, textvariable=self.fserver, width=15)
        fserver_label.grid(row=0, column=0, padx=8, pady=5)
        self.fserver_txtbox.grid(row=0, column=1, padx=8, pady=5)

        #     Host User Input Box
        fuser_label = Label(ftp_frame, text='Username:')
        self.fuser_txtbox = Entry(ftp_frame, textvariable=self.fuser, width=15)
        fuser_label.grid(row=0, column=2, padx=8, pady=5)
        self.fuser_txtbox.grid(row=0, column=3, padx=8, pady=5)

        #     Host Pass Input Box
        fpass_label = Label(ftp_frame, text='Password:')
        self.fpass_txtbox = Entry(ftp_frame, textvariable=self.fpass, width=15)
        fpass_label.grid(row=0, column=4, padx=8, pady=5)
        self.fpass_txtbox.grid(row=0, column=5, padx=8, pady=5)
        self.fpass_txtbox.bind('<KeyRelease>', self.fhide_pass)

        # Apply Email Labels & Input boxes to the EConn_Frame
        #     Email Server Input Box
        eserver_label = Label(econn_frame, text='Email Server:')
        self.eserver_txtbox = Entry(econn_frame, textvariable=self.email_server)
        eserver_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        self.eserver_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     Email Port Input Box
        eport_label = Label(econn_frame, text='Email Port:')
        self.eport_txtbox = Entry(econn_frame, textvariable=self.email_port)
        eport_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        self.eport_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     Email User Name Input Box
        euname_label = Label(econn_frame, text='Email User Name:')
        self.euname_txtbox = Entry(econn_frame, textvariable=self.email_user_name)
        euname_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        self.euname_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        #     Email User Pass Input Box
        eupass_label = Label(econn_frame, text='Email User Pass:')
        self.eupass_txtbox = Entry(econn_frame, textvariable=self.email_user_pass)
        eupass_label.grid(row=3, column=0, padx=8, pady=5, sticky='w')
        self.eupass_txtbox.grid(row=3, column=1, padx=13, pady=5, sticky='e')
        self.eupass_txtbox.bind('<KeyRelease>', self.ehide_pass)

        # Apply Email Labels & Input boxes to the EMsg_Frame
        #     From Email Address Input Box
        efrom_label = Label(emsg_frame, text='From Addr:')
        self.efrom_txtbox = Entry(emsg_frame, textvariable=self.email_from)
        efrom_label.grid(row=0, column=0, padx=8, pady=5, sticky='w')
        self.efrom_txtbox.grid(row=0, column=1, padx=13, pady=5, sticky='e')

        #     To Email Address Input Box
        eto_label = Label(emsg_frame, text='To Addr:')
        self.eto_txtbox = Entry(emsg_frame, textvariable=self.email_to)
        eto_label.grid(row=1, column=0, padx=8, pady=5, sticky='w')
        self.eto_txtbox.grid(row=1, column=1, padx=13, pady=5, sticky='e')

        #     CC Email Address Input Box
        ecc_label = Label(emsg_frame, text='CC Addr:')
        self.ecc_txtbox = Entry(emsg_frame, textvariable=self.email_cc)
        ecc_label.grid(row=2, column=0, padx=8, pady=5, sticky='w')
        self.ecc_txtbox.grid(row=2, column=1, padx=13, pady=5, sticky='e')

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text='Save Settings', width=15, command=self.save_settings)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        self.cancel_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        self.cancel_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        #     Upload Settings Button
        self.upload_button = Button(self.main, text='Upload Settings', width=15, command=self.upload_settings)
        self.upload_button.pack(in_=button_frame, side=TOP, padx=10, pady=5)

        # Fill Textboxes with settings
        self.fill_gui()

        # Show GUI Dialog
        self.main.mainloop()

        # Function to fill GUI textbox fields

    def fill_gui(self):
        self.fill_textbox('Settings', self.server, 'Server')
        self.fill_textbox('Settings', self.database, 'Database')
        self.fill_textbox('Local_Settings', self.fserver, 'FTP Host')
        self.fill_textbox('Local_Settings', self.fuser, 'FTP User')
        self.fill_textbox('Settings', self.email_server, 'Email_Server')
        self.fill_textbox('Settings', self.email_port, 'Email_Port')
        self.fill_textbox('Settings', self.email_user_name, 'Email_User')
        self.fill_textbox('Local_Settings', self.email_from, 'Email_From')
        self.fill_textbox('Local_Settings', self.email_to, 'Email_To')
        self.fill_textbox('Local_Settings', self.email_cc, 'Email_CC')

        if self.epass_obj:
            self.email_user_pass.set('*' * len(self.epass_obj.decrypt_text()))

        if self.fpass_obj:
            self.fpass.set('*' * len(self.fpass_obj.decrypt_text()))

        if not self.email_port.get():
            self.email_port.set("587")

        if not self.server.get() or not self.database.get() or not self.test_sql_conn():
            self.save_button.configure(state=DISABLED)
            self.upload_button.configure(state=DISABLED)
            self.fserver_txtbox.configure(stat=DISABLED)
            self.fuser_txtbox.configure(stat=DISABLED)
            self.fpass_txtbox.configure(stat=DISABLED)
            self.eserver_txtbox.configure(stat=DISABLED)
            self.eport_txtbox.configure(stat=DISABLED)
            self.euname_txtbox.configure(stat=DISABLED)
            self.eupass_txtbox.configure(stat=DISABLED)
            self.efrom_txtbox.configure(stat=DISABLED)
            self.eto_txtbox.configure(stat=DISABLED)
            self.ecc_txtbox.configure(stat=DISABLED)

    # Function to check network settings if populated
    def check_network(self, event):
        flag = False

        if self.server.get() and self.database.get() and \
                (global_objs['Settings'].grab_item('Server') != self.server.get() or
                 global_objs['Settings'].grab_item('Database') != self.database.get()):
            self.asql.change_config(server=self.server.get(), database=self.database.get())

            if self.test_sql_conn():
                self.add_setting('Settings', self.server.get(), 'Server')
                self.add_setting('Settings', self.database.get(), 'Database')
                self.save_button.configure(state=NORMAL)
                self.upload_button.configure(state=NORMAL)
                self.fserver_txtbox.configure(stat=NORMAL)
                self.fuser_txtbox.configure(stat=NORMAL)
                self.fpass_txtbox.configure(stat=NORMAL)
                self.eserver_txtbox.configure(stat=NORMAL)
                self.eport_txtbox.configure(stat=NORMAL)
                self.euname_txtbox.configure(stat=NORMAL)
                self.eupass_txtbox.configure(stat=NORMAL)
                self.efrom_txtbox.configure(stat=NORMAL)
                self.eto_txtbox.configure(stat=NORMAL)
                self.ecc_txtbox.configure(stat=NORMAL)
            else:
                flag = True
        else:
            flag = True

        if flag:
            self.save_button.configure(state=DISABLED)
            self.upload_button.configure(state=DISABLED)
            self.fserver_txtbox.configure(stat=DISABLED)
            self.fuser_txtbox.configure(stat=DISABLED)
            self.fpass_txtbox.configure(stat=DISABLED)
            self.eserver_txtbox.configure(stat=DISABLED)
            self.eport_txtbox.configure(stat=DISABLED)
            self.euname_txtbox.configure(stat=DISABLED)
            self.eupass_txtbox.configure(stat=DISABLED)
            self.efrom_txtbox.configure(stat=DISABLED)
            self.eto_txtbox.configure(stat=DISABLED)
            self.ecc_txtbox.configure(stat=DISABLED)

    def fhide_pass(self, event):
        if self.fpass_obj:
            currpass = self.fpass_obj.decrypt_text()

            if len(self.fpass.get()) > len(currpass):
                i = 0

                for letter in self.fpass.get():
                    if letter != '*':
                        if i > len(currpass) - 1:
                            currpass += letter
                        else:
                            mytext = list(currpass)
                            mytext.insert(i, letter)
                            currpass = ''.join(mytext)
                    i += 1
            elif len(self.fpass.get()) > 0:
                i = 0

                for letter in self.fpass.get():
                    if letter != '*':
                        mytext = list(currpass)
                        mytext[i] = letter
                        currpass = ''.join(mytext)
                    i += 1

                if len(currpass) - i > 0:
                    currpass = currpass[:i]
            else:
                currpass = None

            if currpass:
                self.fpass_obj.encrypt_text(currpass)
                self.fpass.set('*' * len(self.fpass_obj.decrypt_text()))
            else:
                self.fpass_obj = None
                self.fpass.set("")
        else:
            self.fpass_obj = CryptHandle()
            self.fpass_obj.encrypt_text(self.fpass.get())
            self.fpass.set('*' * len(self.fpass_obj.decrypt_text()))

    def ehide_pass(self, event):
        if self.epass_obj:
            currpass = self.epass_obj.decrypt_text()

            if len(self.email_user_pass.get()) > len(currpass):
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        if i > len(currpass) - 1:
                            currpass += letter
                        else:
                            mytext = list(currpass)
                            mytext.insert(i, letter)
                            currpass = ''.join(mytext)
                    i += 1
            elif len(self.email_user_pass.get()) > 0:
                i = 0

                for letter in self.email_user_pass.get():
                    if letter != '*':
                        mytext = list(currpass)
                        mytext[i] = letter
                        currpass = ''.join(mytext)
                    i += 1

                if len(currpass) - i > 0:
                    currpass = currpass[:i]
            else:
                currpass = None

            if currpass:
                self.epass_obj.encrypt_text(currpass)
                self.email_user_pass.set('*' * len(self.epass_obj.decrypt_text()))
            else:
                self.epass_obj = None
                self.email_user_pass.set("")
        else:
            self.epass_obj = CryptHandle()
            self.epass_obj.encrypt_text(self.email_user_pass.get())
            self.email_user_pass.set('*' * len(self.epass_obj.decrypt_text()))

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if not self.fserver.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'FTP Server textbox is empty',
                                 parent=self.main)
        elif not self.fuser.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'FTP User textbox is empty',
                                 parent=self.main)
        elif not self.fpass_obj.decrypt_text():
            messagebox.showerror('Txtbox Empty Error!',
                                 'FTP Pass textbox is empty',
                                 parent=self.main)
        elif not self.email_server.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'Email Server textbox is empty',
                                 parent=self.main)
        elif not self.email_user_name.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'Email User Name textbox is empty',
                                 parent=self.main)
        elif not self.email_user_pass.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'Email User Pass textbox is empty',
                                 parent=self.main)
        elif not self.email_to.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'Email To textbox is empty',
                                 parent=self.main)
        elif not self.email_from.get():
            messagebox.showerror('Txtbox Empty Error!',
                                 'Email From textbox is empty',
                                 parent=self.main)
        else:
            error = 0

            if not self.email_port.get():
                self.email_port.set("587")

            try:
                server = smtplib.SMTP(str(self.email_server.get()), int(self.email_port.get()))

                try:
                    server.ehlo()
                    server.starttls()
                    server.ehlo()
                    server.login(self.email_user_name.get(), self.epass_obj.decrypt_text())
                    self.add_setting('Settings', self.email_server.get(), "Email_Server")
                    self.add_setting('Settings', self.email_port.get(), "Email_Port")
                    self.add_setting('Settings', self.email_user_name.get(), "Email_User")
                    self.add_setting('Settings', self.epass_obj.decrypt_text(), "Email_Pass")
                    self.add_setting('Local_Settings', self.email_to.get(), "Email_To")
                    self.add_setting('Local_Settings', self.email_from.get(), "Email_From")
                    self.add_setting('Local_Settings', self.email_cc.get(), "Email_CC")
                except:
                    error = 1
                    pass
                finally:
                    server.close()
            except:
                error = 2
                pass

            try:
                ftp = ftplib.FTP(self.fserver.get())

                try:
                    ftp.login(self.fuser.get(), self.fpass_obj.decrypt_text())
                    self.add_setting('Local_Settings', self.fserver.get(), 'FTP Host')
                    self.add_setting('Local_Settings', self.fuser.get(), 'FTP User')
                    self.add_setting('Local_Settings', self.fpass_obj.decrypt_text(), 'FTP Password')
                    self.cancel()
                except:
                    error = 3
                    pass
                finally:
                    ftp.quit()
            except:
                error = 4
                pass

            if error == 1:
                messagebox.showerror('Invalid Credentials Error!',
                                     'Unable to log into e-mail server & port with user name & pass',
                                     parent=self.main)
            elif error == 2:
                messagebox.showerror('Host Conn Error!',
                                     'Unable to connect to e-mail server & port',
                                     parent=self.main)
            elif error == 3:
                messagebox.showerror('Invalid Credentials Error!',
                                     'Unable to log into ftp server with user name & pass',
                                     parent=self.main)
            elif error == 4:
                messagebox.showerror('Host Conn Error!',
                                     'Unable to connect to ftp server',
                                     parent=self.main)

    # Function to launch change upload settings GUI
    def upload_settings(self):
        if self.change_upload_settings_obj:
            self.change_upload_settings_obj.cancel()

        self.change_upload_settings_obj = ChangeAccSettings(self.main, self.sql_tables)
        self.change_upload_settings_obj.build_gui()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        if self.main:
            self.main.destroy()


class AccSettingsGUI:
    acc_table = None
    acc_cols = None
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
    sql_tbl_truncate_checkbox = None
    cancel_button = None
    complete_sql_tbl_list = pd.DataFrame()

    # Function that is executed upon creation of SettingsGUI class
    def __init__(self, class_obj=None, root=None, config=None, sql_tables=None):
        self.header_text = 'Welcome to STC Upload Settings!\nSettings can be changed below.\nPress save when finished'
        self.complete_sql_tbl_list = sql_tables

        if config:
            self.config = config
        else:
            self.config = None

        if class_obj and root:
            self.class_obj = class_obj
            self.main = Toplevel(root)
        else:
            self.main = Tk()
            self.class_obj = None

        self.main.iconbitmap(icon_path)

        # GUI Variables
        self.acc_tbl_name = StringVar()
        self.sql_tbl_name = StringVar()
        self.sql_tbl_truncate = IntVar()

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
        assert (key and setting_list)

        global_objs[setting_list].del_item(key)

        if val:
            global_objs[setting_list].add_item(key=key, val=val, encrypt=encrypt)
            global_objs[setting_list].write_shelf()

    def check_tbl(self, table):
        for t_data in self.complete_sql_tbl_list:
            if t_data[0] == 'TABLE' and table.lower() == '{0}.{1}'.format(t_data[1][1].lower(), t_data[1][2].lower()):
                return True

    # Function to build GUI for settings
    def build_gui(self, header=None, acc_table=None, acc_cols=None):
        # Change to custom header title if specified
        if header:
            self.header_text = header

        self.acc_table = acc_table
        self.acc_cols = acc_cols

        # Set GUI Geometry and GUI Title
        self.main.geometry('540x636+500+70')
        if self.config:
            self.main.title('Change STC Upload Setting')
        else:
            self.main.title('Add STC Upload Setting')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        add_upload_frame = LabelFrame(self.main, text='Add Upload Settings', width=444, height=210)
        acc_tbl_frame = LabelFrame(add_upload_frame, text='Access Table', width=444, height=105)
        acc_tbl_top_frame = Frame(acc_tbl_frame)
        acc_tbl_bottom_left_frame = LabelFrame(acc_tbl_frame, text='Column List')
        acc_tbl_bottom_right_frame = LabelFrame(acc_tbl_frame, text='Column Select List')
        acc_tbl_bottom_middle_frame = Frame(acc_tbl_frame)
        sql_tbl_frame = LabelFrame(add_upload_frame, text='SQL Table', width=444, height=105)
        sql_tbl_top_frame = Frame(sql_tbl_frame)
        sql_tbl_bottom_left_frame = LabelFrame(sql_tbl_frame, text='Column List')
        sql_tbl_bottom_right_frame = LabelFrame(sql_tbl_frame, text='Column Select List')
        sql_tbl_bottom_middle_frame = Frame(sql_tbl_frame)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        add_upload_frame.pack(fill="both")
        acc_tbl_frame.grid(row=0)
        acc_tbl_top_frame.grid(row=0, column=0, columnspan=3, sticky=W+E)
        acc_tbl_bottom_left_frame.grid(row=1, column=0)
        acc_tbl_bottom_right_frame.grid(row=1, column=2, padx=5)
        acc_tbl_bottom_middle_frame.grid(row=1, column=1, padx=5)
        sql_tbl_frame.grid(row=1)
        sql_tbl_top_frame.grid(row=0, column=0, columnspan=3, sticky=W + E)
        sql_tbl_bottom_left_frame.grid(row=1, column=0)
        sql_tbl_bottom_right_frame.grid(row=1, column=2, padx=5)
        sql_tbl_bottom_middle_frame.grid(row=1, column=1, padx=5)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=510, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to the Acc_TBL_Frame
        #     Access Table Name Input Box
        acc_tbl_name_label = Label(acc_tbl_top_frame, text='TBL Name:')
        acc_tbl_name_txtbox = Entry(acc_tbl_top_frame, textvariable=self.acc_tbl_name, width=64)
        acc_tbl_name_label.grid(row=0, column=0, padx=8, pady=5)
        acc_tbl_name_txtbox.grid(row=0, column=1, padx=5, pady=5)
        acc_tbl_name_txtbox.configure(state=DISABLED)

        #     Access Table Column List
        atc_xscrollbar = Scrollbar(acc_tbl_bottom_left_frame, orient='horizontal')
        atc_yscrollbar = Scrollbar(acc_tbl_bottom_left_frame, orient='vertical')
        self.atc_list_box = Listbox(acc_tbl_bottom_left_frame, selectmode=SINGLE, width=30,
                                    yscrollcommand=atc_yscrollbar, xscrollcommand=atc_xscrollbar)
        atc_xscrollbar.config(command=self.atc_list_box.xview)
        atc_yscrollbar.config(command=self.atc_list_box.yview)
        self.atc_list_box.grid(row=0, column=0, padx=8, pady=5)
        atc_xscrollbar.grid(row=1, column=0, sticky=W+E)
        atc_yscrollbar.grid(row=0, column=1, sticky=N+S)
        self.atc_list_box.bind("<Down>", self.atc_list_down)
        self.atc_list_box.bind("<Up>", self.atc_list_up)
        self.atc_list_box.bind('<<ListboxSelect>>', self.atc_select)

        #     Access Column Migration Buttons
        self.acc_right_button = Button(acc_tbl_bottom_middle_frame, text='>', width=5,
                                       command=self.acc_right_migrate)
        self.acc_all_right_button = Button(acc_tbl_bottom_middle_frame, text='>>', width=5,
                                           command=self.acc_all_right_migrate)
        self.acc_left_button = Button(acc_tbl_bottom_middle_frame, text='<', width=5,
                                      command=self.acc_left_migrate)
        self.acc_all_left_button = Button(acc_tbl_bottom_middle_frame, text='<<', width=5,
                                          command=self.acc_all_left_migrate)
        self.acc_right_button.grid(row=0, column=2, padx=7, pady=7)
        self.acc_all_right_button.grid(row=1, column=2, padx=7, pady=7)
        self.acc_left_button.grid(row=2, column=2, padx=7, pady=7)
        self.acc_all_left_button.grid(row=3, column=2, padx=7, pady=7)

        #     Access Table Column Selection List
        atcs_xscrollbar = Scrollbar(acc_tbl_bottom_right_frame, orient='horizontal')
        atcs_yscrollbar = Scrollbar(acc_tbl_bottom_right_frame, orient='vertical')
        self.atcs_list_box = Listbox(acc_tbl_bottom_right_frame, selectmode=SINGLE, width=30,
                                     yscrollcommand=atcs_yscrollbar, xscrollcommand=atcs_xscrollbar)
        atcs_xscrollbar.config(command=self.atcs_list_box.xview)
        atcs_yscrollbar.config(command=self.atcs_list_box.yview)
        self.atcs_list_box.grid(row=0, column=3, padx=8, pady=5)
        atcs_xscrollbar.grid(row=1, column=3, sticky=W + E)
        atcs_yscrollbar.grid(row=0, column=4, sticky=N + S)
        self.atcs_list_box.bind("<Down>", self.atcs_list_down)
        self.atcs_list_box.bind("<Up>", self.atcs_list_up)
        self.atcs_list_box.bind('<<ListboxSelect>>', self.atcs_select)

        # Apply Widgets to the SQL_TBL_Frame
        #     SQL Table Name Input Box
        sql_tbl_name_label = Label(sql_tbl_top_frame, text='TBL Name:')
        self.sql_tbl_name_txtbox = Entry(sql_tbl_top_frame, textvariable=self.sql_tbl_name, width=52)
        sql_tbl_name_label.grid(row=0, column=0, padx=8, pady=5)
        self.sql_tbl_name_txtbox.grid(row=0, column=1, padx=5, pady=5)
        self.sql_tbl_name_txtbox.bind('<KeyRelease>', self.check_tbl_name)

        #     Truncate Check Box
        self.sql_tbl_truncate_checkbox = Checkbutton(sql_tbl_top_frame, text='Truncate TBL',
                                                     variable=self.sql_tbl_truncate)
        self.sql_tbl_truncate_checkbox.grid(row=0, column=2, padx=8, pady=5)

        #     SQL Table Column List
        stc_xscrollbar = Scrollbar(sql_tbl_bottom_left_frame, orient='horizontal')
        stc_yscrollbar = Scrollbar(sql_tbl_bottom_left_frame, orient='vertical')
        self.stc_list_box = Listbox(sql_tbl_bottom_left_frame, selectmode=SINGLE, width=30,
                                    yscrollcommand=stc_yscrollbar, xscrollcommand=stc_xscrollbar)
        stc_xscrollbar.config(command=self.stc_list_box.xview)
        stc_yscrollbar.config(command=self.stc_list_box.yview)
        self.stc_list_box.grid(row=0, column=0, padx=8, pady=5)
        stc_xscrollbar.grid(row=1, column=0, sticky=W + E)
        stc_yscrollbar.grid(row=0, column=1, sticky=N + S)
        self.stc_list_box.bind("<Down>", self.stc_list_down)
        self.stc_list_box.bind("<Up>", self.stc_list_up)
        self.stc_list_box.bind('<<ListboxSelect>>', self.stc_select)

        #     SQL Column Migration Buttons
        self.sql_right_button = Button(sql_tbl_bottom_middle_frame, text='>', width=5,
                                       command=self.sql_right_migrate)
        self.sql_all_right_button = Button(sql_tbl_bottom_middle_frame, text='>>', width=5,
                                           command=self.sql_all_right_migrate)
        self.sql_left_button = Button(sql_tbl_bottom_middle_frame, text='<', width=5,
                                      command=self.sql_left_migrate)
        self.sql_all_left_button = Button(sql_tbl_bottom_middle_frame, text='<<', width=5,
                                          command=self.sql_all_left_migrate)
        self.sql_right_button.grid(row=0, column=2, padx=7, pady=7)
        self.sql_all_right_button.grid(row=1, column=2, padx=7, pady=7)
        self.sql_left_button.grid(row=2, column=2, padx=7, pady=7)
        self.sql_all_left_button.grid(row=3, column=2, padx=7, pady=7)

        #     SQL Table Column Selection List
        stcs_xscrollbar = Scrollbar(sql_tbl_bottom_right_frame, orient='horizontal')
        stcs_yscrollbar = Scrollbar(sql_tbl_bottom_right_frame, orient='vertical')
        self.stcs_list_box = Listbox(sql_tbl_bottom_right_frame, selectmode=SINGLE, width=30,
                                     yscrollcommand=stcs_yscrollbar, xscrollcommand=atcs_xscrollbar)
        stcs_xscrollbar.config(command=self.stcs_list_box.xview)
        stcs_yscrollbar.config(command=self.stcs_list_box.yview)
        self.stcs_list_box.grid(row=0, column=3, padx=8, pady=5)
        stcs_xscrollbar.grid(row=1, column=3, sticky=W + E)
        stcs_yscrollbar.grid(row=0, column=4, sticky=N + S)
        self.stcs_list_box.bind("<Down>", self.stcs_list_down)
        self.stcs_list_box.bind("<Up>", self.stcs_list_up)
        self.stcs_list_box.bind('<<ListboxSelect>>', self.stcs_select)

        if self.config:
            button_name = 'Change Setting'
        else:
            button_name = 'Save Setting'

        # Apply Buttons to Button_Frame
        #     Save Button
        self.save_button = Button(self.main, text=button_name, width=15, command=self.save_settings)
        self.save_button.pack(in_=button_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        self.cancel_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        self.cancel_button.pack(in_=button_frame, side=RIGHT, padx=10, pady=5)

        if self.class_obj:
            #     Delete Setting
            delete_button = Button(self.main, text='Delete Setting', width=15, command=self.delete_setting)
            delete_button.pack(in_=button_frame, side=TOP, padx=10, pady=5)

        # Fill Textboxes with settings
        self.fill_gui()

        if not self.class_obj:
            # Show GUI Dialog
            self.main.mainloop()

    # Function to fill GUI textbox fields
    def fill_gui(self):
        disable_sql = True

        if self.acc_table:
            self.acc_tbl_name.set(self.acc_table)

            for col in self.acc_cols:
                self.atc_list_box.insert('end', col)

            self.atc_list_box.select_set(self.stc_list_sel)
        elif self.config:
            if self.config[0]:
                self.acc_tbl_name.set(self.config[0])

            if self.config[1]:
                for col in self.config[1]:
                    self.atc_list_box.insert('end', col)

            if self.config[2]:
                for col in self.config[2]:
                    self.atcs_list_box.insert('end', col)

            if self.config[3]:
                self.sql_tbl_name.set(self.config[3])

                if self.config[4]:
                    for col in self.config[4]:
                        self.stcs_list_box.insert('end', col)

                if self.check_tbl(self.config[3]):
                    self.populate_tbl_lists(True)

                if self.config[5]:
                    self.sql_tbl_truncate.set(1)

                disable_sql = False
        else:
            self.save_button.configure(state=DISABLED)
            self.atc_list_box.configure(state=DISABLED)
            self.atcs_list_box.configure(state=DISABLED)
            self.acc_right_button.configure(state=DISABLED)
            self.acc_all_right_button.configure(state=DISABLED)
            self.acc_left_button.configure(state=DISABLED)
            self.acc_all_left_button.configure(state=DISABLED)
            self.sql_tbl_name_txtbox.configure(state=DISABLED)
            self.sql_tbl_truncate_checkbox.configure(state=DISABLED)

        if disable_sql:
            self.stc_list_box.configure(state=DISABLED)
            self.stcs_list_box.configure(state=DISABLED)
            self.sql_right_button.configure(state=DISABLED)
            self.sql_all_right_button.configure(state=DISABLED)
            self.sql_left_button.configure(state=DISABLED)
            self.sql_all_left_button.configure(state=DISABLED)

    def populate_tbl_lists(self, check_stcs_list=False):
        mytbl = self.sql_tbl_name.get().split('.')
        asql = global_objs['SQL']

        try:
            asql.connect('alch')
            myresults = asql.execute('''
                SELECT
                    Column_Name
                    
                FROM INFORMATION_SCHEMA.COLUMNS
        
                WHERE
                    Table_Schema = '{0}'
                        AND
                    Table_Name = '{1}'
            '''.format(mytbl[0], mytbl[1]))
        except:
            myresults = pd.DataFrame()
        finally:
            asql.close_conn()

        if not myresults.empty:
            if self.stcs_list_box.size() > 0:
                stcs = self.stcs_list_box.get(0, self.stcs_list_box.size() - 1)
            else:
                stcs = None

            for col in myresults['Column_Name'].tolist():
                skipadd = False

                if check_stcs_list and stcs:
                    for col2 in stcs:
                        if col.lower() == col2.lower():
                            skipadd = True
                            break

                if not skipadd:
                    self.stc_list_box.insert('end', col)

            self.atc_list_box.select_set(self.stc_list_sel)
        else:
            messagebox.showerror('No Columns Error!', 'Table has no columns in SQL Server', parent=self.main)

    def check_tbl_name(self, event):
        if self.stc_list_box.size() > 0:
            self.stc_list_sel = 0
            self.stc_list_box.delete(0, self.stc_list_box.size() - 1)

        if self.stcs_list_box.size() > 0:
            self.stcs_list_sel = 0
            self.stcs_list_box.delete(0, self.stcs_list_box.size() - 1)

        if (self.config or self.acc_table) and self.sql_tbl_name.get() and self.check_tbl(self.sql_tbl_name.get()):
            self.stc_list_box.configure(state=NORMAL)
            self.stcs_list_box.configure(state=NORMAL)
            self.sql_right_button.configure(state=NORMAL)
            self.sql_all_right_button.configure(state=NORMAL)
            self.sql_left_button.configure(state=NORMAL)
            self.sql_all_left_button.configure(state=NORMAL)
            self.populate_tbl_lists()
        elif str(self.stc_list_box['state']) != 'disabled':
            self.stc_list_box.configure(state=DISABLED)
            self.stcs_list_box.configure(state=DISABLED)
            self.sql_right_button.configure(state=DISABLED)
            self.sql_all_right_button.configure(state=DISABLED)
            self.sql_left_button.configure(state=DISABLED)
            self.sql_all_left_button.configure(state=DISABLED)

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
        if self.atc_list_box.curselection():
            self.atcs_list_box.insert('end', self.atc_list_box.get(self.atc_list_box.curselection()))
            self.atc_list_box.delete(self.atc_list_box.curselection(), self.atc_list_box.curselection())

            if self.atc_list_box.size() > 0:
                if self.atc_list_sel > 0:
                    self.atc_list_sel -= 1
                self.atc_list_box.select_set(self.atc_list_sel)
            elif self.atc_list_sel > 0:
                self.atc_list_sel = -1
            else:
                self.atc_list_sel = 0
                self.atcs_list_sel = 0
                self.atcs_list_box.select_set(self.atcs_list_sel)

    # Button to migrate single record to right list (Access TBL Section)
    def acc_all_right_migrate(self):
        if self.atc_list_box.size() > 0:
            for i in range(self.atc_list_box.size()):
                self.atcs_list_box.insert('end', self.atc_list_box.get(i))

            self.atc_list_box.delete(0, self.atc_list_box.size() - 1)
            self.atc_list_sel = 0
            self.atcs_list_sel = 0
            self.atcs_list_box.select_set(self.atcs_list_sel)

    # Button to migrate single record to right list (Access TBL Section)
    def acc_left_migrate(self):
        if self.atcs_list_box.curselection():
            self.atc_list_box.insert('end', self.atcs_list_box.get(self.atcs_list_box.curselection()))
            self.atcs_list_box.delete(self.atcs_list_box.curselection(), self.atcs_list_box.curselection())

            if self.atcs_list_box.size() > 0:
                if self.atcs_list_sel > 0:
                    self.atcs_list_sel -= 1
                self.atcs_list_box.select_set(self.atcs_list_sel)
            elif self.atcs_list_sel > 0:
                self.atcs_list_sel = -1
            else:
                self.atcs_list_sel = 0
                self.atc_list_sel = 0
                self.atc_list_box.select_set(self.atc_list_sel)

    # Button to migrate single record to right list (Access TBL Section)
    def acc_all_left_migrate(self):
        if self.atcs_list_box.size() > 0:
            for i in range(self.atcs_list_box.size()):
                self.atc_list_box.insert('end', self.atcs_list_box.get(i))

            self.atcs_list_box.delete(0, self.atcs_list_box.size() - 1)
            self.atcs_list_sel = 0
            self.atc_list_sel = 0
            self.atc_list_box.select_set(self.atc_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_right_migrate(self):
        if self.stc_list_box.curselection():
            self.stcs_list_box.insert('end', self.stc_list_box.get(self.stc_list_box.curselection()))
            self.stc_list_box.delete(self.stc_list_box.curselection(), self.stc_list_box.curselection())

            if self.stc_list_box.size() > 0:
                if self.stc_list_sel > 0:
                    self.stc_list_sel -= 1
                self.stc_list_box.select_set(self.stc_list_sel)
            elif self.stc_list_sel > 0:
                self.stc_list_sel = -1
            else:
                self.stc_list_sel = 0
                self.stcs_list_sel = 0
                self.stcs_list_box.select_set(self.stcs_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_all_right_migrate(self):
        if self.stc_list_box.size() > 0:
            for i in range(self.stc_list_box.size()):
                self.stcs_list_box.insert('end', self.stc_list_box.get(i))

            self.stc_list_box.delete(0, self.stc_list_box.size() - 1)
            self.stc_list_sel = 0
            self.stcs_list_sel = 0
            self.stcs_list_box.select_set(self.stcs_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_left_migrate(self):
        if self.stcs_list_box.curselection():
            self.stc_list_box.insert('end', self.stcs_list_box.get(self.stcs_list_box.curselection()))
            self.stcs_list_box.delete(self.stcs_list_box.curselection(), self.stcs_list_box.curselection())

            if self.stcs_list_box.size() > 0:
                if self.stcs_list_sel > 0:
                    self.stcs_list_sel -= 1
                self.stcs_list_box.select_set(self.stcs_list_sel)
            elif self.stcs_list_sel > 0:
                self.stcs_list_sel = -1
            else:
                self.stcs_list_sel = 0
                self.stc_list_sel = 0
                self.stc_list_box.select_set(self.stc_list_sel)

    # Button to migrate single record to right list (SQL TBL Section)
    def sql_all_left_migrate(self):
        if self.stcs_list_box.size() > 0:
            for i in range(self.stcs_list_box.size()):
                self.stc_list_box.insert('end', self.stcs_list_box.get(i))

            self.stcs_list_box.delete(0, self.stcs_list_box.size() - 1)
            self.stcs_list_sel = 0
            self.stc_list_sel = 0
            self.stc_list_box.select_set(self.stc_list_sel)

    # Function to save settings when the Save Settings button is pressed
    def save_settings(self):
        if self.acc_table or self.class_obj:
            if self.atcs_list_box.size() < 1:
                messagebox.showerror('List Empty Error!',
                                     'Access Table Select Column Listbox is empty. Please migrate columns',
                                     parent=self.main)
            elif not self.sql_tbl_name.get():
                messagebox.showerror('Input Box Empty Error!',
                                     'SQL Table input field is empty. Please populate and migrate columns thereafter',
                                     parent=self.main)
            elif self.stcs_list_box.size() < 1:
                messagebox.showerror('List Empty Error!',
                                     'SQL Table Select Column Listbox is empty. Please migrate columns',
                                     parent=self.main)
            elif self.atcs_list_box.size() != self.stcs_list_box.size():
                messagebox.showerror('List Comparison Error!',
                                     'SQL Table Select Column Listbox size != Access Table Select Listbox size',
                                     parent=self.main)
            else:
                if self.check_tbl(self.sql_tbl_name.get()):
                    messagebox.showerror('Invalid SQL TBL!', 'SQL TBL does not exist in sql server', parent=self.main)
                else:
                    configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')
                    if configs:
                        for config in configs:
                            if config[0] == self.acc_tbl_name.get():
                                configs.remove(config)
                                break
                    else:
                        configs = []

                    if self.sql_tbl_truncate.get() == 1:
                        configs.append([self.acc_tbl_name.get(),
                                        self.atc_list_box.get(0, self.atc_list_box.size() - 1),
                                        self.atcs_list_box.get(0, self.atcs_list_box.size() - 1),
                                        self.sql_tbl_name.get(),
                                        self.stcs_list_box.get(0, self.stcs_list_box.size() - 1),
                                        True])
                    else:
                        configs.append([self.acc_tbl_name.get(),
                                        self.atc_list_box.get(0, self.atc_list_box.size() - 1),
                                        self.atcs_list_box.get(0, self.atcs_list_box.size() - 1),
                                        self.sql_tbl_name.get(),
                                        self.stcs_list_box.get(0, self.stcs_list_box.size() - 1),
                                        False])
                    self.add_setting('Local_Settings', configs, 'Accdb_Configs', False)
                    self.main.destroy()

    # Function to delete setting when Delete button is pressed
    def delete_setting(self):
        myresponse = messagebox.askokcancel(
            'Delete Notice!',
            'Deleting this setting will lose this setting forever. Would you like to proceed?',
            parent=self.main)
        if myresponse:
            configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')

            if configs:
                for config in configs:
                    if config[0] == self.acc_tbl_name.get():
                        configs.remove(config)
                        break

                self.add_setting('Local_Settings', configs, 'Accdb_Configs', False)
                self.main.destroy()

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        if self.main:
            self.main.destroy()


class ChangeAccSettings:
    change_button = None
    upload_button = None
    list_box = None
    change_setting_obj = None
    manual_upload_obj = None
    list_sel = 0

    def __init__(self, root, sql_tables):
        self.main = Toplevel(root)
        self.main.iconbitmap(icon_path)
        self.header_text = 'Welcome to STC Upload Config Settings!\nPlease modify/manually upload config'
        self.configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')
        self.sql_tables = sql_tables

    # Function to build GUI for Extract Shelf
    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('375x285+630+290')
        self.main.title('STC Upload Config Settings')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        list_frame = LabelFrame(self.main, text='Settings List', width=444, height=140)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        list_frame.pack(fill="both")
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        #     Access Setting List Populate
        xscrollbar = Scrollbar(list_frame, orient='horizontal')
        yscrollbar = Scrollbar(list_frame, orient='vertical')
        self.list_box = Listbox(list_frame, selectmode=SINGLE, width=55,
                                yscrollcommand=yscrollbar, xscrollcommand=xscrollbar)
        xscrollbar.config(command=self.list_box.xview)
        yscrollbar.config(command=self.list_box.yview)
        self.list_box.grid(row=0, column=3, padx=8, pady=5)
        xscrollbar.grid(row=1, column=3, sticky=W + E)
        yscrollbar.grid(row=0, column=4, sticky=N + S)
        self.list_box.bind("<Down>", self.list_down)
        self.list_box.bind("<Up>", self.list_up)
        self.list_box.bind('<<ListboxSelect>>', self.list_select)

        # Apply Buttons to Button_Frame
        #     Change Button
        self.change_button = Button(self.main, text='Change Setting', width=15, command=self.change_setting)
        self.change_button.pack(in_=button_frame, side=LEFT, padx=5, pady=5)

        #     Cancel Button
        cancel_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        cancel_button.pack(in_=button_frame, side=RIGHT, padx=5, pady=5)

        # Apply Buttons to Button_Frame
        #     Upload Button
        self.upload_button = Button(self.main, text='Manual Upload', width=15, command=self.manual_upload)
        self.upload_button.pack(in_=button_frame, side=TOP, padx=5, pady=5)

        self.load_gui_fields()

    def load_gui_fields(self, set_id=0):
        if self.configs:
            for config in self.configs:
                self.list_box.insert('end', config[0])

            if set_id > -1:
                self.list_box.select_set(set_id)
        else:
            self.list_box.configure(state=DISABLED)
            self.upload_button.configure(state=DISABLED)
            self.change_button.configure(state=DISABLED)

    def list_down(self, event):
        if self.list_sel < self.list_box.size() - 1:
            self.list_box.select_clear(self.list_sel)
            self.list_sel += 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user presses up key (ATCS List)
    def list_up(self, event):
        if self.list_sel > 0:
            self.list_box.select_clear(self.list_sel)
            self.list_sel -= 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user clicks item (STC List)
    def list_select(self, event):
        if self.list_box and self.list_box.curselection() \
                and -1 < self.list_sel < self.list_box.size() - 1:
            self.list_sel = self.list_box.curselection()[0]

    def manual_upload(self):
        if self.list_box.curselection() and self.configs:
            if self.manual_upload_obj:
                self.manual_upload_obj.cancel()
                self.manual_upload_obj = None

            self.manual_upload_obj = ManualUploadGUI(self.main, self.list_box.get(self.list_box.curselection()))

            if self.manual_upload_obj:
                self.manual_upload_obj.build_gui()
                self.manual_upload_obj = None

    def change_setting(self):
        if self.list_box.curselection() and self.configs:
            if self.change_setting_obj:
                self.change_setting_obj.cancel()
                self.change_setting_obj = None

            self.configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')

            for config in self.configs:
                if config[0] == self.list_box.get(self.list_box.curselection()):
                    self.change_setting_obj = AccSettingsGUI(self, self.main, config, self.sql_tables)
                    break

            if self.change_setting_obj:
                self.change_setting_obj.build_gui()
                self.change_setting_obj = None

            if self.list_box.size() > 0:
                self.configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')
                self.list_box.delete(0, self.list_box.size() - 1)
                self.load_gui_fields(-1)

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        if self.main:
            self.main.destroy()


class ManualUploadGUI:
    list_box = None
    upload_button = None
    list_sel = -1

    def __init__(self, root=None, table=None):
        self.main = Toplevel(root)
        self.main.iconbitmap(icon_path)
        self.header_text = 'Welcome to Manual Upload Config!\nPlease choose a date to upload'
        self.table = table
        self.accdb_paths = []

    # Function to build GUI for Manual Upload
    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('255x285+630+290')
        self.main.title('STC Manual Upload Config')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        list_frame = LabelFrame(self.main, text='Upload Date List', width=444, height=140)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        list_frame.pack(fill="both")
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        #     Upload Date List Box Widget
        xscrollbar = Scrollbar(list_frame, orient='horizontal')
        yscrollbar = Scrollbar(list_frame, orient='vertical')
        self.list_box = Listbox(list_frame, selectmode=SINGLE, width=35,
                                yscrollcommand=yscrollbar, xscrollcommand=xscrollbar)
        xscrollbar.config(command=self.list_box.xview)
        yscrollbar.config(command=self.list_box.yview)
        self.list_box.grid(row=0, column=3, padx=8, pady=5)
        xscrollbar.grid(row=1, column=3, sticky=W + E)
        yscrollbar.grid(row=0, column=4, sticky=N + S)
        self.list_box.bind("<Down>", self.list_down)
        self.list_box.bind("<Up>", self.list_up)
        self.list_box.bind('<<ListboxSelect>>', self.list_select)

        # Apply Buttons to Button_Frame
        #     Upload Button
        self.upload_button = Button(self.main, text='Upload Accdb', width=15, command=self.manual_upload)
        self.upload_button.pack(in_=button_frame, side=LEFT, padx=5, pady=5)

        #     Cancel Button
        cancel_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        cancel_button.pack(in_=button_frame, side=RIGHT, padx=5, pady=5)

        self.load_gui_fields()

    def load_gui_fields(self):
        self.grab_accdbs()

        if len(self.accdb_paths) > 0:
            for apath in self.accdb_paths:
                self.list_box.insert('end', os.path.basename(os.path.dirname(apath)))

            self.list_box.select_set(0)
        else:
            self.list_box.configure(state=DISABLED)
            self.upload_button.configure(state=DISABLED)

    def list_down(self, event):
        if self.list_sel < self.list_box.size() - 1:
            self.list_box.select_clear(self.list_sel)
            self.list_sel += 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user presses up key (ATCS List)
    def list_up(self, event):
        if self.list_sel > 0:
            self.list_box.select_clear(self.list_sel)
            self.list_sel -= 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user clicks item (STC List)
    def list_select(self, event):
        if self.list_box and self.list_box.curselection() \
                and -1 < self.list_sel < self.list_box.size() - 1:
            self.list_sel = self.list_box.curselection()[0]

    def grab_accdbs(self):
        for filename in os.listdir(ProcessedDir):
            file_path = os.path.join(ProcessedDir, filename)

            if os.path.isdir(file_path):
                self.accdb_paths += list(pl.Path(file_path).glob('*.accdb')) +\
                                    list(pl.Path(file_path).glob('*.mdb'))

        if len(self.accdb_paths) > 0:
            rem_list = []
            trim_date = datetime.now()
            trim_date = (trim_date - relativedelta.relativedelta(months=3)).__format__("%Y%m%d")

            for file_path in self.accdb_paths:
                if os.path.basename(os.path.dirname(file_path)) < trim_date:
                    rem_list.append(file_path)
                else:
                    fnd_tbl = False
                    asql = SQLHandle(logobj=global_objs['Event_Log'], accdb_file=file_path)

                    try:
                        asql.connect('accdb', )
                        tables = asql.tables()

                        for table in tables:
                            if table[0] == 'TABLE' and table[1][2].lower() == self.table.lower():
                                fnd_tbl = True
                                break

                        if not fnd_tbl:
                            rem_list.append(file_path)
                    finally:
                        asql.close_conn()
                        del asql

            if len(rem_list) > 0:
                for rem in rem_list:
                    self.accdb_paths.remove(rem)

    def manual_upload(self):
        if self.list_box.curselection():
            batch = self.list_box.get(self.list_box.curselection())

            for apath in self.accdb_paths:
                if batch == os.path.basename(os.path.dirname(apath)):
                    global_objs['Event_Log'].log_mode(True, self.main)
                    global_objs['Event_Log'].write_log('Validating table [{0}]'.format(self.table))
                    myobj = AccdbHandle(apath, batch)

                    if myobj.validate(self.table) and myobj.process(self.table):
                        myobj.apply_features(self.table)
                        stats = myobj.get_success_stats(self.table)
                        global_objs['Event_Log'].write_log('Manual Upload [{0}] => [{1}] {2} Records uploaded'.format(
                                                               stats[0], stats[1], stats[2]))
                        messagebox.showinfo('Manual Upload Success', '[{0}] => [{1}] {2} Records uploaded'.format(
                            stats[0], stats[1], stats[2]
                        ), parent=global_objs['Event_Log'].log_gui_root())
                    else:
                        messagebox.showerror('Manual Upload Failed', 'Wasnt able to validate and process table',
                                             parent=global_objs['Event_Log'].log_gui_root())

                    global_objs['Event_Log'].log_mode(False)
                    break

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        if self.main:
            self.main.destroy()


class ConfigReviewGUI:
    list_box = None
    create_config_button = None
    settings_obj = None
    list_sel = -1

    def __init__(self, config_review=None):
        self.main = Tk()
        self.main.iconbitmap(icon_path)
        self.header_text = 'Welcome to Config Review!\nPlease review or create a new config below:'
        self.config_review = config_review

    # Function to build GUI for Manual Upload
    def build_gui(self):
        # Set GUI Geometry and GUI Title
        self.main.geometry('255x285+630+290')
        self.main.title('Config Review')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        list_frame = LabelFrame(self.main, text='Config Review List', width=444, height=140)
        button_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack()
        list_frame.pack(fill="both")
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=self.header_text, width=375, justify=CENTER)
        header.pack(in_=header_frame)

        #     Upload Date List Box Widget
        xscrollbar = Scrollbar(list_frame, orient='horizontal')
        yscrollbar = Scrollbar(list_frame, orient='vertical')
        self.list_box = Listbox(list_frame, selectmode=SINGLE, width=35,
                                yscrollcommand=yscrollbar, xscrollcommand=xscrollbar)
        xscrollbar.config(command=self.list_box.xview)
        yscrollbar.config(command=self.list_box.yview)
        self.list_box.grid(row=0, column=3, padx=8, pady=5)
        xscrollbar.grid(row=1, column=3, sticky=W + E)
        yscrollbar.grid(row=0, column=4, sticky=N + S)
        self.list_box.bind("<Down>", self.list_down)
        self.list_box.bind("<Up>", self.list_up)
        self.list_box.bind('<<ListboxSelect>>', self.list_select)

        # Apply Buttons to Button_Frame
        #     Create Config Button
        self.create_config_button = Button(self.main, text='Upload Accdb', width=15, command=self.create_config)
        self.create_config_button.pack(in_=button_frame, side=LEFT, padx=5, pady=5)

        #     Cancel Button
        cancel_button = Button(self.main, text='Cancel', width=15, command=self.cancel)
        cancel_button.pack(in_=button_frame, side=RIGHT, padx=5, pady=5)

        self.main.mainloop()

        self.load_gui_fields(False)

    def load_gui_fields(self, refresh_config=True):
        if refresh_config:
            global_objs['Local_Settings'].read_shelf()
            self.config_review = global_objs['Local_Settings'].grab_item('Config_Review')

        if len(self.config_review) > 0:
            for config in self.config_review:
                self.list_box.insert('end', config[1])

            self.list_box.select_set(0)
        else:
            self.list_box.configure(state=DISABLED)
            self.create_config_button.configure(state=DISABLED)

    def list_down(self, event):
        if self.list_sel < self.list_box.size() - 1:
            self.list_box.select_clear(self.list_sel)
            self.list_sel += 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user presses up key (ATCS List)
    def list_up(self, event):
        if self.list_sel > 0:
            self.list_box.select_clear(self.list_sel)
            self.list_sel -= 1
            self.list_box.select_set(self.list_sel)

    # Function adjusts selection of item when user clicks item (STC List)
    def list_select(self, event):
        if self.list_box and self.list_box.curselection() \
                and -1 < self.list_sel < self.list_box.size() - 1:
            self.list_sel = self.list_box.curselection()[0]

    def create_config(self):
        if self.list_box.curselection():
            config_name = self.list_box.get(self.list_box.curselection())

            for config in self.config_review:
                if config[1] == config_name:
                    if config[4]:
                        acc_obj = AccSettingsGUI(class_obj=self, root=self.main)
                        acc_obj.build_gui(config[3], config[1], config[2])
                    else:
                        acc_obj = AccSettingsGUI(class_obj=self, root=self.main, config=config[5])
                        acc_obj.build_gui(config[3])

                    self.load_gui_fields()
                    break

    # Function to destroy GUI when Cancel button is pressed
    def cancel(self):
        if self.main:
            self.main.destroy()


# Main loop routine to create GUI Settings
if __name__ == '__main__':
    cr = global_objs['Local_Settings'].grab_item('Config_Review')

    if cr:
        obj = ConfigReviewGUI()
        obj.build_gui()

    obj = SettingsGUI()
    obj.build_gui()

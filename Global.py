from pandas.io import sql
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus
from cryptography.fernet import Fernet
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from sqlalchemy.exc import SQLAlchemyError
from contextlib import closing
from tkinter import filedialog
from tkinter import messagebox
from tkinter import *

import threading
import traceback
import xml.etree.ElementTree as ET
import pathlib as pl
import pandas as pd
import sqlalchemy as mysql
import shelve_lock
import pyodbc
import os
import datetime
import logging
import base64
import random
import string
import shutil

# Include this immport of dumb from dbm_lock for pyinstaller's Exe compiler
from dbm_lock import dumb


def grabobjs(scriptdir, filename=None):
    if scriptdir and os.path.exists(scriptdir):
        myobjs = dict()
        myobjs['Local_Settings'] = ShelfHandle(os.path.join(scriptdir, 'Script_Settings'))

        if len(list(pl.Path(scriptdir).glob('Script_Settings.*'))) > 0:
            myobjs['Local_Settings'].read_shelf()
            mydir = myobjs['Local_Settings'].grab_item('General_Settings_Path')
        else:
            mydir = None

        if not mydir or not os.path.exists(mydir):
            obj = GeneralSettingsGUI(scriptdir)
            obj.build_gui()
            myobjs['Local_Settings'].read_shelf()
            mydir = myobjs['Local_Settings'].grab_item('General_Settings_Path')

        if mydir and os.path.exists(mydir):
            myobjs['Settings'] = ShelfHandle(os.path.join(myobjs['Local_Settings'].grab_item(
                'General_Settings_Path'), 'General_Settings'))
            myobjs['Settings'].read_shelf()
            myobjs['Event_Log'] = LogHandle(scriptdir, filename)
            myobjs['SQL'] = SQLHandle(logobj=myobjs['Event_Log'], settingsobj=myobjs['Settings'])
            myobjs['Errors'] = ErrHandle(myobjs['Event_Log'])

            return myobjs
        else:
            raise Exception('No General Settings were established. Please re-run')
    else:
        raise Exception('Invalid script path provided')


class GeneralSettingsGUI:
    def __init__(self, script_dir):
        self.main = Tk()
        self.gen_dir = StringVar()
        self.script_dir = script_dir

    def build_gui(self):
        file_path = os.path.join(self.script_dir, 'Script_Settings_backup')
        header_text = "Settings will need to be setup\nPlease specify the location or to set the location of General Settings"

        # Set GUI Geometry and GUI Title
        self.main.geometry('400x130+500+90')
        self.main.title('Settings Setup')
        self.main.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self.main)
        settings_frame = LabelFrame(self.main, text='Settings Setup', width=508, height=70)
        buttons_frame = Frame(self.main)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        settings_frame.pack(fill="both")
        buttons_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self.main, text=header_text, width=400, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Directory Entry Box Widget
        dir_label = Label(settings_frame, text='Directory:')
        dir_txtbox = Entry(settings_frame, textvariable=self.gen_dir, width=42)
        dir_label.grid(row=0, column=0, padx=4, pady=5)
        dir_txtbox.grid(row=0, column=1, padx=4, pady=5)

        # Apply Directory Finder Button Widget
        find_dir_button = Button(settings_frame, text='Find Dir', width=7, command=self.find_dir)
        find_dir_button.grid(row=0, column=2, padx=4, pady=5)

        # Apply Buttons to the Buttons Frame
        #     Save Settings
        create_button = Button(buttons_frame, text='Create Settings', width=15, command=self.create_settings)
        create_button.grid(row=0, column=0, pady=6, padx=9)

        #     Restore Button
        restore_button = Button(buttons_frame, text='Restore Settings', width=15, command=self.restore_settings)
        restore_button.grid(row=0, column=1, pady=6, padx=9)

        if os.path.exists('%s.dir' % file_path) and os.path.exists('%s.dat' % file_path)\
                and os.path.exists('%s.bak' % file_path):
            restore_button.configure(state=NORMAL)
        else:
            restore_button.configure(state=DISABLED)

        #     Cancel Button
        cancel_button = Button(buttons_frame, text='Cancel', width=15, command=self.cancel)
        cancel_button.grid(row=0, column=2, pady=6, padx=9)

        # Show dialog
        self.main.mainloop()

    def find_dir(self):
        if self.gen_dir.get() and os.path.exists(self.gen_dir.get()):
            init_dir = os.path.dirname(self.gen_dir.get())
        else:
            init_dir = '/'

        file = filedialog.askdirectory(initialdir=init_dir, title='Select Directory', parent=self.main)

        if file:
            self.gen_dir.set(file)

    def create_settings(self):
        gen_dir = self.gen_dir.get()

        if not gen_dir:
            messagebox.showerror('Field Empty Error!', 'No value has been inputed for Directory',
                                 parent=self.main)
        elif not os.path.exists(gen_dir):
            messagebox.showerror('Invalid Path Error!', 'Directory Path does not exist',
                                 parent=self.main)
        else:
            obj = ShelfHandle(os.path.join(self.script_dir, 'Script_Settings'))
            obj.add_item('General_Settings_Path', self.gen_dir.get())
            obj.write_shelf()
            del obj
            self.main.destroy()

    def restore_settings(self):
        backup_file = os.path.join(self.script_dir, 'Script_Settings_backup')
        setting_file = os.path.join(self.script_dir, 'Script_Settings')

        if os.path.exists('%s.dat' % backup_file) \
                and os.path.exists('%s.dir' % backup_file) \
                and os.path.exists('%s.bak' % backup_file):
            shutil.copy2('%s.dat' % backup_file, '%s.dat' % setting_file)
            shutil.copy2('%s.dir' % backup_file, '%s.dir' % setting_file)
            shutil.copy2('%s.bak' % backup_file, '%s.bak' % setting_file)
            self.cancel()

    def cancel(self):
        self.main.destroy()


class CryptHandle:
    key = None
    encrypted_text = None

    @staticmethod
    def random_text():
        digits = "".join([random.choice(string.digits+string.ascii_letters) for i in range(15)])
        return digits

    def create_key(self, text=None):
        if not text:
            text = self.random_text()

        etext = text.encode()
        salt = os.urandom(16)

        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
            backend=default_backend()
        )

        self.key = base64.urlsafe_b64encode(kdf.derive(etext))

    @staticmethod
    def code_method(obj):
        if isinstance(obj, int):
            return obj.to_bytes(4, byteorder='big', signed=True)
        elif isinstance(obj, str):
            return obj.encode()
        else:
            return obj.decode()

    def encrypt_text(self, item):
        if isinstance(item, int) or isinstance(item, str):
            if not self.key:
                self.create_key()

            crypt_obj = Fernet(self.key)
            self.encrypted_text = crypt_obj.encrypt(self.code_method(item))
        else:
            raise Exception('Invalid data type for encryption')

    def decrypt_text(self):
        if self.key and self.encrypted_text:
            crypt_obj = Fernet(self.key)
            return self.code_method(crypt_obj.decrypt(self.encrypted_text))

    def grab_items(self):
        return [self.key, self.encrypted_text]

    def compare_text(self, compare_key, encrypt_text):
        if self.key and self.encrypted_text and compare_key and encrypt_text:
            crypt_obj = Fernet(compare_key)
            crypt_obj2 = Fernet(self.key)

            if crypt_obj.decrypt(self.encrypted_text) == crypt_obj2.decrypt(encrypt_text):
                return True
            else:
                return False
        elif not compare_key:
            raise Exception('No comparison key has been provided in parameter')
        elif not encrypt_text:
            raise Exception('No encrypted text has been provided in parameter')
        elif not self.key:
            raise Exception('No Key has been generated. Please generate a key')
        else:
            raise Exception('No text has been encrypted. Please encrypt a text')


class ShelfHandle:
    def __init__(self, filepath=None):
        if os.path.exists(os.path.split(filepath)[0]):
            self.file = filepath
            self.shelf_data = dict()
            self.rem_keys = list()
            self.add_keys = list()
        else:
            raise Exception('Invalid filepath has been provided')

    def get_shelf_path(self):
        return self.file

    def change_config(self, filepath):
        if os.path.exists(os.path.split(self.file)[0]):
            self.file = filepath

    def backup(self):
        file = '%s.bak' % self.file
        file2 = '%s.dat' % self.file
        file3 = '%s.dir' % self.file

        if os.path.exists(file) and os.stat(file).st_size > 0 and os.path.exists(file2) and os.stat(file2).st_size > 0 \
                and os.path.exists(file3) and os.stat(file3).st_size > 0:
            shutil.copy2(file, '%s_backup.bak' % self.file)
            shutil.copy2(file2, '%s_backup.dat' % self.file)
            shutil.copy2(file3, '%s_backup.dir' % self.file)

    def read_shelf(self):
        self.shelf_data.clear()
        self.rem_keys.clear()
        self.add_keys.clear()

        shelf = shelve_lock.open(self.file)

        try:
            if len(shelf) > 0:
                for k, v in shelf.items():
                    self.shelf_data[k] = v
        finally:
            shelf.close()

    def write_shelf(self):
        shelf = shelve_lock.open(self.file)

        try:
            if len(self.rem_keys) > 0:
                for k in self.rem_keys:
                    if k in shelf.keys():
                        del shelf[k]

            if len(self.add_keys) > 0:
                for key in self.add_keys:
                    shelf[key] = self.shelf_data[key]

                self.shelf_data.clear()

            if len(shelf) > 0:
                for k, v in shelf.items():
                    self.shelf_data[k] = v

            self.rem_keys.clear()
            self.add_keys.clear()
        finally:
            shelf.close()

    def empty_shelf(self):
        shelf = shelve_lock.open(self.file)
        shelf.clear()
        shelf.close()

    def get_keys(self):
        return self.shelf_data.keys()

    def grab_item(self, key):
        if key and key in self.shelf_data.keys():
            return self.shelf_data[key]

    def add_item(self, key, val=None, inputmsg=None, encrypt=False):
        if key:
            while not val:
                if inputmsg:
                    print(inputmsg)
                else:
                    print("Please input value for {}:".format(key))

                val = input()

            if key in self.rem_keys:
                self.rem_keys.remove(key)

            if key not in self.add_keys:
                self.add_keys.append(key)

            if encrypt:
                myobj = CryptHandle()
                myobj.encrypt_text(val)
                self.shelf_data[key] = myobj
            else:
                self.shelf_data[key] = val

    def del_item(self, key):
        if key and key in self.shelf_data.keys():
            if key in self.add_keys:
                self.add_keys.remove(key)

            self.rem_keys.append(key)
            del self.shelf_data[key]

    def grab_list(self):
        return self.shelf_data

    def empty_list(self):
        if len(self.shelf_data) > 0:
            for key in self.shelf_data.keys():
                self.rem_keys.append(key)

            self.add_keys.clear()
            self.shelf_data.clear()

    def add_list(self, dict_list):
        if isinstance(dict_list, dict) and len(dict_list) > 0:
            for k, v in dict_list.items():
                if k in self.rem_keys:
                    self.rem_keys.remove(k)

                self.add_keys.append(k)
                self.shelf_data[k] = v


class FakeConsole(Frame):
    def __init__(self, root, *args, **kargs):
        Frame.__init__(self, root, *args, **kargs)

        # white text on black background,
        # for extra versimilitude
        self.text = Text(self, bg="black", fg="white")
        self.text.pack()

        # list of things not yet printed
        self.printQueue = []

        # one thread will be adding to the print queue,
        # and another will be iterating through it.
        # better make sure one doesn't interfere with the other.
        self.printQueueLock = threading.Lock()

        self.after(5, self.on_idle)

    # check for new messages every five milliseconds
    def on_idle(self):
        with self.printQueueLock:
            for msg in self.printQueue:
                self.text.insert(END, msg)
                self.text.see(END)
            self.printQueue = []
        self.after(5, self.on_idle)

    # print msg to the console
    def show(self, msg, sep="\n"):
        with self.printQueueLock:
            self.printQueue.append(str(msg) + sep)


class LogHandle:
    new_console = False
    console = None
    root = None

    def __init__(self, scriptpath, filename=None):
        if scriptpath:
            if not os.path.exists(os.path.join(scriptpath, '01_Event_Logs')):
                os.makedirs(os.path.join(scriptpath, '01_Event_Logs'))

            self.EventLog_Dir = os.path.join(scriptpath, '01_Event_Logs')

            if filename:
                self.filename = filename
            else:
                self.filename = os.path.basename(os.path.dirname(os.path.abspath(__file__)))
        else:
            raise Exception('Settings object was not passed through')

    def log_mode(self, new_console=False, root=None):
        if self.console and hasattr(self.console, 'destroy'):
            self.root.destroy()
            self.console = None
            self.root = None

        self.root = root
        self.new_console = new_console

    def log_gui_root(self):
        return self.root

    def write_log(self, message, action='info'):
        filepath = os.path.join(self.EventLog_Dir,
                                "{0}_{1}_Log.txt".format(datetime.datetime.now().__format__("%Y%m%d"), self.filename))

        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)

        logging.basicConfig(filename=filepath,
                            level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

        if self.new_console:
            if not self.console:
                if self.root:
                    self.root = Toplevel(self.root)
                    self.console = FakeConsole(self.root)
                    threading.Thread(target=self.console.pack).start()
                else:
                    self.root = Tk()
                    self.console = FakeConsole(self.root)
                    self.console.pack()
                    threading.Thread(target=self.root.mainloop).start()

            self.console.show(message)
        else:
            print('{0} - {1} - {2}'.format(datetime.datetime.now(), action.upper(), message))

        if action == 'debug':
            logging.debug(message)
        elif action == 'info':
            logging.info(message)
        elif action == 'warning':
            logging.warning(message)
        elif action == 'error':
            logging.error(message)
        elif action == 'critical':
            logging.critical(message)


class SQLHandle:
    def __init__(self, logobj=None, settingsobj=None, server=None, database=None, dsn=None, accdb_file=None):
        self.server = None
        self.database = None
        self.dsn = None
        self.accdb_file = None
        self.conn_type = None
        self.raw_engine = None
        self.engine = None
        self.session = None
        self.cursor = None
        self.dataset = []

        self.change_config(settingsobj, server, database, dsn, accdb_file)
        self.logobj = logobj

    def __connect_str(self):
        if self.conn_type == 'alch':
            assert(self.server and self.database)
            p = quote_plus(
                'DRIVER={};PORT={};SERVER={};DATABASE={};Trusted_Connection=yes;'.format(
                    '{SQL Server Native Client 11.0}', '1433', self.server, self.database))

            return '{}+pyodbc:///?odbc_connect={}'.format('mssql', p)
        elif self.conn_type == 'sql':
            assert(self.server and self.database)
            return 'driver={0};server={1};database={2};autocommit=True;Trusted_Connection=yes'\
                .format('{SQL Server}', self.server, self.database)
        elif self.conn_type == 'accdb':
            assert self.accdb_file
            return 'DRIVER={};DBQ={};Exclusive=1'.format('{Microsoft Access Driver (*.mdb, *.accdb)}',
                                                         self.accdb_file)
        elif self.conn_type == 'dsn':
            assert self.dsn
            return 'DSN={};DATABASE=default;Trusted_Connection=Yes;'.format(self.dsn)
        else:
            raise Exception('Invalid conn_type specified')

    def __proc_errors(self, err_code=None, err_desc=None):
        if self.logobj:
            if err_code and err_desc:
                self.logobj.write_log('[Error Code {0}] - {1}'.format(err_code, err_desc))
            else:
                self.logobj.write_log(traceback.format_exc(), 'critical')
        else:
            if err_code and err_desc:
                print('[Error Code {0}] - {1}'.format(err_code, err_desc))
            else:
                print(traceback.format_exc())

    def __commit(self, engine):
        try:
            if engine and hasattr(engine, 'commit'):
                engine.commit()
        except:
            self.__rollback(engine)

    def __rollback(self, engine):
        try:
            if engine and hasattr(engine, 'rollback'):
                engine.rollback()
        except:
            self.close_conn()

    def __store_dataset(self, res_rowset):
        try:
            data = [tuple(t) for t in res_rowset.fetchall()]
            cols = [column[0] for column in res_rowset.description]
            self.dataset.append(pd.DataFrame(data, columns=cols))
        except:
            pass

    def change_config(self, settingsobj=None, server=None, database=None, dsn=None, accdb_file=None):
        if settingsobj:
            if settingsobj.grab_item('Server'):
                self.server = settingsobj.grab_item('Server').decrypt_text()
            if settingsobj.grab_item('Database'):
                self.database = settingsobj.grab_item('Database').decrypt_text()
            if settingsobj.grab_item('DSN'):
                self.dsn = settingsobj.grab_item('DSN').decrypt_text()
        elif server and database:
            self.server = server
            self.database = database
        elif dsn:
            self.dsn = dsn
        elif accdb_file:
            self.accdb_file = accdb_file
        else:
            raise Exception('Invalid connection variables passed')

    def connect(self, conn_type, test_conn=False, session=False, conn_timeout=3, query_time_out=0):
        self.conn_type = conn_type
        self.session = session
        conn_str = self.__connect_str()

        try:
            if self.conn_type == 'alch':
                if not self.raw_engine:
                    self.raw_engine = mysql.create_engine(
                        conn_str, connect_args={'timeout': conn_timeout, 'connect_timeout': conn_timeout,
                                                'options': '-c statement_timeout=%s' % query_time_out})

                    try:
                        self.raw_engine.connect()
                    except:
                        self.__proc_errors()
                        self.close_conn()
                        return False
                    else:
                        self.raw_engine = self.raw_engine.raw_connection()

                if not self.engine:
                    self.engine = mysql.create_engine(
                        conn_str, connect_args={'timeout': conn_timeout, 'connect_timeout': conn_timeout,
                                                'options': '-c statement_timeout=%s' % query_time_out})

                    try:
                        self.engine.connect()
                    except:
                        self.__proc_errors()
                        self.close_conn()
                        return False
                    else:
                        if self.session:
                            self.engine = sessionmaker(bind=self.engine)
                            self.engine = self.engine()
                            self.engine._model_changes = {}

                if test_conn:
                    self.close_conn()
            elif not self.engine and not self.raw_engine:
                self.raw_engine = pyodbc.connect(
                        conn_str, connect_args={'timeout': conn_timeout, 'connect_timeout': conn_timeout,
                                                'options': '-c statement_timeout=%s' % query_time_out})
                try:
                    self.raw_engine.commit()
                except:
                    self.__proc_errors()
                    self.close_conn()
                    return False
        except Exception as e:
            self.__proc_errors(err_code=type(e).__name__, err_desc=str(e))
            self.close_conn()
            return False
        else:
            if test_conn:
                self.close_conn()

            return True

    def close_cursor(self):
        if self.cursor and hasattr(self.cursor, 'cancel'):
            self.cursor.cancel()

        self.__rollback(self.cursor)
        self.cursor = None

    def close_conn(self):
        self.close_cursor()

        if self.raw_engine and hasattr(self.raw_engine, 'close'):
            self.raw_engine.close()

        if self.raw_engine and hasattr(self.raw_engine, 'dispose'):
            self.raw_engine.dispose()

        if self.engine and hasattr(self.engine, 'close'):
            self.engine.close()

        if self.engine and hasattr(self.engine, 'dispose'):
            self.engine.dispose()

        self.engine = None
        self.raw_engine = None
        self.session = False

    def grab_sql_objs(self):
        if self.engine:
            return [self.engine, self.raw_engine]
        elif self.raw_engine:
            return self.raw_engine

    def tables(self):
        if self.raw_engine:
            try:
                with closing(self.raw_engine.cursor()) as cursor:
                    if self.conn_type == 'accdb':
                        tables = [[t.table_type, [t.table_cat, t.table_schem, t.table_name]]
                                  for t in cursor.tables() if 'msys' not in t.table_name.lower()]
                    else:
                        tables = [[t.table_type, [t.table_cat, t.table_schem, t.table_name]] for t in cursor.tables()]
            except:
                self.__proc_errors()
            else:
                return tables

    def create_table(self, df, table):
        if self.conn_type == 'alch' and not self.session and self.engine:
            try:
                df.to_sql(
                    table,
                    self.engine,
                    if_exists='replace'
                )
            except:
                self.__proc_errors()
            else:
                return True

    def upload_df(self, df, table, if_exists='append', index=True, index_label='linenumber'):
        if self.conn_type == 'alch' and not self.session and self.engine:
            tbl = table.split('.')

            if len(tbl) == 2:
                schema = tbl[0]
                tbl_name = tbl[1]
            else:
                schema = None
                tbl_name = table

            try:
                df.to_sql(
                    tbl_name,
                    self.engine,
                    schema=schema,
                    if_exists=if_exists,
                    index=index,
                    index_label=index_label,
                    chunksize=1000
                )
            except:
                self.__proc_errors()
                self.close_conn()
            else:
                return True

    def execute(self, str_txt, params=None, execute=False, proc=False, ret_err=False):
        df = pd.DataFrame()

        try:
            if proc and params:
                str_txt = 'EXEC {0} {1}'.format(str_txt, params)
            elif proc:
                str_txt = 'EXEC {0}'.format(str_txt)

            if execute and self.raw_engine:
                with closing(self.raw_engine.cursor()) as self.cursor:
                    self.dataset = []
                    result = self.cursor.execute(str_txt)
                    self.__store_dataset(result)

                    while result.nextset():
                        self.__store_dataset(result)

                    if len(self.dataset) == 1 and not self.dataset[0].empty:
                        df = self.dataset[0]
                    elif len(self.dataset) > 1:
                        df = self.dataset

                    self.__commit(self.cursor)

                self.cursor = None
            elif not execute:
                if self.conn_type == 'alch' and self.engine:
                    self.cursor = self.engine.execute(mysql.text(str_txt))

                    if self.cursor and self.cursor._saved_cursor.arraysize > 0:
                        df = pd.DataFrame(self.cursor.fetchall(), columns=self.cursor._metadata.keys)

                    self.cursor = None
                elif self.conn_type != 'alch' and self.raw_engine:
                    df = sql.read_sql(str_txt, self.raw_engine)
        except SQLAlchemyError as e:
            err = [df, e.code, str(e.__dict__['orig'])]
            self.__rollback(self.cursor)
            self.cursor = None

            if ret_err:
                return err
            else:
                self.__proc_errors(err_code=err[1], err_desc=err[2])
        except pyodbc.Error as e:
            err = [df, e.args[0], e.args[1]]
            self.__rollback(self.cursor)
            self.cursor = None

            if ret_err:
                return err
            else:
                self.__proc_errors(err_code=err[1], err_desc=err[2])
        except (AttributeError, Exception) as e:
            err = [df, type(e).__name__, str(e)]
            self.__rollback(self.cursor)
            self.cursor = None

            if ret_err:
                return err
            else:
                self.__proc_errors(err_code=err[1], err_desc=err[2])
        else:
            if ret_err:
                return [df, None, None]
            else:
                return df


class ErrHandle:
    errors = dict()

    def __init__(self, logobj):
        if not logobj:
            raise Exception('Event Log object not included in parameter')

        self.logobj = logobj

    @staticmethod
    def trim_df(df_to_trim, df_to_compare):
        if isinstance(df_to_trim, pd.DataFrame) and isinstance(df_to_compare, pd.DataFrame)\
                and not df_to_trim.empty and not df_to_compare.empty:
            df_to_trim.drop(df_to_compare.index, inplace=True)

    @staticmethod
    def concat_dfs(df_list):
        if isinstance(df_list, list) and len(df_list) > 0:
            dfs = []

            for df in df_list:
                if isinstance(df, pd.DataFrame):
                    dfs.append(df)

            if len(dfs) > 0:
                return pd.concat(dfs, ignore_index=True, sort=False).drop_duplicates().reset_index(drop=True)

    def append_errors(self, err_items, key=None):
        if isinstance(err_items, list) and len(err_items) > 0:
            self.logobj.write_log('Error(s) found. Appending to virtual list', 'warning')

            if key and key in self.errors.keys():
                self.errors[key].append(err_items)
            elif key:
                self.errors[key] = [err_items]
            elif 'default' in self.errors.keys():
                self.errors['default'].append(err_items)
            else:
                self.errors['default'] = [err_items]

    def grab_errors(self, key=None):
        if key and key in self.errors.keys():
            mylist = self.errors[key]
            del self.errors[key]
            return mylist
        elif not key and 'default' in self.errors.keys():
            mylist = self.errors['default']
            del self.errors['default']
            return mylist
        else:
            return None


class XMLParseClass:
    def __init__(self, file):
        try:
            tree = ET.parse(file)
            self.root = tree.getroot()
        except AssertionError as a:
            print('\t[-] {} : Parse failed.'.format(a))
            pass

    def parseelement(self, element, parsed=None):
        if parsed is None:
            parsed = dict()

        if element.keys():
            for key in element.keys():
                if key not in parsed:
                    parsed[key] = element.attrib.get(key)

                if element.text and element.tag not in parsed:
                    parsed[element.tag] = element.text

        elif element.text and element.tag not in parsed:
            parsed[element.tag] = element.text

        for child in list(element):
            self.parseelement(child, parsed)
        return parsed

    def parsexml(self, findpath, dictvar=None):
        if isinstance(dictvar, dict):
            for item in self.root.findall(findpath):
                dictvar = self.parseelement(item, dictvar)

            return dictvar
        else:
            parsed = [self.parseelement(item) for item in self.root.findall(findpath)]
            df = pd.DataFrame(parsed)

            return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)


class XMLAppendClass:
    def __init__(self, file):
        self.file = file

    def write_xml(self, df):
        with open(self.file, 'w') as xmlFile:
            xmlFile.write(
                '<?xml version="1.0" encoding="UTF-8"?>\n'
            )
            xmlFile.write('<records>\n')

            xmlFile.write(
                '\n'.join(df.apply(self.xml_encode, axis=1))
            )

            xmlFile.write('\n</records>')

    @staticmethod
    def xml_encode(row):
        xmlitem = ['  <record>']

        for field in row.index:
            if row[field]:
                xmlitem \
                    .append('    <var var_name="{0}">{1}</var>' \
                            .format(field, row[field]))

        xmlitem.append('  </record>')

        return '\n'.join(xmlitem)

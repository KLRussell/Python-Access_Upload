from pandas.io import sql
from sqlalchemy.orm import sessionmaker
from urllib.parse import quote_plus
from cryptography.fernet import Fernet
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

import traceback
import xml.etree.ElementTree as ET
import pathlib as pl
import pandas as pd
import sqlalchemy as mysql
import shelve
import pyodbc
import os
import datetime
import logging
import base64
import random
import string


def grabobjs(scriptdir, filename=None):
    if scriptdir and os.path.exists(scriptdir):
        myobjs = dict()
        myinput = None

        if len(list(pl.Path(scriptdir).glob('Script_Settings.*'))) > 0:
            myobjs['Local_Settings'] = ShelfHandle(os.path.join(scriptdir, 'Script_Settings'))
            mydir = myobjs['Local_Settings'].grab_item('General_Settings_Path')

            if mydir and os.path.exists(mydir):
                myobjs['Settings'] = ShelfHandle(os.path.join(myobjs['Local_Settings'].grab_item(
                    'General_Settings_Path'), 'General_Settings'))
            else:
                while not myinput:
                    print("Please input a directory path where to setup general settings at:")
                    myinput = input()

                    if myinput and not os.path.exists(myinput):
                        myinput = None
        else:
            while not myinput:
                print("Please input a directory path where to setup general settings at:")
                myinput = input()

                if myinput and not os.path.exists(myinput):
                    myinput = None

        if myinput:
            myobjs['Local_Settings'] = ShelfHandle(os.path.join(scriptdir, 'Script_Settings'))
            myobjs['Local_Settings'].add_item('General_Settings_Path', myinput)
            myobjs['Settings'] = ShelfHandle(os.path.join(myinput, 'General_Settings'))

        myobjs['Event_Log'] = LogHandle(scriptdir, filename)
        myobjs['SQL'] = SQLHandle(logobj=myobjs['Event_Log'], settingsobj=myobjs['Settings'])
        myobjs['Errors'] = ErrHandle(myobjs['Event_Log'])

        return myobjs
    else:
        raise Exception('Invalid script path provided')


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
            self.filepath = filepath
            sfile = shelve.open(filepath)
            type(sfile)
            sfile.close()
        else:
            raise Exception('Invalid filepath has been provided')

    def change_config(self, filepath):
        if os.path.exists(os.path.split(self.filepath)[0]):
            self.filepath = filepath

    def get_shelf_path(self):
        return self.filepath

    def get_keys(self):
        mykeys = []
        sfile = shelve.open(self.filepath)
        type(sfile)

        for key in sfile.keys():
            mykeys.append(key)

        sfile.close()

        return mykeys

    def grab_item(self, key):
        sfile = shelve.open(self.filepath)
        type(sfile)

        if key in sfile.keys():
            myitem = sfile[key]
        else:
            myitem = None

        sfile.close()

        return myitem

    def add_item(self, key, val=None, inputmsg=None, encrypt=False):
        if key:
            sfile = shelve.open(self.filepath)
            type(sfile)

            if key not in sfile.keys():
                if encrypt:
                    myobj = CryptHandle()
                else:
                    myobj = None

                if not val:
                    myinput = None

                    while not myinput:
                        if inputmsg:
                            print(inputmsg)
                        else:
                            print("Please input value for {}:".format(key))
                        myinput = input()

                    if myobj:
                        myobj.encrypt_text(myinput)
                        sfile[key] = myobj
                    else:
                        sfile[key] = myinput
                elif myobj:
                    myobj.encrypt_text(val)
                    sfile[key] = myobj
                else:
                    sfile[key] = val

            sfile.close()

    def del_item(self, key):
        if key:
            sfile = shelve.open(self.filepath)
            type(sfile)

            if key in sfile.keys():
                del sfile[key]

            sfile.close()

    def grab_list(self):
        mydict = dict()
        sfile = shelve.open(self.filepath)
        type(sfile)

        for k, v in sfile.items():
            mydict[k] = v

        sfile.close()

        return mydict

    def empty_list(self):
        sfile = shelve.open(self.filepath)
        type(sfile)

        for k in sfile.keys():
            del sfile[k]

        sfile.close()

    def add_list(self, dict_list):
        if len(dict_list) > 0 and isinstance(dict_list, dict):
            sfile = shelve.open(self.filepath)
            type(sfile)

            for k, v in dict_list.items():
                sfile[k] = v

            sfile.close()


class LogHandle:
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

    def write_log(self, message, action='info'):
        filepath = os.path.join(self.EventLog_Dir,
                                "{0}_{1}_Log.txt".format(datetime.datetime.now().__format__("%Y%m%d"), self.filename))

        logging.basicConfig(filename=filepath,
                            level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

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
    server = None
    database = None
    dsn = None
    conn_type = None
    conn_str = None
    session = False
    engine = None
    conn = None
    cursor = None

    def __init__(self, logobj=None, settingsobj=None, server=None, database=None, dsn=None, accdb_file=None):
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

        if logobj:
            self.logobj = logobj

    def change_config(self, settingsobj=None, server=None, database=None, dsn=None, accdb_file=None):
        if settingsobj:
            self.server = settingsobj.grab_item('Server').decrypt_text()
            self.database = settingsobj.grab_item('Database').decrypt_text()
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

    def create_conn_str(self):
        if self.conn_type == 'alch':
            assert(self.server and self.database)
            p = quote_plus(
                'DRIVER={};PORT={};SERVER={};DATABASE={};Trusted_Connection=yes;'
                    .format('{SQL Server Native Client 11.0}', '1433', self.server, self.database))

            self.conn_str = '{}+pyodbc:///?odbc_connect={}'.format('mssql', p)
        elif self.conn_type == 'sql':
            assert(self.server and self.database)
            self.conn_str = 'driver={0};server={1};database={2};autocommit=True;Trusted_Connection=yes'\
                .format('{SQL Server}', self.server, self.database)
        elif self.conn_type == 'accdb':
            assert self.accdb_file
            self.conn_str = 'DRIVER={};DBQ={};Exclusive=1'.format('{Microsoft Access Driver (*.mdb, *.accdb)}',
                                                                  self.accdb_file)
        elif self.conn_type == 'dsn':
            assert self.dsn
            self.conn_str = 'DSN={};DATABASE=default;Trusted_Connection=Yes;'.format(self.dsn)
        else:
            raise Exception('Invalid conn_type specified')

    def test_conn(self, conn_type=None):
        assert(conn_type or self.conn_type)
        myreturn = False

        if conn_type:
            self.conn_type = conn_type

        self.create_conn_str()

        myquery = "SELECT 1 from sys.sysprocesses"

        try:
            if self.conn_type == 'alch':
                self.engine = mysql.create_engine(self.conn_str)
                obj = self.engine.execute(mysql.text(myquery))
                if obj._saved_cursor.arraysize > 0:
                    myreturn = True
            else:
                self.conn = pyodbc.connect(self.conn_str)
                self.cursor = self.conn.cursor()
                self.conn.commit()

                if self.conn_type == 'accdb' and len(self.get_accdb_tables()) > 0:
                    myreturn = True
                else:
                    df = sql.read_sql(myquery, self.conn)

                    if len(df) > 0:
                        myreturn = True
        except:
            pass
        finally:
            self.close()

        return myreturn

    def get_accdb_tables(self):
        if self.conn_type == 'accdb':
            mylist = []
            ct = self.cursor.tables

            for row in ct():
                if 'msys' not in row.table_name.lower():
                    mylist.append(row.table_name)

            return mylist

    def connect(self, conn_type):
        assert (conn_type or self.conn_type)
        self.conn_type = conn_type

        if self.test_conn():
            try:
                if self.conn_type == 'alch':
                    self.engine = mysql.create_engine(self.conn_str)
                else:
                    self.conn = pyodbc.connect(self.conn_str, autocommit=True)
                    self.cursor = self.conn.cursor()
                    self.conn.commit()
            except:
                self.close()
                if self.logobj:
                    self.logobj.write_log(traceback.format_exc(), 'critical')
                else:
                    print(traceback.format_exc())
                raise Exception('Stopping script')
        elif self.conn_type in ('alch', 'sql'):
            self.server = None
            self.database = None
            raise Exception(
                'Error 1 - Failed test connection to SQL Server. Server name {0} or database name {1} is incorrect'
                    .format(self.server, self.database))
        elif self.conn_type == 'accdb':
            self.accdb_file = None
            raise Exception(
                'Error 1 - Failed test connection to access databse file {0}'.format(
                    self.accdb_file))
        else:
            self.dsn = None
            raise Exception('Error 1 - Failed test connection to DSN connection {0}'.format(self.dsn))

    def close(self):
        if self.conn_type == 'alch':
            if self.engine:
                self.engine.dispose()
        else:
            if self.cursor:
                self.cursor.close()

            if self.conn:
                self.conn.close()

    def createsession(self):
        if self.conn_type == 'alch':
            try:
                self.engine = sessionmaker(bind=self.engine)
                self.engine = self.engine()
                self.engine._model_changes = {}
                self.session = True
            except:
                self.close()
                if self.logobj:
                    self.logobj.write_log(traceback.format_exc(), 'critical')
                else:
                    print(traceback.format_exc())
                raise Exception('Stopping script')

    def createtable(self, dataframe, sqltable):
        if self.conn_type == 'alch' and not self.session:
            try:
                dataframe.to_sql(
                    sqltable,
                    self.engine,
                    if_exists='replace',
                )
            except:
                self.close()
                if self.logobj:
                    self.logobj.write_log(traceback.format_exc(), 'critical')
                else:
                    print(traceback.format_exc())
                raise Exception('Stopping script')

    def grabengine(self):
        if self.conn_type == 'alch':
            return self.engine
        else:
            return [self.cursor, self.conn]

    def upload(self, dataframe, sqltable, index=True, index_label='linenumber'):
        if self.conn_type == 'alch' and not self.session:
            mytbl = sqltable.split(".")
            try:
                if len(mytbl) > 1:
                    dataframe.to_sql(
                        mytbl[1],
                        self.engine,
                        schema=mytbl[0],
                        if_exists='append',
                        index=index,
                        index_label=index_label,
                        chunksize=1000
                    )
                else:
                    dataframe.to_sql(
                        mytbl[0],
                        self.engine,
                        if_exists='replace',
                        index=False,
                        chunksize=1000
                    )
                return True
            except:
                self.close()
                if self.logobj:
                    self.logobj.write_log(traceback.format_exc(), 'critical')
                else:
                    print(traceback.format_exc())
                raise Exception('Stopping script')

    def query(self, query):
        try:
            if self.conn_type == 'alch':
                obj = self.engine.execute(mysql.text(query))

                if obj._saved_cursor.arraysize > 0:
                    data = obj.fetchall()
                    columns = obj._metadata.keys

                    return pd.DataFrame(data, columns=columns)

            else:
                df = sql.read_sql(query, self.conn)
                return df

        except:
            self.close()
            if self.logobj:
                self.logobj.write_log(traceback.format_exc(), 'critical')
            else:
                print(traceback.format_exc())
            raise Exception('Stopping script')

    def execute(self, query):
        try:
            if self.conn_type == 'alch':
                self.engine.execution_options(autocommit=True).execute(mysql.text(query))
            else:
                self.cursor.execute(query)

        except:
            self.close()
            if self.logobj:
                self.logobj.write_log(traceback.format_exc(), 'critical')
            else:
                print(traceback.format_exc())
            raise Exception('Stopping script')


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

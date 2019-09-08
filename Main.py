# If python 32-bit and MS Access engine is 64-bit, use this 32-bit MS Access driver
# http://www.microsoft.com/en-us/download/details.aspx?id=13255
# Download 7-zip from website below, add to Path Var by going into Control Panal -> System -> Advanced System Settings
# -> Environmental Variables. Double-click Path under System Variables and add directory to C:\Program Files\7-Zip
# https://www.7-zip.org/download.html

from Global import grabobjs
from Global import SQLHandle
from Settings import SettingsGUI
from Settings import AccSettingsGUI
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

import subprocess
import rarfile
import traceback
import ftplib
import os
import pathlib as pl
import pandas as pd
import smtplib
import zipfile
import shutil
import sys

if getattr(sys, 'frozen', False):
    application_path = sys.executable
else:
    application_path = __file__

curr_dir = os.path.dirname(os.path.abspath(application_path))
main_dir = os.path.dirname(curr_dir)
ProcessDir = os.path.join(main_dir, "02_Process")
ProcessedDir = os.path.join(main_dir, "03_Processed")
SQLDir = os.path.join(main_dir, "05_SQL")
global_objs = grabobjs(main_dir, 'STC_Upload')


class STCFTP:
    ftp = None

    def __init__(self):
        self.host = global_objs['Local_Settings'].grab_item('FTP Host')
        self.username = global_objs['Local_Settings'].grab_item('FTP User')
        self.password = global_objs['Local_Settings'].grab_item('FTP Password')

    def connect(self):
        if self.host and self.username and self.password:
            global_objs['Event_Log'].write_log('Opening socket to FTP server (%s) to find updates' %
                                               self.host.decrypt_text())

            try:
                self.ftp = ftplib.FTP(self.host.decrypt_text())

                try:
                    self.ftp.login(self.username.decrypt_text(), self.password.decrypt_text())
                except:
                    global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
                    self.ftp.quit()
                    self.ftp = None
            except:
                global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
                self.ftp = None
                return
        else:
            ValueError('Error 1 - Missing Host, Username, and/or Password')

    def close_connect(self):
        if self.ftp:
            self.ftp.quit()

    def grab_stc_zips(self):
        if self.ftp:
            global_objs['Event_Log'].write_log('Grabbing updates from FTP server and unzipping files')
            self.ftp.cwd('/To Granite')

            entries = list(self.ftp.mlsd())
            entries.sort(key=lambda entry: "" if entry[0].startswith('SDN') else entry[1]['modify'], reverse=True)
            date = entries[0][1]['modify'][0:8]

            dest = os.path.join(os.path.join(ProcessDir, date), 'Unzipped')

            if not os.path.exists(os.path.join(ProcessDir, date)):
                os.makedirs(os.path.join(ProcessDir, date))

            if not os.path.exists(dest):
                os.makedirs(dest)

            for item in entries:
                if item[1]['modify'][0:8] == date:
                    path = os.path.join(os.path.join(ProcessDir, date), item[0])

                    if not os.path.exists(path):
                        try:
                            self.ftp.retrbinary('RETR %s' % item[0], open(path, 'wb').write, 8 * 1024)

                            if item[0].endswith('.zip'):
                                with zipfile.ZipFile(path, 'r') as zip_ref:
                                    zip_ref.extractall(dest)

                                os.remove(path)
                            elif item[0].endswith('.rar'):
                                rar = rarfile.RarFile(path)
                                rar.extractall(dest)
                            elif item[0].endswith('.7z'):
                                z = subprocess.Popen('7z e "{0}" -o"{1}"'.format(path, dest), shell=True)
                                z.wait()
                                z.kill()
                            else:
                                os.rename(path, os.path.join(dest, item[0]))
                        except:
                            global_objs['Event_Log'].write_log(traceback.format_exc(), 'critical')
                            if os.path.exists(path):
                                os.unlink(path)
                else:
                    break


class AccdbHandle:
    configs = None
    config = None
    accdb_cols = None
    sql_cols = None
    upload_df = pd.DataFrame()

    def __init__(self, file, batch):
        self.file = file
        self.batch = batch
        self.asql = SQLHandle(logobj=global_objs['Event_Log'], settingsobj=global_objs['Settings'])
        self.asql.connect(conn_type='alch')

        if not self.email_port:
            self.email_port = 587

    @staticmethod
    def get_accdb_tables():
        myresults = global_objs['SQL'].get_accdb_tables()

        if len(myresults) > 0:
            return myresults

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
        else:
            return False

    def get_config(self, table):
        self.configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')

        if table and self.configs:
            for config in self.configs:
                if config[0] == table:
                    self.config = config
                    break

    def switch_config(self):
        if self.configs and self.config:
            for config in self.configs:
                if config == self.config:
                    self.configs.remove(config)
                    break

            self.configs.append(self.config)
            global_objs['Local_Settings'].del_item('Accdb_Configs')
            global_objs['Local_Settings'].add_item('Accdb_Configs', self.configs)

    def get_accdb_cols(self, table):
        myresults = global_objs['SQL'].query('''
            SELECT TOP 1 *
            FROM [{0}]'''.format(table))

        if len(myresults) < 1:
            raise Exception('File {0} has no valid columns in table {1}'.format(
                os.path.basename(self.file), table))

        self.accdb_cols = myresults.columns.tolist()

    def get_sql_cols(self, table):
        myresults = self.asql.query('''
            SELECT
                Column_Name

            FROM INFORMATION_SCHEMA.COLUMNS

            WHERE
                TABLE_SCHEMA = '{0}'
                    AND
                TABLE_NAME = '{1}'
        '''.format(table.split('.')[0], table.split('.')[1]))

        if len(myresults) < 1:
            raise Exception('SQL Table {} has no valid tables in the access database file'.format(self.to_table))

        self.sql_cols = myresults['Column_Name'].tolist()

    def validate_sql_table(self, table):
        if table:
            splittable = table.split('.')

            if len(splittable) == 2:
                results = self.asql.query('''
                    select 1
                    from information_schema.tables
                    where
                        table_schema = '{0}'
                            and
                        table_name = '{1}'
                '''.format(splittable[0], splittable[1]))

                if results.empty:
                    return False
                else:
                    return True
        else:
            return False

    def validate(self, table):
        self.get_accdb_cols(table)
        self.get_config(table)

        if not self.config:
            header_text = 'Welcome to STC Upload!\nThere is no configuration for table.\nPlease add configuration setting below:'
            self.config_gui(table, header_text)

        if self.config and not self.validate_sql_table(self.config[3]):
            self.config[3] = None
            header_text = 'Welcome to STC Upload!\nSQL Server TBL does not exist.\nPlease fix configuration in Upload Settings:'
            self.config_gui(table, header_text, False)

        if self.config and not self.validate_cols(self.config[2], self.accdb_cols):
            self.config[2] = None
            self.switch_config()
            header_text = 'Welcome to STC Upload!\nOne or more column does not exist.\nPlease redo config for access table columns:'
            self.config_gui(table, header_text, False)

        if self.config:
            self.get_sql_cols(self.config[3])

            if not self.validate_cols(self.config[4], self.sql_cols):
                self.config[4] = None
                self.switch_config()
                header_text = 'Welcome to STC Upload!\nOne or more column does not exist.\nPlease redo config for sql table columns:'
                self.config_gui(table, header_text, False)

            if not self.validate_cols(['Source_File'], self.sql_cols):
                self.switch_config()
                header_text = 'Welcome to STC Upload!\nSQL Table does not have a Source_File column:'
                self.config_gui(table, header_text, False)

                if not self.validate_cols(['Source_File'], self.sql_cols):
                    global_objs['Event_Log'].write_log('Selected SQL table does not have Source_File column. Closing',
                                                       'error')
                    self.config = False

        if self.config:
            if self.updates_not_exists():
                return True
            else:
                global_objs['Event_Log'].write_log('Updates already exist for SQL table %s' % self.config[3], 'warning')
                return False
        else:
            return False

    def updates_not_exists(self):
        return self.asql.query('''
            SELECT TOP 1 1
            FROM {0}
            WHERE
                Source_File = 'Updated Records {1}'
        '''.format(self.config[3], self.batch)).empty

    def config_gui(self, table, header_text, insert=True):
        obj = AccSettingsGUI()

        if insert:
            obj.build_gui(header_text, table, self.accdb_cols)
            self.get_config(table)
        else:
            old_config = self.config
            obj.build_gui(header_text)
            self.get_config(table)

            if old_config == self.config:
                self.config = None

    def process(self, table):
        global_objs['Event_Log'].write_log('Reading data from accdb table [{0}]'.format(table))

        self.upload_df = global_objs['SQL'].query('''
            SELECT
                [{0}],
                'Updated Records {2}' As Source_File

            FROM [{1}]
        '''.format('], ['.join(self.config[2]), table, self.batch))

        if not self.upload_df.empty:
            self.upload_df.columns = self.config[4] + ('Source_File',)

            if self.config[5]:
                global_objs['Event_Log'].write_log('Truncating table [{0}]'.format(self.config[3]))
                self.asql.execute('truncate table {0}'.format(self.config[3]))

            global_objs['Event_Log'].write_log('Uploading data to sql table {0}'.format(self.config[3]))
            self.asql.upload(self.upload_df, self.config[3], index=False, index_label=None)
            global_objs['Event_Log'].write_log('Data successfully uploaded from table [{0}] to sql table {1}'
                                               .format(table, self.config[3]))
            return True
        else:
            global_objs['Event_Log'].write_log('Failed to grab data from access table [{0}]. No update made'
                                               .format(table), 'error')
            return False

    def apply_features(self, table):
        if self.validate_cols(['Edit_DT'], self.sql_cols):
            self.asql.execute('''
                UPDATE {0}
                SET
                    Edit_DT = GETDATE()
            '''.format(self.config[3]))
        elif self.validate_cols(['Edit_Date'], self.sql_cols):
            self.asql.execute('''
                UPDATE {0}
                SET
                    Edit_Date = GETDATE()
            '''.format(self.config[3]))

        path = os.path.join(SQLDir, table)

        if os.path.exists(path):
            for file in list(pl.Path(path).glob('*.sql')):
                with open(str(file), 'r') as f:
                    self.asql.execute(f.read())
                    f.close()

    def close_asql(self):
        self.asql.close()

    def get_success_stats(self):
        return [self.config[2], self.config[3], len(self.upload_df)]


def email_results(batch, upload_results):
    message = MIMEMultipart()
    email_server = global_objs['Settings'].grab_item('Email_Server')
    email_user = global_objs['Settings'].grab_item('Email_User')
    email_pass = global_objs['Settings'].grab_item('Email_Pass')
    email_port = global_objs['Settings'].grab_item('Email_Port')
    email_from = global_objs['Local_Settings'].grab_item('Email_From')
    email_to = global_objs['Local_Settings'].grab_item('Email_To')
    email_cc = global_objs['Local_Settings'].grab_item('Email_CC')

    message['From'] = email_from.decrypt_text()
    message['To'] = email_to.decrypt_text()
    message['Cc'] = email_cc.decrypt_text()
    message['Date'] = formatdate(localtime=True)

    body = 'Happy Friday DART,\n\nThe following items have been successfully uploaded to SQL Server:\n\n'

    for result in upload_results:
        body += '\t* {0} -> {1} ({2} records)'.format(result[0], result[1], result[2])

    body += "\n\nYours Truly\n\nThe CDA's"

    try:
        server = smtplib.SMTP(str(email_server.decrypt_text()),
                              int(email_port))
        try:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(email_user.decrypt_text(), email_pass.decrypt_text())
            message['Subject'] = 'STC Updated Records %s' % batch
            message.attach(MIMEText(body))
            server.sendmail(email_from.decrypt_text(), email_to.decrypt_text(), str(message))
        except:
            global_objs['Event_Log'].write_log('Failed to log into email server to send e-mail', 'error')
        finally:
            server.close()
    except:
        global_objs['Event_Log'].write_log('Failed to connect to email server to send e-mail', 'error')


def check_for_updates():
    file_paths = []

    for root, dirs, files in os.walk(ProcessDir):
        if "Unzipped" in dirs and len(dirs) == 1:
            dir_path = os.path.join(root, dirs[0])
            file_paths += list(pl.Path(dir_path).glob('*.accdb')) + list(pl.Path(dir_path).glob('*.mdb'))

    return file_paths


def process_updates(files):
    paths_to_remove = []
    upload_results = []
    batch = None

    for file in files:
        batch = os.path.basename(os.path.dirname(os.path.dirname(file)))
        global_objs['Event_Log'].write_log('Processing file [{0}]'.format(os.path.basename(file)))
        myobj = AccdbHandle(file, batch)

        global_objs['SQL'].change_config(accdb_file=file)
        global_objs['SQL'].connect('accdb')

        try:
            for table in myobj.get_accdb_tables():
                global_objs['Event_Log'].write_log('Validating table [{0}]'.format(table))

                if myobj.validate(table) and myobj.process(table):
                    myobj.apply_features(table)
                    upload_results.append(myobj.get_success_stats())

        finally:
            myobj.close_asql()
            global_objs['SQL'].close()

        batch_dir = os.path.join(ProcessedDir, batch)
        my_path = os.path.join(batch_dir, os.path.basename(file))

        if not os.path.exists(batch_dir):
            os.makedirs(batch_dir)

        if os.path.exists(my_path):
            os.remove(file)
        else:
            os.rename(file, my_path)

        if not os.path.dirname(os.path.dirname(file)) in paths_to_remove:
            paths_to_remove.append(os.path.dirname(os.path.dirname(file)))

    for path in paths_to_remove:
        shutil.rmtree(path)

    if upload_results:
        email_results(batch, upload_results)


def check_settings():
    header_text = None
    my_return = False
    obj = SettingsGUI()

    if not os.path.exists(ProcessDir):
        os.makedirs(ProcessDir)

    if not os.path.exists(ProcessedDir):
        os.makedirs(ProcessedDir)

    if not global_objs['Settings'].grab_item('Server') \
            or not global_objs['Settings'].grab_item('Database') \
            or not global_objs['Local_Settings'].grab_item('FTP Host') \
            or not global_objs['Local_Settings'].grab_item('FTP User') \
            or not global_objs['Local_Settings'].grab_item('FTP Password') \
            or not global_objs['Settings'].grab_item('Email_Server') \
            or not global_objs['Settings'].grab_item('Email_User') \
            or not global_objs['Settings'].grab_item('Email_Pass') \
            or not global_objs['Local_Settings'].grab_item('Email_To') \
            or not global_objs['Local_Settings'].grab_item('Email_From'):
        header_text = 'Welcome to STC Upload!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
    else:
        try:
            if not obj.sql_connect():
                header_text = 'Welcome to STC Upload!\nNetwork settings are invalid.\nPlease fix the network settings below:'
            else:
                if global_objs['Settings'].grab_item('Email_Port'):
                    port = global_objs['Settings'].grab_item('Email_Port').decrypt_text()
                else:
                    port = 587

                try:
                    server = smtplib.SMTP(str(global_objs['Settings'].grab_item('Email_Server').decrypt_text()),
                                          int(port))
                    try:
                        server.ehlo()
                        server.starttls()
                        server.ehlo()
                        server.login(global_objs['Settings'].grab_item('Email_User').decrypt_text(),
                                     global_objs['Settings'].grab_item('Email_Pass').decrypt_text())

                        try:
                            ftp = ftplib.FTP(global_objs['Local_Settings'].grab_item('FTP Host').decrypt_text())

                            try:
                                ftp.login(global_objs['Local_Settings'].grab_item('FTP User').decrypt_text(),
                                          global_objs['Local_Settings'].grab_item('FTP Password').decrypt_text())
                                my_return = True
                            except:
                                header_text = 'Welcome to STC Upload!\nFTP User and/or Pass are invalid.\nPlease fix below:'
                            finally:
                                ftp.quit()
                        except:
                            header_text = 'Welcome to STC Upload!\nFTP server does not exist.\nPlease fix below:'
                    except:
                        header_text = 'Welcome to STC Upload!\nEmail User and/or Pass are invalid.\nPlease fix below:'
                    finally:
                        server.close()
                except:
                    header_text = 'Welcome to STC Upload!\nEmail server does not exist.\nPlease fix below:'
        finally:
            obj.sql_close()

    if header_text:
        obj.build_gui(header_text)

    obj.cancel()
    del obj

    return my_return


if __name__ == '__main__':
    if check_settings():
        mobj = STCFTP()

        try:
            mobj.connect()
            mobj.grab_stc_zips()
        finally:
            mobj.close_connect()

        has_updates = check_for_updates()

        if has_updates:
            global_objs['Event_Log'].write_log('Found {} files to process'.format(len(has_updates)))
            process_updates(has_updates)
        else:
            global_objs['Event_Log'].write_log('Found no files to process', 'warning')
    else:
        global_objs['Event_Log'].write_log('Settings Mode was established. Need to re-run script', 'warning')

    os.system('pause')

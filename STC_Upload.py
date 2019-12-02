# If python 32-bit and MS Access engine is 64-bit, use this 32-bit MS Access driver
# http://www.microsoft.com/en-us/download/details.aspx?id=13255
# Download 7-zip from website below, add to Path Var by going into Control Panal -> System -> Advanced System Settings
# -> Environmental Variables. Double-click Path under System Variables and add directory to C:\Program Files\7-Zip
# https://www.7-zip.org/download.html

from Global import grabobjs
from STC_Upload_Settings import SettingsGUI
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from STC_Upload_Settings import AccdbHandle

import datetime
import subprocess
import rarfile
import traceback
import ftplib
import os
import pathlib as pl
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
        self.log = global_objs['Event_Log']
        self.settings = global_objs['Local_Settings']
        self.host = self.settings.grab_item('FTP Host')
        self.username = self.settings.grab_item('FTP User')
        self.password = self.settings.grab_item('FTP Password')

    def connect(self):
        if self.host and self.username and self.password:
            self.log.write_log('Opening socket to FTP server (%s) to find updates' % self.host.decrypt_text())

            try:
                self.ftp = ftplib.FTP(self.host.decrypt_text())

                try:
                    self.ftp.login(self.username.decrypt_text(), self.password.decrypt_text())
                except:
                    self.log.write_log(traceback.format_exc(), 'critical')
                    self.ftp.quit()
                    self.ftp = None
            except:
                self.log.write_log(traceback.format_exc(), 'critical')
                self.ftp = None
                return
        else:
            ValueError('Error 1 - Missing Host, Username, and/or Password')

    def close_connect(self):
        if self.ftp:
            self.ftp.quit()

    def grab_stc_zips(self):
        if self.ftp:
            self.log.write_log('Grabbing updates from FTP server and unzipping files')
            self.ftp.cwd('/To Granite')

            entries = list(self.ftp.mlsd())
            entries.sort(key=lambda entry: "" if entry[0].startswith('SDN') or entry[1]['type'] == 'dir' or not (
                    entry[0].endswith('.zip') or entry[0].endswith('.rar') or entry[0].endswith('.7z')
                    or entry[0].endswith('.accdb') or entry[0].endswith('.mdb')) else entry[1]['modify'], reverse=True)
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
                            self.log.write_log(traceback.format_exc(), 'critical')
                            if os.path.exists(path):
                                os.unlink(path)
                else:
                    break


def email_results(batch, upload_results):
    message = MIMEMultipart()
    body = []
    body2 = []

    email_server = global_objs['Settings'].grab_item('Email_Server')
    email_user = global_objs['Settings'].grab_item('Email_User')
    email_pass = global_objs['Settings'].grab_item('Email_Pass')
    email_port = global_objs['Settings'].grab_item('Email_Port').decrypt_text()
    email_from = global_objs['Local_Settings'].grab_item('Email_From')
    email_to = global_objs['Local_Settings'].grab_item('Email_To')
    email_cc = global_objs['Local_Settings'].grab_item('Email_CC')

    if not email_port:
        email_port = 587

    message['From'] = email_from.decrypt_text()
    message['To'] = email_to.decrypt_text()
    message['Cc'] = email_cc.decrypt_text()
    message['Date'] = formatdate(localtime=True)

    body.append('Happy %s DART,' % datetime.datetime.today().strftime("%A"))
    body.append('The following items have been successfully uploaded to SQL Server:')

    for result in upload_results:
        body2.append('\t\u2022  {0} => {1} ({2} records)'.format(result[0], result[1], result[2]))

    body.append('\n'.join(body2))
    body.append("Yours Truly,")
    body.append("The CDA's")

    try:
        server = smtplib.SMTP(str(email_server.decrypt_text()),
                              int(email_port))
        try:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(email_user.decrypt_text(), email_pass.decrypt_text())
            message['Subject'] = 'STC Updated Records %s' % batch
            message.attach(MIMEText('\n\n'.join(body)))
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


def rem_proc_elements():
    global_objs['Event_Log'].write_log('Cleaning Process directory')

    for filename in os.listdir(ProcessDir):
        file_path = os.path.join(ProcessDir, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            global_objs['Event_Log'].write_log('Failed to delete %s. Reason: %s' % (file_path, e), 'warning')
            pass


def process_updates(files):
    batch = None
    upload_results = []

    for file in files:
        batch = os.path.basename(os.path.dirname(os.path.dirname(file)))
        global_objs['Event_Log'].write_log('Processing file [{0}]'.format(os.path.basename(file)))
        myobj = AccdbHandle(file, batch)

        try:
            for table in myobj.get_accdb_tables():
                if table[0] == 'TABLE':
                    global_objs['Event_Log'].write_log('Validating table [{0}]'.format(table[1][2]))

                    if myobj.validate(table[1][2]) and myobj.process(table[1][2]):
                        myobj.apply_features(table[1][2])
                        upload_results.append(myobj.get_success_stats(table[1][2]))
        finally:
            myobj.close_conn()
            del myobj

        batch_dir = os.path.join(ProcessedDir, batch)
        my_path = os.path.join(batch_dir, os.path.basename(file))

        if not os.path.exists(batch_dir):
            os.makedirs(batch_dir)

        if os.path.exists(my_path):
            os.remove(file)
        else:
            os.rename(file, my_path)

    if len(upload_results) > 0:
        email_results(batch, upload_results)


def check_settings(stop_gui=False):
    server = global_objs['Settings'].grab_item('Server')
    database = global_objs['Settings'].grab_item('Database')
    ftp_host = global_objs['Local_Settings'].grab_item('FTP Host')
    ftp_user = global_objs['Local_Settings'].grab_item('FTP User')
    ftp_pass = global_objs['Local_Settings'].grab_item('FTP Password')
    email_server = global_objs['Settings'].grab_item('Email_Server')
    email_user = global_objs['Settings'].grab_item('Email_User')
    email_pass = global_objs['Settings'].grab_item('Email_Pass')
    email_to = global_objs['Local_Settings'].grab_item('Email_To')
    email_from = global_objs['Local_Settings'].grab_item('Email_From')
    email_port = global_objs['Settings'].grab_item('Email_Port')
    header_text = None
    obj = SettingsGUI()

    if not os.path.exists(ProcessDir):
        os.makedirs(ProcessDir)

    if not os.path.exists(ProcessedDir):
        os.makedirs(ProcessedDir)

    if not server or not database or not ftp_host or not ftp_user or not ftp_pass or not email_server\
            or not email_user or not email_pass or not email_to or not email_from:
        header_text = ['Welcome to STC Upload!', 'Settings haven''t been established.',
                       'Please fill out all empty fields below:']
    else:
        try:
            if not obj.test_sql_conn():
                header_text = ['Welcome to STC Upload!', 'Network settings are invalid.',
                               'Please fix the network settings below:']
            else:
                if email_port:
                    port = email_port.decrypt_text()
                else:
                    port = 587

                try:
                    server = smtplib.SMTP(str(email_server.decrypt_text()), int(port))

                    try:
                        server.ehlo()
                        server.starttls()
                        server.ehlo()
                        server.login(email_user.decrypt_text(), email_pass.decrypt_text())

                        try:
                            ftp = ftplib.FTP(ftp_host.decrypt_text())

                            try:
                                ftp.login(ftp_user.decrypt_text(), ftp_pass.decrypt_text())
                            except:
                                header_text = ['Welcome to STC Upload!', 'FTP User and/or Pass are invalid.',
                                               'Please fix below:']
                            finally:
                                ftp.quit()
                        except:
                            header_text = ['Welcome to STC Upload!', 'FTP server does not exist.', 'Please fix below:']
                    except:
                        header_text = ['Welcome to STC Upload!', 'Email User and/or Pass are invalid.',
                                       'Please fix below:']
                    finally:
                        server.close()
                except:
                    header_text = ['Welcome to STC Upload!', 'Email server does not exist.', 'Please fix below:']
        except:
            pass

    if header_text and not stop_gui:
        obj.build_gui(header_text)

    obj.cancel()
    del obj

    if header_text and not stop_gui:
        return check_settings(True)
    elif not header_text:
        return True


if __name__ == '__main__':
    if check_settings():
        rem_proc_elements()
        mobj = STCFTP()

        try:
            mobj.connect()
            mobj.grab_stc_zips()
        finally:
            mobj.close_connect()

        has_updates = check_for_updates()

        if has_updates:
            global_objs['Event_Log'].write_log('Found {0} files to process'.format(len(has_updates)))
            process_updates(has_updates)
            rem_proc_elements()
        else:
            global_objs['Event_Log'].write_log('Found no files to process', 'warning')
    else:
        global_objs['Event_Log'].write_log('Settings was not established. Exiting program', 'fatal')

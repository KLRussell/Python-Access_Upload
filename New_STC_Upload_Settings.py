from KGlobal import Toolbox
from os.path import join, dirname, abspath
from tkinter.messagebox import showerror
from tkinter import *

import sys

if getattr(sys, 'frozen', False):
    app_path = sys.executable
else:
    app_path = __file__

curr_dir = dirname(abspath(app_path))
main_dir = dirname(curr_dir)
logging_dir = join(main_dir, "01_Event_Logs")
process_dir = join(main_dir, "02_Process")
processed_dir = join(main_dir, "03_Processed")
sql_dir = join(main_dir, "05_SQL")
tool = Toolbox(app_path, logging_dir=logging_dir, logging_base_name="STC_Upload")
sql = tool.default_sql_conn()
email_engine = tool.default_exchange_conn()
local_config = tool.local_config


class FG(Tk):
    def __init__(self, header, ftp_server, ftp_user, ftp_pass):
        Tk.__init__(self)
        from KGlobal.data import CryptHandle

        if not isinstance(ftp_pass, (CryptHandle, type(None))):
            raise ValueError("'ftp_pass' %r is not an instance of CryptHandle" % ftp_pass)

        self.__ftp_server = StringVar()
        self.__ftp_user = StringVar()
        self.__ftp_pass = StringVar()
        self.ftp_server = ftp_server
        self.ftp_user = ftp_user

        if ftp_pass:
            self.ftp_pass = ftp_pass.peak()

        if ftp_pass:
            self.__ftp_pass_enc = ftp_pass
        else:
            self.__ftp_pass_enc = CryptHandle(alias='FTP_Pass', private=True)

        if header:
            self.__header = [header, 'Please fill out the information below:']
        else:
            self.__header = ['Welcome to setting up STC Upload FTP Setup!', 'Please fill out the information below:']

        self.__build()

    @property
    def ftp_server(self):
        return self.__ftp_server.get()

    @ftp_server.setter
    def ftp_server(self, ftp_server):
        if ftp_server is None:
            self.__ftp_server.set('')
        else:
            self.__ftp_server.set(ftp_server)

    @property
    def ftp_user(self):
        return self.__ftp_user.get()

    @ftp_user.setter
    def ftp_user(self, ftp_user):
        if ftp_user is None:
            self.__ftp_user.set('')
        else:
            self.__ftp_user.set(ftp_user)

    @property
    def ftp_pass(self):
        return self.__ftp_pass.get()

    @ftp_pass.setter
    def ftp_pass(self, ftp_pass):
        if ftp_pass is None:
            self.__ftp_pass.set('')
        else:
            self.__ftp_pass.set(ftp_pass)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('280x185+550+150')
        self.title('FTP Setup')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        ftp_frame = LabelFrame(self, text='FTP Settings', width=508, height=70)
        buttons_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        ftp_frame.pack(fill="both")
        buttons_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(self.__header), width=500, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to ftp_frame
        #    Apply FTP Server Entry Box Widget
        ftp_server_label = Label(ftp_frame, text='Server:')
        ftp_server_entry = Entry(ftp_frame, textvariable=self.__ftp_server, width=35)
        ftp_server_label.grid(row=0, column=0, padx=4, pady=5)
        ftp_server_entry.grid(row=0, column=1, columnspan=3, padx=4, pady=5)

        #    Apply FTP User Entry Box Widget
        ftp_user_label = Label(ftp_frame, text='User:')
        ftp_user_entry = Entry(ftp_frame, textvariable=self.__ftp_user, width=35)
        ftp_user_label.grid(row=1, column=0, padx=4, pady=5)
        ftp_user_entry.grid(row=1, column=1, padx=4, pady=5)

        #    Apply FTP Pass Entry Box Widget
        ftp_pass_label = Label(ftp_frame, text='Pass:')
        ftp_pass_entry = Entry(ftp_frame, textvariable=self.__ftp_pass, width=35)
        ftp_pass_label.grid(row=2, column=0, padx=4, pady=5)
        ftp_pass_entry.grid(row=2, column=1, padx=4, pady=5)
        ftp_pass_entry.bind('<KeyRelease>', self.__hide_ftp_pass)

        # Apply Widgets to buttons_frame
        #     Save button
        save_button = Button(self, text='Save', width=13, command=self.__save)
        save_button.pack(in_=buttons_frame, side=LEFT, padx=7, pady=7)

        #     Cancel button
        cancel_button = Button(self, text='Cancel', width=13, command=self.destroy)
        cancel_button.pack(in_=buttons_frame, side=RIGHT, padx=7, pady=7)

    def __hide_ftp_pass(self, event):
        curr_pass = self.__ftp_pass_enc.decrypt()

        if not curr_pass:
            curr_pass = ''

        curr_pass_len = len(curr_pass)

        if len(self.ftp_pass) > 0:
            curr_pass = self.__adjust_pass(curr_pass)

            if curr_pass_len > len(self.ftp_pass):
                curr_pass = curr_pass[:len(self.ftp_pass)]
        else:
            curr_pass = self.ftp_pass

        self.__ftp_pass_enc.encrypt(curr_pass)
        self.ftp_pass = self.__ftp_pass_enc.peak()

    def __adjust_pass(self, curr_pass):
        curr_pass_len = len(curr_pass)

        for pos, letter in enumerate(self.ftp_pass):
            if letter != '*':
                if pos > curr_pass_len - 1:
                    curr_pass += letter
                else:
                    my_text = list(curr_pass)
                    my_text.insert(pos, letter)
                    curr_pass = ''.join(my_text)

        return curr_pass

    def __save(self):
        if not self.ftp_server:
            showerror('Field Empty Error!', 'No value has been inputed for FTP Server', parent=self)
        elif not self.ftp_user:
            showerror('Field Empty Error!', 'No value has been inputed for FTP User', parent=self)
        elif not self.ftp_pass:
            showerror('Field Empty Error!', 'No value has been inputed for FTP Pass', parent=self)
        else:
            error = ftp_check(self.ftp_server, self.ftp_user, self.__ftp_pass_enc)

            if error == 1:
                showerror('FTP Connection Error!', 'Unable to connect to Server')
            elif error == 2:
                showerror(('FTP Credentials Error!', 'Credentials to FTP is incorrect'))
            else:
                local_config.setcrypt('FTP_Server', self.ftp_server)
                local_config.setcrypt('FTP_User', self.ftp_user)
                local_config['FTP_Pass'] = self.__ftp_pass_enc
                local_config.sync()
                self.destroy()


class FTPGUI(FG):
    def __init__(self, header=None, ftp_server=None, ftp_user=None, ftp_pass=None):
        FG.__init__(self, header=header, ftp_server=ftp_server, ftp_user=ftp_user, ftp_pass=ftp_pass)
        self.mainloop()


class ESG(Tk):
    def __init__(self, header, to_email, cc_email):
        Tk.__init__(self)

        if header:
            self.__header = [header, 'Please fill out the information below:']
        else:
            self.__header = ['Welcome to setting up Email Distro Setup!', 'Please fill out the information below:']

        self.__cc_email = StringVar()
        self.__to_email = StringVar()
        self.to_email = to_email
        self.cc_email = cc_email

        self.__build()

    @property
    def to_email(self):
        return self.__to_email.get()

    @to_email.setter
    def to_email(self, to_email):
        if to_email is None:
            self.__to_email.set('')
        else:
            self.__to_email.set(to_email)

    @property
    def cc_email(self):
        return self.__cc_email.get()

    @cc_email.setter
    def cc_email(self, cc_email):
        if cc_email is None:
            self.__cc_email.set('')
        else:
            self.__cc_email.set(cc_email)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('280x155+550+150')
        self.title('Email Distro Setup')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        email_frame = LabelFrame(self, text='Email Distro Settings', width=508, height=70)
        buttons_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        email_frame.pack(fill="both")
        buttons_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(self.__header), width=500, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to email_frame
        #    Apply To Email Entry Box Widget
        to_email_label = Label(email_frame, text='To Email:')
        to_email_entry = Entry(email_frame, textvariable=self.__to_email)
        to_email_label.grid(row=0, column=0, padx=4, pady=5)
        to_email_entry.grid(row=0, column=1, columnspan=3, padx=4, pady=5)

        #    Apply To Email Entry Box Widget
        cc_email_label = Label(email_frame, text='Cc Email:')
        cc_email_entry = Entry(email_frame, textvariable=self.__cc_email)
        cc_email_label.grid(row=1, column=0, padx=4, pady=5)
        cc_email_entry.grid(row=1, column=1, columnspan=3, padx=4, pady=5)

        # Apply Widgets to buttons_frame
        #     Save button
        save_button = Button(self, text='Save', width=13, command=self.__save)
        save_button.pack(in_=buttons_frame, side=LEFT, padx=7, pady=7)

        #     Cancel button
        cancel_button = Button(self, text='Cancel', width=13, command=self.destroy)
        cancel_button.pack(in_=buttons_frame, side=RIGHT, padx=7, pady=7)

    def __save(self):
        if not self.to_email:
            showerror('Field Empty Error!', 'No value has been inputed for To Email', parent=self)
        elif self.to_email.find('@') < 0:
            showerror('Email To Address Error!', '@ is not in the email to address field', parent=self)
        elif self.cc_email and self.cc_email.find('@') < 0:
            showerror('Email CC Address Error!', '@ is not in the email cc address field', parent=self)
        else:
            if self.to_email:
                local_config.setcrypt(key='Email_To', val=self.to_email)

            if self.cc_email:
                local_config.setcrypt(key='Email_Cc', val=self.cc_email)

            local_config.sync()
            self.destroy()


class EmailSetupGUI(ESG):
    def __init__(self, header=None, to_email=None, cc_email=None):
        ESG.__init__(self, header=header, to_email=to_email, cc_email=cc_email)


def ftp_check(ftp_server, ftp_user, ftp_pass):
    from KGlobal.data import CryptHandle

    if not ftp_server:
        raise ValueError("'ftp_server' is missing")
    elif not ftp_user:
        raise ValueError("'ftp_user' is missing")
    elif not ftp_pass:
        raise ValueError("'ftp_pass' is missing")
    elif not isinstance(ftp_pass, CryptHandle):
        raise ValueError("'ftp_pass' %r is not an instance of CryptHandle" % ftp_pass)
    else:
        from ftplib import FTP
        error = 0

        try:
            ftp = FTP(ftp_server)

            try:
                ftp.login(ftp_user, ftp_pass.decrypt())
            except:
                error = 2
                pass
            finally:
                ftp.quit()
        except:
            error = 1
            pass

        return error


def check_settings():
    if not local_config['FTP_Server'] or not local_config['FTP_User'] or not local_config['FTP_Pass']:
        FTPGUI()
    else:
        error = ftp_check(local_config['FTP_Server'].decrypt(), local_config['FTP_User'].decrypt(),
                          local_config['FTP_Pass'])

        if error == 1:
            FTPGUI('Welcome! Was unable to connect to server. Please re-check server address!')
        elif error == 2:
            FTPGUI('Welcome! Was unable to authenticate with credentials. Please check credentials!')

    if not local_config['Email_To']:
        EmailSetupGUI()

    if local_config['FTP_Server'] and local_config['FTP_User'] and local_config['FTP_Pass']\
            and local_config['Email_To']:
        return True



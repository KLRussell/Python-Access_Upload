from tkinter.messagebox import showerror, askokcancel
from tkinter import *
from New_STC_Upload_Settings import local_config, sql, tool
from New_STC_Upload_Class import check_processed


class AccProfileGUI(Toplevel):
    __acc_tbl_entry = None
    __acc_col_list = None
    __acc_col_sel_list = None
    __sql_tbl_entry = None
    __sql_col_list = None
    __sql_col_sel_list = None
    __prev_sql_tbl_txt = None

    def __init__(self, parent, sql_tbl_list, profile, grandparent=None):
        Toplevel.__init__(self)

        if not isinstance(sql_tbl_list, list):
            raise ValueError("'sql_tbl_list' %r is not an instance of List" % sql_tbl_list)

        if not isinstance(profile, dict):
            raise ValueError("'profile' %r is not an instance of Dict" % profile)

        self.__grandparent = grandparent
        self.__parent = parent
        self.__sql_tbl_list = sql_tbl_list
        self.__profile = profile
        self.__acc_tbl = StringVar()
        self.__sql_tbl = StringVar()
        self.__truncate_tbl = IntVar()

        if 'Header' in profile.keys():
            self.__header = [profile['Header'], 'Please fill out the information below:']
        else:
            self.__header = ['Welcome to setting up Access Profile Setup!', 'Please fill out the information below:']

        self.__build()
        self.__fill_gui()

    @property
    def acc_tbl(self):
        return self.__acc_tbl.get()

    @property
    def acc_col_sel(self):
        return self.__acc_col_sel_list.get(0, self.__acc_col_sel_list.size() - 1)

    @property
    def sql_tbl(self):
        return self.__sql_tbl.get()

    @property
    def sql_col_list(self):
        return self.__sql_col_list.get(0, self.__sql_col_list.size() - 1) + \
               self.__sql_col_sel_list.get(0, self.__sql_col_sel_list.size() - 1)

    @property
    def sql_col_sel(self):
        return self.__sql_col_sel_list.get(0, self.__sql_col_sel_list.size() - 1)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('540x612+500+70')
        self.title('Access Profile Setup')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        acc_frame = LabelFrame(self, text='Access Settings', width=508, height=70)
        acc_tbl_frame = Frame(acc_frame)
        acc_col_frame = Frame(acc_frame)
        acc_col_list_frame = LabelFrame(acc_col_frame, text='Column List')
        acc_col_button_frame = Frame(acc_col_frame)
        acc_col_sel_list_frame = LabelFrame(acc_col_frame, text='Column Selected List')
        sql_frame = LabelFrame(self, text='SQL Settings', width=508, height=70)
        sql_tbl_frame = Frame(sql_frame)
        sql_col_frame = Frame(sql_frame)
        sql_col_list_frame = LabelFrame(sql_col_frame, text='Column List')
        sql_col_button_frame = Frame(sql_col_frame)
        sql_col_sel_list_frame = LabelFrame(sql_col_frame, text='Column Selected List')
        buttons_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        acc_frame.pack(fill="both")
        acc_tbl_frame.pack(fill="both")
        acc_col_frame.pack(fill="both")
        acc_col_list_frame.grid(row=0, column=0)
        acc_col_button_frame.grid(row=0, column=1, padx=5)
        acc_col_sel_list_frame.grid(row=0, column=2, padx=5)
        sql_frame.pack(fill="both")
        sql_tbl_frame.pack(fill="both")
        sql_col_frame.pack(fill="both")
        sql_col_list_frame.grid(row=0, column=0)
        sql_col_button_frame.grid(row=0, column=1, padx=5)
        sql_col_sel_list_frame.grid(row=0, column=2, padx=5)
        buttons_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(self.__header), width=500, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply Widgets to the Acc_TBL_Frame
        #     Access Table Name Input Box
        acc_tbl_label = Label(acc_tbl_frame, text='Acc TBL:')
        self.__acc_tbl_entry = Entry(acc_tbl_frame, textvariable=self.__acc_tbl, width=64)
        acc_tbl_label.grid(row=0, column=0, padx=8, pady=5)
        self.__acc_tbl_entry.grid(row=0, column=1, padx=5, pady=5)
        self.__acc_tbl_entry.configure(state=DISABLED)

        #     Access Column List
        xbar = Scrollbar(acc_col_list_frame, orient='horizontal')
        ybar = Scrollbar(acc_col_list_frame, orient='vertical')
        self.__acc_col_list = Listbox(acc_col_list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                      xscrollcommand=xbar)
        xbar.config(command=self.__acc_col_list.xview)
        ybar.config(command=self.__acc_col_list.yview)
        self.__acc_col_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W+E)
        ybar.grid(row=0, column=1, sticky=N+S)
        self.__acc_col_list.bind("<Down>", self.__list_action)
        self.__acc_col_list.bind("<Up>", self.__list_action)

        #     Access Column Migration Buttons
        acc_right_button = Button(acc_col_button_frame, text='>', width=5)
        acc_all_right_button = Button(acc_col_button_frame, text='>>', width=5)
        acc_left_button = Button(acc_col_button_frame, text='<', width=5)
        acc_all_left_button = Button(acc_col_button_frame, text='<<', width=5)
        acc_right_button.grid(row=0, column=2, padx=7, pady=7)
        acc_all_right_button.grid(row=1, column=2, padx=7, pady=7)
        acc_left_button.grid(row=2, column=2, padx=7, pady=7)
        acc_all_left_button.grid(row=3, column=2, padx=7, pady=7)
        acc_right_button.bind("<1>", self.__migrate_acc)
        acc_all_right_button.bind("<1>", self.__migrate_acc)
        acc_left_button.bind("<1>", self.__migrate_acc)
        acc_all_left_button.bind("<1>", self.__migrate_acc)

        #     Access Column Selected List
        xbar = Scrollbar(acc_col_sel_list_frame, orient='horizontal')
        ybar = Scrollbar(acc_col_sel_list_frame, orient='vertical')
        self.__acc_col_sel_list = Listbox(acc_col_sel_list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                          xscrollcommand=xbar)
        xbar.config(command=self.__acc_col_sel_list.xview)
        ybar.config(command=self.__acc_col_sel_list.yview)
        self.__acc_col_sel_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__acc_col_sel_list.bind("<Down>", self.__list_action)
        self.__acc_col_sel_list.bind("<Up>", self.__list_action)

        # Apply Widgets to the SQL_TBL_Frame
        #     SQL Table Name Input Box
        sql_tbl_label = Label(sql_tbl_frame, text='SQL TBL:')
        self.__sql_tbl_entry = Entry(sql_tbl_frame, textvariable=self.__sql_tbl, width=52)
        sql_tbl_label.grid(row=0, column=0, padx=8, pady=5)
        self.__sql_tbl_entry.grid(row=0, column=1, padx=5, pady=5)
        self.__sql_tbl_entry.bind('<KeyRelease>', self.__fill_sql_cols)

        tbl_trunc_chkbox = Checkbutton(sql_tbl_frame, text='Truncate TBL', variable=self.__truncate_tbl)
        tbl_trunc_chkbox.grid(row=0, column=2, padx=5, pady=5)

        #     SQL Column List
        xbar = Scrollbar(sql_col_list_frame, orient='horizontal')
        ybar = Scrollbar(sql_col_list_frame, orient='vertical')
        self.__sql_col_list = Listbox(sql_col_list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                      xscrollcommand=xbar)
        xbar.config(command=self.__sql_col_list.xview)
        ybar.config(command=self.__sql_col_list.yview)
        self.__sql_col_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__sql_col_list.bind("<Down>", self.__list_action)
        self.__sql_col_list.bind("<Up>", self.__list_action)

        #     SQL Column Migration Buttons
        sql_right_button = Button(sql_col_button_frame, text='>', width=5)
        sql_all_right_button = Button(sql_col_button_frame, text='>>', width=5)
        sql_left_button = Button(sql_col_button_frame, text='<', width=5)
        sql_all_left_button = Button(sql_col_button_frame, text='<<', width=5)
        sql_right_button.grid(row=0, column=2, padx=7, pady=7)
        sql_all_right_button.grid(row=1, column=2, padx=7, pady=7)
        sql_left_button.grid(row=2, column=2, padx=7, pady=7)
        sql_all_left_button.grid(row=3, column=2, padx=7, pady=7)
        sql_right_button.bind("<1>", self.__migrate_sql)
        sql_all_right_button.bind("<1>", self.__migrate_sql)
        sql_left_button.bind("<1>", self.__migrate_sql)
        sql_all_left_button.bind("<1>", self.__migrate_sql)

        #     SQL Column Selected List
        xbar = Scrollbar(sql_col_sel_list_frame, orient='horizontal')
        ybar = Scrollbar(sql_col_sel_list_frame, orient='vertical')
        self.__sql_col_sel_list = Listbox(sql_col_sel_list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                          xscrollcommand=xbar)
        xbar.config(command=self.__sql_col_sel_list.xview)
        ybar.config(command=self.__sql_col_sel_list.yview)
        self.__sql_col_sel_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__sql_col_sel_list.bind("<Down>", self.__list_action)
        self.__sql_col_sel_list.bind("<Up>", self.__list_action)

        # Apply Buttons to Button_Frame
        #     Save Button
        save_button = Button(self, text='Save', width=15, command=self.__save_settings)
        save_button.pack(in_=buttons_frame, side=LEFT, padx=10, pady=5)

        #     Cancel Button
        cancel_button = Button(self, text='Cancel', width=15, command=self.destroy)
        cancel_button.pack(in_=buttons_frame, side=RIGHT, padx=10, pady=5)

    def __fill_gui(self):
        if 'Acc_TBL' in self.__profile.keys():
            self.__acc_tbl.set(self.__profile['Acc_TBL'])

        if 'Acc_Cols' in self.__profile.keys():
            if 'Acc_Cols_Sel' in self.__profile.keys():
                acc_cols_sel = [col.lower() for col in self.__profile['Acc_Cols_Sel']]
            else:
                acc_cols_sel = [None]

            for col in self.__profile['Acc_Cols']:
                if col.lower() in acc_cols_sel:
                    self.__acc_col_sel_list.insert('end', col)
                else:
                    self.__acc_col_list.insert('end', col)

        if 'SQL_TBL' in self.__profile.keys():
            self.__prev_sql_tbl_txt = self.__profile['SQL_TBL']
            self.__sql_tbl.set(self.__profile['SQL_TBL'])

        if 'SQL_TBL_Trunc' in self.__profile.keys():
            self.__truncate_tbl.set(self.__profile['SQL_TBL_Trunc'])

        if 'SQL_Cols' in self.__profile.keys():
            if 'SQL_Cols_Sel' in self.__profile.keys():
                sql_cols_sel = [col.lower() for col in self.__profile['SQL_Cols_Sel']]
            else:
                sql_cols_sel = [None]

            for col in self.__profile['SQL_Cols']:
                if col.lower() in sql_cols_sel:
                    self.__sql_col_sel_list.insert('end', col)
                else:
                    self.__sql_col_list.insert('end', col)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and (event.keysym == 'Down' and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and event.keysm in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

    def __fill_sql_cols(self, event):
        widget = event.widget

        if not self.__prev_sql_tbl_txt or widget.get().lower() != self.__prev_sql_tbl_txt.lower():
            self.__prev_sql_tbl_txt = widget.get()

            if self.__sql_col_list.size():
                self.__sql_col_list.delete(0, self.__sql_col_list.size() - 1)

            if self.__sql_col_sel_list.size():
                self.__sql_col_sel_list.delete(0, self.__sql_col_sel_list.size() - 1)

            if widget.get().lower() in self.__sql_tbl_list:
                table = widget.get().split('.')
                query = '''
                    SELECT
                        Column_Name
    
                    FROM INFORMATION_SCHEMA.COLUMNS
    
                    WHERE
                        Table_Schema = '{0}'
                            AND
                        Table_Name = '{1}'
                '''.format(table[0], table[1])

                results = sql.sql_execute(query_str=query, execute=False)

                if results and results.results:
                    for col in results.results[0]['Column_Name'].tolist():
                        self.__sql_col_list.insert('end', col)

    def __migrate_acc(self, event):
        widget = event.widget

        if widget.cget('text') == '>' and self.__acc_col_list.curselection():
            self.__acc_col_sel_list.insert('end', self.__acc_col_list.get(self.__acc_col_list.curselection()))
            self.__acc_col_list.delete(self.__acc_col_list.curselection(), self.__acc_col_list.curselection())
            self.after_idle(self.__acc_col_sel_list.select_set, self.__acc_col_sel_list.size() - 1)
        elif widget.cget('text') == '<' and self.__acc_col_sel_list.curselection():
            self.__acc_col_list.insert('end', self.__acc_col_sel_list.get(self.__acc_col_sel_list.curselection()))
            self.__acc_col_sel_list.delete(self.__acc_col_sel_list.curselection(),
                                           self.__acc_col_sel_list.curselection())
            self.after_idle(self.__acc_col_list.select_set, self.__acc_col_list.size() - 1)
        elif widget.cget('text') == '>>' and self.__acc_col_list.size() > 0:
            for i in range(self.__acc_col_list.size()):
                self.__acc_col_sel_list.insert('end', self.__acc_col_list.get(i))

            self.__acc_col_list.delete(0, self.__acc_col_list.size() - 1)
            self.after_idle(self.__acc_col_sel_list.select_set, self.__acc_col_sel_list.size() - 1)
        elif widget.cget('text') == '<<' and self.__acc_col_sel_list.size() > 0:
            for i in range(self.__acc_col_sel_list.size()):
                self.__acc_col_list.insert('end', self.__acc_col_sel_list.get(i))

            self.__acc_col_sel_list.delete(0, self.__acc_col_sel_list.size() - 1)
            self.after_idle(self.__acc_col_list.select_set, self.__acc_col_list.size() - 1)

    def __migrate_sql(self, event):
        widget = event.widget

        if widget.cget('text') == '>' and self.__sql_col_list.curselection():
            self.__sql_col_sel_list.insert('end', self.__sql_col_list.get(self.__sql_col_list.curselection()))
            self.__sql_col_list.delete(self.__sql_col_list.curselection(), self.__sql_col_list.curselection())
            self.after_idle(self.__sql_col_sel_list.select_set, self.__sql_col_sel_list.size() - 1)
        elif widget.cget('text') == '<' and self.__sql_col_sel_list.curselection():
            self.__sql_col_list.insert('end', self.__sql_col_sel_list.get(self.__sql_col_sel_list.curselection()))
            self.__sql_col_sel_list.delete(self.__sql_col_sel_list.curselection(),
                                           self.__sql_col_sel_list.curselection())
            self.after_idle(self.__sql_col_list.select_set, self.__sql_col_list.size() - 1)
        elif widget.cget('text') == '>>' and self.__sql_col_list.size() > 0:
            for i in range(self.__sql_col_list.size()):
                self.__sql_col_sel_list.insert('end', self.__sql_col_list.get(i))

            self.__sql_col_list.delete(0, self.__sql_col_list.size() - 1)
            self.after_idle(self.__sql_col_sel_list.select_set, self.__sql_col_sel_list.size() - 1)
        elif widget.cget('text') == '<<' and self.__sql_col_sel_list.size() > 0:
            for i in range(self.__sql_col_sel_list.size()):
                self.__sql_col_list.insert('end', self.__sql_col_sel_list.get(i))

            self.__sql_col_sel_list.delete(0, self.__sql_col_sel_list.size() - 1)
            self.after_idle(self.__sql_col_list.select_set, self.__sql_col_list.size() - 1)

    def __save_settings(self):
        if not self.acc_col_sel:
            showerror('List Empty Error!', 'No Access columns were selected', parent=self)
        elif not self.sql_tbl:
            showerror('Field Empty Error!', 'No value has been inputed for SQL TBL', parent=self)
        elif not self.sql_col_sel:
            showerror('List Empty Error!', 'No SQL columns were selected', parent=self)
        elif len(self.acc_col_sel) != len(self.sql_col_sel):
            showerror('List Col Unequal Error!', 'Access Selected Columns dont match SQL Selected Columns', parent=self)
        else:
            if 'Header' in self.__profile.keys():
                del self.__profile['Header']

            self.__profile['Acc_Cols_Sel'] = self.acc_col_sel
            self.__profile['SQL_TBL'] = self.sql_tbl
            self.__profile['SQL_TBL_Trunc'] = self.__truncate_tbl.get()
            self.__profile['SQL_Cols'] = self.sql_col_list
            self.__profile['SQL_Cols_Sel'] = self.sql_col_sel
            profiles = local_config['Profiles']

            if not profiles:
                profiles = dict()

            profiles[self.__profile['Acc_TBL']] = self.__profile
            local_config['Profiles'] = profiles

            if local_config['Err_Profiles']:
                profiles = local_config['Err_Profiles']

                if self.__profile['Acc_TBL'] in profiles.keys():
                    del profiles[self.__profile['Acc_TBL']]

                    if len(profiles) > 0:
                        local_config['Err_Profiles'] = profiles
                    else:
                        del local_config['Err_Profiles']

            local_config.sync()
            if self.__grandparent:
                self.__grandparent.load_gui()

            self.__parent.load_gui()
            self.destroy()


class EmailModify(Toplevel):
    def __init__(self):
        Toplevel.__init__(self)

        self.__header = ['Welcome to setting up Email Distro Setup!', 'Please fill out the information below:']
        self.__cc_email = StringVar()
        self.__to_email = StringVar()

        self.__build()
        self.__load_gui()

    @property
    def to_email(self):
        if self.__to_email.get():
            return self.__to_email.get()
        else:
            return None

    @to_email.setter
    def to_email(self, to_email):
        if to_email is None:
            self.__to_email.set('')
        else:
            self.__to_email.set(to_email)

    @property
    def cc_email(self):
        if self.__cc_email.get():
            return self.__cc_email.get()
        else:
            return None

    @cc_email.setter
    def cc_email(self, cc_email):
        if cc_email is None:
            self.__cc_email.set('')
        else:
            self.__cc_email.set(cc_email)

    def __build(self):
        # Set GUI Geometry and GUI Title
        self.geometry('280x155+670+350')
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

    def __load_gui(self):
        if local_config['Email_To']:
            self.to_email = local_config['Email_To'].decrypt()

        if local_config['Email_Cc']:
            self.cc_email = local_config['Email_Cc'].decrypt()

    def __save(self):
        if not self.to_email:
            showerror('Field Empty Error!', 'No value has been inputed for To Email', parent=self)
        elif self.to_email.find('@') < 0:
            showerror('Email To Address Error!', '@ is not in the email to address field', parent=self)
        elif self.cc_email and self.cc_email.find('@') < 0:
            showerror('Email CC Address Error!', '@ is not in the email cc address field', parent=self)
        else:
            if local_config['Email_To']:
                local_config['Email_To'] = self.to_email
            elif self.to_email or self.cc_email:
                if self.to_email:
                    local_config.setcrypt(key='Email_To', val=self.to_email)

                if self.cc_email:
                    local_config.setcrypt(key='Email_Cc', val=self.cc_email)

                local_config.sync()

            self.destroy()


class ManualUpload(Toplevel):
    __upload_list = None

    def __init__(self, table_name):
        Toplevel.__init__(self)
        self.__table_name = table_name
        self.__build()
        self.__load_gui()

    def __build(self):
        header = ['Welcome to Manual Upload!', 'Feel free to manually upload a previous batch:']

        # Set GUI Geometry and GUI Title
        self.geometry('355x285+630+290')
        self.title('Manual Upload')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        main_frame = LabelFrame(self, text='Batch List', width=444, height=140)
        list_frame = Frame(main_frame)
        button_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        list_frame.grid(row=0, column=0)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(header), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply List Widget to list_frame
        #     Access Profiles List
        xbar = Scrollbar(list_frame, orient='horizontal')
        ybar = Scrollbar(list_frame, orient='vertical')
        self.__upload_list = Listbox(list_frame, selectmode=SINGLE, width=50, yscrollcommand=ybar,
                                     xscrollcommand=xbar)
        xbar.config(command=self.__upload_list.xview)
        ybar.config(command=self.__upload_list.yview)
        self.__upload_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__upload_list.bind("<Down>", self.__list_action)
        self.__upload_list.bind("<Up>", self.__list_action)

        # Apply button Widgets to button_frame
        #     Cancel Button
        button = Button(button_frame, text='Upload', width=20, command=self.__upload)
        button.grid(row=0, column=0, padx=10, pady=5)

        #     Cancel Button
        button = Button(button_frame, text='Cancel', width=20, command=self.destroy)
        button.grid(row=0, column=1, padx=10, pady=5)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and (event.keysym == 'Down' and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and event.keysm in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

    def __load_gui(self):
        processed = local_config['Processed']

        if self.__table_name in processed.keys():
            for proc_list in processed[self.__table_name]:
                self.__upload_list.insert('end', proc_list[0])

    def __upload(self):
        if self.__upload_list.size() > 0 and self.__upload_list.curselection():
            file_path = None
            processed = local_config['Processed']
            batch = self.__upload_list.get(self.__upload_list.curselection()[0])

            if self.__table_name in processed.keys():
                for proc_list in processed[self.__table_name]:
                    if proc_list[0] == batch:
                        file_path = proc_list[1]
                        break

            if self.__table_name and file_path and batch:
                from New_STC_Upload_Class import AccToSQL
                from threading import Thread

                tool.gui_console(title="Manual Upload Log")
                acc_obj = AccToSQL()
                Thread(target=acc_obj.manual_process, kwargs={'table': self.__table_name, 'accdb_file': file_path,
                                                              'batch': batch}).start()


class ErrorProfiles(Toplevel):
    __sql_tables = None
    __acc_profiles_list = None
    __error_button = None

    def __init__(self, parent, sql_tables):
        Toplevel.__init__(self)
        self.sql_tables = sql_tables
        self.__parent = parent
        self.__build()
        self.load_gui()

    @property
    def sql_tables(self):
        return self.__sql_tables

    @sql_tables.setter
    def sql_tables(self, sql_tables):
        self.__sql_tables = sql_tables

    def __build(self):
        header = ['Welcome to Error Profile List!', 'Feel free to add or delete a profile:']

        # Set GUI Geometry and GUI Title
        self.geometry('355x285+630+290')
        self.title('Error Profile List')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        main_frame = LabelFrame(self, text='Error Profiles', width=444, height=140)
        list_frame = Frame(main_frame)
        control_frame = Frame(main_frame)
        button_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        list_frame.grid(row=0, column=0)
        control_frame.grid(row=0, column=1, padx=5)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(header), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply List Widget to list_frame
        #     Access Profiles List
        xbar = Scrollbar(list_frame, orient='horizontal')
        ybar = Scrollbar(list_frame, orient='vertical')
        self.__acc_profiles_list = Listbox(list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                           xscrollcommand=xbar)
        xbar.config(command=self.__acc_profiles_list.xview)
        ybar.config(command=self.__acc_profiles_list.yview)
        self.__acc_profiles_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__acc_profiles_list.bind("<Down>", self.__list_action)
        self.__acc_profiles_list.bind("<Up>", self.__list_action)

        # Apply button Widgets to control_frame
        #     Access Column Migration Buttons
        modify_button = Button(control_frame, text='Add', width=15)
        delete_button = Button(control_frame, text='Delete', width=15)
        modify_button.grid(row=0, column=2, padx=7, pady=7)
        delete_button.grid(row=1, column=2, padx=7, pady=7)
        modify_button.bind("<1>", self.__profile_action)
        delete_button.bind("<ButtonRelease-1>", self.__profile_action)

        # Apply button Widgets to button_frame
        #     Cancel Button
        button = Button(button_frame, text='Cancel', width=46, command=self.destroy)
        button.grid(row=0, column=2, padx=10, pady=5)

    def load_gui(self):
        if self.__acc_profiles_list.size() > 0:
            self.__acc_profiles_list.delete(0, self.__acc_profiles_list.size() - 1)

        profiles = local_config['Err_Profiles']

        if profiles:
            for profile_name in profiles.keys():
                self.__acc_profiles_list.insert('end', profile_name)

            self.after_idle(self.__acc_profiles_list.select_set, 0)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and (event.keysym == 'Down' and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and event.keysm in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

    def __profile_action(self, event):
        selections = self.__acc_profiles_list.curselection()

        if selections:
            widget = event.widget
            profiles = local_config['Err_Profiles']
            profile_name = self.__acc_profiles_list.get(selections[0])

            if widget.cget('text') == 'Add' and profile_name in profiles.keys():
                AccProfileGUI(grandparent=self.__parent, parent=self, sql_tbl_list=self.sql_tables,
                              profile=profiles[profile_name])
            elif widget.cget('text') == 'Delete':
                myresponse = askokcancel(
                    'Delete Notice!',
                    'Deleting this profile will lose this profile forever. Would you like to proceed?',
                    parent=self)

                if myresponse:
                    if profile_name in profiles.keys():
                        del profiles[profile_name]

                        if len(profiles) > 0:
                            local_config['Err_Profiles'] = profiles
                        else:
                            del local_config['Err_Profiles']

                    self.__acc_profiles_list.delete(selections[0], selections[0])

                    if self.__acc_profiles_list.size() > 0 and selections[0] == 0:
                        self.after_idle(widget.select_set, 0)
                    elif self.__acc_profiles_list.size() > 0:
                        self.after_idle(widget.select_set, selections[0] - 1)


class APL(Tk):
    __sql_tables = None
    __acc_profiles_list = None
    __error_button = None

    def __init__(self):
        Tk.__init__(self)
        self.sql_tables = sql.sql_tables()
        self.__build()
        self.load_gui()

    @property
    def sql_tables(self):
        return self.__sql_tables

    @sql_tables.setter
    def sql_tables(self, sql_tables):
        from KGlobal.sql.cursor import SQLCursor

        if isinstance(sql_tables, SQLCursor):
            if sql_tables.errors:
                raise ValueError("Error Code '{0}', {1}".format(sql_tables.errors[0], sql_tables.errors[1]))
            else:
                df = sql_tables.results[0]
                df = df.loc[(df['Table_Name'] != 'msys') & (df['Table_Type'] == 'TABLE')]
                df['Table'] = df['Table_Schema'].str.cat(df['Table_Name'], sep=".")
                self.__sql_tables = [tbl.lower() for tbl in df['Table'].tolist()]
        elif isinstance(sql_tables, type(None)):
            self.__sql_tables = None
        else:
            raise ValueError("'sql_tables' %r is not an instance of SQLCursor" % sql_tables)

    def __build(self):
        header = ['Welcome to Access Profile List!', 'Feel free to modify, delete, or change settings:']

        # Set GUI Geometry and GUI Title
        self.geometry('355x285+630+290')
        self.title('Access Profile List')
        self.resizable(False, False)

        # Set GUI Frames
        header_frame = Frame(self)
        main_frame = LabelFrame(self, text='Access Profiles', width=444, height=140)
        list_frame = Frame(main_frame)
        control_frame = Frame(main_frame)
        button_frame = Frame(self)

        # Apply Frames into GUI
        header_frame.pack(fill="both")
        main_frame.pack(fill="both")
        list_frame.grid(row=0, column=0)
        control_frame.grid(row=0, column=1, padx=5)
        button_frame.pack(fill="both")

        # Apply Header text to Header_Frame that describes purpose of GUI
        header = Message(self, text='\n'.join(header), width=375, justify=CENTER)
        header.pack(in_=header_frame)

        # Apply List Widget to list_frame
        #     Access Profiles List
        xbar = Scrollbar(list_frame, orient='horizontal')
        ybar = Scrollbar(list_frame, orient='vertical')
        self.__acc_profiles_list = Listbox(list_frame, selectmode=SINGLE, width=30, yscrollcommand=ybar,
                                           xscrollcommand=xbar)
        xbar.config(command=self.__acc_profiles_list.xview)
        ybar.config(command=self.__acc_profiles_list.yview)
        self.__acc_profiles_list.grid(row=0, column=0, padx=8, pady=5)
        xbar.grid(row=1, column=0, sticky=W + E)
        ybar.grid(row=0, column=1, sticky=N + S)
        self.__acc_profiles_list.bind("<Down>", self.__list_action)
        self.__acc_profiles_list.bind("<Up>", self.__list_action)

        # Apply button Widgets to control_frame
        #     Access Column Migration Buttons
        manual_upload_button = Button(control_frame, text='Manual Upload', width=15)
        modify_button = Button(control_frame, text='Modify', width=15)
        delete_button = Button(control_frame, text='Delete', width=15)
        manual_upload_button.grid(row=0, column=2, padx=7, pady=7)
        modify_button.grid(row=1, column=2, padx=7, pady=7)
        delete_button.grid(row=2, column=2, padx=7, pady=7)
        manual_upload_button.bind("<1>", self.__profile_action)
        modify_button.bind("<1>", self.__profile_action)
        delete_button.bind("<ButtonRelease-1>", self.__profile_action)

        # Apply button Widgets to button_frame
        #     Mail Settings Button
        button = Button(button_frame, text='Mail Settings', width=13, command=self.__mail_settings)
        button.grid(row=0, column=0, padx=10, pady=5)

        #     Error Profiles Button
        self.__error_button = Button(button_frame, text='Error Profiles', width=12, command=self.__error_profiles)
        self.__error_button.grid(row=0, column=1, padx=10, pady=5)

        #     Cancel Button
        button = Button(button_frame, text='Cancel', width=13, command=self.destroy)
        button.grid(row=0, column=2, padx=10, pady=5)

    def load_gui(self):
        if self.__acc_profiles_list.size() > 0:
            self.__acc_profiles_list.delete(0, self.__acc_profiles_list.size() - 1)

        profiles = local_config['Profiles']

        if profiles:
            for profile_name in profiles.keys():
                self.__acc_profiles_list.insert('end', profile_name)

            self.after_idle(self.__acc_profiles_list.select_set, 0)

    def __list_action(self, event):
        widget = event.widget

        if widget.size() > 0:
            selections = widget.curselection()

            if selections and (event.keysym == 'Up' and selections[0] > 0):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] - 1)
            elif selections and (event.keysym == 'Down' and selections[0] < widget.size() - 1):
                self.after_idle(widget.select_clear, selections[0])
                self.after_idle(widget.select_set, selections[0] + 1)
            elif not selections and event.keysm in ('Up', 'Down'):
                self.after_idle(widget.select_set, 0)

    def __profile_action(self, event):
        selections = self.__acc_profiles_list.curselection()

        if selections:
            widget = event.widget
            profiles = local_config['Profiles']
            profile_name = self.__acc_profiles_list.get(selections[0])

            if widget.cget('text') == 'Manual Upload' and profile_name in profiles.keys():
                ManualUpload(profile_name)
            elif widget.cget('text') == 'Modify' and profile_name in profiles.keys():
                AccProfileGUI(parent=self, sql_tbl_list=self.sql_tables,
                              profile=profiles[profile_name])
            elif widget.cget('text') == 'Delete':
                myresponse = askokcancel(
                    'Delete Notice!',
                    'Deleting this profile will lose this profile forever. Would you like to proceed?',
                    parent=self)

                if myresponse:
                    if profile_name in profiles.keys():
                        del profiles[profile_name]

                        if len(profiles) > 0:
                            local_config['Profiles'] = profiles
                        else:
                            del local_config['Profiles']

                        local_config.sync()

                    self.__acc_profiles_list.delete(selections[0], selections[0])

                    if self.__acc_profiles_list.size() > 0 and selections[0] == 0:
                        self.after_idle(self.__acc_profiles_list.select_set, 0)
                    elif self.__acc_profiles_list.size() > 0:
                        self.after_idle(self.__acc_profiles_list.select_set, selections[0] - 1)

    @staticmethod
    def __mail_settings():
        EmailModify()

    def __error_profiles(self):
        ErrorProfiles(parent=self, sql_tables=self.sql_tables)


class AccProfileList(APL):
    def __init__(self):
        APL.__init__(self)
        self.mainloop()


if __name__ == '__main__':
    check_processed()
    obj = AccProfileList()

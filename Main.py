# If python 32-bit and MS Access engine is 64-bit, use this 32-bit MS Access driver
# http://www.microsoft.com/en-us/download/details.aspx?id=13255

from Global import grabobjs
from Global import SQLHandle
from Settings import SettingsGUI

import os
import pathlib as pl
import datetime

CurrDir = os.path.dirname(os.path.abspath(__file__))
AccdbDir = os.path.join(CurrDir, "02_Process")
ProcessedDir = os.path.join(CurrDir, "03_Processed")
global_objs = grabobjs(CurrDir, 'AccessDB')


class AccdbHandle:
    config = None
    accdb_cols = None
    sql_cols = None

    def __init__(self, file):
        self.file = file
        self.asql = SQLHandle(global_objs['Settings'])
        self.asql.connect(conn_type='alch')
        self.configs = global_objs['Local_Settings'].grab_item('Accdb_Configs')

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

        if not self.configs:
            header_text = 'Welcome to Access DB Upload!\nThere is no configuration for table ({0}) in file ({1}).\nPlease add configuration setting below:'
            self.config_gui(table, header_text)
        else:
            self.get_config(table)

        if self.config and not self.validate_sql_table(self.config[3]):
            header_text = 'Welcome to Access DB Upload!\nTable ({0}) does not exist in SQL Server.\nPlease fix configuration in Upload Settings:'
            self.config_gui(table, header_text, False)

        if self.config and not self.validate_cols(self.accdb_cols, self.config[2]):
            self.config[1] = self.accdb_cols
            self.config[2] = None
            self.switch_config()
            header_text = 'Welcome to Access DB Upload!\nOne or more columns for access columns does not exist anymore.\nPlease redo configuration for access table columns:'
            self.config_gui(table, header_text, False)

        if self.config:
            self.get_sql_cols(self.config[3])

            if not self.validate_cols(self.sql_cols, self.config[4]):
                self.config[4] = None
                self.switch_config()
                header_text = 'Welcome to Access DB Upload!\nOne or more columns for sql table columns does not exist anymore.\nPlease redo configuration for sql table columns:'
                self.config_gui(table, header_text, False)

        if self.config:
            return True
        else:
            return False

    def config_gui(self, table, header_text, insert=True):
        obj = SettingsGUI()

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
        global_objs['Event_Log'].write_log('Uploading data from table [{0}] to sql table {1}'
                                           .format(table, self.config[3]))

        myresults = global_objs['SQL'].query('''
            SELECT
                [{0}]
            
            FROM [{1}]
        '''.format('], ['.join(self.config[2]), table))

        if not myresults.empty:
            myresults.columns = self.config[4]

            if self.config[5]:
                global_objs['Event_Log'].write_log('Truncating table [{0}]'.format(self.config[3]))
                self.asql.execute('truncate table {0}'.format(self.config[3]))

            self.asql.upload(myresults, self.config[3], index=False, index_label=None)
            global_objs['Event_Log'].write_log('Data successfully uploaded from table [{0}] to sql table {1}'
                                               .format(table, self.config[3]))
            return True
        else:
            global_objs['Event_Log'].write_log('Failed to grab data from access table [{0}]. No update made'
                                               .format(table), 'error')
            return False

    def close_asql(self):
        self.asql.close()


def check_for_updates():
    f = list(pl.Path(AccdbDir).glob('*.accdb'))

    if f:
        return f

    f = list(pl.Path(AccdbDir).glob('*.mdb'))

    if f:
        return f


def process_updates(files):
    for file in files:
        processed = False
        global_objs['Event_Log'].write_log('Processing file {0}'.format(os.path.basename(file)))
        myobj = AccdbHandle(file)

        global_objs['SQL'].connect('accdb', accdb_file=file)

        try:
            for table in myobj.get_accdb_tables():
                global_objs['Event_Log'].write_log('Validating table [{0}]'.format(table))

                if myobj.validate(table) and myobj.process(table):
                    processed = True

        finally:
            myobj.close_asql()
            global_objs['SQL'].close()

        if processed:
            filename = os.path.basename(file)
            os.rename(file, os.path.join(ProcessedDir, '{0}_{1}{2}'.format(
                datetime.datetime.now().__format__("%Y%m%d"), os.path.split(filename)[0],
                os.path.split(filename)[1])))


def check_settings():
    my_return = False
    obj = SettingsGUI()

    if not os.path.exists(AccdbDir):
        os.makedirs(AccdbDir)

    if not os.path.exists(ProcessedDir):
        os.makedirs(ProcessedDir)

    if not global_objs['Settings'].grab_item('Server') \
            or not global_objs['Settings'].grab_item('Database'):
        header_text = 'Welcome to Access DB Upload!\nSettings haven''t been established.\nPlease fill out all empty fields below:'
        obj.build_gui(header_text)
    else:
        try:
            if not obj.sql_connect():
                header_text = 'Welcome to Access DB Upload!\nNetwork settings are invalid.\nPlease fix the network settings below:'
                obj.build_gui(header_text)
            else:
                my_return = True
        finally:
            obj.sql_close()

    del obj
    return my_return


if __name__ == '__main__':
    if check_settings():

        has_updates = check_for_updates()

        if has_updates:
            global_objs['Event_Log'].write_log('Found {} files to process'.format(len(has_updates)))

            process_updates(has_updates)
        else:
            global_objs['Event_Log'].write_log('Found no files to process', 'warning')
    else:
        global_objs['Event_Log'].write_log('Settings Mode was established. Need to re-run script', 'warning')

    os.system('pause')

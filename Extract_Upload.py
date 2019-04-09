# If python 32-bit and MS Access engine is 64-bit, use this 32-bit MS Access driver
# http://www.microsoft.com/en-us/download/details.aspx?id=13255

from Global import grabobjs
from Global import SQLHandle

import os
import pathlib as pl
import datetime

CurrDir = os.path.dirname(os.path.abspath(__file__))
AccdbDir = os.path.join(CurrDir, "02_Process")
ProcessedDir = os.path.join(CurrDir, "03_Processed")
Global_Objs = grabobjs(CurrDir)


class AccdbHandle:
    accdb_cols = None
    sql_cols = None
    from_cols = None
    to_table = None
    to_cols = None

    def __init__(self, file):
        self.file = file
        self.asql = SQLHandle(Global_Objs['Settings'])
        self.asql.connect(conn_type='alch')

    @staticmethod
    def get_accdb_tables():
        myresults = Global_Objs['SQL'].get_accdb_tables()

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

    def get_accdb_cols(self, table):
        myresults = Global_Objs['SQL'].query('''
            SELECT TOP 1 *
            FROM [{0}]'''.format(table))

        if len(myresults) < 1:
            raise Exception('File {0} has no valid columns in table {1}'.format(
                os.path.basename(self.file), table))

        self.accdb_cols = myresults.columns.tolist()

    def get_sql_cols(self):
        myresults = self.asql.query('''
            SELECT
                Column_Name
            
            FROM INFORMATION_SCHEMA.COLUMNS
            
            WHERE
                TABLE_SCHEMA = '{0}'
                    AND
                TABLE_NAME = '{1}'
        '''.format(self.to_table.split('.')[0], self.to_table.split('.')[1]))

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
        configs = Global_Objs['Local_Settings'].grab_item('Accdb_Configs')

        if not configs:
            if self.add_config(table):
                Global_Objs['Local_Settings'].add_item('Accdb_Configs', [[table, self.from_cols, self.to_table,
                                                                         self.to_cols]])
                return True
            else:
                return False
        else:
            for config in configs:
                if table == config[0]:
                    self.from_cols = config[1]
                    self.to_table = config[2]
                    self.to_cols = config[3]
                    self.get_sql_cols()

                    if not self.validate_cols(self.from_cols, self.accdb_cols):
                        Global_Objs['Event_Log'].write_log(
                            'Stored settings for {0} has one or more columns that is not found in file {1} '.format(
                                table, os.path.basename(self.file)), 'error')
                        return False

                    if not self.validate_sql_table(self.to_table):
                        Global_Objs['Event_Log'].write_log(
                            'SQL TBL {0} has stored settings, but table does not exist anymore'.format(
                                self.to_table), 'error')
                        return False

                    if not self.validate_cols(self.to_cols, self.sql_cols):
                        Global_Objs['Event_Log'].write_log(
                            'Stored settings for SQL TBL {0} has one or more columns that does not exist'.format(
                                self.to_table), 'error')
                        return False

            if not self.from_cols and not self.to_table and not self.to_cols:
                if self.add_config(table):
                    configs.append([table, self.from_cols, self.to_table, self.to_cols])
                    Global_Objs['Local_Settings'].del_item('Accdb_Configs')
                    Global_Objs['Local_Settings'].add_item('Accdb_Configs', configs)
                    return True
                else:
                    return False
            else:
                return True

    def add_config(self, table):
        myanswer = None

        while not myanswer:
            print('There is no configuration for table ({0}) in file ({1}). Would you like to add configuration? (yes, no)'
                  .format(table, os.path.basename(self.file)))
            myanswer = input()

            if myanswer.lower() not in ['yes', 'no']:
                myanswer = None

        if myanswer.lower() == 'no':
            return False

        myanswer = None

        while not myanswer:
            print('From these accdb columns, what columns would you like to pull information from:')
            print('[{}]'.format(', '.join(self.accdb_cols)))
            myanswer = input()

            if myanswer:
                myanswer = myanswer.replace(', ', ',').split(',')

                if not self.validate_cols(myanswer, self.accdb_cols):
                    myanswer = None

        self.from_cols = myanswer
        myanswer = None

        while not myanswer:
            print('Please input the SQL server (schema).(table) that this information will be appending to:')
            myanswer = input()

            if not self.validate_sql_table(myanswer):
                myanswer = None

        self.to_table = myanswer
        self.get_sql_cols()
        myanswer = None

        while not myanswer:
            print('From these sql table columns, please choose the columns that corresponds to the Access Table Columns where information will be inserted:')
            print('SQL Table Cols: [{}]'.format(', '.join(self.sql_cols)))
            print('Access Table Cols: [{}]'.format(', '.join(self.from_cols)))
            myanswer = input()

            if myanswer:
                myanswer = myanswer.replace(', ', ',').split(',')

                if not self.validate_cols(myanswer, self.sql_cols):
                    myanswer = None

        self.to_cols = myanswer

        return True

    def process(self, table):
        Global_Objs['Event_Log'].write_log('Uploading data from table {0} to sql table {1}'
                                           .format(table, self.to_table))

        myresults = Global_Objs['SQL'].query('''
            SELECT
                [{0}]
            
            FROM [{1}]
        '''.format('], ['.join(self.from_cols), table))

        if not myresults.empty:
            myresults.columns = self.to_cols
            self.asql.upload(myresults, self.to_table, index=False, index_label=None)
            Global_Objs['Event_Log'].write_log('Data successfully uploaded from table {0} to sql table {1}'
                                               .format(table, self.to_table))
        else:
            Global_Objs['Event_Log'].write_log('Failed to grab data from access table {0}. No update made'
                                               .format(table), 'error')

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
        Global_Objs['Event_Log'].write_log('Processing file {0}'.format(os.path.basename(file)))
        myobj = AccdbHandle(file)

        Global_Objs['SQL'].connect('accdb', accdb_file=file)

        for table in myobj.get_accdb_tables():
            Global_Objs['Event_Log'].write_log('Validating table {0}'.format(table))
            if myobj.validate(table):
                myobj.process(table)
                filename = os.path.basename(file)
                os.rename(file, os.path.join(ProcessedDir,
                                             datetime.datetime.now().__format__("%Y%m%d") + os.path.split(filename)[0]
                                             + os.path.split(filename)[1]))

        myobj.close_asql()
        Global_Objs['SQL'].close()


if __name__ == '__main__':
    if not os.path.exists(AccdbDir):
        os.makedirs(AccdbDir)

    if not os.path.exists(ProcessedDir):
        os.makedirs(ProcessedDir)

    has_updates = check_for_updates()

    if has_updates:
        Global_Objs['Event_Log'].write_log('Found {} files to process'.format(len(has_updates)))

        process_updates(has_updates)
    else:
        Global_Objs['Event_Log'].write_log('Found no files to process', 'warning')

    os.system('pause')

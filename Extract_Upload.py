from Global import grabobjs
from Global import SQLHandle

import os
import pathlib as pl
import pandas as pd

CurrDir = os.path.dirname(os.path.abspath(__file__))
AccdbDir = os.path.join(CurrDir, "01_Process")
Global_Objs = grabobjs(CurrDir)


class AccdbHandle:
    accdb_cols = None
    sql_cols = None
    from_cols = None
    to_table = None
    to_cols = None

    def __init__(self, file):
        self.file = file
        self.asql = SQLHandle(Global_Objs['Settings']).connect(conn_type='alch')

    @staticmethod
    def get_accdb_tables():
        myresults = Global_Objs['SQL'].query('''
                SELECT MSysObjects.Name AS table_name
                FROM MSysObjects
                WHERE (((Left([Name],1))<>"~") 
                        AND ((Left([Name],4))<>"MSys") 
                        AND ((MSysObjects.Type) In (1,4,6)))''')

        if len(myresults) > 0:
            return myresults['table_name'].tolist()

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
                Global_Objs['Local_Settings'].add_item('Accdb_Configs', [table, self.from_cols, self.to_table,
                                                                         self.to_cols])
                return True
            else:
                return False
        else:
            for config in configs:
                if table == config[0]:
                    self.from_cols = config[1]
                    self.to_table = config[2]
                    self.to_cols = config[3]

            if not self.from_cols and not self.to_table and not self.to_cols:
                if self.add_config(table):
                    Global_Objs['Local_Settings'].add_item('Accdb_Configs', [table, self.from_cols, self.to_table,
                                                                             self.to_cols])
                    return True
                else:
                    return False
            else:
                return True

    def add_config(self, table):
        myanswer = None

        while not myanswer:
            print('There is no configuration for table ({0}) in file ({1}). Would you like to add configuration? (yes, no)'.format(table,
                                                                                                                         os.path.basename(self.file)))
            myanswer = input()

            if myanswer.lower() not in ['yes', 'no']:
                myanswer = None

        if myanswer.lower() == 'no':
            return False

        myanswer = None

        while not myanswer:
            print('From these accdb columns, what columns would you like to pull information from:')
            print('[{}]'.format(self.accdb_cols.join(', ')))
            myanswer = input()

            if self.validate_cols(myanswer, self.accdb_cols):
                myanswer = None

        self.from_cols = myanswer
        myanswer = None

        while not myanswer:
            print('Please input the SQL server (schema).(table) that this information will be appending to:')
            myanswer = input()

            if self.validate_sql_table(myanswer):
                myanswer = None

        self.to_table = myanswer
        self.get_sql_cols()
        myanswer = None

        while not myanswer:
            print('From these sql table columns, please choose the columns that corresponds to the Access Table Columns where information will be inserted:')
            print('SQL Table Cols: [{}]'.format(self.sql_cols.join(', ')))
            print('Access Table Cols: [{}]'.format(self.accdb_cols.join(', ')))
            myanswer = input()

            if self.validate_cols(myanswer, self.sql_cols):
                myanswer = None

        self.to_cols = myanswer

    def process(self):


def check_for_updates():
    f = list(pl.Path(AccdbDir).glob('*.accdb'))

    if f:
        return f

    f = list(pl.Path(AccdbDir).glob('*.mdb'))

    if f:
        return f


def process_updates(files):
    for file in files:
        myobj = AccdbHandle(file)

        Global_Objs['SQL'].connect(accdb_file=file)

        for table in myobj.get_accdb_tables():
            if myobj.validate(table):
                myobj.process()

        Global_Objs['SQL'].close()


if __name__ == '__main__':
    if not os.path.exists(AccdbDir):
        os.makedirs(AccdbDir)

    has_updates = check_for_updates()

    if has_updates:
        Global_Objs['Event_Log'].write_log('Found {} files to process'.format(len(has_updates)))

        process_updates(has_updates)
    else:
        Global_Objs['Event_Log'].write_log('Found no files to process', 'warning')

    os.system('pause')

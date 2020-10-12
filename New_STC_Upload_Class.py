from New_STC_Upload_Settings import tool, local_config, processed_dir, sql, sql_dir, email_engine, process_dir
from ftplib import FTP
from os.path import join, exists, basename, dirname, isfile, islink, isdir
from os import makedirs, unlink, rename, walk, remove, listdir
from shutil import rmtree
from zipfile import ZipFile
from rarfile import RarFile
from subprocess import Popen
from pandas import DataFrame
from pathlib import Path
from portalocker import Lock
from datetime import datetime

import traceback


class STCFTP(FTP):
    __entries = None
    __upload_dt = None
    __upload_dest_sub_dir = None
    __upload_dest_dir = None

    def __init__(self):
        try:
            FTP.__init__(self, local_config['FTP_Server'].decrypt())

            try:
                self.login(local_config['FTP_User'].decrypt(), local_config['FTP_Pass'].decrypt())
            except Exception as err:
                self.quit()
                tool.write_to_log(traceback.format_exc(), 'critical')
                tool.write_to_log("Credentials - Error Code '{0}', {1}".format(type(err).__name__, str(err)))
        except Exception as err:
            tool.write_to_log(traceback.format_exc(), 'critical')
            tool.write_to_log("Connect - Error Code '{0}', {1}".format(type(err).__name__, str(err)))

    def setup_ftp(self):
        tool.write_to_log('Navigating FTP & Setup Folders')
        self.cwd('/To Granite')
        self.__entries = list(self.mlsd())
        self.__entries.sort(key=lambda entry: "" if entry[0].startswith('SDN') or entry[1]['type'] == 'dir' or not (
                    entry[0].endswith('.zip') or entry[0].endswith('.rar') or entry[0].endswith('.7z')
                    or entry[0].endswith('.accdb') or entry[0].endswith('.mdb')) else entry[1]['modify'], reverse=True)
        self.__upload_dt = self.__entries[0][1]['modify'][0:8]
        self.__upload_dest_sub_dir = join(process_dir, self.__upload_dt)
        self.__upload_dest_dir = join(self.__upload_dest_sub_dir, 'Unzipped')

        if not exists(self.__upload_dest_sub_dir):
            makedirs(self.__upload_dest_sub_dir)

        if not exists(self.__upload_dest_dir):
            makedirs(self.__upload_dest_dir)

    def ftp_download(self):
        for item in self.__entries:
            if item[1]['modify'][0:8] == self.__upload_dt:
                download_path = join(self.__upload_dest_sub_dir, item[0])

                if not exists(download_path):
                    tool.write_to_log("Downloading '%s' from FTP" % item[0])

                    try:
                        self.retrbinary('RETR %s' % item[0], open(download_path, 'wb').write, 8 * 1024)
                        self.__unzip_file(download_path)
                    except Exception as err:
                        if exists(download_path):
                            unlink(download_path)

                        tool.write_to_log(traceback.format_exc(), 'critical')
                        tool.write_to_log("Download - Error Code '{0}', {1}".format(type(err).__name__, str(err)))
                        pass
            else:
                break

    def __unzip_file(self, zipped_fp):
        if zipped_fp.endswith('.zip'):
            tool.write_to_log("Unzipping '%s' file" % basename(zipped_fp))

            with ZipFile(zipped_fp, 'r') as zip_ref:
                zip_ref.extractall(self.__upload_dest_dir)
        elif zipped_fp.endswith('.rar'):
            tool.write_to_log("Unrarring '%s' file" % basename(zipped_fp))

            rar = RarFile(zipped_fp)
            rar.extractall(self.__upload_dest_dir)
        elif zipped_fp.endswith('.7z'):
            tool.write_to_log("Un7zipping '%s' file" % basename(zipped_fp))

            z = Popen('7z e "{0}" -o"{1}"'.format(zipped_fp, self.__upload_dest_dir), shell=True)
            z.wait()
            z.kill()
        else:
            tool.write_to_log("Migrating '%s' file" % basename(zipped_fp))
            rename(zipped_fp, join(self.__upload_dest_dir, basename(zipped_fp)))

    def __del__(self):
        self.quit()


class EmailClass(object):
    __email = None
    __batch = None
    __email_failures = dict()
    __email_success = dict()

    def __init__(self):
        from KGlobal.exchangelib import Exchange
        self.__email_engine = email_engine
        if local_config['Email_To']:
            self.__to_email = local_config['Email_To'].decrypt().replace(' ', '').replace(';', ',').split(',')
        else:
            self.__to_email = None

        if local_config['Email_Cc']:
            self.__cc_email = local_config['Email_Cc'].decrypt().replace(' ', '').replace(';', ',').split(',')
        else:
            self.__cc_email = None

        if not isinstance(self.__email_engine, Exchange):
            raise ValueError("%r is not an instance of Exchange" % self.__email_engine)

    @property
    def batch(self):
        return self.__batch

    @batch.setter
    def batch(self, batch):
        self.__batch = batch

    def record_results(self, table, error_msg=None, records=0, to_table=None):
        if table:
            if error_msg:
                if to_table:
                    tool.write_to_log("{0} => {1} Completed ({2})".format(table, to_table, error_msg), 'warning')
                else:
                    tool.write_to_log("{0} ({1})".format(table, error_msg), 'warning')

                self.__email_failures[table] = [to_table, error_msg]
            else:
                tool.write_to_log("{0} => {1} Completed ({2} records)".format(table, to_table, records))
                self.__email_success[table] = [to_table, records]

    def process_email(self, manual=False):
        if len(self.__email_success) > 0 or len(self.__email_failures) > 0:
            tool.write_to_log('Sending email to distros')
            from exchangelib import Message

            self.__email = Message(account=self.__email_engine)

            if self.__to_email:
                self.__email.to_recipients = self.__gen_email_list(self.__to_email)

            if self.__cc_email:
                self.__email.cc_recipients = self.__gen_email_list(self.__cc_email)

            if manual:
                self.__email.subject = 'Manual STC Records Upload %s' % self.batch
            else:
                self.__email.subject = 'STC Updated Records %s' % self.batch

            self.__email.body = self.__gen_body()
            self.__email.send()

    def __gen_body(self):
        body = list()
        body2 = list()
        body3 = list()
        names = [email.split('@')[0] for email in self.__to_email]

        for table, vals in self.__email_success.items():
            body2.append('\t\u2022  {0} => {1} ({2} records)'.format(table, vals[0], vals[1]))

        for table, vals in self.__email_failures.items():
            if vals[0]:
                body3.append('\t\u2022  {0} => {1} ({2})'.format(table, vals[0], vals[1]))
            else:
                body3.append('\t\u2022  {0} ({1})'.format(table, vals[1]))

        body.append('Happy {0} {1},'.format(datetime.today().strftime("%A"), '; '.join(names)))

        if body2:
            body.append('The following items have been successfully uploaded to SQL Server:')
            body.append('\n'.join(body2))

        if body3:
            body.append("The following items were unsuccesful:")
            body.append('\n'.join(body3))

        body.append("Yours Truly,")
        body.append("The BI Team")
        return '\n\n'.join(body)

    @staticmethod
    def __gen_email_list(emails):
        from exchangelib import Mailbox
        objs = list()

        for email in emails:
            objs.append(Mailbox(email_address=email))

        return objs


class SQLValidate(EmailClass):
    __sql_cols = None
    __sql_tables = None
    __profile = None
    __accdb_df = DataFrame()

    def __init__(self):
        EmailClass.__init__(self)
        self.sql_tables = sql.sql_tables()

    @property
    def profile(self):
        return self.__profile

    @profile.setter
    def profile(self, profile):
        if isinstance(profile, (type(None), dict)):
            self.__profile = profile
        else:
            raise ValueError("'profile' %r is not an instance of dict" % profile)

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

    @property
    def sql_cols(self):
        return self.__sql_cols

    @sql_cols.setter
    def sql_cols(self, sql_cols):
        from KGlobal.sql.cursor import SQLCursor

        if isinstance(sql_cols, (type(None), SQLCursor)):
            self.__sql_cols = [col.lower() for col in sql_cols.results[0]['Column_Name'].tolist()]
        else:
            raise ValueError("'sql_cols' %r is not an instance of SQLCursor" % sql_cols)

    @property
    def accdb_df(self):
        return self.__accdb_df

    @accdb_df.setter
    def accdb_df(self, accdb_df):
        if isinstance(accdb_df, (type(None), DataFrame)):
            self.__accdb_df = accdb_df
        else:
            raise ValueError("'accdb_df' %r is not an instance of DataFrame" % accdb_df)

    def val_sql_tbl(self, table):
        if table.lower() in self.sql_tables:
            return True

    def val_sql_col(self, column):
        if column.lower() in self.sql_cols:
            return True

    def sql_tbl_cols(self, table):
        query = '''
            SELECT
                Column_Name

            FROM INFORMATION_SCHEMA.COLUMNS

            WHERE
                TABLE_SCHEMA = '{0}'
                    AND
                TABLE_NAME = '{1}'
        '''.format(table.split('.')[0], table.split('.')[1])
        self.sql_cols = sql.sql_execute(query_str=query)

    def tbl_is_updated(self):
        if self.profile:
            query = '''
                SELECT COUNT(*) As Rows
                FROM {0}
                WHERE
                    Source_File = 'Updated Records {1}'
            '''.format(self.profile['SQL_TBL'], self.batch)
            results = sql.sql_execute(query_str=query)

            if results.results:
                return results.results[0].iloc[0, 0]
            elif results.errors:
                error_msg = "Error Code '{0}', {1}".format(results.errors[0], results.errors[1])
                self.record_results(self.profile['Acc_TBL'], error_msg, 0, self.profile['SQL_TBL'])
                return -1
        return 0

    def upload_df(self):
        if self.profile and not self.accdb_df.empty:
            tool.write_to_log("Uploading {0} records to {1}".format(len(self.accdb_df),
                                                                    self.profile['SQL_TBL']))
            table = self.profile['SQL_TBL'].split('.')

            if self.profile['SQL_TBL_Trunc'] == 1:
                query = "TRUNCATE TABLE %s" % self.profile['SQL_TBL']
                sql.sql_execute(query_str=query, execute=True)

            results = sql.sql_upload(dataframe=self.accdb_df, table_name=table[1], table_schema=table[0],
                                     if_exists='append', index=False)

            if results.errors:
                error_msg = "Error Code '{0}', {1}".format(results.errors[0], results.errors[1])
                self.record_results(self.profile['Acc_TBL'], error_msg, 0, self.profile['SQL_TBL'])
            else:
                self.record_results(self.profile['Acc_TBL'], None, len(self.accdb_df), self.profile['SQL_TBL'])
                self.__append_features()

    def __append_features(self):
        tool.write_to_log("Appending features to %s" % self.profile['SQL_TBL'])

        if self.val_sql_col('Edit_DT'):
            query = '''
                UPDATE {0}
                    SET
                        Edit_DT = GETDATE()
                WHERE
                    Source_File = 'Updated Records {1}'
            '''.format(self.profile['SQL_TBL'], self.batch)
            sql.sql_execute(query_str=query, execute=True)
        elif self.val_sql_col('Edit_Date'):
            query = '''
                UPDATE {0}
                    SET
                        Edit_Date = GETDATE()
                WHERE
                    Source_File = 'Updated Records {1}'
            '''.format(self.profile['SQL_TBL'], self.batch)
            sql.sql_execute(query_str=query, execute=True)

        path = join(sql_dir, self.profile['Acc_TBL'])

        if exists(path):
            for file in list(Path(path).glob('*.sql')):
                with Lock(str(file), 'r') as f:
                    query = f.read()

                sql.sql_execute(query_str=query, execute=True)


class AccValidate(SQLValidate):
    __acc_tables = None
    __acc_cols = None
    __acc_engine = None

    def __init__(self):
        SQLValidate.__init__(self)

    @property
    def acc_engine(self):
        return self.__acc_engine

    @acc_engine.setter
    def acc_engine(self, acc_engine):
        from KGlobal.sql import SQLEngineClass

        if isinstance(acc_engine, SQLEngineClass):
            self.__acc_engine = acc_engine
        else:
            raise ValueError("'acc_engine' %r is not an instance of SQLEngineClas" % acc_engine)

    @property
    def acc_tables(self):
        return self.__acc_tables

    @acc_tables.setter
    def acc_tables(self, acc_tables):
        from KGlobal.sql.cursor import SQLCursor

        if isinstance(acc_tables, SQLCursor):
            if acc_tables.errors:
                raise ValueError("Error Code '{0}', {1}".format(acc_tables.errors[0], acc_tables.errors[1]))
            else:
                df = acc_tables.results[0]
                df = df[(df['Table_Name'] != 'msys') & (df['Table_Type'] == 'TABLE')]
                self.__acc_tables = df['Table_Name'].tolist()
        elif isinstance(acc_tables, type(None)):
            self.__acc_tables = None
        else:
            raise ValueError("'acc_tables' %r is not an instance of SQLCursor" % acc_tables)

    @property
    def acc_cols(self):
        return self.__acc_cols

    @acc_cols.setter
    def acc_cols(self, acc_cols):
        if isinstance(acc_cols, (type(None), list)):
            self.__acc_cols = acc_cols
        else:
            raise ValueError("'acc_cols' %r is not an instance of List" % acc_cols)

    def db_connect(self, sql_config):
        from KGlobal.sql.engine import SQLEngineClass
        self.acc_engine = tool.config_sql_conn(sql_config=sql_config)

        if isinstance(self.__acc_engine, SQLEngineClass):
            self.acc_tables = self.__acc_engine.sql_tables()

    def retire_acc_engine(self):
        if self.acc_engine:
            self.acc_engine.close_connections(destroy_self=True)
            self.__acc_engine = None

    def val_acc_col(self, column):
        if column.lower() in [col.lower() for col in self.acc_cols]:
            return True

    def acc_tbl_cols(self, table):
        from KGlobal.sql.cursor import SQLCursor
        query = '''
            SELECT TOP 1 *
            FROM [{0}]
        '''.format(table)

        results = self.acc_engine.sql_execute(query_str=query)

        if results and isinstance(results, SQLCursor):
            if results.errors:
                raise Exception("Error '{0}', {1}".format(results.errors[0], results.errors[1]))
            else:
                self.acc_cols = results.results[0].columns.tolist()
        else:
            raise Exception("No valid columns were found for table %s" % table)

    def download_df(self):
        self.accdb_df = DataFrame()

        if self.acc_engine and self.profile:
            tool.write_to_log("Downloading '%s' to a dataframe" % self.profile['Acc_TBL'])
            query = '''
                SELECT
                    [{0}],
                    'Updated Records {2}' As Source_File

                FROM [{1}]
            '''.format('], ['.join(self.profile['Acc_Cols_Sel']), self.profile['Acc_TBL'], self.batch)

            results = self.acc_engine.sql_execute(query_str=query)

            if results.results:
                self.accdb_df = results.results[0]
                self.accdb_df.columns = self.profile['SQL_Cols_Sel'] + ('Source_File',)
            elif results.errors:
                error_msg = "Error Code '{0}', {1}".format(results.errors[0], results.errors[1])
                self.record_results(self.profile['Acc_TBL'], error_msg, 0, self.profile['SQL_TBL'])


class AccToSQL(AccValidate):
    def __init__(self):
        AccValidate.__init__(self)
        self.__accdb_profiles = local_config['Profiles']
        self.__accdb_processed = local_config['Processed']

        if not self.__accdb_processed:
            self.__accdb_processed = dict()

    def manual_process(self, table, accdb_file, batch):
        from KGlobal.sql import SQLConfig

        tool.write_to_log("Processing '%s' file" % basename(accdb_file))
        self.batch = batch
        sql_config = SQLConfig(conn_type='accdb', accdb_fp=str(accdb_file))
        self.db_connect(sql_config)

        if self.acc_engine:
            try:
                if self.__accdb_profiles and table in self.__accdb_profiles.keys():
                    self.profile = self.__accdb_profiles[table]
                else:
                    self.profile = None

                self.acc_tbl_cols(table)
                error = self.__validate(table)

                if error:
                    self.__package_err(table, error)
                else:
                    self.__upload_profile()
            finally:
                self.retire_acc_engine()
                self.process_email(manual=True)
                tool.write_to_log("File '%s' has finished being processed" % basename(accdb_file))
                tool.gui_console(turn_off=True)

    def process_file(self, accdb_file):
        from KGlobal.sql import SQLConfig

        tool.write_to_log("Processing '%s' file" % basename(accdb_file))
        self.batch = basename(dirname(dirname(accdb_file)))
        sql_config = SQLConfig(conn_type='accdb', accdb_fp=str(accdb_file))
        self.db_connect(sql_config)

        if self.acc_engine:
            try:
                for table in self.acc_tables:
                    if self.__accdb_profiles and table in self.__accdb_profiles.keys():
                        self.profile = self.__accdb_profiles[table]
                    else:
                        self.profile = None

                    self.__processed_list(accdb_file, table)
                    self.acc_tbl_cols(table)
                    error = self.__validate(table)

                    if error:
                        self.__package_err(table, error)
                    else:
                        self.__upload_profile()
            finally:
                self.retire_acc_engine()

    def migrate_file(self, source_fp):
        dest_dir = join(processed_dir, self.batch)
        dest_fp = join(dest_dir, basename(source_fp))

        if not exists(dest_dir):
            makedirs(dest_dir)

        if exists(dest_fp):
            remove(source_fp)
        else:
            rename(source_fp, dest_fp)

    def save_processed(self):
        local_config['Processed'] = self.__accdb_processed
        local_config.sync()

    def __processed_list(self, accdb_file, table):
        dest_dir = join(processed_dir, self.batch)
        dest_fp = join(dest_dir, basename(accdb_file))

        if table not in self.__accdb_processed.keys():
            self.__accdb_processed[table] = list()

        if [self.batch, dest_fp] not in self.__accdb_processed[table]:
            self.__accdb_processed[table].append([self.batch, dest_fp])

    def __validate(self, table):
        tool.write_to_log("Validating access table '%s'" % basename(table))
        if not self.profile:
            return [1, 'Acc table %s profile was never established' % table, 'Please establish new profile']

        if not self.__validate_cols(self.profile['Acc_Cols'], False):
            return [2, 'Acc table %s column validation failed' % table, 'Please re-do assignment']

        if len(self.profile['Acc_Cols']) != len(self.acc_cols):
            return [3, 'Acc table %s has new columns unassigned' % table, 'Please assign']

        if not self.val_sql_tbl(self.profile['SQL_TBL']):
            return [4, 'SQL table %s is not a valid table anymore' % self.profile['SQL_TBL'],
                    'Please choose new sql table']

        self.sql_tbl_cols(self.profile['SQL_TBL'])

        if not self.__validate_cols(self.profile['SQL_Cols'], True):
            return [5, 'SQL table %s column validation failed' % self.profile['SQL_TBL'], 'Please re-do assignment']

        # if len(self.profile['SQL_Cols']) != len(self.sql_cols):
        #    return [6, 'SQL table %s has new columns unassigned' % self.profile['SQL_TBL'], 'Please assign']

    def __upload_profile(self):
        rows_updated = self.tbl_is_updated()

        if rows_updated > 0:
            self.record_results(self.profile['Acc_TBL'], None, rows_updated, self.profile['SQL_TBL'])
        elif rows_updated == 0:
            self.download_df()
            self.upload_df()

    def __package_err(self, table, error):
        err_profile = dict()
        self.record_results(table, "Error Code '{0}', {1}".format(error[0], error[1]), 0, None)
        err_profile['Header'] = '{0}. {1}'.format(error[1], error[2])
        err_profile['Acc_TBL'] = table
        err_profile['Acc_Cols'] = self.acc_cols

        if error[0] in [2, 3, 5, 6]:
            err_profile['SQL_TBL'] = self.profile['SQL_TBL']
            err_profile['SQL_Cols'] = self.profile['SQL_Cols']

        if error[0] in [3, 4, 5, 6]:
            err_profile['Acc_Cols_Sel'] = self.profile['Acc_Cols_Sel']

        if error[0] in [3, 6]:
            err_profile['SQL_Cols_Sel'] = self.profile['SQL_Cols_Sel']

        err_profiles = local_config['Err_Profiles']

        if not err_profiles:
            err_profiles = dict()

        err_profiles[table] = err_profile
        local_config['Err_Profiles'] = err_profiles
        local_config.sync()

    def __validate_cols(self, col_list, is_sql=True):
        for col in col_list:
            if (is_sql and not self.val_sql_col(col)) or (not is_sql and not self.val_acc_col(col)):
                return False

        return True


def grab_accdb():
    file_paths = []

    for root, dirs, files in walk(process_dir):
        if "Unzipped" in dirs and len(dirs) == 1:
            dir_path = join(root, dirs[0])
            file_paths += list(Path(dir_path).glob('*.accdb')) + list(Path(dir_path).glob('*.mdb'))

    return file_paths


def rem_proc_elements():
    tool.write_to_log('Cleaning Process directory')

    for filename in listdir(process_dir):
        file_path = join(process_dir, filename)
        try:
            if isfile(file_path) or islink(file_path):
                unlink(file_path)
            elif isdir(file_path):
                rmtree(file_path)
        except Exception as e:
            tool.write_to_log('Failed to delete %s. Reason: %s' % (file_path, e), 'warning')
            pass


def check_processed():
    if not local_config['Processed']:
        acc_paths = []

        for filename in listdir(processed_dir):
            file_path = join(processed_dir, filename)

            if isdir(file_path):
                acc_paths += list(Path(file_path).glob('*.accdb')) + list(Path(file_path).glob('*.mdb'))

        if len(acc_paths) > 0:
            from KGlobal.sql.config import SQLConfig
            from KGlobal.sql.engine import SQLEngineClass
            from KGlobal.sql.cursor import SQLCursor
            processed = dict()

            for accdb_file in acc_paths:
                sql_config = SQLConfig(conn_type='accdb', accdb_fp=str(accdb_file))
                acc_engine = tool.config_sql_conn(sql_config=sql_config)

                try:
                    if isinstance(acc_engine, SQLEngineClass):
                        acc_tables = acc_engine.sql_tables()

                        if isinstance(acc_tables, SQLCursor):
                            if acc_tables.results:
                                df = acc_tables.results[0]
                                df = df[(df['Table_Name'] != 'msys') & (df['Table_Type'] == 'TABLE')]

                                for table in df['Table_Name'].tolist():
                                    if table not in processed.keys():
                                        processed[table] = list()

                                    processed[table].append([basename(dirname(accdb_file)), accdb_file])
                finally:
                    acc_engine.close_connections(destroy_self=True)

            if processed:
                local_config['Processed'] = processed
                local_config.sync()

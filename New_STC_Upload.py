from New_STC_Upload_Settings import check_settings, tool
from New_STC_Upload_Class import STCFTP, AccToSQL, grab_accdb, rem_proc_elements, check_processed

import traceback


def stc_ftp():
    try:
        ftp_obj = STCFTP()
        ftp_obj.setup_ftp()
        ftp_obj.ftp_download()
    finally:
        del ftp_obj


def proc_accdbs():
    files = grab_accdb()

    if files:
        obj = AccToSQL()

        for file in files:
            obj.process_file(file)
            obj.migrate_file(file)

        obj.save_processed()
        obj.process_email()
        del obj


if __name__ == '__main__':
    try:
        check_processed()

        if check_settings():
            rem_proc_elements()
            stc_ftp()
            proc_accdbs()
            rem_proc_elements()
        else:
            tool.write_to_log("Settings - Error Code 'Check_Settings', Settings was not establish thus program exit",
                              action="error")
    except Exception as e:
        tool.write_to_log(traceback.format_exc(), 'critical')
        tool.write_to_log("Main Loop - Error Code '{0}', {1}".format(type(e).__name__, str(e)))

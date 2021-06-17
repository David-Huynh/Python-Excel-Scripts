"""Applying Excel Macro to Multiple XLSM Files

This script operates on all permanent xlsm files in the given DIRECTORY, 
and applies the macro identified by MACRO_WORKBOOK_NAME, MODULE_NAME, and MACRO_NAME

MACRO_WORKBOOK must always be placed in the same folder as the script
DIRECTORY must exist for the script to execute properly and is ALWAYS relative to directory of
the script

This script requires `pywin32` to be installed in the Python environment you are running the
script in.

Author: David Huynh
"""
import os
import win32com.client

MACRO_WORKBOOK_NAME="macro_workbook.xlsm"
MODULE_NAME="Module1"
MACRO_NAME="ProcessFiles"

def excel_macro_repeated(macro_file, module_name, macro_name):
    """
    Activates a macro in the macro_file workbook,
    that cycles through xls files in a
    Files directory relative to macro_file
    Aimed to be used as an efficient way to update bloomberg functions
    daily
    Adapted from https://stackoverflow.com/questions/14766238/run-same-excel-macro-on-multiple-excel-files
    in order to run bloomberg refresh
    
    @type macro_file: str
    @param macro_file: the full filename of the workbook that contains the macro
    @type module_name: str
    @param module_name: the name of the module for the macro
    @type macro_name: str
    @param macro_name: the name of the macro
    @return: True if process finished successfully and false otherwise
    """
    ##Starts excel window to operate on
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        macro_workbook = excel.Workbooks.Open(os.path.abspath("./"+macro_file))
        excel.Application.Run("{}!{}.{}".format(macro_file, module_name, macro_name))
    except:
        print("ERROR: No file" + os.path.abspath("./"+macro_file) +"found, OR macro_workbook/macro is invalid")
    finally:
        macro_workbook.Close()
        excel.Application.Quit()
        del excel
    print("Task completed successfully")
    return True


excel_macro_repeated(MACRO_WORKBOOK_NAME, MODULE_NAME, MACRO_NAME)

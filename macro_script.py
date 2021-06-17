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

DIRECTORY="./"
MACRO_WORKBOOK_NAME="script_final_cpi.xlsm"
MODULE_NAME="Module1"
MACRO_NAME="BloombergRefresh"

def excel_macro_repeated(directory, macro_file, module_name, macro_name):
    """
    Operates on all permanent xlsm files in the given directory, 
    using the macro_file workbook and the module/macro_name given

    @type directory: str
    @param directory: the relative directory of the spreadsheets to be operated on
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
    
    macro_workbook = excel.Workbooks.Add(os.path.abspath("./"+macro_file))
    try:
        #Identifies files to be operated on
        for file in os.listdir(directory):
            ##Ignores temporary files created automatically that start with ~ 
            if file.endswith(".xlsm") and not file.startswith("~") and not file==macro_file:
                workbook = excel.Workbooks.Add(os.path.abspath(directory+"/"+file))
                try:
                    ##Runs the macro given by macro_name from macro_file, on the 'excel' application 
                    excel.Application.Run("{}!{}.{}".format(macro_file, module_name, macro_name))
                except:
                    print("Invalid macro workbook or macro")
                    return False
                workbook.Save()
                workbook.Close()
        macro_workbook.Close()
        excel.Application.Quit()
    except:
        print("Invalid directory")
        return False
    return True


excel_macro_repeated(DIRECTORY, MACRO_WORKBOOK_NAME, MODULE_NAME, MACRO_NAME)
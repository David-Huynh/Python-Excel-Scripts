"""Converts XLSX to XLSM (macro enabled excel file)

This script operates on all permanent xlsx files in the given WORKING_DIRECTORY, 
and applies the macro identified by MACRO_WORKBOOK_NAME, MODULE_NAME, and MACRO_NAME

WORKING/DESTINATION_DIRECTORY is always relative to the directory of the script
and both folders must exist in order for the script to execute properly

This script requires `pywin32` to be installed in the Python environment you are running the
script in.

Author: David Huynh
"""
import os
import win32com.client

WORKING_DIRECTORY="./"
DESTINATION_DIRECTORY="./output"

def xlsx_to_xlsm(working_directory, destination_directory):
    """
    Operates on all permanent xlsx files in the given working_directory, 
    converts them all into xlsm files in the destination_directory
    WARNING:Both directories must exist

    @type working_directory: str
    @param working_directory: the relative directory of the xlsx files to be converted
    @type destination_directory: str
    @param destination_directory: the relative output directory for the xlsm files
    @return: True if process finished successfully and false otherwise
    """
    ##Starts excel window to operate on
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    #Identifies files to be converted
    for file in os.listdir(working_directory):
        ##Ignores temporary files created automatically that start with ~ 
        if file.endswith(".xlsx") and not file.startswith("~"):
            workbook = excel.Workbooks.Open(os.path.abspath(working_directory+"/"+file))
            try:
                #Adapted from https://stackoverflow.com/questions/39292179/how-to-convert-xlsm-macro-enabled-excel-file-to-xlsx-using-python for opposite case and for multiple files
                ##converts every xlsx file in working_directory to xlsm files in destination_directory
                excel.DisplayAlerts = False
                workbook.DoNotPromptForConvert = True
                workbook.CheckCompatibility = False
                workbook.SaveAs(os.path.abspath(destination_directory) +"/"+ file.split(".")[0] + ".xlsm", FileFormat=52, ConflictResolution=2)
            except:
                print("Invalid destination directory")
                return False
    excel.Application.Quit()
    return True

xlsx_to_xlsm(WORKING_DIRECTORY, DESTINATION_DIRECTORY)
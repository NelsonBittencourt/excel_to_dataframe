# ****************************************************************
# excel_to_pandas.py
# Python script to test Excel to Pandas Dataframe DLL
# Created by Nelson Rossi Bittencourt.
# Github: https://github.com/nelsonbittencourt/excel_to_dataframe
# Last change: 04/11/2024 (dd/mm/yyyy).
# Version: 0.2.59
# License: MIT
# ****************************************************************


# *********** Imports **********
import ctypes as ct         # For dll access.
import pandas as pd         # Pandas (needs openpyxl to open Excel files).
from io import StringIO     # To convert csv to binary.
import os				    # To get path of this file and dll.
import platform             # To check for Windows or Linux.


# *********** Setup **********
# Sets up full path to 'excel_to_df.dll' or 'excel_to_df.so'

is_windows = True
dll_file = 'excel_to_df.dll'

if (platform.uname()[0]!="Windows"): 
    is_windows = False
    dll_file = 'excel_to_df.so'

abs_path = os.path.abspath(os.path.dirname(__file__))
dll_path = os.path.join(abs_path,dll_file)

# *********** Initializations ********** 

# Loads dll. Tries to load installed version. Otherwise, tries current dir.
try:
    if (is_windows):
        wsdf_dll = ct.CDLL(dll_path)        
    else:
         wsdf_dll = ct.cdll.LoadLibrary(dll_path)         
except OSError as e:
    print('Error loading dll file! Error:' , e)
    exit(-1)
    

# 'Instantianting' dll functions

# Opens Excel function
dll_open_excel = wsdf_dll.openExcel
dll_open_excel.argtypes = [ct.c_char_p]
dll_open_excel.restype = ct.c_int

# Gets sheet data function
dll_get_sheet_mt = wsdf_dll.getSheetData
dll_get_sheet_mt.argtypes = [ct.c_char_p]
dll_get_sheet_mt.restype = ct.c_char_p

# Closes Excel function
dll_close_excel = wsdf_dll.closeExcel
dll_close_excel.restype = ct.c_int

# Gets dll version
dll_version = wsdf_dll.version
dll_version.restype = ct.c_char_p


# *********** Functions ********** 

def version():
    """
    Gets dll version as string.

    Arguments:
    None

    Returns:
    string
    
    Requires:
    excel_to_df.dll

    Version 0.2.52

    """             
    tmp = dll_version()       
    return tmp.decode('utf-8')


def open_excel(excel_file_name):
    """
    Opens an Excel file and loads shared strings, styles (for dates only) and worksheet names.

    Arguments:
    excel_file_name - string with full path to Excel file.

    Returns:
     0  - success;
    -1  - file not found or
    -2  - file found, invalid Excel.
    
    Requires:
    excel_to_df.dll or excel_to_df.so

    Version 0.2.52

    """        
    
    return dll_open_excel(excel_file_name.encode('utf-8'))    

def close_excel():
    """
    Closes an Excel file.

    Arguments:
    None.

    Returns:
    None.

    Requires:
    excel_to_df.dll or excel_to_df.so    

    Version 0.1

    """        
    
    return dll_close_excel()


def ws_to_df(sheet_name):
    """
    Loads an Excel worksheet and converts to pandas dataframe.

    Arguments:
    sheet_name  -   (string) An existing worksheet name.                
    
    Returns:
    Pandas dataframe.

    Requires:
    Pandas, ctypes, io.StringIO, excel_to_df.dll

    Version 0.2.51

    """
    # Gets worksheet data
    
    tmp = dll_get_sheet_mt(sheet_name.encode('utf-8'))
    
    # If data exists, converts to Pandas dataframe
    if  (tmp!=None and len(tmp)>0):
        return pd.read_csv(StringIO(tmp.decode('utf-8')),lineterminator='\n',header=None,sep=';',low_memory=False)
    else:    
        print("Error trying to convert sheet to Pandas Dataframe")
        return -1


def split_df(df, split_string, col_search, header_offset=0):
    """
    Splits a dataframe to 'x' dataframes considering 'split_string' as table separator.

    Arguments:
    df              -   Pandas dataframe to split;
    split_string    -   String to define boundaries of tables;
    col_search      -   (Integer) Column to search for 'split_string' (zero-based) and
    header_offset   -   (Integer)(Optional) Offset from 'split_string' row and header row.

    Returns:
    list of pandas dataframes.
    
    Requires:
    pandas.

    Version 0.2.51

    """

    # Gets a list of rows that contains 'split_string'
    sp_rows = df[df[df.columns[col_search]].str.contains(split_string, regex=True)==True].index.to_list()
    sp_rows.append(df.shape[0])

    # Creates a empty list to splitted dataframes
    ld = []

    # Loops for eath row in splitter_rows
    for row in range(0,len(sp_rows)-1):
        
        # Creates a sub dataframe, reset index and change columns
        df2 = df.iloc[sp_rows[row]+header_offset:sp_rows[row+1],:]
        df2 = df2.reset_index(drop=True)
        df2.columns = df.iloc[[sp_rows[row]+header_offset], :].values.tolist()
        
        # Adds sub dataframe to list of dataframes and destroy df2
        ld.append(df2)        
        del df2

    return ld


# *********** Entry point **********
if __name__ == "__main__":
    pass 
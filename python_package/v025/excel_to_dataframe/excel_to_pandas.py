# ***************************************************
# Python script to test Excel to Pandas Dataframe DLL
# Created by Nelson Rossi Goulart Bittencourt.
# Github: https://github.com/nelsonbittencourt
# 14/12/2022 (dd/mm/yyyy).
# ***************************************************

# TODO: 1) compare dataframes; 2) rename dll functions and 3) change separator in dll string

# *********** Imports **********

import ctypes as ct         # For access dll
import pandas as pd         # Pandas
from io import StringIO     # To convert csv to binary
import timeit    			# Benchmarks
import pathlib				# To get path of this file


# *********** Setup **********

# Sets up full path to excel_to_df.dll
dll_path = "{}\\{}".format(pathlib.Path(__file__).parent.resolve() ,'excel_to_df.dll')



# *********** Initializations ********** 

# Loads dll
wsdf_dll = ct.CDLL(dll_path,winmode=0x8)

# 'Instantianting' dll functions

# Opens Excel function
dll_open_excel = wsdf_dll.openExcel
dll_open_excel.argtypes = [ct.c_char_p]
dll_open_excel.restype = ct.c_int

# Gets sheet data function (single thread)
dll_get_sheet = wsdf_dll.getSheet
dll_get_sheet.argtypes = [ct.c_char_p]
dll_get_sheet.restype = ct.c_char_p

# Gets sheet data function (multi-thread)
dll_get_sheet_mt = wsdf_dll.getSheetMT
dll_get_sheet_mt.argtypes = [ct.c_char_p]
dll_get_sheet_mt.restype = ct.c_char_p

# Closes Excel function
dll_close_excel = wsdf_dll.closeExcel
dll_close_excel.restype = ct.c_int

# Gets dll version
dll_version = wsdf_dll.version
dll_version.restype = ct.c_char_p



# *********** Pandas Functions ********** 

def get_dll_version():
    """
    Gets dll version.

    Arguments:
    None

    Returns:
    string
    
    Requires:
    excel_to_df.dll

    """             
    tmp = dll_version()       
    return tmp.decode('utf-8')



def open_excel(excel_file_name):
    """
    Opens an Excel file and loads shared strings, styles (for dates only) and sheets names.

    Arguments:
    excel_file_name - string with full path to Excel file.

    Returns:
    None    - success;
    -1      - file not found or
    -2      - file found, invalid Excel.
    
    Requires:
    excel_to_df.dll

    """        
    return dll_open_excel(excel_file_name.encode())    



def ws_to_df(sheet_name, multi_thread=0):
    """
    Loads an Excel worksheet and converts to pandas dataframe.

    Arguments:
    sheet_name  -   (string) An existing worksheet name.                
    
    Returns:
    Pandas dataframe.

    Requires:
    Pandas, ctypes, io.StringIO, excel_to_df.dll

    """
    # Gets worksheet data
    
    if (multi_thread==0):
        tmp = dll_get_sheet(sheet_name.encode())
    else:
        tmp = dll_get_sheet_mt(sheet_name.encode())
    
    # If data exists, converts to Pandas dataframe
    if (tmp!=None):
        return pd.read_csv(StringIO(tmp.decode('utf-8')),lineterminator='\n',header=None,sep=';',low_memory=False)
    else:    
        return -1



def split_df(df, split_string, col_search, header_offset=0):
    """
    Splits a dataframe to x dataframes considering 'split_string' as table separator.

    Arguments:
    df              -   Pandas dataframe to split;
    split_string    -   String to define boundaries of tables;
    col_search      -   (Integer) Column to search for 'split_string' (zero-based) and
    header_offset   -   (Integer)(Optional) Offset from 'split_string' row and header row.

    Returns:
    list of pandas dataframes.
    
    Requires:
    pandas.

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



def test_dll(excel_file_full_path, excel_worksheet_name, multi_thread=0):
    """
    Using dll, opens a Excel file and converts a worksheet to dataframe
    This rotine is used by benchmark function.
    
    Arguments:
    excel_file_full_path    - (string) full path for Excel file and
    excel_worksheet_name    - (string) worksheet name for data extration.

    Returns:
    None

    Requires:
    ctypes, excel_to_csv.dll

    """
    ret = open_excel(excel_file_full_path)
    
    # If Excel file is loaded, tries to load worksheet.
    if (ret==0):        
        tmp_df = ws_to_df(excel_worksheet_name,multi_thread)
    else:
        print('Error on load Excel! Error number:{:d}'.format(ret))
        
    dll_close_excel()    



def test_pure_pandas(excel_file_full_path,excel_worksheet_name):
    """
    Using Pandas, opens a Excel file and converts a worksheet to dataframe.
    This rotine is used by benchmark function.

    Arguments:
    excel_file_full_path    - (string) full path for Excel file and
    excel_worksheet_name    - (string) worksheet name for data extration.

    Returns:
    None.

    Requires:
    Pandas.

    """

    # Creates an Excel object.
    excel = pd.ExcelFile(excel_file_full_path)

    # Gets dataframe from specified worksheet.
    tmp_df = pd.read_excel(excel, sheet_name=excel_worksheet_name)    
    
    # Garbage collection.
    del excel



def benchmarks(save_csvs=False):
    """
    Runs benchmarks with two large files from CCEE (Brazilian Chamber for the Commercialisation of Electrical Energy).
    
    Arguments:
    None.

    Returns:
    Strings to console.

    Requires:
    Pandas, excel_to_df.dll, test_dll, test_pure_pandas.

    """

    # Excel files and worksheets for benchmarking
    excel_files = ['benchmarks/infomercado_individuais_2022.xlsx','benchmarks/infomercado_contratos.xlsx']
    sheet_names  = ['003 Consumo','Contratos Distribuidoras']

    for a in range(0,2):
        excel_file = excel_files[a]
        sheet_name = sheet_names[a]

        print('**************************************************************************************************')
        print('Benchmark test {} - Excel File \'{}\', Sheet \'{}\''.format(a+1,excel_file,sheet_name))
        print('')        
        
        # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        tmp_test = test_dll

        # Single thread
        dll_time_st = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1,globals=locals())
        
        # Multi-thread
        dll_time_mt = timeit.timeit("tmp_test(excel_file,sheet_name,1)", number=1,globals=locals())
        
        # # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        # tmp_test = test_pure_pandas
        # pd_time = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1, globals=locals())
        pd_time = 400
        
        # Calculates the ratio pandas time / dll time.
        ratio_st = pd_time/dll_time_st
        ratio_mt = pd_time/dll_time_mt
        
        # Prints resutls.
        print('')
        print('**************************************************************************************************')
        print('Results:')
        print('')
        print('A) Single Thread')
        print('Pandas..................... {:.2f} seconds.'.format(pd_time))
        print('DLL (single thread)........ {:.2f} seconds.'.format(dll_time_st))
        print('--> DLL is................. {:.2f} times faster than Pandas.'.format(ratio_st))
        print('')
        print('B) Multi Thread')
        print('Pandas..................... {:.2f} seconds.'.format(pd_time))
        print('DLL (multi-thread)......... {:.2f} seconds.'.format(dll_time_mt))
        print('--> DLL is................. {:.2f} times faster than Pandas.'.format(ratio_mt))
        print('')
        print('C) Single vrs Multi Thread')
        print('DLL (single thread)........ {:.2f} seconds.'.format(dll_time_st))
        print('DLL (multi-thread)......... {:.2f} seconds.'.format(dll_time_mt))
        print('--> Multi/Single........... {:.2f}'.format(dll_time_mt/dll_time_st))



def get_csvs(excel_file_full_path,excel_worksheet_name):
    """

    Opens an Excel file name, loads a worksheet data and saves to csv.

    Arguments:
    excel_file_full_path    - (string) Excel file full path and
    excel_worksheet_name    - (string) Excel worksheet name.

    Returns:
    None.

    Requires:
    Pandas, excel_to_df.dll.
    
    """
    
    # Opens Excel file. 
    ret = open_excel(excel_file_full_path)
    
    # Converts worksheet data to Pandas dataframe (dll single thread).
    tmp_df = ws_to_df(excel_worksheet_name,multi_thread=0)
    tmp_df.to_csv(excel_worksheet_name + '_st.csv',sep=';',decimal=',')
    dll_close_excel()    
    del tmp_df

    # Converts worksheet data to Pandas dataframe (dll multi thread).
    ret = open_excel(excel_file_full_path)
    tmp_df = ws_to_df(excel_worksheet_name,multi_thread=1)
    tmp_df.to_csv(excel_worksheet_name + '_mt.csv',sep=';',decimal=',')
    dll_close_excel()    
    del tmp_df
    
    # Converts worksheet data to Pandas dataframe (Pandas)
    excel = pd.ExcelFile(excel_file_full_path)    
    tmp_df = pd.read_excel(excel, sheet_name=excel_worksheet_name)   
    tmp_df.to_csv(excel_worksheet_name + '_pd.csv',sep=';',decimal=',')
    del tmp_df



# *********** Entry point **********
if __name__ == "__main__":
    pass 
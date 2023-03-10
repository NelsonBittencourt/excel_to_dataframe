# ****************************************************************
# example.py
# Python script showing dll usage and benchmarks.
# Created by Nelson Rossi Goulart Bittencourt.
# Github: https://github.com/nelsonbittencourt/excel_to_dataframe
# Last change: 18/02/2023 (dd/mm/yyyy).
# Version: 0.2.56
# License: MIT
# ****************************************************************

# *********** Functions list **********
# main              - Entry-point with simple dll usage example;
# benchmarks        - Benchmark routine (need 'test_dll', 'test_pure_pandas' and file into benchmarks folder);
# test_dll          - Auxiliary routine to convert worksheets to Pandas dataframe using dll;
# test_pure_pandas  - Auxiliary routine to convert worksheets to Pandas dataframe using Pandas;
# get_csvs          - Converts a Excel to csv files using dll.


# *********** Imports **********
import pandas as pd      
import timeit  
import os     					

# Tries to import from installed packet (PyPI or Anaconda/Miniconda)
# or from local file (direct download files from Github).
try:
   import excel_to_dataframe.excel_to_pandas as etd    
except ImportError as e:
   import excel_to_pandas as etd	


# *********** Functions **********

def main():
    """
    Simple dll Usage example. Use breakpoints and debuger mode to view the results.
    
    Arguments:
    None.

    Returns:
    None.

    Requires:
    excel_to_df.dll.
    
    Version 0.2.5
    """

    # Excel file and worksheet name to work with.
    excel_file = os.path.join(absolute_path,'benchmarks/infomercado_individuais_2022.xlsx')
    sheet_name = '003 Consumo'
    
    # Open Excel file.
    etd.open_excel(excel_file)
    
    # Gets a dataframe from worksheet.
    df = etd.ws_to_df(sheet_name,mult_thread=0)
        
    # Closes Excel.
    etd.dll_close_excel()

    # Splits dataframe based on string search.
    # ldf - list of dataframes.
    ldf = etd.split_df(df,'Tabela ',col_search=1,header_offset=1)

    print('Done.')


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

    Version 0.2.51

    """
    ret = etd.open_excel(excel_file_full_path)
    
    # If Excel file is loaded, tries to load worksheet.
    if (ret==0):        
        tmp_df = etd.ws_to_df(excel_worksheet_name,multi_thread)
    else:
        print('Error on load Excel! Error number:{:d}'.format(ret))
        
    etd.dll_close_excel()    


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

    Version 0.2.4

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
    save_csvs: (bool) if true, saves csvs files to disk.

    Returns:
    Strings to console.

    Requires:
    Pandas, excel_to_df.dll, test_dll, test_pure_pandas.

    Version 0.2.51

    """

    # Excel files and worksheets for benchmarking
    excel_files = ['benchmarks/infomercado_individuais_2022.xlsx','benchmarks/infomercado_contratos.xlsx']
    sheet_names  = ['003 Consumo','Contratos Distribuidoras']

    for a in range(0,2):
        excel_file = os.path.join(absolute_path, excel_files[a])
        sheet_name = sheet_names[a]

        print('**************************************************************************************************')
        print('Benchmark test {} - Excel File \'{}\', Sheet \'{}\''.format(a+1,excel_files[a],sheet_name))
        print('')        
        
        # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        tmp_test = test_dll

        # Single thread
        dll_time_st = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1,globals=locals())
        
        # Multi-thread
        dll_time_mt = timeit.timeit("tmp_test(excel_file,sheet_name,1)", number=1,globals=locals())
        
        # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        tmp_test = test_pure_pandas
        pd_time = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1, globals=locals())
                
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

    Version 0.2.51
    
    """
    
    # Opens Excel file. 
    ret = etd.open_excel(excel_file_full_path)
    
    # Converts worksheet data to Pandas dataframe (dll single thread).
    tmp_df = etd.ws_to_df(excel_worksheet_name,multi_thread=0)
    tmp_df.to_csv(excel_worksheet_name + '_st.csv',sep=';',decimal=',')
    etd.dll_close_excel()    
    del tmp_df

    # Converts worksheet data to Pandas dataframe (dll multi thread).
    ret = etd.open_excel(excel_file_full_path)
    tmp_df = etd.ws_to_df(excel_worksheet_name,multi_thread=1)
    tmp_df.to_csv(excel_worksheet_name + '_mt.csv',sep=';',decimal=',')
    etd.dll_close_excel()    
    del tmp_df
    
    # Converts worksheet data to Pandas dataframe (Pandas)
    excel = pd.ExcelFile(excel_file_full_path)    
    tmp_df = pd.read_excel(excel, sheet_name=excel_worksheet_name)   
    tmp_df.to_csv(excel_worksheet_name + '_pd.csv',sep=';',decimal=',')
    del tmp_df


# *********** Entry point **********
if __name__ == "__main__":
    
    # Gets and prints dll version.
    print (etd.get_dll_version())
    #exit(0)
        
    # Runs benchmarks
    benchmarks()

    #  Saves csvs to comparisons;
    print('')
    print('**************************************************************************************************')
    print('Converting worksheets to csvs:')
    get_csvs(os.path.join(absolute_path,'benchmarks/infomercado_individuais_2022.xlsx'), '003 Consumo')    
    get_csvs(os.path.join(absolute_path, 'benchmarks/infomercado_contratos.xlsx'), 'Contratos Distribuidoras')
    
    print('Done!')
    
    # # Another dll usage example.
    # main()

    
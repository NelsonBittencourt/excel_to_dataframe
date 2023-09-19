# ****************************************************************
# example.py
# Python script showing dll usage and benchmarks.
# Created by Nelson Rossi Goulart Bittencourt.
# Github: https://github.com/nelsonbittencourt/excel_to_dataframe
# Last change: 19/09/2023 (dd/mm/yyyy).
# Version: 0.2.58
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

#Tries to import from installed packet (PyPI or Anaconda/Miniconda)
#or from local file (direct download files from Github).
try:
    import excel_to_dataframe.excel_to_pandas as etd    
except ModuleNotFoundError as e1:        
    import excel_to_pandas as etd	
except ModuleNotFoundError as e2:    
    print("Error trying to import excel_to_dataframe. Error:", e2)

absolute_path = os.path.abspath(os.path.dirname(__file__))

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
    excel_file = os.path.join(absolute_path,'benchmarks//infomercado_individuais_20221.xlsx')
   
    sheet_name = '003 Consumo'
    
    # Open Excel file.
    ret = etd.open_excel(excel_file)
    
    # Gets a dataframe from worksheet.
    df = etd.ws_to_df(sheet_name,mult_thread=0)
        
    # Closes Excel.
    etd.close_excel()

    # Splits dataframe based on string search.
    # ldf - list of dataframes.
    ldf = etd.split_df(df,'Tabela ',col_search=1,header_offset=1)

    print('Done.')


def test_dll(excel_file_full_path, excel_worksheet_name):
    """
    Using dll, opens a Excel file and converts a worksheet to dataframe
    This rotine is used by benchmark function.
    
    Arguments:
    excel_file_full_path    - (string) full path for Excel file and
    
    Returns:
    None

    Requires:
    ctypes, excel_to_csv.dll

    Version 0.2.52

    """
    
    ret = etd.open_excel(excel_file_full_path)
           
    # If Excel file is loaded, tries to load worksheet.
    if (ret==0):        
        tmp_df = etd.ws_to_df(excel_worksheet_name)
    else:
        print('Error on load Excel! Error number:{:d}'.format(ret))
        
    etd.close_excel()
    
       

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

    Version 0.2.53

    """

    # Excel files and worksheets for benchmarking
    # excel_files = ['benchmarks//infomercado_individuais_2022.xlsx','benchmarks//infomercado_contratos.xlsx']
    excel_files = ['benchmarks//infomercado_contratos.xlsx','benchmarks//infomercado_individuais_2022.xlsx']
    # sheet_names  = ['003 Consumo','Contratos Distribuidoras']
    sheet_names  = ['Contratos Distribuidoras','003 Consumo']

    for a in range(0,2):
        excel_file = os.path.join(absolute_path, excel_files[a])
        sheet_name = sheet_names[a]

        print('**************************************************************************************************')
        print('Benchmark test {} - Excel File \'{}\', Sheet \'{}\''.format(a+1,excel_files[a],sheet_name))
        print('')        
        
        # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        tmp_test = test_dll
        dll_time_mt = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1,globals=locals())
        
        # 'tmp_test' is a local pointer to the function. Allows usage of 'timeit' inside functions.
        tmp_test = test_pure_pandas
        pd_time = timeit.timeit("tmp_test(excel_file,sheet_name)", number=1, globals=locals())
                
        # Calculates the ratio pandas time / dll time.                
        ratio_mt = pd_time/dll_time_mt
        
        # Prints resutls.
        print('')
        print('**************************************************************************************************')
        print('Results:')
        print('')
        print('Pandas..................... {:.2f} seconds.'.format(pd_time))
        print('DLL........................ {:.2f} seconds.'.format(dll_time_mt))
        print('--> DLL is................. {:.2f} times faster than Pandas.'.format(ratio_mt))
        print('')
        


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

    Version 0.2.52
    
    """
    
    # Converts worksheet data to Pandas dataframe
    ret = etd.open_excel(excel_file_full_path)
    tmp_df = etd.ws_to_df(excel_worksheet_name)

    if (isinstance(tmp_df,pd.DataFrame)):
        tmp_df.to_csv(excel_worksheet_name + '.csv',sep=';',decimal=',')
    else:
        print("Error: Can't not open sheet: " + excel_worksheet_name)
    
    a = etd.close_excel  
    
    del tmp_df
    
    # # Converts worksheet data to Pandas dataframe (Pandas)
    # excel = pd.ExcelFile(excel_file_full_path)    
    # tmp_df = pd.read_excel(excel, sheet_name=excel_worksheet_name)   
    # tmp_df.to_csv(excel_worksheet_name + '_pd.csv',sep=';',decimal=',')
    # del tmp_df



# *********** Entry point **********
if __name__ == "__main__":
    
   
    # Gets and prints dll version.
    print (etd.version())
    
    # # Runs benchmarks
    benchmarks()
    exit(0)

    #  Saves csvs to comparisons;
    print('')
    print('**************************************************************************************************')
    print('Converting worksheets to csvs:')
    # get_csvs(os.path.join(absolute_path,'benchmarks/infomercado_individuais_2022.xlsx'), '003 Consumo')    
    # get_csvs(os.path.join(absolute_path, 'benchmarks/infomercado_contratos.xlsx'), 'Notas Explicativas')                                                                    
    # get_csvs(os.path.join(absolute_path, 'benchmarks/infomercado_contratos.xlsx'), 'Contratos Distribuidoras')
    # print('Done!')
    
    # # Another dll usage example.
    # main()

    
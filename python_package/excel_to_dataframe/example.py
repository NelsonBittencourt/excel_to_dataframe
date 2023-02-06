# ************************************************
# Python script showing dll usage and benchmarks.
# Created by Nelson Rossi Goulart Bittencourt.
# Github: https://github.com/nelsonbittencourt
# 14/12/2022 (dd/mm/yyyy).
# ************************************************


# *********** Imports **********

import excel_to_df as etd       # Excel to Pandas functions
import pandas as pd             # Pandas



# *********** Functions **********

def main():
    """
    Single dll Usage example. Use breakpoints and debuger mode to view the results.
    
    Arguments:
    None.

    Returns:
    None.

    Requires:
    excel_to_df.dll.
    
    """

    # Excel file and worksheet name to work with.
    excel_file = 'benchmarks/infomercado_individuais_2022.xlsx'
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



# *********** Entry point **********
if __name__ == "__main__":
    
    # Gets and prints dll version.
    print (etd.get_dll_version())
    
    # Runs benchmarks
    etd.benchmarks()

    # # Saves csvs to comparisons;
    # etd.get_csvs('benchmarks/infomercado_individuais_2022.xlsx', '003 Consumo')
    # etd.get_csvs('benchmarks/infomercado_contratos.xlsx', 'Contratos Distribuidoras')
    
    # # Dll usage example.
    # main()

    
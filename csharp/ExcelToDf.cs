using Microsoft.Data.Analysis;              // To use dataframe class.
using System.Runtime.InteropServices;       // Interops with dll.

namespace csExcelToDf
{
    class excelToDF
    {
        
        // Flag to indicate if Excel is open.
        private static bool isExcelOpened = false;

        // Path to dll.
        private const string dllPath = "excel_to_df_cs.dll";

        // Prototypes to access dll functions.
        [DllImport(dllPath,CallingConvention = CallingConvention.Cdecl)]
        private static extern string version();

        [DllImport(dllPath,CallingConvention = CallingConvention.Cdecl)]
        private static extern int openExcel(string ExcelFileName);

        [DllImport(dllPath,CallingConvention = CallingConvention.Cdecl)]
        private static extern int closeExcel();

        [DllImport(dllPath,CallingConvention = CallingConvention.Cdecl)]
        private static extern string getSheetMT(string sheetName); 
        
        // Open an Excel file.
        public static void OpenExcelFile(string excelFile)
        {            
            if (!isExcelOpened)
            {
                int ret = openExcel(excelFile);
                if (ret==0)
                {
                    isExcelOpened=true;
                }
                else
                {
                    Console.WriteLine("Open Excel error. Error number: " + ret.ToString());
                }
            }            
        }

        // Closes an Excel file.
        public static void CloseExcel()
        {
            isExcelOpened = false;
            closeExcel();
        }
        
        // Gets a DataFrame from worksheet "sheetName" of opened Excel file.
        public static DataFrame GetSheetData(string sheetName) 
        {   
            // Gets worksheet data as a string.
            string dataframe_text = getSheetMT(sheetName);
                        
            // Tries to convert string to DataFrame.
            DataFrame df1 = DataFrame.LoadCsvFromString(dataframe_text,separator:';',header:false,guessRows:1);

            return df1;    
        }

        // Prints dll information.
        public static string GetDllVersion()
        {
            return version();
        }


    }
}
using Microsoft.Data.Analysis;  // To use dataframe class.  

namespace csExcelToDf
{
    class Program
    {
        public static void Main()
        {
            
            // Gets and prints dll information.
            string version = excelToDF.GetDllVersion();
            Console.WriteLine(version);

            // Opens Excel file.
            excelToDF.OpenExcelFile("benchmarks/infomercado_individuais_2022.xlsx");
            
            // Gets worksheet data.
            Console.WriteLine("Loading worksheet data...");
            DataFrame df = excelToDF.GetSheetData("003 Consumo");
            
            // Closes Excel file.            
            excelToDF.CloseExcel();
            
            // Saves dataframe to csv format.
            Console.WriteLine("Saving csv...");
            DataFrame.SaveCsv(df,"test.csv",separator:';');

            Console.WriteLine("Done.");
        }
    }
}
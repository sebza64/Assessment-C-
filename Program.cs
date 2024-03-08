using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static async Task Main()
    {
        string url = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";
        string localFolderPath = "C:\\Users\\smsomi\\Desktop\\Assessment";

        await ProcessFilesAsync(url, localFolderPath);

        Console.WriteLine("Files processed successfully!");



    }

    static async Task ProcessFilesAsync(string url, string localFolderPath)
    {
        // ... (previous code)

        // Inside the existing if (!downloadedFiles.Contains(fileName)) block
        if (!downloadedFiles.Contains(fileName))
        {
            // Read and process the Excel file
            await ProcessExcelFile(localFilePath);

            // ... (rest of the existing code)
        }
    }

    static async Task ProcessExcelFile(string filePath)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
        Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Assuming data is on the first sheet

        // Get contract details from the Excel file and save to SQL database
        SaveContractDetailsToDatabase(worksheet);

        // Release Excel resources
        Marshal.ReleaseComObject(worksheet);
        workbook.Close(false);
        Marshal.ReleaseComObject(workbook);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);
    }

    static void SaveContractDetailsToDatabase(Excel.Worksheet worksheet)
    {
        // Implement your logic to extract and save contract details to SQL database
        // For example, you can iterate through rows and columns to get cell values
        // and then insert them into the SQL database.

        // Example:
        string contractName = ((Excel.Range)worksheet.Cells[1, 1]).Value?.ToString();
        string contractValue = ((Excel.Range)worksheet.Cells[1, 2]).Value?.ToString();

        // Connect to SQL database and execute the insert query
         SqlConnection connection = new SqlConnection("YData Source=SMSOMI650G4\FLOWCENTRIC;User ID=smsomi;Password=***********");
         connection.Open();
        SqlCommand command = new SqlCommand($"INSERT INTO DailyMTM (ContractName, ContractValue) VALUES ('{contractName}', '{contractValue}')", connection);
         command.ExecuteNonQuery();
        connection.Close();
    }
}

using System;
using OfficeOpenXml;


// Path to your Excel file
string filePath = @"C:\Users\sadra\source\repos\ReadHierarchicalExcelFile\ReadHierarchicalExcelFile\test.xlsx";

// Load the Excel file
using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
{
    // Assuming the data is in the first worksheet
    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

    // Get the number of rows in the worksheet
    int rowCount = worksheet.Dimension.Rows;
    int columnsCount = worksheet.Dimension.Columns;

    // Read hierarchical data from Excel cells
    for (int i = 1; i <= rowCount; i++)
    {
        // Assuming the hierarchical data is in the first column
        string hierarchicalData = worksheet.Cells[i, 2].Value?.ToString();

        // Output hierarchical data (you can process it further as needed)
        Console.WriteLine(hierarchicalData);
    }
}

Console.WriteLine("Press any key to exit...");
Console.ReadKey();

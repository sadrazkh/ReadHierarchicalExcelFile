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

    var topics = new List<Topic>();

    // Read hierarchical data from Excel cells
    for (int i = 2; i <= rowCount; i++)
    {
        for (int j = columnsCount; j > 1; j--)
        {
            if (!string.IsNullOrEmpty(worksheet.Cells[i, j].Value?.ToString()))
            {
                topics.Add(new Topic() { Name = worksheet.Cells[i, j].Value?.ToString(), Row = i, Columns = j, Children = null, Show = true});
            }
            //else
            //{
            //    topics.Add((new Topic() { Name = worksheet.Cells[i, j].Value?.ToString(), Row = i, Columns = j, Children = null, Show = true}));
            //}
        }
        // Assuming the hierarchical data is in the first column

        // Output hierarchical data (you can process it further as needed)
        //Console.WriteLine(hierarchicalData);
    }

    foreach (var topic in topics.OrderBy(c => c.Columns).ToList())
    {
        var columns = topic.Columns + 1;
        if (columns > columnsCount)
            continue;
        
        var rows = topic.Row;
        var maxRows = topic.Row;


        var nextTopic = topics.Where(c => c.Columns == topic.Columns).FirstOrDefault(c => c.Row > topic.Row);
        if (nextTopic != null)
        {
            maxRows = nextTopic.Row;
        }
        else
        {
            maxRows = rowCount;
        }

        var childrens = topics.Where(c => c.Columns == columns).Where(c => c.Row >= rows).Where(c => c.Row < maxRows)
            .ToList();


        topic.Children = childrens;
    }


    topics = topics.OrderBy(c => c.Columns).ToList();
}


Console.WriteLine("Press any key to exit...");
Console.ReadKey();



public class Topic
{
    public int Row { get; set; }
    public int Columns { get; set; }
    public string Name { get; set; }
    public bool Show { get; set; }
    public List<Topic> Children { get; set; }
}
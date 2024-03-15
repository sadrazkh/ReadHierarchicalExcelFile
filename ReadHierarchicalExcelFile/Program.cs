using System;
using OfficeOpenXml;


string filePath = @"C:\Users\sadra\source\repos\ReadHierarchicalExcelFile\ReadHierarchicalExcelFile\test.xlsx";

using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePath)))
{
    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

    int rowCount = worksheet.Dimension.Rows;
    int columnsCount = worksheet.Dimension.Columns;

    var topics = new List<Topic>();

    for (int i = 2; i <= rowCount; i++)
    {
        for (int j = columnsCount; j > 1; j--)
        {
            if (!string.IsNullOrEmpty(worksheet.Cells[i, j].Value?.ToString()))
            {
                topics.Add(new Topic() { Name = worksheet.Cells[i, j].Value?.ToString(), Row = i, Columns = j, Children = null, Show = true});
            }
        }
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
            maxRows = rowCount + 1;
        }

        var childrens = topics.Where(c => c.Columns == columns).Where(c => c.Row >= rows).Where(c => c.Row < maxRows)
            .ToList();

        topics.Where(c => c.Columns == columns).Where(c => c.Row >= rows).Where(c => c.Row < maxRows).ToList().ForEach(c => c.Show = false);
        topic.Children = childrens;
    }


    topics = topics.OrderBy(c => c.Columns).Where(c => c.Show).ToList();
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
using System;
using System.Data;
using Aspose.Cells;

public class DataTableToExcelExporter
{
    /// <summary>
    /// Exports the content of a DataTable to an XLSX workbook while preserving formatting.
    /// </summary>
    /// <param name="dataTable">The DataTable containing the data to export.</param>
    /// <param name="templatePath">
    /// Optional path to an existing Excel template. If null or empty, a new workbook is created.
    /// </param>
    /// <param name="outputPath">Full file path where the resulting XLSX file will be saved.</param>
    public static void Export(DataTable dataTable, string? templatePath, string outputPath)
    {
        // Create or load the workbook
        Workbook workbook = string.IsNullOrWhiteSpace(templatePath)
            ? new Workbook()
            : new Workbook(templatePath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Write column headers
        for (int c = 0; c < dataTable.Columns.Count; c++)
        {
            cells[0, c].PutValue(dataTable.Columns[c].ColumnName);
        }

        // Write data rows
        for (int r = 0; r < dataTable.Rows.Count; r++)
        {
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                cells[r + 1, c].PutValue(dataTable.Rows[r][c]);
            }
        }

        // Save the workbook to the specified output file
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }

    // Sample entry point for testing
    public static void Main()
    {
        // Create sample DataTable
        DataTable dt = new DataTable("Sample");
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Name", typeof(string));
        dt.Columns.Add("Date", typeof(DateTime));

        dt.Rows.Add(1, "Alice", DateTime.Now);
        dt.Rows.Add(2, "Bob", DateTime.Now.AddDays(1));
        dt.Rows.Add(3, "Charlie", DateTime.Now.AddDays(2));

        // Export without a template
        string outputPath = "ExportedData.xlsx";
        Export(dt, null, outputPath);

        Console.WriteLine($"Data exported to {outputPath}");
    }
}
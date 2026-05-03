using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare sample data in a DataTable
        DataTable table = new DataTable();
        table.Columns.Add("Product", typeof(string));
        table.Columns.Add("Quantity", typeof(int));
        table.Columns.Add("Price", typeof(double));

        table.Rows.Add("Apple", 10, 0.5);
        table.Rows.Add("Banana", 20, 0.3);
        table.Rows.Add("Cherry", 15, 1.2);

        // Manually import the DataTable into the worksheet (including column names)
        int startRow = 0;
        int startColumn = 0;

        // Write column headers
        for (int col = 0; col < table.Columns.Count; col++)
        {
            cells[startRow, startColumn + col].PutValue(table.Columns[col].ColumnName);
        }

        // Write data rows
        int currentRow = startRow + 1;
        foreach (DataRow dr in table.Rows)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                cells[currentRow, startColumn + col].PutValue(dr[col]);
            }
            currentRow++;
        }

        // Save the workbook as an XLSX file
        workbook.Save("DataGridImportDemo.xlsx");
    }
}
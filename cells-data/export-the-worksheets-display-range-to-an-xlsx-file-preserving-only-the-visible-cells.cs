using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsVisibleExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a source workbook and fill it with sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Populate some data
            sourceSheet.Cells["A1"].PutValue("Header1");
            sourceSheet.Cells["B1"].PutValue("Header2");
            sourceSheet.Cells["C1"].PutValue("Header3");
            sourceSheet.Cells["A2"].PutValue("Data1");
            sourceSheet.Cells["B2"].PutValue("Data2");
            sourceSheet.Cells["C2"].PutValue("Data3");
            sourceSheet.Cells["A3"].PutValue("Data4");
            sourceSheet.Cells["B3"].PutValue("Data5");
            sourceSheet.Cells["C3"].PutValue("Data6");

            // Hide a row and a column to demonstrate visibility filtering
            sourceSheet.Cells.HideRow(1);      // Hide second row (index 1)
            sourceSheet.Cells.HideColumn(1);  // Hide second column (index 1)

            // Determine the used range size
            int totalRows = sourceSheet.Cells.MaxDisplayRange.RowCount;
            int totalCols = sourceSheet.Cells.MaxDisplayRange.ColumnCount;

            // Set export options to include only visible cells
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                PlotVisibleCells = true,
                PlotVisibleRows = true,
                PlotVisibleColumns = true,
                ExportColumnName = true
            };

            // Export the visible portion of the worksheet to a DataTable
            DataTable visibleData = sourceSheet.Cells.ExportDataTable(
                0,               // start row
                0,               // start column
                totalRows,       // number of rows
                totalCols,       // number of columns
                exportOptions);

            // Create a new workbook to hold the exported visible data
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];

            // Manually import the DataTable into the new worksheet (including column names)
            int startRow = 0;
            // Insert column names as the first row
            for (int c = 0; c < visibleData.Columns.Count; c++)
            {
                destSheet.Cells[startRow, c].PutValue(visibleData.Columns[c].ColumnName);
            }
            startRow++;

            // Insert data rows
            for (int r = 0; r < visibleData.Rows.Count; r++)
            {
                for (int c = 0; c < visibleData.Columns.Count; c++)
                {
                    destSheet.Cells[startRow + r, c].PutValue(visibleData.Rows[r][c]);
                }
            }

            // Save the result as XLSX, containing only the visible cells
            destWorkbook.Save("VisibleCellsExport.xlsx", SaveFormat.Xlsx);
        }
    }
}
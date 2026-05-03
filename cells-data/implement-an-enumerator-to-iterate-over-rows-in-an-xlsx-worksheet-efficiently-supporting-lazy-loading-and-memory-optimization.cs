using System;
using System.Collections;
using Aspose.Cells;

class LazyRowEnumeratorDemo
{
    static void Main()
    {
        // Load a large XLSX file with memory‑optimized settings.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference; // minimize memory footprint

        // Replace "LargeFile.xlsx" with the path to your source workbook.
        Workbook workbook = new Workbook("LargeFile.xlsx", loadOptions);

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain a synchronized enumerator for rows.
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator(false, true);

        // Iterate through rows lazily. Rows and their cells are materialized only when accessed.
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;

            // Access the first cell of the current row only when needed.
            Cell firstCell = row.GetCellOrNull(0);
            string cellValue = firstCell != null ? firstCell.StringValue : "<empty>";

            Console.WriteLine($"Row {row.Index}: {cellValue}");
        }

        // Save the workbook (unchanged) using the standard Save method.
        workbook.Save("Processed.xlsx");
    }
}
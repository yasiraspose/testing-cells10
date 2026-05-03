using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();                     // lifecycle: create
        Worksheet sheet = workbook.Worksheets[0];              // get first worksheet
        Cells cells = sheet.Cells;                             // access cells collection

        // ---------- Write values to individual cells ----------
        // Using row/column indexer (zero‑based)
        cells[0, 0].PutValue("Product");   // A1
        cells[0, 1].PutValue("Price");     // B1

        // Using A1 style address
        cells["A2"].PutValue("Laptop");
        cells["B2"].PutValue(1200.50);
        cells["A3"].PutValue("Phone");
        cells["B3"].PutValue(799.99);

        // ---------- Read values back ----------
        Console.WriteLine($"{cells[0, 0].StringValue}\t{cells[0, 1].StringValue}");
        for (int r = 1; r <= 2; r++)
        {
            string product = cells[r, 0].StringValue;
            double price = cells[r, 1].DoubleValue;
            Console.WriteLine($"{product}\t{price}");
        }

        // ---------- Apply a simple style to the header row ----------
        Style headerStyle = workbook.CreateStyle();            // create a new style
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        cells[0, 0].SetStyle(headerStyle);
        cells[0, 1].SetStyle(headerStyle);

        // ---------- Save the workbook ----------
        workbook.Save("Products.xlsx");                         // lifecycle: save

        // ---------- Load the saved file ----------
        Workbook loaded = new Workbook("Products.xlsx");        // lifecycle: load
        Worksheet loadedSheet = loaded.Worksheets[0];
        Cells loadedCells = loadedSheet.Cells;

        // ---------- Modify a cell in the loaded workbook ----------
        loadedCells["B2"].PutValue(1150.00); // discounted price for Laptop

        // ---------- Save the modified workbook ----------
        loaded.Save("Products_Updated.xlsx");
    }
}
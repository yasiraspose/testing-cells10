using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data using the row/column indexer
        cells[0, 0].PutValue("Product");   // A1
        cells[0, 1].PutValue("Price");     // B1
        cells[1, 0].PutValue("Apple");     // A2
        cells[1, 1].PutValue(1.20);        // B2
        cells[2, 0].PutValue("Banana");    // A3
        cells[2, 1].PutValue(0.80);        // B3

        // Retrieve a cell by its address (name indexer) and read its value
        Cell priceCell = cells["B2"]; // B2 corresponds to Apple price
        double originalPrice = priceCell.DoubleValue;
        Console.WriteLine($"Original price of Apple: {originalPrice}");

        // Modify the cell value
        priceCell.PutValue(originalPrice + 0.30); // increase price by 0.30
        Console.WriteLine($"Updated price of Apple: {priceCell.DoubleValue}");

        // Apply a style to the header row using the name indexer
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        cells["A1"].SetStyle(headerStyle);
        cells["B1"].SetStyle(headerStyle);

        // Save the workbook (lifecycle: save)
        workbook.Save("Products.xlsx");
    }
}
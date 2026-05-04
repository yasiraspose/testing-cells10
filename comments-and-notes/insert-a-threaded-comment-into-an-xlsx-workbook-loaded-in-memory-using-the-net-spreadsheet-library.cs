using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook in memory
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "John Doe",               // Author name
            "john.doe@example.com",   // User ID / email
            "EXAMPLE_PROVIDER");      // Provider ID (can be empty)

        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 using the string overload
        worksheet.Comments.AddThreadedComment("B2", "This is a threaded comment.", author);

        // Retrieve and display the threaded comments for verification
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments("B2");
        foreach (ThreadedComment comment in threadedComments)
        {
            Console.WriteLine($"Author: {comment.Author.Name}, Text: {comment.Notes}");
        }

        // Save the workbook to a memory stream (in‑memory XLSX)
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            // The MemoryStream now contains the XLSX file bytes.
        }
    }
}
using System;
using Aspose.Cells;

namespace ThreadedCommentTimestampDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Access the comments collection of the current worksheet
                CommentCollection comments = worksheet.Comments;

                // Determine the used range to limit the search for threaded comments
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve threaded comments for the current cell (by row and column)
                        ThreadedCommentCollection threadedComments = comments.GetThreadedComments(row, col);

                        // If there are any threaded comments, output their creation timestamps
                        if (threadedComments != null && threadedComments.Count > 0)
                        {
                            string cellName = worksheet.Cells[row, col].Name;
                            Console.WriteLine($"Cell {cellName} contains {threadedComments.Count} threaded comment(s):");

                            for (int i = 0; i < threadedComments.Count; i++)
                            {
                                ThreadedComment tc = threadedComments[i];
                                Console.WriteLine($"  Comment {i + 1} created at: {tc.CreatedTime}");
                            }
                        }
                    }
                }
            }

            // Save the workbook (optional – here we simply rewrite the same file)
            workbook.Save("output.xlsx");
        }
    }
}
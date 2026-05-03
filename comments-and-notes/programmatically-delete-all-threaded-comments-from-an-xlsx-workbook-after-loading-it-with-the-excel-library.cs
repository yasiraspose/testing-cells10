using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DeleteAllThreadedComments
    {
        public static void Run()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear all comments (including threaded comments) from the worksheet
                sheet.ClearComments();
            }

            // Path to save the modified workbook
            string outputPath = "output_without_threaded_comments.xlsx";

            // Save the workbook after removing comments
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DeleteAllThreadedComments.Run();
        }
    }
}
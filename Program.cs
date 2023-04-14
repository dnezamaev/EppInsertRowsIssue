using OfficeOpenXml;

namespace EppInsertRowsIssue
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var eppPackage = new ExcelPackage();

            using (var stream = File.OpenRead("insert_rows_test.xlsx"))
            {
                eppPackage.Load(stream);
            }

            // We have 2 rows with formulas in C column.
            var book = eppPackage.Workbook;
            var eppWorksheet = book.Worksheets[0];

            // Insert row(-s) after first one. 
            // New row one-based index is 2.
            // Second row index is 3 now.
            eppWorksheet.InsertRow(2, 1);

            // Formula updated fine =A3*B3.
            Console.WriteLine(eppWorksheet.Cells["C3"].Formula);

            // Insert row(-s) after second one (it has index 3 now).
            eppWorksheet.InsertRow(4, 1);

            // Formula should not be updated, because row 3 is above.
            // But now its =A2*B2. Why?
            Console.WriteLine(eppWorksheet.Cells["C3"].Formula);

            eppPackage.SaveAs("out.xlsx");
        }
    }
}
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PiramidReportBuilder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string sourseFilePath = "";
            string destinationFilePath = "";

            using (XLWorkbook sourseWorkbook = new XLWorkbook(sourseFilePath))
            {
                IXLRange sourseRange = sourseWorkbook.Worksheet(1).Worksheet.RangeUsed();
                sourseRange = sourseRange.Range($"F6:N{sourseRange.LastRowUsed().RowNumber()}");

                using (XLWorkbook destinationWorkbook = new XLWorkbook(destinationFilePath))
                {
                    sourseRange.CopyTo(destinationWorkbook.AddWorksheet("New Sheet").Cell(1, 1));

                    destinationWorkbook.SaveAs("");
                }
            }
        }
    }
}
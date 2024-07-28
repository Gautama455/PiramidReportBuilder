using ClosedXML.Excel;

namespace PiramidReportBuilder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (File.Exists("")) //filepath 
            {
                try
                {
                    using (XLWorkbook workbook = new XLWorkbook("")) //filepath 
                    {
                        workbook.Worksheets.Add("полнота сбора");
                        
                        workbook.SaveAs(""); //filepath 
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine("Не удалось открыть файл. Возможно, он уже открыт в другой программе.");
                    Console.WriteLine("Ошибка: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Файл не существует");
            }

        }

        public IXLRange Parse()
        {
            IXLRange currentRange = null;

            if (File.Exists("")) //filepath
            {
                using (XLWorkbook workbook = new XLWorkbook(""))  //filepath 
                {
                    IXLRange xLRange = workbook.Worksheet(1).Worksheet.RangeUsed();
                    currentRange = xLRange.Range($"F6:N{xLRange.LastRow().RowNumber()}");
                }
            }

            return currentRange;
        }
    }
}
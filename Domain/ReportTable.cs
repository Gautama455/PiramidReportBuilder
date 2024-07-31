using ClosedXML.Excel;

namespace PiramidReportBuilder.Domain
{
    internal abstract class ReportTable : IReportTable
    {
        protected readonly string _tableName;

        public ReportTable(string tableName)
        {
            _tableName = tableName;
        }

        public string Name => _tableName;

        //private void ReadData()
        //{

        //}
    }
}

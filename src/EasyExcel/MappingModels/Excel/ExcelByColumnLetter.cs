namespace EasyExcel.MappingModels.Excel
{
    public class ExcelByColumnLetter : Base.ByColumnLetter
    {
        private string _columnHeader { get; set; }

        public ExcelByColumnLetter(string columnLetter,
            string attributeName,
            string columnHeader)
            : base(attributeName, columnLetter)
        {
            _columnHeader = columnHeader;
        }

        public string ColumnHeader => _columnHeader;
    }
}

namespace EasyExcel.MappingModels.Excel
{
    public class ExcelByColumnIndex : Base.ByColumnIndex
    {
        private string _columnHeader { get; set; }

        public ExcelByColumnIndex(int columnIndex,
            string attributeName,
            string columnHeader)
            : base(attributeName, columnIndex)
        {
            _columnHeader = columnHeader;
        }

        public string ColumnHeader => _columnHeader;
    }
}

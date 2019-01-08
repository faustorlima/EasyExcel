namespace EasyExcel.MappingModels.Base
{
    public abstract class ByColumnIndex : Common
    {
        private int _columnIndex { get; }

        public ByColumnIndex(string attributeName,
            int columnIndex)
            : base(attributeName)
        {
            _columnIndex = columnIndex;
        }

        public int ColumnIndex => _columnIndex;
    }
}

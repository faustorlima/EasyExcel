namespace EasyExcel.MappingModels.Base
{
    public abstract class ByColumnLetter : Common
    {
        private string _columnLetter { get; }

        public ByColumnLetter(string attributeName,
            string columnLetter)
            : base(attributeName)
        {
            _columnLetter = columnLetter;
        }

        public string ColumnLetter => _columnLetter;
    }
}

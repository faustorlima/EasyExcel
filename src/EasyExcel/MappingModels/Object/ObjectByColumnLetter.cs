namespace EasyExcel.MappingModels.Object
{
    public class ObjectByColumnLetter : Base.ByColumnLetter
    {
        private bool _required { get; }

        public ObjectByColumnLetter(
            string attributeName,
            string columnLetter,
            bool required)
            : base(attributeName, columnLetter)
        {
            _required = required;
        }

        public bool Required => _required;
    }
}

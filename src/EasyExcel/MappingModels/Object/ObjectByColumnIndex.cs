namespace EasyExcel.MappingModels.Object
{
    public class ObjectByColumnIndex : Base.ByColumnIndex
    {
        private bool _required { get; set; }

        public ObjectByColumnIndex(int columnIndex,
            string attributeName,
            bool required)
            : base(attributeName, columnIndex)
        {
            _required = required;
        }

        public bool Required => _required;
    }
}

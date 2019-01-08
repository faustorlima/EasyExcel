namespace EasyExcel.MappingModels.Base
{
    public abstract class Common
    {
        private string _attributeName { get; set; }

        public Common(string attributeName)
        {
            _attributeName = attributeName;
        }

        public string AttributeName => _attributeName;
    }
}

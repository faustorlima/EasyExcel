using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using EasyExcel.Helper;
using OfficeOpenXml;
using EasyExcel.MappingModels.Object;

namespace EasyExcel
{
    public class ExcelToObjectConverter
    {
        public static IEnumerable<T> ToObjectCollection<T>(string spreadsheetFilePath, IEnumerable<ObjectByColumnLetter> columnsMapping) where T : new()
        {
            var columnIndexMapping = ConverterHelper.GetMappingReadByColumnIndex(columnsMapping);
            return ToObjectCollection<T>(spreadsheetFilePath, columnIndexMapping);
        }

        public static IEnumerable<T> ToObjectCollection<T>(string spreadsheetFilePath, IEnumerable<ObjectByColumnIndex> columnsMapping) where T : new()
        {
            using (var stream = File.OpenRead(spreadsheetFilePath))
                return ToObjectCollection<T>(stream, columnsMapping);
        }

        public static IEnumerable<T> ToObjectCollection<T>(Stream spreadsheet, IEnumerable<ObjectByColumnLetter> columnsMapping) where T : new()
        {
            var columnIndexMapping = ConverterHelper.GetMappingReadByColumnIndex(columnsMapping);
            return ToObjectCollection<T>(spreadsheet, columnIndexMapping);
        }

        public static IEnumerable<T> ToObjectCollection<T>(Stream spreadsheet, IEnumerable<ObjectByColumnIndex> columnsMapping) where T : new()
        {
            var r = new List<T>();
            var ws = new ExcelPackage(spreadsheet).Workbook.Worksheets.FirstOrDefault();
            var hasData = true;
            var spreadsheetLine = 2;

            while (hasData)
            {
                var hasAnyData = false;
                foreach (var columnMapping in columnsMapping)
                {
                    if (ws.Cells[spreadsheetLine, columnMapping.ColumnIndex].Value != null)
                    {
                        hasAnyData = true;
                        break;
                    }
                }

                if (!hasAnyData) return r;

                var properties = typeof(T).GetProperties();
                var newItem = new T();

                foreach (var property in properties)
                {
                    var columnMapping = columnsMapping.FirstOrDefault(c => c.AttributeName == property.Name);
                    if (columnMapping != null)
                    {
                        var value = ws.Cells[spreadsheetLine, columnMapping.ColumnIndex].Value;

                        if (value == null)
                        {
                            if (columnMapping.Required)
                                throw new Exception(string.Format("The required field {0} is empty at line {1} of spreadsheet.", property.Name, spreadsheetLine));
                            else
                                continue;
                        }

                        try
                        {
                            Type type = null;
                            if (Nullable.GetUnderlyingType(property.PropertyType) != null)
                            {
                                if (property.PropertyType.IsEnum)
                                    type = typeof(Nullable<Int32>);
                                else
                                    type = Nullable.GetUnderlyingType(property.PropertyType);
                            }
                            else if (property.PropertyType.IsEnum)
                                type = typeof(Int32);
                            else
                                type = property.PropertyType;
                        
                            property.SetValue(newItem, Convert.ChangeType(value, type));
                        }
                        catch
                        {
                            throw new Exception(string.Format("It was not possible to convert value: '{0}' to attribute {1}({2}) at line {3} and column '{4}' of spreadsheet.", value, property.Name, property.PropertyType, spreadsheetLine, columnMapping.ColumnIndex));
                        }
                    }
                }

                r.Add(newItem);

                spreadsheetLine++;
            }

            return r;
        }
    }
}

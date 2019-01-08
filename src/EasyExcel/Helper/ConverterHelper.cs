using EasyExcel.MappingModels.Excel;
using EasyExcel.MappingModels.Object;
using System;
using System.Collections.Generic;

namespace EasyExcel.Helper
{
    public class ConverterHelper
    {
        public static IEnumerable<ExcelByColumnIndex> GetMappingWriteByColumnIndex(IEnumerable<ExcelByColumnLetter> columnsMapping)
        {
            var columnIndexMapping = new List<ExcelByColumnIndex>();
            foreach (var columnMapping in columnsMapping)
            {
                columnIndexMapping.Add(new ExcelByColumnIndex(
                    ExcelColumnNameToNumber(columnMapping.ColumnLetter),
                    columnMapping.AttributeName,
                    columnMapping.ColumnHeader));
            }
            return columnIndexMapping;
        }

        public static IEnumerable<ObjectByColumnIndex> GetMappingReadByColumnIndex(IEnumerable<ObjectByColumnLetter> columnsMapping)
        {
            IList<ObjectByColumnIndex> columnIndexMapping = new List<ObjectByColumnIndex>();
            foreach (var columnMapping in columnsMapping)
            {
                columnIndexMapping.Add(new ObjectByColumnIndex(
                    ExcelColumnNameToNumber(columnMapping.ColumnLetter),
                    columnMapping.AttributeName,
                    columnMapping.Required));
            }
            return columnIndexMapping;
        }

        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
    }
}

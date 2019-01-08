using System.IO;
using System.Collections.Generic;
using EasyExcel.MappingModels.Excel;
using OfficeOpenXml;
using EasyExcel.Helper;
using System;

namespace EasyExcel
{
    public sealed class ObjectToExcelConverter 
    {
        public static void CreateFileFromObjectCollection<T>(IEnumerable<ExcelByColumnLetter> columnsMapping, IEnumerable<T> data, string targetSpreadsheetPath)
        {
            var columnIndexMapping = ConverterHelper.GetMappingWriteByColumnIndex(columnsMapping);
            CreateFileFromObjectCollection<T>(columnIndexMapping, data, targetSpreadsheetPath);
        }
        
        public static void CreateFileFromObjectCollection<T>(IEnumerable<ExcelByColumnIndex> columnsMapping, IEnumerable<T> data, string targetSpreadsheetPath)
        {
            var stream = FromObjectCollection<T>(columnsMapping, data);
            using (FileStream output = new FileStream(targetSpreadsheetPath, FileMode.Create))
            {
                stream.CopyTo(output);
            }
        }
        
        public static Stream FromObjectCollection<T>(IEnumerable<ExcelByColumnLetter> columnsMapping, IEnumerable<T> data)
        {
            var columnIndexMapping = ConverterHelper.GetMappingWriteByColumnIndex(columnsMapping);
            return FromObjectCollection<T>(columnIndexMapping, data);
        }
        
        public static Stream FromObjectCollection<T>(IEnumerable<ExcelByColumnIndex> columnsMapping, IEnumerable<T> data)
        {
            var s = new MemoryStream();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Data");

            foreach (var columnMapping in columnsMapping)
            {
                ws.Cells[1, columnMapping.ColumnIndex].Value = columnMapping.ColumnHeader;
            }

            var spreadsheetLine = 2;

            foreach (var item in data)
            {
                foreach (var column in columnsMapping)
                {
                    if (typeof(T).GetProperty(column.AttributeName).PropertyType.IsEnum)
                        ws.Cells[spreadsheetLine, column.ColumnIndex].Value = (int)typeof(T).GetProperty(column.AttributeName).GetValue(item);
                    else if (typeof(T).GetProperty(column.AttributeName).PropertyType.Name.Equals("DateTime"))
                        ws.Cells[spreadsheetLine, column.ColumnIndex].Value = Convert.ToDateTime(typeof(T).GetProperty(column.AttributeName).GetValue(item)).ToString("yyyy-MM-dd HH:mm:ss");
                    else
                        ws.Cells[spreadsheetLine, column.ColumnIndex].Value = typeof(T).GetProperty(column.AttributeName).GetValue(item);
                }

                spreadsheetLine += 1;
            }

            pck.Save();

            pck.SaveAs(s);
            s.Position = 0;

            return s;
        }
    }
}

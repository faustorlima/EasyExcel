using Bogus;
using Bogus.DataSets;
using EasyExcel.MappingModels.Excel;
using EasyExcel.Tests.TestModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace EasyExcel.Tests
{
    public class ObjectToExcelConverterTests
    {
        private readonly Randomizer random = new Randomizer();

        [Fact]
        public void ObjectList_GetSpreadSheetStream_Sucess()
        {
            // Arrange
            var employees = new Faker<Employee>("en")
                 .CustomInstantiator(f => new Employee(
                     f.Name.FullName(Name.Gender.Male),
                     Gender.Male,
                     f.Date.Past(60, DateTime.Now.AddYears(-16)),
                     decimal.Round(random.Number(5, 7) + random.Decimal(0, 1), 2),
                     decimal.Round(random.Number(160, 500) + random.Decimal(0, 1), 2)
                     )).Generate(100);

            var columnsMapping = new List<ExcelByColumnIndex> {
                new ExcelByColumnIndex(1, "Name", "Name"),
                new ExcelByColumnIndex(2, "Gender", "Gender"),
                new ExcelByColumnIndex(3, "DateOfBirth", "Date of Birth"),
                new ExcelByColumnIndex(4, "Height", "Height"),
                new ExcelByColumnIndex(5, "Weight", "Weight")
            };

            // Act
            var spreadsheetStream = ObjectToExcelConverter.FromObjectCollection(columnsMapping, employees);

            // Assert
            var spreadsheet = new ExcelPackage(spreadsheetStream).Workbook.Worksheets.FirstOrDefault();

            foreach(var map in columnsMapping)
            {
                Assert.Equal(map.ColumnHeader, spreadsheet.Cells[1, map.ColumnIndex].Value);
            }
            
            var line = 2;
            foreach (var employee in employees)
            {
                Assert.Equal(employee.Name, spreadsheet.Cells[line, columnsMapping.Where(c => c.AttributeName == "Name").First().ColumnIndex].Value);
                Assert.Equal((int)employee.Gender, Convert.ToInt32(spreadsheet.Cells[line, columnsMapping.Where(c => c.AttributeName == "Gender").First().ColumnIndex].Value));
                Assert.Equal(employee.Height, Convert.ToDecimal(spreadsheet.Cells[line, columnsMapping.Where(c => c.AttributeName == "Height").First().ColumnIndex].Value));
                Assert.Equal(employee.Weight, Convert.ToDecimal(spreadsheet.Cells[line, columnsMapping.Where(c => c.AttributeName == "Weight").First().ColumnIndex].Value));
                Assert.Equal(employee.DateOfBirth.ToString("yyyy-MM-dd HH:mm:ss"), spreadsheet.Cells[line, columnsMapping.Where(c => c.AttributeName == "DateOfBirth").First().ColumnIndex].Value);

                line += 1;
            }
        }

        [Fact]
        public void ObjectList_GetSpreadSheetFile_Sucess()
        {
            // Arrange
            var employees = new Faker<Employee>("en")
                 .CustomInstantiator(f => new Employee(
                     f.Name.FullName(Name.Gender.Male),
                     Gender.Male,
                     f.Date.Past(60, DateTime.Now.AddYears(-16)),
                     decimal.Round(random.Number(5, 7) + random.Decimal(0, 1), 2),
                     decimal.Round(random.Number(160, 500) + random.Decimal(0, 1), 2)
                     )).Generate(100);

            var columnsMapping = new List<ExcelByColumnIndex> {
                new ExcelByColumnIndex(1, "Name", "Name"),
                new ExcelByColumnIndex(2, "Gender", "Gender"),
                new ExcelByColumnIndex(3, "DateOfBirth", "Date of Birth"),
                new ExcelByColumnIndex(4, "Height", "Height"),
                new ExcelByColumnIndex(5, "Weight", "Weight")
            };

            // Act
            var spreadsheetFilePath = Path.Combine(Directory.GetParent(Path.GetDirectoryName(Environment.CurrentDirectory)).Parent.ToString(), "TestFiles\\Employees.xlsx");
            ObjectToExcelConverter.CreateFileFromObjectCollection(columnsMapping, employees, spreadsheetFilePath);            
        }
    }
}

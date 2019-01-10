using EasyExcel.Exceptions;
using EasyExcel.MappingModels.Object;
using EasyExcel.Tests.TestModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace EasyExcel.Tests
{
    public class ExcelToObjectConverterTests
    {
        [Fact]
        public void ExcelSpreadSheetFile_GetObjectCollection_Sucess()
        {
            // Arrange
            var spreadsheetFilePath = Path.Combine(Directory.GetParent(Path.GetDirectoryName(Environment.CurrentDirectory)).Parent.ToString(), "TestFiles\\Employees.xlsx");

            var columnsMapping = new List<ObjectByColumnIndex> {
                new ObjectByColumnIndex(1, "Name", true),
                new ObjectByColumnIndex(2, "Gender", true),
                new ObjectByColumnIndex(3, "DateOfBirth", true),
                new ObjectByColumnIndex(4, "Height", true),
                new ObjectByColumnIndex(5, "Weight", true)
            };

            // Act
            var employees = ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetFilePath, columnsMapping);

            // Assert
            ExcelWorksheet spreadsheet = null;
            using (var file = File.OpenRead(spreadsheetFilePath))
            {
                spreadsheet = new ExcelPackage(file).Workbook.Worksheets.FirstOrDefault();
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
        public void ExcelSpreadSheetStream_GetObjectCollection_Sucess()
        {
            // Arrange
            var spreadsheetFilePath = Path.Combine(Directory.GetParent(Path.GetDirectoryName(Environment.CurrentDirectory)).Parent.ToString(), "TestFiles\\Employees.xlsx");

            var spreadsheetStream = File.OpenRead(spreadsheetFilePath);

            var columnsMapping = new List<ObjectByColumnIndex> {
                new ObjectByColumnIndex(1, "Name", true),
                new ObjectByColumnIndex(2, "Gender", true),
                new ObjectByColumnIndex(3, "DateOfBirth", true),
                new ObjectByColumnIndex(4, "Height", true),
                new ObjectByColumnIndex(5, "Weight", true)
            };

            // Act
            var employees = ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetStream, columnsMapping);

            // Assert
            var spreadsheet = new ExcelPackage(spreadsheetStream).Workbook.Worksheets.FirstOrDefault();
            spreadsheetStream.Dispose();

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
        public void ExcelSpreadSheetFile_GetObjectCollection_SpreadsheetValueConversionException()
        {
            // Arrange
            var spreadsheetFilePath = Path.Combine(Directory.GetParent(Path.GetDirectoryName(Environment.CurrentDirectory)).Parent.ToString(), "TestFiles\\EmployeesConversionException.xlsx");

            var columnsMapping = new List<ObjectByColumnIndex> {
                new ObjectByColumnIndex(1, "Name", true),
                new ObjectByColumnIndex(2, "Gender", true),
                new ObjectByColumnIndex(3, "DateOfBirth", true),
                new ObjectByColumnIndex(4, "Height", true),
                new ObjectByColumnIndex(5, "Weight", true)
            };

            // Assert
            var ex = Assert.Throws<SpreadsheetValueConversionException>(() => ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetFilePath, columnsMapping));
            Assert.Equal("It was not possible to convert value: 'Very Tall' to attribute Height(System.Decimal) at line 101 and column '4' of spreadsheet.", ex.Message);
        }

        [Fact]
        public void ExcelSpreadSheetFile_GetObjectCollection_SpreadsheetEmptyRequiredFieldException()
        {
            // Arrange
            var spreadsheetFilePath = Path.Combine(Directory.GetParent(Path.GetDirectoryName(Environment.CurrentDirectory)).Parent.ToString(), "TestFiles\\EmployeesRequiredFieldException.xlsx");

            var columnsMapping = new List<ObjectByColumnIndex> {
                new ObjectByColumnIndex(1, "Name", true),
                new ObjectByColumnIndex(2, "Gender", true),
                new ObjectByColumnIndex(3, "DateOfBirth", true),
                new ObjectByColumnIndex(4, "Height", true),
                new ObjectByColumnIndex(5, "Weight", true)
            };

            // Assert
            var ex = Assert.Throws<SpreadsheetEmptyRequiredFieldException>(() => ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetFilePath, columnsMapping));
            Assert.Equal("The required field Gender is empty at line 43 of spreadsheet.", ex.Message);
        }
    }
}

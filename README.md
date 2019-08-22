# What's EasyExcel?
> EasyExcel is a simple but very powerful tool that easily converts objects to excel and excel to objects.

## What exactly you can do using EasyExcel?
- Convert an excel file into a list of objects;
- Convert a list object into an excel file;

The artfacts below will be used on next examples:


### Excel Spreadsheet

```
| Name                  | Gender | Date of Birth | Height | Weight |
| --------------------- | ------ | ------------- | ------ | ------ | 
| Adam Braun            | 0      | 1952-05-02    | 7.41   | 162.81 |
| Allen Swaniawski      | 0      | 1974-03-24    | 6.62   | 387.32 |
| Guy Tromp             | 0      | 1992-02-18    | 7.45   | 161.84 |
| Edgar Haag            | 0      | 1978-04-22    | 7.55   | 398.99 |
```


### Class (Object)

```
    public class Employee
    {
        public string Name { get; set; }
        public Gender Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        public decimal Height { get; set; }
        public decimal Weight { get; set; }

        public Employee() { }

        public Employee(
            string name,
            Gender gender,
            DateTime dateOfBirth,
            decimal height,
            decimal weight)
        {
            Name = name;
            Gender = gender;
            DateOfBirth = dateOfBirth;
            Height = height;
            Weight = weight;
        }
    }

    public enum Gender
    {
        Male,
        Female
    }

```

## Converting a list object into an excel file

1 - Create an ExcelByColumnIndex object list that maps the object attributes to excel column, as shown below

```
var columnsMapping = new List<ExcelByColumnIndex> {
    new ExcelByColumnIndex(1, "Name", "Name"),
    new ExcelByColumnIndex(2, "Gender", "Gender"),
    new ExcelByColumnIndex(3, "DateOfBirth", "Date of Birth"),
    new ExcelByColumnIndex(4, "Height", "Height"),
    new ExcelByColumnIndex(5, "Weight", "Weight")
};
```
> You can you ExcelByColumnLetter instead of ExcelByColumnIndex in order to inform A, B, C ... to indicate the columns

> The columns mapping object list will be responsible to tell which attributes go in wich column and also tells what's the header text


2 - Convert object list to file
```
var employees = new List<Employee> {
    new Employee("Adam Braun", Gender.Male, DateTime.Parse("1952-05-02"),  7.41m, 162.81m),
    new Employee("Allen Swaniawski", Gender.Male, DateTime.Parse("1974-03-24"),  6.62m, 387.32m),
    new Employee("Guy Tromp", Gender.Male, DateTime.Parse("1992-02-18"), 7.45m, 161.84m),
    new Employee("Edgar Haag", Gender.Male, DateTime.Parse("1978-04-22"), 7.55m, 398.99m),
    ...
};

var spreadsheetFilePath = "<your excel file path>";
ObjectToExcelConverter.CreateFileFromObjectCollection(columnsMapping, employees, spreadsheetFilePath);  
```` 

## Converting excel files into a list of objects
1 - Create an ObjectByColumnIndex object list that maps the excel column to object attributes, as shown below

```
var columnsMapping = new List<ObjectByColumnIndex> {
    new ObjectByColumnIndex(1, "Name", true),
    new ObjectByColumnIndex(2, "Gender", true),
    new ObjectByColumnIndex(3, "DateOfBirth", true),
    new ObjectByColumnIndex(4, "Height", true),
    new ObjectByColumnIndex(5, "Weight", true)
};
```
> You can you ObjectByColumnLetter instead of ObjectByColumnIndex in order to inform A, B, C ... to indicate the columns

> The columns mapping object list will be responsible to tell in which column is the data for each attribute and also tells if that value is required.

2 - Convert excel to object list
2.1 - From a excel file path
```
var spreadsheetFilePath = "<your excel file path>";
var employees = ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetFilePath, columnsMapping);
```

2.2 - From a excel file stream
```
var spreadsheetFilePath = "<your excel file path>";
var spreadsheetStream = File.OpenRead(spreadsheetFilePath);
var employees = ExcelToObjectConverter.ToObjectCollection<Employee>(spreadsheetStream, columnsMapping);
```` 


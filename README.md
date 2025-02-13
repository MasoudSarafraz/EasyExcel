# EasyExcel - A Simple Excel Utility for C#

**EasyExcel** is a lightweight utility for reading from and writing to Excel files in C#. It is built using only the built-in features of .NET Core(OpenXML), without relying on any external libraries or packages. This makes it a great choice for projects where you want to avoid additional dependencies.

## Features

- **Read Excel Files**: Convert Excel files into a list of strongly-typed objects.
- **Write Excel Files**: Export a list of objects or a `DataTable` to an Excel file.
- **No External Dependencies**: Uses only the built-in features of .NET Core.
- **Flexible**: Supports multiple sheets and custom column names via attributes.
- **Lightweight**: Minimal overhead and easy to integrate into existing projects.

## Installation

Since **EasyExcel** is a single-file utility, you can simply copy the `EasyExcel.cs` file into your project and start using it.

## Usage

### Reading from Excel

To read data from an Excel file into a list of objects, use the `ReadExcelFile` method:

```csharp
[ExcelSheetName(SheetName ="Sheet1")]
public class Person
{
[ExcelColumnName(ColumnName = "CustomerName")]
    public string Name { get; set; }
[ExcelColumnName(ColumnName = "CustomerAge")]
    public int Age { get; set; }
[ExcelColumnName(ColumnName = "BirthDate")]
    public DateTime BirthDate { get; set; }
}

using (var stream = new FileStream("path_to_excel_file.xlsx", FileMode.Open))
{
    var people = EasyExcel.ReadExcelFile<Person>(stream);
    foreach (var person in people)
    {
        Console.WriteLine($"Name: {person.Name}, Age: {person.Age}, BirthDate: {person.BirthDate}");
    }
}
```
### Writing to Excel

To write a list of objects to an Excel file, use the ExportToExcel method:

```csharp
public class Person
{
[ExcelExport(DisplayName = "Person Name")]
    public string Name { get; set; }
[ExcelExport(DisplayName = "Person Age")]
    public int Age { get; set; }
    public DateTime BirthDate { get; set; }
}
var people = new List<Person>
{
    new Person { Name = "John Doe", Age = 30, BirthDate = new DateTime(1990, 1, 1) },
    new Person { Name = "Jane Doe", Age = 25, BirthDate = new DateTime(1995, 5, 5) }
};

var excelBytes = EasyExcel.ExportToExcel(people, "People");
File.WriteAllBytes("output.xlsx", excelBytes);
```
### Customizing Column Names and Sheet Names

You can customize the column names and sheet names using attributes:

```csharp
[ExcelSheetName("Employees")]
public class Employee
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelColumnName("Age in Years")]
    public int Age { get; set; }

    [ExcelColumnName("Date of Birth")]
    public DateTime BirthDate { get; set; }
}
```
### Customizing Column Names and Sheet Names

You can customize the column names and sheet names using attributes:
```csharp
[ExcelSheetName("Employees")]
public class Employee
{
    [ExcelColumnName("Full Name")]
    public string Name { get; set; }

    [ExcelColumnName("Age in Years")]
    public int Age { get; set; }

    [ExcelColumnName("Date of Birth")]
    public DateTime BirthDate { get; set; }
}
```

### Handling Multiple Sheets
You can read data from multiple sheets in a single Excel file:
```csharp
using (var stream = new FileStream("path_to_excel_file.xlsx", FileMode.Open))
{
    var (people, employees) = EasyExcel.ReadExcelFile<Person, Employee>(stream);
}
```
Attributes
----------

*   **ExcelColumnNameAttribute**: Specifies the column name in the Excel file.
    
*   **ExcelSheetNameAttribute**: Specifies the sheet name in the Excel file.
    
*   **ExcelExportAttribute**: Specifies whether a property should be included in the export and its display name.
    

Limitations
-----------

*   **Excel Format**: Only supports .xlsx files (Excel 2007 and later).
    
*   **Data Types**: Supports basic data types like int, double, decimal, bool, DateTime, and string.

Contributing
------------

Contributions are welcome! If you find a bug or have a feature request, please open an issue or submit a pull request.

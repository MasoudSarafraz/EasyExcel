using System;

namespace EasyExcelTools
{
    [System.AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnNameAttribute : System.Attribute
    {        
        public string ColumnName { get; set; }
        public ExcelColumnNameAttribute(string columnName) => ColumnName = columnName;

    }

    [System.AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetNameAttribute : System.Attribute
    {
        public string SheetName { get; set; }
        public ExcelSheetNameAttribute(string sheetName) => SheetName = sheetName;

    }
}

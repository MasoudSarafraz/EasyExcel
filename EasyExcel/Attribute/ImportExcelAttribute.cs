using System;

namespace EasyExcelTools
{
    [System.AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnNameAttribute : System.Attribute
    {        
        public string ColumnName { get; set; }
    }

    [System.AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetNameAttribute : System.Attribute
    {
        public string SheetName { get; set; }
    }
}

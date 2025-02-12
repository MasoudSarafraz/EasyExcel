using System;

namespace EasyExcelTools
{
    [System.AttributeUsage(AttributeTargets.Property)]
    public class ExcelExportAttribute : System.Attribute
    {        
        public string DisplayName { get; set; }
        public int ColumnOrder { get; set; }
    }
}

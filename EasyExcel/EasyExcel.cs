using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace EasyExcelTools
{
    public static class EasyExcel
    {
        public static List<T> ReadExcelFile<T>(Stream stream) where T : new()
        {
            var result = new List<T>();
            var properties = typeof(T).GetProperties().ToDictionary(p => GetExcelColumnName(p), p => p);
            using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart == null) { return result; }
                var sheetName = GetExcelSheetName<T>();
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name?.Value.Trim() == sheetName);
                if (sheet == null) { return result; }
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                if (sheetData == null) { return result; }
                var headers = sheetData.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Select(c => GetCellValue(c, workbookPart).Trim()).ToList();
                if (headers == null || !headers.Any()) { return result; }
                foreach (var row in sheetData.Elements<Row>().Skip(1))
                {
                    var item = new T();
                    var cells = row.Elements<Cell>().ToList();
                    for (int i = 0; i < headers.Count && i < cells.Count; i++)
                    {
                        var header = headers[i];
                        var cellValue = GetCellValue(cells[i], workbookPart);
                        if (properties.ContainsKey(header))
                        {
                            var property = properties[header];
                            try
                            {
                                var convertedValue = ConvertValue(cellValue, property.PropertyType);
                                property.SetValue(item, convertedValue);
                            }
                            catch { /* Ignore conversion errors */ }
                        }
                    }
                    result.Add(item);
                }
            }
            return result;
        }
        public static (List<T1>, List<T2>) ReadExcelFile<T1, T2>(Stream stream) where T1 : new() where T2 : new()
        {
            return new ValueTuple<List<T1>, List<T2>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream));
        }
        public static (List<T1>, List<T2>, List<T3>) ReadExcelFile<T1, T2, T3>(Stream stream) where T1 : new() where T2 : new() where T3 : new()
        {
            return new ValueTuple<List<T1>, List<T2>, List<T3>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream));
        }
        public static (List<T1>, List<T2>, List<T3>, List<T4>) ReadExcelFile<T1, T2, T3, T4>(Stream stream) where T1 : new() where T2 : new() where T3 : new() where T4 : new()
        {
            return new ValueTuple<List<T1>, List<T2>, List<T3>, List<T4>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream), ReadExcelFile<T4>(stream));
        }
        public static (List<T1>, List<T2>, List<T3>, List<T4>, List<T5>) ReadExcelFile<T1, T2, T3, T4, T5>(Stream stream) where T1 : new() where T2 : new() where T3 : new() where T4 : new() where T5 : new()
        {
            return new ValueTuple<List<T1>, List<T2>, List<T3>, List<T4>, List<T5>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream), ReadExcelFile<T4>(stream), ReadExcelFile<T5>(stream));
        }
        public static byte[] ExportToExcel<T>(IEnumerable<T> data, string sheetName = "Sheet1") where T : new()
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName });
                    WriteDataTableToWorksheet(worksheetPart.Worksheet, ToFilteredDataTable(data));
                    EnsureWorkbookStylesPart(workbookPart);
                    workbookPart.Workbook.Save();
                }
                return memoryStream.ToArray();
            }
        }
        public static byte[] ExportToExcel<T>(DataTable datatable, string sheetName = "Sheet1") where T : new()
        {
            IEnumerable<T> data = ConvertDataTableToIEnumerable<T>(datatable);
            return ExportToExcel(data, sheetName);            
        }
        private static IEnumerable<T> ConvertDataTableToIEnumerable<T>(DataTable dataTable) where T : new()
        {
            foreach (DataRow row in dataTable.Rows)
            {
                T item = new T();
                foreach (PropertyInfo prop in typeof(T).GetProperties())
                {
                    if (dataTable.Columns.Contains(prop.Name) && row[prop.Name] != DBNull.Value)
                    {
                        object value = Convert.ChangeType(row[prop.Name], prop.PropertyType);
                        prop.SetValue(item, value, null);
                    }
                }
                yield return item;
            }
        }
        private static DataTable ToFilteredDataTable<T>(IEnumerable<T> data)
        {
            var dataTable = new DataTable();
            var properties = typeof(T).GetProperties().Where(p => p.GetCustomAttribute<ExcelExportAttribute>() != null).OrderBy(p => GetColumnOrder(p)).ToList();
            foreach (var property in properties)
            {
                var attribute = property.GetCustomAttribute<ExcelExportAttribute>();
                var columnName = attribute?.DisplayName ?? property.Name;
                dataTable.Columns.Add(columnName, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }
            foreach (var item in data)
            {
                var row = new object[properties.Count];
                for (int i = 0; i < properties.Count; i++)
                {
                    row[i] = properties[i].GetValue(item) ?? DBNull.Value;
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        private static int GetColumnOrder(PropertyInfo property)
        {
            var attribute = property.GetCustomAttribute<ExcelExportAttribute>();
            return attribute?.ColumnOrder ?? int.MaxValue;
        }
        private static DataTable ToDataTable<T>(IEnumerable<T> data)
        {
            var dataTable = new DataTable();
            var properties = typeof(T).GetProperties();
            foreach (var property in properties)
            {
                dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }
            foreach (var item in data)
            {
                var row = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    row[i] = properties[i].GetValue(item) ?? DBNull.Value;
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }
        private static void WriteDataTableToWorksheet(Worksheet worksheet, DataTable dataTable)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
            int rowIndex = 1;
            int colIndex = 0;
            var headerRow = new Row { RowIndex = (uint)rowIndex++ };
            foreach (var column in dataTable.Columns.Cast<DataColumn>())
            {
                var cell = CreateTextCell(GetCellReference(colIndex++, rowIndex - 1), column.ColumnName);
                headerRow.Append(cell);
            }
            sheetData.AppendChild(headerRow);
            foreach (DataRow row in dataTable.Rows)
            {
                var dataRow = new Row { RowIndex = (uint)rowIndex++ };
                colIndex = 0;
                foreach (var item in row.ItemArray)
                {
                    var cell = CreateTypedCell(GetCellReference(colIndex++, rowIndex - 1), item);
                    dataRow.Append(cell);
                }
                sheetData.AppendChild(dataRow);
            }
        }
        private static Cell CreateTextCell(string cellReference, string value)
        {
            return new Cell { CellReference = cellReference, DataType = CellValues.String, CellValue = new CellValue(value) };
        }
        private static Cell CreateTypedCell(string cellReference, object value)
        {
            if (value == null) { return new Cell { CellReference = cellReference }; }
            var cell = new Cell { CellReference = cellReference };
            if (value is int || value is double || value is decimal)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(value.ToString());
            }
            else if (value is DateTime dateTime)
            {
                cell.DataType = CellValues.Date;
                cell.CellValue = new CellValue(dateTime.ToString("o"));
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(value.ToString());
            }
            return cell;
        }
        private static string GetCellReference(int columnIndex, int rowIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var columnLetter = string.Empty;
            while (columnIndex >= 0)
            {
                var remainder = columnIndex % 26;
                columnLetter = letters[remainder] + columnLetter;
                columnIndex = (columnIndex / 26) - 1;
            }
            return $"{columnLetter}{rowIndex}";
        }
        private static void EnsureWorkbookStylesPart(WorkbookPart workbookPart)
        {
            if (workbookPart.WorkbookStylesPart == null)
            {
                workbookPart.AddNewPart<WorkbookStylesPart>();
            }
            var stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart.Stylesheet == null)
            {
                stylesPart.Stylesheet = new Stylesheet();
            }
            var fonts = new Fonts(new Font());
            fonts.Count = (uint)fonts.ChildElements.Count;
            stylesPart.Stylesheet.Append(fonts);
            var fills = new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            fills.Count = (uint)fills.ChildElements.Count;
            stylesPart.Stylesheet.Append(fills);
            var borders = new Borders(new Border());
            borders.Count = (uint)borders.ChildElements.Count;
            stylesPart.Stylesheet.Append(borders);
            var cellFormats = new CellFormats(new CellFormat(), new CellFormat { FormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;
            stylesPart.Stylesheet.Append(cellFormats);
            stylesPart.Stylesheet.Save();
        }
        private static CellValues GetCellDataType(object value)
        {
            if (value == null || value is string) { return CellValues.String; }
            if (value is int || value is double || value is decimal) { return CellValues.Number; }
            if (value is DateTime) { return CellValues.Date; }
            return CellValues.String;
        }
        private static string GetExcelColumnName(PropertyInfo property)
        {
            var attribute = property.GetCustomAttribute<ExcelColumnNameAttribute>();
            return attribute?.ColumnName ?? property.Name;
        }
        private static string GetExcelSheetName<T>()
        {
            var attribute = typeof(T).GetCustomAttribute<ExcelSheetNameAttribute>();
            return attribute?.SheetName ?? "Sheet1";
        }
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null) { return string.Empty; }
            var value = cell.CellValue?.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTable != null && int.TryParse(value, out int index))
                {
                    value = sharedStringTable.SharedStringTable.ElementAt(index)?.InnerText;
                }
            }
            return value ?? string.Empty;
        }
        private static object ConvertValue(string value, Type type)
        {
            if (string.IsNullOrEmpty(value)) { return null; }
            if (type == typeof(int)) { return int.Parse(value); }
            if (type == typeof(double)) { return double.Parse(value); }
            if (type == typeof(decimal)) { return decimal.Parse(value); }
            if (type == typeof(bool)) { return bool.Parse(value); }
            if (type == typeof(DateTime)) { return DateTime.Parse(value); }
            return value;
        }
        //public static class EasyExcel
        //{
        //    public static List<T> ReadExcelFile<T>(Stream stream) where T : new()
        //    {
        //        var result = new List<T>();
        //        var properties = typeof(T).GetProperties()
        //            .ToDictionary(p => GetExcelColumnName(p), p => p);

        //        using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
        //        {
        //            var workbookPart = spreadsheetDocument.WorkbookPart;
        //            if (workbookPart == null) return result;

        //            var sheetName = GetExcelSheetName<T>();
        //            var sheet = workbookPart.Workbook.Descendants<Sheet>()
        //                .FirstOrDefault(s => s.Name?.Value == sheetName);
        //            if (sheet == null) return result;

        //            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        //            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
        //            if (sheetData == null) return result;

        //            var headers = sheetData.Elements<Row>().FirstOrDefault()?.Elements<Cell>()
        //                .Select(c => GetCellValue(c, workbookPart)).ToList();

        //            if (headers == null || !headers.Any()) return result;

        //            foreach (var row in sheetData.Elements<Row>().Skip(1)) // رد سربرگ
        //            {
        //                var item = new T();
        //                var cells = row.Elements<Cell>().ToList();

        //                for (int i = 0; i < headers.Count && i < cells.Count; i++)
        //                {
        //                    var header = headers[i];
        //                    var cellValue = GetCellValue(cells[i], workbookPart);

        //                    if (properties.ContainsKey(header))
        //                    {
        //                        var property = properties[header];
        //                        try
        //                        {
        //                            var convertedValue = ConvertValue(cellValue, property.PropertyType);
        //                            property.SetValue(item, convertedValue);
        //                        }
        //                        catch { /* Ignore conversion errors */ }
        //                    }
        //                }

        //                result.Add(item);
        //            }
        //        }

        //        return result;
        //    }
        //    public static (List<T1>, List<T2>) ReadExcelFile<T1, T2>(Stream stream)
        //        where T1 : new()
        //        where T2 : new()
        //    {
        //        return new ValueTuple<List<T1>, List<T2>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream));
        //    }
        //    public static (List<T1>, List<T2>, List<T3>) ReadExcelFile<T1, T2, T3>(Stream stream)
        //        where T1 : new()
        //        where T2 : new()
        //        where T3 : new()
        //    {
        //        return new ValueTuple<List<T1>, List<T2>, List<T3>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream));
        //    }
        //    public static (List<T1>, List<T2>, List<T3>, List<T4>) ReadExcelFile<T1, T2, T3, T4>(Stream stream)
        //        where T1 : new()
        //        where T2 : new()
        //        where T3 : new()
        //        where T4 : new()
        //    {
        //        return new ValueTuple<List<T1>, List<T2>, List<T3>, List<T4>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream), ReadExcelFile<T4>(stream));
        //    }
        //    public static (List<T1>, List<T2>, List<T3>, List<T4>, List<T5>) ReadExcelFile<T1, T2, T3, T4, T5>(Stream stream)
        //        where T1 : new()
        //        where T2 : new()
        //        where T3 : new()
        //        where T4 : new()
        //        where T5 : new()
        //    {
        //        return new ValueTuple<List<T1>, List<T2>, List<T3>, List<T4>, List<T5>>(ReadExcelFile<T1>(stream), ReadExcelFile<T2>(stream), ReadExcelFile<T3>(stream), ReadExcelFile<T4>(stream), ReadExcelFile<T5>(stream));
        //    }
        //    public static byte[] ExportToExcel<T>(IEnumerable<T> data, string sheetName = "Sheet1")
        //    {
        //        using (var memoryStream = new MemoryStream())
        //        {
        //            using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        //            {
        //                var workbookPart = spreadsheetDocument.AddWorkbookPart();
        //                workbookPart.Workbook = new Workbook();

        //                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        //                worksheetPart.Worksheet = new Worksheet(new SheetData());

        //                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        //                sheets.Append(new Sheet()
        //                {
        //                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
        //                    SheetId = 1,
        //                    Name = sheetName
        //                });
        //                WriteDataTableToWorksheet(worksheetPart.Worksheet, ToFilteredDataTable(data));
        //                EnsureWorkbookStylesPart(workbookPart);
        //                workbookPart.Workbook.Save();
        //            }
        //            return memoryStream.ToArray();
        //        }
        //    }
        //    private static DataTable ToFilteredDataTable<T>(IEnumerable<T> data)
        //    {
        //        var dataTable = new DataTable();
        //        var properties = typeof(T).GetProperties()
        //            .Where(p => p.GetCustomAttribute<ExcelExportAttribute>() != null) // فیلتر ویژگی‌ها
        //            .OrderBy(p => GetColumnOrder(p)) // مرتب‌سازی بر اساس ColumnOrder
        //            .ToList();
        //        foreach (var property in properties)
        //        {
        //            var attribute = property.GetCustomAttribute<ExcelExportAttribute>();
        //            var columnName = attribute?.DisplayName ?? property.Name; // اگر DisplayName وجود ندارد، نام ویژگی استفاده می‌شود
        //            dataTable.Columns.Add(columnName, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
        //        }
        //        foreach (var item in data)
        //        {
        //            var row = new object[properties.Count];
        //            for (int i = 0; i < properties.Count; i++)
        //            {
        //                row[i] = properties[i].GetValue(item) ?? DBNull.Value;
        //            }
        //            dataTable.Rows.Add(row);
        //        }
        //        return dataTable;
        //    }
        //    private static int GetColumnOrder(PropertyInfo property)
        //    {
        //        var attribute = property.GetCustomAttribute<ExcelExportAttribute>();
        //        return attribute?.ColumnOrder ?? int.MaxValue; // اگر ColumnOrder وجود ندارد، آخرین ستون در نظر گرفته می‌شود
        //    }
        //    private static DataTable ToDataTable<T>(IEnumerable<T> data)
        //    {
        //        var dataTable = new DataTable();
        //        var properties = typeof(T).GetProperties();

        //        foreach (var property in properties)
        //        {
        //            dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
        //        }

        //        foreach (var item in data)
        //        {
        //            var row = new object[properties.Length];
        //            for (int i = 0; i < properties.Length; i++)
        //            {
        //                row[i] = properties[i].GetValue(item) ?? DBNull.Value;
        //            }
        //            dataTable.Rows.Add(row);
        //        }

        //        return dataTable;
        //    }
        //    private static void WriteDataTableToWorksheet(Worksheet worksheet, DataTable dataTable)
        //    {
        //        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

        //        int rowIndex = 1;
        //        int colIndex = 0;

        //        // نوشتن سربرگ‌ها
        //        var headerRow = new Row { RowIndex = (uint)rowIndex++ };
        //        foreach (var column in dataTable.Columns.Cast<DataColumn>())
        //        {
        //            var cell = CreateTextCell(GetCellReference(colIndex++, rowIndex - 1), column.ColumnName);
        //            headerRow.Append(cell);
        //        }
        //        sheetData.AppendChild(headerRow);

        //        // نوشتن داده‌ها
        //        foreach (DataRow row in dataTable.Rows)
        //        {
        //            var dataRow = new Row { RowIndex = (uint)rowIndex++ };
        //            colIndex = 0;

        //            foreach (var item in row.ItemArray)
        //            {
        //                var cell = CreateTypedCell(GetCellReference(colIndex++, rowIndex - 1), item);
        //                dataRow.Append(cell);
        //            }

        //            sheetData.AppendChild(dataRow);
        //        }
        //    }
        //    private static Cell CreateTextCell(string cellReference, string value)
        //    {
        //        return new Cell
        //        {
        //            CellReference = cellReference,
        //            DataType = CellValues.String,
        //            CellValue = new CellValue(value)
        //        };
        //    }
        //    private static Cell CreateTypedCell(string cellReference, object value)
        //    {
        //        if (value == null) return new Cell { CellReference = cellReference };

        //        var cell = new Cell { CellReference = cellReference };
        //        if (value is int || value is double || value is decimal)
        //        {
        //            cell.DataType = CellValues.Number;
        //            cell.CellValue = new CellValue(value.ToString());
        //        }
        //        else if (value is DateTime dateTime)
        //        {
        //            cell.DataType = CellValues.Date;
        //            cell.CellValue = new CellValue(dateTime.ToString("o")); // ISO 8601 Format
        //        }
        //        else
        //        {
        //            cell.DataType = CellValues.String;
        //            cell.CellValue = new CellValue(value.ToString());
        //        }

        //        return cell;
        //    }
        //    private static string GetCellReference(int columnIndex, int rowIndex)
        //    {
        //        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        //        var columnLetter = string.Empty;

        //        while (columnIndex >= 0)
        //        {
        //            var remainder = columnIndex % 26;
        //            columnLetter = letters[remainder] + columnLetter;
        //            columnIndex = (columnIndex / 26) - 1;
        //        }

        //        return $"{columnLetter}{rowIndex}";
        //    }
        //    private static void EnsureWorkbookStylesPart(WorkbookPart workbookPart)
        //    {
        //        if (workbookPart.WorkbookStylesPart == null)
        //        {
        //            workbookPart.AddNewPart<WorkbookStylesPart>();
        //        }

        //        var stylesPart = workbookPart.WorkbookStylesPart;
        //        if (stylesPart.Stylesheet == null)
        //        {
        //            stylesPart.Stylesheet = new Stylesheet();
        //        }

        //        // ایجاد بخش Font
        //        var fonts = new Fonts(
        //            new Font() // Font پیش‌فرض
        //        );
        //        fonts.Count = (uint)fonts.ChildElements.Count; // تنظیم تعداد فونت‌ها
        //        stylesPart.Stylesheet.Append(fonts);

        //        // ایجاد بخش Fill
        //        var fills = new Fills(
        //        new Fill(new PatternFill { PatternType = PatternValues.None }), // Fill پیش‌فرض
        //            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }) // Fill جدید
        //        );
        //        fills.Count = (uint)fills.ChildElements.Count; // تنظیم تعداد Fill‌ها
        //        stylesPart.Stylesheet.Append(fills);

        //        // ایجاد بخش Border
        //        var borders = new Borders(
        //            new Border() // Border پیش‌فرض
        //        );
        //        borders.Count = (uint)borders.ChildElements.Count; // تنظیم تعداد Border‌ها
        //        stylesPart.Stylesheet.Append(borders);

        //        // ایجاد بخش CellFormat با فعال کردن RTL
        //        var cellFormats = new CellFormats(
        //            new CellFormat(), // فرمت پیش‌فرض
        //            new CellFormat
        //            {
        //                //TextDirection = new TextDirection { Val = TextDirectionValues.Rtl }, // فعال کردن RTL
        //                FormatId = 0, // ارجاع به فرمت پیش‌فرض
        //                FontId = 0, // ارجاع به Font پیش‌فرض
        //                FillId = 0, // ارجاع به Fill پیش‌فرض
        //                BorderId = 0 // ارجاع به Border پیش‌فرض
        //            }
        //        );
        //        cellFormats.Count = (uint)cellFormats.ChildElements.Count; // تنظیم تعداد CellFormat‌ها
        //        stylesPart.Stylesheet.Append(cellFormats);

        //        stylesPart.Stylesheet.Save();
        //    }
        //    private static CellValues GetCellDataType(object value)
        //    {
        //        if (value == null || value is string)
        //            return CellValues.String;
        //        if (value is int || value is double || value is decimal)
        //            return CellValues.Number;
        //        if (value is DateTime)
        //            return CellValues.Date;
        //        return CellValues.String;
        //    }
        //    private static string GetExcelColumnName(PropertyInfo property)
        //    {
        //        var attribute = property.GetCustomAttribute<ExcelColumnNameAttribute>();
        //        return attribute?.ColumnName ?? property.Name;
        //    }
        //    private static string GetExcelSheetName<T>()
        //    {
        //        var attribute = typeof(T).GetCustomAttribute<ExcelSheetNameAttribute>();
        //        return attribute?.SheetName ?? "Sheet1";
        //    }
        //    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        //    {
        //        if (cell == null) return string.Empty;

        //        var value = cell.CellValue?.Text;
        //        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //        {
        //            var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        //            if (sharedStringTable != null && int.TryParse(value, out int index))
        //            {
        //                value = sharedStringTable.SharedStringTable.ElementAt(index)?.InnerText;
        //            }
        //        }

        //        return value ?? string.Empty;
        //    }
        //    private static object ConvertValue(string value, Type type)
        //    {
        //        if (string.IsNullOrEmpty(value)) return null;

        //        if (type == typeof(int)) return int.Parse(value);
        //        if (type == typeof(double)) return double.Parse(value);
        //        if (type == typeof(decimal)) return decimal.Parse(value);
        //        if (type == typeof(bool)) return bool.Parse(value);
        //        if (type == typeof(DateTime)) return DateTime.Parse(value);

        //        return value;
        //    }
        //}
    }
}
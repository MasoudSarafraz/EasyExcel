using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace EasyExcelTools
{
    public static class EasyExcel
    {
        public static List<T> ReadExcelFile<T>(Stream oStream) where T : new()
        {
            var oResult = new List<T>();
            var oProperties = typeof(T).GetProperties().ToDictionary(p => GetExcelColumnName(p), p => p);
            using (var oMemoryStream = new MemoryStream())
            {
                if (oStream.CanSeek) oStream.Position = 0; else throw new ArgumentException("Stream must be seekable.");
                oStream.CopyTo(oMemoryStream);
                oMemoryStream.Position = 0;
                using (var oSpreadsheetDocument = SpreadsheetDocument.Open(oMemoryStream, false))
                {
                    var oWorkbookPart = oSpreadsheetDocument.WorkbookPart;
                    if (oWorkbookPart == null) return oResult;
                    var sSheetName = GetExcelSheetName<T>();
                    var oSheet = oWorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name?.Value.Trim() == sSheetName);
                    if (oSheet == null) return oResult;
                    var oWorksheetPart = (WorksheetPart)oWorkbookPart.GetPartById(oSheet.Id);
                    var oSheetData = oWorksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    if (oSheetData == null) return oResult;
                    var oSharedStringTable = oWorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;
                    var oHeaders = oSheetData.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Select(c => GetCellValue(c, oSharedStringTable).Trim()).ToList();
                    if (oHeaders == null || !oHeaders.Any()) return oResult;
                    foreach (var oRow in oSheetData.Elements<Row>().Skip(1))
                    {
                        var oItem = new T();
                        var oCells = oRow.Elements<Cell>().ToList();
                        for (int iIndex = 0; iIndex < oHeaders.Count && iIndex < oCells.Count; iIndex++)
                        {
                            var sHeader = oHeaders[iIndex];
                            var sCellValue = GetCellValue(oCells[iIndex], oSharedStringTable);
                            if (oProperties.TryGetValue(sHeader, out var oProperty))
                            {
                                try { var oConvertedValue = ConvertValue(sCellValue, oProperty.PropertyType); oProperty.SetValue(oItem, oConvertedValue); }
                                catch { }
                            }
                        }
                        oResult.Add(oItem);
                    }
                }
            }
            return oResult;
        }

        public static (List<T1>, List<T2>) ReadExcelFile<T1, T2>(Stream oStream) where T1 : new() where T2 : new()
        {
            var oAllSheetsData = ReadAllSheetsData(oStream);
            var oList1 = new List<T1>(); var oList2 = new List<T2>();
            MapDataToType(oAllSheetsData, oList1); MapDataToType(oAllSheetsData, oList2);
            return (oList1, oList2);
        }
        public static (List<T1>, List<T2>, List<T3>) ReadExcelFile<T1, T2, T3>(Stream oStream) where T1 : new() where T2 : new() where T3 : new()
        {
            var oAllSheetsData = ReadAllSheetsData(oStream);
            var oList1 = new List<T1>(); var oList2 = new List<T2>(); var oList3 = new List<T3>();
            MapDataToType(oAllSheetsData, oList1); MapDataToType(oAllSheetsData, oList2); MapDataToType(oAllSheetsData, oList3);
            return (oList1, oList2, oList3);
        }
        public static (List<T1>, List<T2>, List<T3>, List<T4>) ReadExcelFile<T1, T2, T3, T4>(Stream oStream) where T1 : new() where T2 : new() where T3 : new() where T4 : new()
        {
            var oAllSheetsData = ReadAllSheetsData(oStream);
            var oList1 = new List<T1>(); var oList2 = new List<T2>(); var oList3 = new List<T3>(); var oList4 = new List<T4>();
            MapDataToType(oAllSheetsData, oList1); MapDataToType(oAllSheetsData, oList2); MapDataToType(oAllSheetsData, oList3); MapDataToType(oAllSheetsData, oList4);
            return (oList1, oList2, oList3, oList4);
        }
        public static (List<T1>, List<T2>, List<T3>, List<T4>, List<T5>) ReadExcelFile<T1, T2, T3, T4, T5>(Stream oStream) where T1 : new() where T2 : new() where T3 : new() where T4 : new() where T5 : new()
        {
            var oAllSheetsData = ReadAllSheetsData(oStream);
            var oList1 = new List<T1>(); var oList2 = new List<T2>(); var oList3 = new List<T3>(); var oList4 = new List<T4>(); var oList5 = new List<T5>();
            MapDataToType(oAllSheetsData, oList1); MapDataToType(oAllSheetsData, oList2); MapDataToType(oAllSheetsData, oList3); MapDataToType(oAllSheetsData, oList4); MapDataToType(oAllSheetsData, oList5);
            return (oList1, oList2, oList3, oList4, oList5);
        }
        private static Dictionary<string, List<Dictionary<string, string>>> ReadAllSheetsData(Stream oStream)
        {
            var oAllSheetsData = new Dictionary<string, List<Dictionary<string, string>>>();
            using (var oMemoryStream = new MemoryStream())
            {
                if (oStream.CanSeek) { oStream.Position = 0; } else { throw new ArgumentException("Input stream must be seekable.", nameof(oStream)); }
                oStream.CopyTo(oMemoryStream);
                oMemoryStream.Position = 0;
                using (var oSpreadsheetDocument = SpreadsheetDocument.Open(oMemoryStream, false))
                {
                    var oWorkbookPart = oSpreadsheetDocument.WorkbookPart;
                    if (oWorkbookPart == null) return oAllSheetsData;
                    var oSharedStringTable = oWorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;
                    foreach (var oSheet in oWorkbookPart.Workbook.Descendants<Sheet>())
                    {
                        var oWorksheetPart = (WorksheetPart)oWorkbookPart.GetPartById(oSheet.Id);
                        var oSheetData = oWorksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        if (oSheetData == null) continue;
                        var oHeaders = oSheetData.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Select(c => GetCellValue(c, oSharedStringTable).Trim()).ToList();
                        if (oHeaders == null || !oHeaders.Any()) continue;
                        var oSheetRows = new List<Dictionary<string, string>>();
                        foreach (var oRow in oSheetData.Elements<Row>().Skip(1))
                        {
                            var oRowData = new Dictionary<string, string>();
                            var oCells = oRow.Elements<Cell>().ToList();
                            for (int iIndex = 0; iIndex < oHeaders.Count && iIndex < oCells.Count; iIndex++) { oRowData[oHeaders[iIndex]] = GetCellValue(oCells[iIndex], oSharedStringTable); }
                            oSheetRows.Add(oRowData);
                        }
                        oAllSheetsData[oSheet.Name.Value] = oSheetRows;
                    }
                }
            }
            return oAllSheetsData;
        }
        private static void MapDataToType<T>(Dictionary<string, List<Dictionary<string, string>>> oAllSheetsData, List<T> oTargetList) where T : new()
        {
            var sSheetName = GetExcelSheetName<T>();
            if (!oAllSheetsData.TryGetValue(sSheetName, out var oSheetData)) return;
            var oProperties = typeof(T).GetProperties().ToDictionary(p => GetExcelColumnName(p), p => p);
            foreach (var oRowData in oSheetData)
            {
                var oItem = new T();
                foreach (var oKvp in oRowData)
                {
                    if (oProperties.TryGetValue(oKvp.Key, out var oProperty)) { try { var oConvertedValue = ConvertValue(oKvp.Value, oProperty.PropertyType); oProperty.SetValue(oItem, oConvertedValue); } catch { } }
                }
                oTargetList.Add(oItem);
            }
        }
        public static byte[] ExportToExcel<T>(IEnumerable<T> oData, string sSheetName = "Sheet1") where T : new()
        {
            if (sSheetName == "Sheet1") { sSheetName = GetExcelSheetName<T>(); }
            using (var oMemoryStream = new MemoryStream())
            {
                using (var oSpreadsheetDocument = SpreadsheetDocument.Create(oMemoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var oWorkbookPart = oSpreadsheetDocument.AddWorkbookPart(); oWorkbookPart.Workbook = new Workbook();
                    var oWorksheetPart = oWorkbookPart.AddNewPart<WorksheetPart>(); oWorksheetPart.Worksheet = new Worksheet(new SheetData());
                    var oSheets = oSpreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                    oSheets.Append(new Sheet() { Id = oSpreadsheetDocument.WorkbookPart.GetIdOfPart(oWorksheetPart), SheetId = 1, Name = sSheetName });
                    WriteDataTableToWorksheet(oWorksheetPart.Worksheet, ToFilteredDataTable(oData));
                    EnsureWorkbookStylesPart(oWorkbookPart);
                    oWorkbookPart.Workbook.Save();
                }
                return oMemoryStream.ToArray();
            }
        }
        public static byte[] ExportToExcel<T>(DataTable oDatatable, string sSheetName = "Sheet1") where T : new()
        {
            if (sSheetName == "Sheet1") { sSheetName = GetExcelSheetName<T>(); }
            IEnumerable<T> oData = ConvertDataTableToIEnumerable<T>(oDatatable);
            return ExportToExcel(oData, sSheetName);
        }
        private static IEnumerable<T> ConvertDataTableToIEnumerable<T>(DataTable oDatatable) where T : new()
        {
            var oProperties = typeof(T).GetProperties().ToDictionary(p => p.Name, p => p);
            foreach (DataRow oRow in oDatatable.Rows) { T oItem = new T(); foreach (DataColumn oColumn in oDatatable.Columns) if (oProperties.TryGetValue(oColumn.ColumnName, out var oProperty) && oRow[oColumn] != DBNull.Value) { try { object oValue = Convert.ChangeType(oRow[oColumn], oProperty.PropertyType); oProperty.SetValue(oItem, oValue, null); } catch { } } yield return oItem; }
        }
        private static DataTable ToFilteredDataTable<T>(IEnumerable<T> oData)
        {
            var oDataTable = new DataTable();
            var oProperties = typeof(T).GetProperties().Where(p => p.GetCustomAttribute<ExcelExportAttribute>() != null).OrderBy(p => GetColumnOrder(p)).ToList();
            foreach (var oProperty in oProperties)
            {
                var sColumnName = GetExcelColumnName(oProperty);
                oDataTable.Columns.Add(sColumnName, Nullable.GetUnderlyingType(oProperty.PropertyType) ?? oProperty.PropertyType);
            }
            foreach (var oItem in oData) { var oRow = new object[oProperties.Count]; for (int iIndex = 0; iIndex < oProperties.Count; iIndex++) { oRow[iIndex] = oProperties[iIndex].GetValue(oItem) ?? DBNull.Value; } oDataTable.Rows.Add(oRow); }
            return oDataTable;
        }
        private static int GetColumnOrder(PropertyInfo oProperty) { var oAttribute = oProperty.GetCustomAttribute<ExcelExportAttribute>(); return oAttribute?.ColumnOrder ?? int.MaxValue; }
        private static void WriteDataTableToWorksheet(Worksheet oWorksheet, DataTable oDataTable)
        {
            var oSheetData = oWorksheet.GetFirstChild<SheetData>() ?? oWorksheet.AppendChild(new SheetData());
            int iRowIndex = 1; int iColIndex = 0;
            var oHeaderRow = new Row { RowIndex = (uint)iRowIndex++ };
            foreach (var oColumn in oDataTable.Columns.Cast<DataColumn>()) { var oCell = CreateTextCell(GetCellReference(iColIndex++, iRowIndex - 1), oColumn.ColumnName); oHeaderRow.Append(oCell); }
            oSheetData.AppendChild(oHeaderRow);
            foreach (DataRow oRow in oDataTable.Rows) { var oDataRow = new Row { RowIndex = (uint)iRowIndex++ }; iColIndex = 0; foreach (var oItem in oRow.ItemArray) { var oCell = CreateTypedCell(GetCellReference(iColIndex++, iRowIndex - 1), oItem); oDataRow.Append(oCell); } oSheetData.AppendChild(oDataRow); }
        }
        private static Cell CreateTextCell(string sCellReference, string sValue) { return new Cell { CellReference = sCellReference, DataType = CellValues.String, CellValue = new CellValue(sValue) }; }
        private static Cell CreateTypedCell(string sCellReference, object oValue)
        {
            if (oValue == null) return new Cell { CellReference = sCellReference };
            var oCell = new Cell { CellReference = sCellReference };
            if (oValue is int || oValue is double || oValue is decimal || oValue is float) { oCell.DataType = CellValues.Number; oCell.CellValue = new CellValue(Convert.ToDouble(oValue).ToString(CultureInfo.InvariantCulture)); }
            else if (oValue is DateTime oDateTime) { oCell.DataType = CellValues.Number; oCell.CellValue = new CellValue(oDateTime.ToOADate().ToString(CultureInfo.InvariantCulture)); oCell.StyleIndex = 1; }
            else if (oValue is bool) { oCell.DataType = CellValues.Boolean; oCell.CellValue = new CellValue((bool)oValue ? "1" : "0"); }
            else { oCell.DataType = CellValues.String; oCell.CellValue = new CellValue(oValue.ToString()); }
            return oCell;
        }
        private static string GetCellReference(int iColumnIndex, int iRowIndex) { const string sLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; var sColumnLetter = string.Empty; while (iColumnIndex >= 0) { var iRemainder = iColumnIndex % 26; sColumnLetter = sLetters[iRemainder] + sColumnLetter; iColumnIndex = (iColumnIndex / 26) - 1; } return $"{sColumnLetter}{iRowIndex}"; }
        private static void EnsureWorkbookStylesPart(WorkbookPart oWorkbookPart)
        {
            if (oWorkbookPart.WorkbookStylesPart == null) oWorkbookPart.AddNewPart<WorkbookStylesPart>();
            var oStylesPart = oWorkbookPart.WorkbookStylesPart;
            if (oStylesPart.Stylesheet == null) oStylesPart.Stylesheet = new Stylesheet();
            var oNumberingFormats = new NumberingFormats(new NumberingFormat { NumberFormatId = 164, FormatCode = "yyyy-mm-dd\\ hh:mm:ss" });
            oNumberingFormats.Count = (uint)oNumberingFormats.ChildElements.Count; oStylesPart.Stylesheet.Append(oNumberingFormats);
            var oFonts = new Fonts(new Font()); oFonts.Count = (uint)oFonts.ChildElements.Count; oStylesPart.Stylesheet.Append(oFonts);
            var oFills = new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })); oFills.Count = (uint)oFills.ChildElements.Count; oStylesPart.Stylesheet.Append(oFills);
            var oBorders = new Borders(new Border()); oBorders.Count = (uint)oBorders.ChildElements.Count; oStylesPart.Stylesheet.Append(oBorders);
            var oCellFormats = new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 164, FontId = 0, FillId = 0, BorderId = 0, ApplyNumberFormat = true });
            oCellFormats.Count = (uint)oCellFormats.ChildElements.Count; oStylesPart.Stylesheet.Append(oCellFormats);
            oStylesPart.Stylesheet.Save();
        }
        private static string GetExcelColumnName(PropertyInfo oProperty) { var oAttribute = oProperty.GetCustomAttribute<ExcelColumnNameAttribute>(); return oAttribute?.ColumnName ?? oProperty.Name; }
        private static string GetExcelSheetName<T>() { var oAttribute = typeof(T).GetCustomAttribute<ExcelSheetNameAttribute>(); return oAttribute?.SheetName ?? "Sheet1"; }
        private static string GetCellValue(Cell oCell, SharedStringTable oSharedStringTable) { if (oCell == null) return string.Empty; var sValue = oCell.CellValue?.Text; if (oCell.DataType != null && oCell.DataType.Value == CellValues.SharedString && oSharedStringTable != null) { if (int.TryParse(sValue, out int iIndex) && iIndex >= 0 && iIndex < oSharedStringTable.Count()) sValue = oSharedStringTable.ElementAt(iIndex).InnerText; } return sValue ?? string.Empty; }
        private static object ConvertValue(string sValue, Type oType)
        {
            if (string.IsNullOrEmpty(sValue)) return null;
            var oCulture = CultureInfo.InvariantCulture;
            if (oType == typeof(int)) return int.Parse(sValue, oCulture);
            if (oType == typeof(double)) return double.Parse(sValue, oCulture);
            if (oType == typeof(decimal)) return decimal.Parse(sValue, oCulture);
            if (oType == typeof(float)) return float.Parse(sValue, oCulture);
            if (oType == typeof(bool)) { if (sValue == "1") return true; if (sValue == "0") return false; return bool.Parse(sValue); }
            if (oType == typeof(DateTime)) return DateTime.FromOADate(double.Parse(sValue, oCulture));
            return sValue;
        }
    }
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Smartmill.Reporting.OpenXml.Data;
using Smartmill.Reporting.OpenXml.Utils;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Smartmill.Reporting.OpenXml
{
    internal class ReportSheetProcessor : DocumentLoader, IReportSheetProcessor
    {
        private static readonly Regex _varDynamicRegex = new Regex(@"^\[(.*?)\]$");
        private static readonly Regex _varStaticRegex = new Regex(@"^_.*?_$");
        private static readonly int _maxRetrievedRows = 30;

        public byte[] Populate(IEnumerable<ReportVariableMapping> values)
        {
            using (var spreadSheet = SpreadsheetDocument.Open(_stream, true, new OpenSettings { AutoSave = true }))
            {
                var shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                var worksheet = spreadSheet.WorkbookPart.WorksheetParts.FirstOrDefault()?.Worksheet;
                var authorizedKeys = values.Select(a => a.Key).Distinct();

                // Process static number / money / string keys
                var rowNumber = 0;
                var cellKeys = GetKeysCell(worksheet, shareStringPart, authorizedKeys);
                FillStaticCells(cellKeys, values, shareStringPart);

                // Process dynamic keys number / money / stirng keys
                rowNumber = 0;
                var keysRow = GetKeysRow(worksheet, shareStringPart, authorizedKeys);
                while (keysRow != null && rowNumber++ < _maxRetrievedRows)
                {
                    FillDynamicSingleKeysRow(worksheet, keysRow, values, shareStringPart, spreadSheet.WorkbookPart);
                    keysRow = GetKeysRow(worksheet, shareStringPart, authorizedKeys);
                }

                spreadSheet.WorkbookPart.Workbook.Save();
                spreadSheet.Save();
            }

            using (var mStream = new MemoryStream())
            {
                _stream.CopyTo(mStream);
                return mStream.ToArray();
            }
        }

        private static void FillStaticCells(IEnumerable<Cell> cellKeys,
           IEnumerable<ReportVariableMapping> values,
           SharedStringTablePart shareStringPart)
        {
            var columns = GetCellsKey(cellKeys, shareStringPart);
            var valuesRows = values?.Where(val => !val.RowNumber.HasValue && columns.Any(a => a.Key.Equals(val.Key)));

            if (valuesRows.IsNullOrEmpty())
            {
                return;
            }

            foreach (var (Key, Reference) in columns)
            {
                var rowValue = valuesRows?.FirstOrDefault(a => a.Key.Equals(Key));
                var cell = cellKeys.FirstOrDefault(a => a.CellReference == Reference);
                SetCellValue(rowValue, cell);
            }
        }

        private static void FillDynamicSingleKeysRow(Worksheet worksheet,
            Row keysRow, IEnumerable<ReportVariableMapping> values,
            SharedStringTablePart shareStringPart, WorkbookPart workbookPart)
        {
            var startIndex = keysRow.RowIndex;
            var columns = GetTableColumnsKey(keysRow, shareStringPart);

            var valuesRows = values?.Where(val => val.RowNumber.HasValue
                                 && columns.Any(a => a.Key.Equals(val.Key)))
               .GroupBy(val => val.RowNumber)
               .Select(valGroup => new
               {
                   Number = valGroup.Key,
                   Row = valGroup?.ToList()
               }).OrderBy(a => a.Number);

            if (valuesRows.IsNullOrEmpty())
            {
                return;
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            worksheet.MoveRowsBelow(startIndex + 1, (uint)valuesRows.Count() - 1);
            var previous = keysRow;
            foreach (var valuesRow in valuesRows)
            {
                var rowClone = (Row)keysRow.CloneNode(true);
                rowClone.UpdateRowIndex(newIndex: startIndex++);

                foreach (var (Key, Reference) in columns)
                {
                    var rowValue = valuesRow.Row?.FirstOrDefault(a => a.Key.Equals(Key));
                    var cellReference = Reference.Replace(keysRow.RowIndex, rowClone.RowIndex);
                    var cell = rowClone.GetCell(cellReference);
                    SetCellValue(rowValue, cell);
                    ApplyStyle(rowValue?.Style, rowClone, cell, sheetData, workbookPart);
                }

                sheetData.InsertAfter(rowClone, previous);
                previous = rowClone;
            }

            sheetData.RemoveChild(keysRow);
        }

        private static void SetCellValue(ReportVariableMapping rowValue, Cell cell)
        {
            if (cell != null)
            {
                if (IsDecimal(rowValue?.Value))
                {
                    cell.CellValue = new CellValue(decimal.Parse(rowValue?.Value));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                }
                else
                {
                    cell.CellValue = new CellValue(rowValue?.Value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
        }

        private static void ApplyStyle(ReportVariableStyle style,
            Row rowClone,
            Cell cell,
            SheetData sheetData,
            WorkbookPart workbookPart)
        {
            if (style == null)
            {
                return;
            }

            if (style.ColSpan > 1)
            {
                var mergeCells = sheetData.Parent?.Elements<MergeCells>()?.FirstOrDefault();
                if (mergeCells is null)
                {
                    mergeCells = new MergeCells();
                    sheetData.Parent?.InsertAfter(mergeCells, sheetData);
                }

                var cells = rowClone.Elements<Cell>()?.ToList();
                var end = cells.ElementAt(cells.IndexOf(cell) + style.ColSpan - 1);
                mergeCells.Append(new MergeCell() { Reference = new StringValue($"{cell.CellReference}:{end.CellReference}") });
            }

            if (style.HasStyleSheet)
            {
                var format = cell.GetCellStyle(workbookPart);
                var font = format.GetFormatFont(workbookPart);
                font.Bold = style.IsBold ? new Bold { Val = true } : font.Bold;
                if (!string.IsNullOrEmpty(style.ForegroundColor))
                {
                    font.RemoveAllChildren<Color>();
                    var color = new Color { Rgb = HexBinaryValue.FromString(style.ForegroundColor) };
                    font.Color = color;
                }
                if (!string.IsNullOrEmpty(style.BackgroundColor))
                {
                    var fill = format.GetFormatFill(workbookPart);
                    format.ApplyFill = true;
                    fill.PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = new ForegroundColor
                        { Rgb = HexBinaryValue.FromString(style.BackgroundColor) }
                    };
                }
            }
        }

        private static Row GetKeysRow(Worksheet worksheet,
            SharedStringTablePart shareStringPart,
            IEnumerable<string> authorizedKeys)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>()
              .FirstOrDefault(r => r.Elements<Cell>()
              .Any(cell =>
              {
                  var shareString = cell.GetValue(shareStringPart).Trim();
                  return _varDynamicRegex.IsMatch(shareString)
                       && authorizedKeys.Contains(shareString.Substring(1, shareString.Length - 1));
              }));
        }

        private static IEnumerable<Cell> GetKeysCell(Worksheet worksheet,
           SharedStringTablePart shareStringPart,
           IEnumerable<string> authorizedKeys)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().SelectMany(r => r.Elements<Cell>())
              .Where(cell =>
              {
                  var shareString = cell.GetValue(shareStringPart).Trim();
                  return _varStaticRegex.IsMatch(shareString)
                             && authorizedKeys.Contains(shareString.Substring(1, shareString.Length - 1));
              });
        }

        private static IEnumerable<(string Key, string Reference)> GetTableColumnsKey(Row keysRow, SharedStringTablePart shareStringPart)
        {
            return GetCellsKey(keysRow?.Descendants<Cell>(), shareStringPart);
        }

        private static IEnumerable<(string Key, string Reference)> GetCellsKey(IEnumerable<Cell> cells, SharedStringTablePart shareStringPart)
        {
            return cells?
                 .Where(cell => !string.IsNullOrEmpty(cell.GetValue(shareStringPart)))
                 .Select(cell =>
                 {
                     var shareString = cell.GetValue(shareStringPart).Trim();
                     return (shareString.Substring(1, shareString.Length - 1), cell.CellReference.Value);
                 });
        }

        private static bool IsDecimal(string str)
        {
            return decimal.TryParse(str, out var _);
        }
    }
}
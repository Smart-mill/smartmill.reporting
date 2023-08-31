using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;

namespace Smartmill.Reporting.OpenXml
{
    internal static class SheetUtils
    {
        internal static string GetValue(this Cell cell, SharedStringTablePart shareStringPart)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString
                && int.TryParse(cell?.InnerText, out var index))
            {
                return shareStringPart.SharedStringTable.ElementAt(index).InnerText;
            }

            return cell?.InnerText;
        }

        internal static Cell GetCell(this Row row, string refrence)
        {
            return row?.Elements<Cell>()?.FirstOrDefault(c => string.Compare
                   (c.CellReference.Value, refrence, true) == 0);
        }

        internal static void UpdateRowIndex(this Row row, UInt32Value newIndex)
        {
            if (row.RowIndex == newIndex)
            {
                return;
            }

            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference.HasValue)
                {
                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(row.RowIndex, newIndex));
                }
            }

            row.RowIndex = newIndex;
        }

        internal static void MoveRowsBelow(this Worksheet worksheet, uint from, uint step)
        {
            if (step <= 0)
            {
                return;
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            var rows = sheetData?.Elements<Row>()?.Where(a => a.RowIndex >= from);
            foreach (var row in rows)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    string cellReference = cell.CellReference.Value;
                    cell.CellReference = new StringValue(cellReference.Replace(row.RowIndex.Value.ToString(), (row.RowIndex + step).ToString()));
                }

                row.RowIndex += step;
            }

            worksheet.UpdateMergedCellReferences(from, step);
        }

        internal static void UpdateMergedCellReferences(this Worksheet worksheet, uint rowIndex, uint step)
        {
            if (worksheet.Elements<MergeCells>().Any())
            {
                var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();

                if (mergeCells != null)
                {
                    // Grab all the merged cells that have a merge cell row index reference equal to or greater than the row index passed in
                    var mergeCellsList = mergeCells.Elements<MergeCell>()
                        .Where(r => r.Reference.HasValue)
                        .Where(r => GetRowIndex(r.Reference.Value.Split(Constants.TwoPoints).ElementAt(0)) >= rowIndex ||
                                      GetRowIndex(r.Reference.Value.Split(Constants.TwoPoints).ElementAt(1)) >= rowIndex).ToList();

                    // Either increment or decrement the row index on the merged cell reference
                    foreach (var mergeCell in mergeCellsList)
                    {
                        string[] cellReference = mergeCell.Reference.Value.Split(Constants.TwoPoints);

                        if (GetRowIndex(cellReference.ElementAt(0)) >= rowIndex)
                        {
                            var columnName = GetColumnName(cellReference.ElementAt(0));
                            cellReference[0] = IncrementCellReference(cellReference.ElementAt(0), step);
                        }

                        if (GetRowIndex(cellReference.ElementAt(1)) >= rowIndex)
                        {
                            var columnName = GetColumnName(cellReference.ElementAt(1));
                            cellReference[1] = IncrementCellReference(cellReference.ElementAt(1), step);
                        }

                        mergeCell.Reference = new StringValue(cellReference[0] + Constants.TwoPoints + cellReference[1]);
                    }
                }
            }
        }

        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            var match = new Regex("[A-Za-z]+").Match(cellName);
            return match.Value;
        }

        public static uint GetRowIndex(string cellReference)
        {
            // Create a regular expression to match the row index portion the cell name.
            var match = new Regex(@"\d+").Match(cellReference);
            return uint.Parse(match.Value);
        }

        public static string IncrementCellReference(string reference, uint step)
        {
            var newReference = reference;

            if (!string.IsNullOrEmpty(reference))
            {
                var parts = Regex.Split(reference, "([A-Z]+)");
                parts[2] = (int.Parse(parts[2]) + step).ToString();
                newReference = parts[1] + parts[2];
            }

            return newReference;
        }

        internal static Font GetFormatFont(this CellFormat format, WorkbookPart workbookPart)
        {
            var styleParts = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            var fonts = styleParts?.Stylesheet.Fonts.Elements<Font>();
            var font = new Font();
            if (int.TryParse(format.FontId, out var index))
            {
                font = (Font)fonts.ElementAt(index).CloneNode(true);
            }

            styleParts.Stylesheet.Fonts.AppendChild(font);
            format.FontId = styleParts.Stylesheet.Fonts.Count;
            styleParts.Stylesheet.Fonts.Count++;
            return font;
        }

        internal static Fill GetFormatFill(this CellFormat format, WorkbookPart workbookPart)
        {
            var styleParts = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            var fills = styleParts.Stylesheet.Fills.Elements<Fill>();
            var fill = new Fill();
            if (int.TryParse(format.FillId, out var index))
            {
                fill = (Fill)fills.ElementAt(index).CloneNode(true);
            }

            styleParts.Stylesheet.Fills.AppendChild(fill);
            format.FillId = styleParts.Stylesheet.Fills.Count;
            styleParts.Stylesheet.Fills.Count++;
            return fill;
        }

        internal static CellFormat GetCellStyle(this Cell cell, WorkbookPart workbookPart)
        {
            var styleParts = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            var cellFormats = styleParts.Stylesheet.CellFormats.Elements<CellFormat>();
            var format = new CellFormat();
            if (int.TryParse(cell?.StyleIndex, out var index))
            {
                format = (CellFormat)cellFormats.ElementAt(index).CloneNode(true);
            }

            styleParts.Stylesheet.CellFormats.AppendChild(format);
            cell.StyleIndex = styleParts.Stylesheet.CellFormats.Count;
            styleParts.Stylesheet.CellFormats.Count++;
            return format;
        }
    }
}
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Smartmill.Reporting.OpenXml.Data;
using Smartmill.Reporting.OpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Smartmill.Reporting.OpenXml
{
    internal class ReportDocumentProcessor : DocumentLoader, IReportDocumentProcessor
    {
        private readonly Regex _varDynamicRegex = new Regex(@"\[((?![\[\]]).*?)\]");
        private readonly Regex _varStaticRegex = new Regex(@"«.*?»");
        private const string MergedField = "MERGEFIELD";

        public ReportDocumentProcessor()
        {
            _varDynamicRegex = new Regex(@"\[((?![\[\]]).*?)\]");
            _varStaticRegex = new Regex(@"«.*?»");
        }

        public byte[] Populate(IEnumerable<ReportVariableMapping> values)
        {
            _ = _stream ?? throw new ArgumentNullException(nameof(_stream));
            using (var wordDocument = WordprocessingDocument.Open(_stream, true, new OpenSettings { AutoSave = true }))
            {
                var mainPart = wordDocument.MainDocumentPart;
                var docBody = mainPart.Document.Body;
                var dynamicTables = GetTables(docBody, _varDynamicRegex);
                if (!dynamicTables.IsNullOrEmpty())
                {
                    foreach (var item in dynamicTables)
                    {
                        FillDynamicTable(item, values);
                    }
                }

                var paragraphs = docBody.Descendants<Paragraph>()
                    .Where(paragraph => !paragraph.InnerText.Contains(MergedField));
                if (!paragraphs.IsNullOrEmpty())
                {
                    foreach (var item in paragraphs)
                    {
                        FillPargraph(item, values, _varStaticRegex);
                    }
                }

                mainPart.Document.Save();
                wordDocument.Save();

                using (var mStream = new MemoryStream())
                {
                    _stream.CopyTo(mStream);
                    return mStream.ToArray();
                }
            }
        }

        private static IEnumerable<Table> GetTables(Body body, Regex regex)
        {
            return body.Descendants<Table>()?
                .Where(table => !string.IsNullOrEmpty(table.InnerText)
                && regex.IsMatch(table.InnerText));
        }

        private void FillDynamicTable(Table table, IEnumerable<ReportVariableMapping> values)
        {
            if (table == null || values.IsNullOrEmpty())
            {
                return;
            }

            var keysRow = GetKeysRow(table);
            var columsIndex = GetTableColumnsKey(keysRow);

            if (keysRow is null || columsIndex.IsNullOrEmpty())
            {
                return;
            }

            var valuesRows = values?.Where(val => val.RowNumber.HasValue && columsIndex.ContainsKey(val.Key))
                .GroupBy(val => val.RowNumber)
                .Select(valGroup => new
                {
                    Number = valGroup.Key,
                    Row = valGroup?.ToList()
                });

            if (!valuesRows.IsNullOrEmpty())
            {
                foreach (var valuesRow in valuesRows)
                {
                    var rowTable = (TableRow)keysRow.CloneNode(true);
                    foreach (var columnIndex in columsIndex)
                    {
                        var rowValue = valuesRow.Row?.FirstOrDefault(a => a.Key.Equals(columnIndex.Key));
                        var run = new Run(new Text(rowValue?.Value));
                        run.PrependChild(GetRunPropertyFromTableCell(keysRow, columnIndex.Value));
                        rowTable.Descendants<TableCell>().ElementAt(columnIndex.Value).RemoveAllChildren<Paragraph>();
                        rowTable.Descendants<TableCell>().ElementAt(columnIndex.Value).Append(new Paragraph(run));
                    }

                    table.InsertAfter(rowTable, keysRow);
                }
            }

            table.RemoveChild(keysRow);
        }

        private void FillPargraph(Paragraph paragraph, IEnumerable<ReportVariableMapping> values, Regex regex)
        {
            if (paragraph is null || values.IsNullOrEmpty())
            {
                return;
            }

            var matchs = regex.Matches(paragraph.InnerText);
            if (matchs == null || matchs.Count <= 0)
            {
                return;
            }

            var properties = GetRunPropertyFromParagraph(paragraph);
            string text = paragraph.InnerText ?? string.Empty;
            foreach (var match in matchs)
            {
                var varValue = values.FirstOrDefault(val => val.Key.Equals(match.ToString().Substring(1, match.ToString().Length - 1)));
                if (varValue != null)
                {
                    text = text.Replace($"{match}", varValue.Value?.ToString());
                }
            }

            var run = new Run(new Text(text));
            run.PrependChild(properties);
            paragraph.RemoveAllChildren<Run>();
            paragraph.Append(run);
        }

        private static Dictionary<string, int> GetTableColumnsKey(TableRow keysRow)
        {
            var index = 0;
            return keysRow?
                 .Descendants<TableCell>()?
                 .Where(cell => !string.IsNullOrEmpty(cell.InnerText))
                 .ToDictionary(cell => cell.InnerText.Trim().Substring(1, cell.InnerText.Trim().Length - 1), cell => index++);
        }

        private TableRow GetKeysRow(Table table)
        {
            return table.Descendants<TableRow>()
                 .FirstOrDefault(tab => !string.IsNullOrEmpty(tab.InnerText)
                     && _varDynamicRegex.IsMatch(tab.InnerText));
        }

        private static RunProperties GetRunPropertyFromTableCell(TableRow rowCopy, int cellIndex)
        {
            return GetRunPropertyFromParagraph(rowCopy.Descendants<TableCell>()
                        .ElementAt(cellIndex).GetFirstChild<Paragraph>());
        }

        private static RunProperties GetRunPropertyFromParagraph(Paragraph paragraph)
        {
            var runProperties = new RunProperties();
            var propos = paragraph?.GetFirstChild<Run>()?.Descendants<RunProperties>();
            if (propos != null)
            {
                foreach (var property in propos)
                {
                    runProperties.AppendChild(property.CloneNode(true));
                }
            }
            return runProperties;
        }
    }
}
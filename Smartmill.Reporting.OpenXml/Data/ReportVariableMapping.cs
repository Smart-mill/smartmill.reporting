namespace Smartmill.Reporting.OpenXml.Data
{
    public class ReportVariableMapping
    {
        public string Key { get; }

        public string Value { get; }

        public int? RowNumber { get; }

        public ReportVariableStyle Style { get; }

        public ReportVariableMapping(string key, string value, int? rowNumber = null,
            ReportVariableStyle style = null)
        {
            Key = key;
            Value = value;
            RowNumber = rowNumber;
            Style = style;
        }
    }
}
namespace Smartmill.Reporting.OpenXml.Data
{
    public class ReportVariableStyle
    {
        public int ColSpan { get; }
        public int RowSpan { get; }
        public string ForegroundColor { get; }
        public string BackgroundColor { get; }
        public bool IsBold { get; }

        public bool HasStyleSheet => !string.IsNullOrEmpty(ForegroundColor) || !string.IsNullOrEmpty(BackgroundColor) || IsBold;

        public ReportVariableStyle(int colSpan = 1, int rowSpan = 1,
            string foregroundColor = null, string backgroundColor = null, bool isBold = false)
        {
            ColSpan = colSpan;
            RowSpan = rowSpan;
            ForegroundColor = foregroundColor;
            BackgroundColor = backgroundColor;
            IsBold = isBold;
        }
    }
}
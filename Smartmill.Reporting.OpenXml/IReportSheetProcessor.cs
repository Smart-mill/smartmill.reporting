using Smartmill.Reporting.OpenXml.Data;
using Smartmill.Reporting.OpenXml.Utils;
using System.Collections.Generic;

namespace Smartmill.Reporting.OpenXml
{
    public interface IReportSheetProcessor : IDocumentLoader
    {
        byte[] Populate(IEnumerable<ReportVariableMapping> values);
    }
}
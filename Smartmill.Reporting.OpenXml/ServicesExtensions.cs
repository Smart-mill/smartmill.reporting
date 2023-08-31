using Microsoft.Extensions.DependencyInjection;

namespace Smartmill.Reporting.OpenXml
{
    public static class ServicesExtensions
    {
        public static IServiceCollection AddSmartmillReporting(this IServiceCollection services)
        {
            return services.AddSingleton<IReportDocumentProcessor, ReportDocumentProcessor>()
                      .AddSingleton<IReportSheetProcessor, ReportSheetProcessor>();
        }
    }
}
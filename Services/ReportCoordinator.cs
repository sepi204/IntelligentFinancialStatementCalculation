using ClosedXML.Excel;
using IntelligentFinancialStatementCalculation.Models;

namespace IntelligentFinancialStatementCalculation.Services;

// هماهنگ‌کنندهٔ تولید همهٔ گزارش‌ها و ترکیب خروجی‌ها
public class ReportCoordinator
{
    private readonly IEnumerable<IReportGenerator> _generators;
    private readonly ILogger<ReportCoordinator> _logger;

    public ReportCoordinator(IEnumerable<IReportGenerator> generators, ILogger<ReportCoordinator> logger)
    {
        _generators = generators;
        _logger = logger;
    }

    public async Task<MemoryStream> GenerateCombinedAsync(WorkbookInput input, CancellationToken cancellationToken)
    {
        using var finalWorkbook = new XLWorkbook();

        foreach (var generator in _generators)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _logger.LogInformation("در حال اجرای گزارش {Generator}", generator.GetType().Name);

            var workbook = await generator
                .GenerateAsync(input, cancellationToken);

            foreach (var sheet in workbook.Worksheets)
            {
                var uniqueName = GetUniqueSheetName(finalWorkbook, sheet.Name);
                sheet.CopyTo(finalWorkbook, uniqueName);
            }
        }

        var stream = new MemoryStream();
        finalWorkbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    private static string GetUniqueSheetName(XLWorkbook workbook, string baseName)
    {
        var name = baseName;
        var counter = 1;
        while (workbook.Worksheets.Any(ws => ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        {
            name = $"{baseName}_{counter++}";
        }
        return name;
    }
}






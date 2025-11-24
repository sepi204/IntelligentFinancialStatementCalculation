using ClosedXML.Excel;
using IntelligentFinancialStatementCalculation.Models;

namespace IntelligentFinancialStatementCalculation.Services.Generators;

// گزارش خلاصه جهت تست
public class SampleSummaryReport : IReportGenerator
{
    public Task<XLWorkbook> GenerateAsync(WorkbookInput input, CancellationToken cancellationToken)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("خلاصه مالی");

        sheet.Cell("A1").Value = "عنوان";
        sheet.Cell("B1").Value = "مقدار";
        sheet.Cell("A2").Value = "درآمد سال";
        sheet.Cell("B2").Value = 1_250_000;
        sheet.Cell("A3").Value = "هزینه سال";
        sheet.Cell("B3").Value = 820_000;
        sheet.Cell("A4").Value = "سود خالص";
        sheet.Cell("B4").Value = 430_000;

        sheet.Columns().AdjustToContents();

        return Task.FromResult(workbook);
    }
}


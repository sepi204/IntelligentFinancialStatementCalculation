using ClosedXML.Excel;
using IntelligentFinancialStatementCalculation.Models;

namespace IntelligentFinancialStatementCalculation.Services.Generators;

// گزارش جزئیات نمونه
public class SampleDetailReport : IReportGenerator
{
    public Task<XLWorkbook> GenerateAsync(WorkbookInput input, CancellationToken cancellationToken)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("جزئیات مالی");

        sheet.Cell("A1").Value = "ردیف";
        sheet.Cell("B1").Value = "شرح";
        sheet.Cell("C1").Value = "مقدار";

        for (var i = 2; i <= 20; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Cell(i, 1).Value = i - 1;
            sheet.Cell(i, 2).Value = $"آیتم مالی {i - 1}";
            sheet.Cell(i, 3).Value = (i - 1) * 10_000;
        }

        sheet.Columns().AdjustToContents();

        return Task.FromResult(workbook);
    }
}


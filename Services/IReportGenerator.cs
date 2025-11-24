using ClosedXML.Excel;
using IntelligentFinancialStatementCalculation.Models;

namespace IntelligentFinancialStatementCalculation.Services;

// قرارداد مشترک تولید گزارش
public interface IReportGenerator
{
    Task<XLWorkbook> GenerateAsync(WorkbookInput input, CancellationToken cancellationToken);
}


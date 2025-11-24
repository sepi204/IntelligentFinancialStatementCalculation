namespace IntelligentFinancialStatementCalculation.Models;

// ورودی مشترک برای تولید گزارش‌ها
public record WorkbookInput(string OriginalFileName, Stream SourceStream);


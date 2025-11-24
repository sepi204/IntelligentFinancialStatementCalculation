using ClosedXML.Excel;
using IntelligentFinancialStatementCalculation.Models;
using System.Collections.Immutable;
using System.Globalization;

namespace IntelligentFinancialStatementCalculation.Services.Generators;

public class FinancialPositionReportGenerator : IReportGenerator
{
    private static readonly ImmutableDictionary<string, BalanceSheetGroup> AccountGroupMap =
        new Dictionary<string, BalanceSheetGroup>
        {
            // دارایی‌های جاری
            { "1110", BalanceSheetGroup.CurrentAssets },
            { "1111", BalanceSheetGroup.CurrentAssets },
            { "1112", BalanceSheetGroup.CurrentAssets },
            { "1120", BalanceSheetGroup.CurrentAssets },

            // دارایی‌های غیرجاری
            { "1210", BalanceSheetGroup.NonCurrentAssets },
            { "1212", BalanceSheetGroup.NonCurrentAssets },
            { "1213", BalanceSheetGroup.NonCurrentAssets }, // استهلاک → منفی
            { "1220", BalanceSheetGroup.NonCurrentAssets },

            // بدهی‌های جاری
            { "2110", BalanceSheetGroup.CurrentLiabilities },
            { "2111", BalanceSheetGroup.CurrentLiabilities },
            { "2120", BalanceSheetGroup.CurrentLiabilities },

            // بدهی‌های بلندمدت
            { "2210", BalanceSheetGroup.NonCurrentLiabilities },
            { "2220", BalanceSheetGroup.NonCurrentLiabilities },

            // حقوق صاحبان سهام
            { "3110", BalanceSheetGroup.Equity },
            { "3111", BalanceSheetGroup.Equity },
            { "3120", BalanceSheetGroup.Equity },
        }.ToImmutableDictionary();

    public async Task<XLWorkbook> GenerateAsync(WorkbookInput input, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        if (!input.SourceStream.CanRead)
            throw new InvalidOperationException("فایل ورودی قابل خواندن نیست.");

        if (input.SourceStream.CanSeek)
            input.SourceStream.Position = 0;

        // 🔹 بارگذاری Workbook از Stream
        using var sourceWorkbook = new XLWorkbook(input.SourceStream);

        // 🔹 تعیین آخرین تاریخ در دفاتر
        DateTime? lastDate = null;
        foreach (var ws in sourceWorkbook.Worksheets)
        {
            var dateCol = FindDateColumn(ws);
            if (!dateCol.HasValue) continue;

            foreach (var item in ws.RowsUsed().Skip(1))
            {
                var dateStr = item.Cell(dateCol.Value).GetString().Trim();
                if (TryParseShamsiDate(dateStr, out var dt) && (!lastDate.HasValue || dt > lastDate.Value))
                    lastDate = dt;
            }
        }

        if (!lastDate.HasValue)
            throw new InvalidOperationException("تاریخ پایان دوره در دفاتر یافت نشد.");

        // 🔹 استخراج مانده حساب‌ها تا آخرین تاریخ
        var accountBalances = ExtractAccountBalances(sourceWorkbook, lastDate.Value);

        // 🔹 تجمیع بر اساس گروه‌ها
        var groups = accountBalances
            .GroupBy(x => x.Group)
            .ToDictionary(g => g.Key, g => g.Sum(x => x.Balance));

        var currentAssets = groups.GetValueOrDefault(BalanceSheetGroup.CurrentAssets, 0);
        var nonCurrentAssets = groups.GetValueOrDefault(BalanceSheetGroup.NonCurrentAssets, 0);
        var totalAssets = currentAssets + nonCurrentAssets;

        var currentLiabilities = groups.GetValueOrDefault(BalanceSheetGroup.CurrentLiabilities, 0);
        var nonCurrentLiabilities = groups.GetValueOrDefault(BalanceSheetGroup.NonCurrentLiabilities, 0);
        var totalLiabilities = currentLiabilities + nonCurrentLiabilities;

        var equity = groups.GetValueOrDefault(BalanceSheetGroup.Equity, 0);
        var totalLiabilitiesAndEquity = totalLiabilities + equity;

        // 🔹 تولید خروجی
        var outputWorkbook = new XLWorkbook();
        var sheet = outputWorkbook.Worksheets.Add("صورت وضعیت مالی");

        var row = 1;
        sheet.Cell(row, 1).Value = $"صورت وضعیت مالی تا پایان دوره {lastDate:yyyy/MM/dd}";
        sheet.Cell(row++, 1).Style.Font.Bold = true;
        row++; // فاصله

        // سرستون‌ها
        sheet.Cell(row, 1).Value = "دارایی‌ها";
        sheet.Cell(row, 5).Value = "بدهی‌ها و حقوق صاحبان سهام";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 5).Style.Font.Bold = true;
        row++;

        // --- دارایی‌ها ---
        sheet.Cell(row, 1).Value = "دارایی‌های جاری";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row++, 2).Value = FormatCurrency(currentAssets);

        sheet.Cell(row, 1).Value = "دارایی‌های غیرجاری";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row++, 2).Value = FormatCurrency(nonCurrentAssets);

        sheet.Cell(row, 1).Value = "جمع کل دارایی‌ها";
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row++, 2).Value = FormatCurrency(totalAssets);
        row++;

        // --- بدهی‌ها و حقوق ---
        sheet.Cell(row, 5).Value = "بدهی‌های جاری";
        sheet.Cell(row, 5).Style.Font.Bold = true;
        sheet.Cell(row++, 6).Value = FormatCurrency(currentLiabilities);

        sheet.Cell(row, 5).Value = "بدهی‌های بلندمدت";
        sheet.Cell(row, 5).Style.Font.Bold = true;
        sheet.Cell(row++, 6).Value = FormatCurrency(nonCurrentLiabilities);

        sheet.Cell(row, 5).Value = "جمع کل بدهی‌ها";
        sheet.Cell(row, 5).Style.Font.Bold = true;
        sheet.Cell(row++, 6).Value = FormatCurrency(totalLiabilities);

        sheet.Cell(row, 5).Value = "حقوق صاحبان سهام";
        sheet.Cell(row, 5).Style.Font.Bold = true;
        sheet.Cell(row++, 6).Value = FormatCurrency(equity);

        sheet.Cell(row, 5).Value = "جمع کل بدهی‌ها و حقوق صاحبان سهام";
        sheet.Cell(row, 5).Style.Font.Bold = true;
        sheet.Cell(row, 6).Value = FormatCurrency(totalLiabilitiesAndEquity);
        if (Math.Abs(totalAssets - totalLiabilitiesAndEquity) > 1)
        {
            sheet.Cell(row, 7).Value = "⚠️ عدم تطابق";
            sheet.Cell(row, 7).Style.Font.FontColor = XLColor.Red;
        }

        // تنظیمات ظاهری
        sheet.Columns(1, 6).Width = 22;
        sheet.Range(1, 1, row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        sheet.Range(1, 1, 1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        sheet.Columns(2, 2).Style.NumberFormat.Format = "#,##0";
        sheet.Columns(6, 6).Style.NumberFormat.Format = "#,##0";

        return outputWorkbook;
    }

    private static int? FindDateColumn(IXLWorksheet sheet)
    {
        var headerRow = sheet.FirstRowUsed();
        if (headerRow == null) return null;
        foreach (var cell in headerRow.CellsUsed())
        {
            var val = cell.GetString().Trim();
            if (val.Contains("تاريخ", StringComparison.OrdinalIgnoreCase) ||
                val.Contains("تاریخ", StringComparison.OrdinalIgnoreCase))
                return cell.Address.ColumnNumber;
        }
        return null;
    }

    private static int? FindColumn(IXLWorksheet sheet, string keyword)
    {
        var headerRow = sheet.FirstRowUsed();
        if (headerRow == null) return null;
        foreach (var cell in headerRow.CellsUsed())
        {
            var val = cell.GetString().Trim();
            if (val.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                return cell.Address.ColumnNumber;
        }
        return null;
    }

    private List<(BalanceSheetGroup Group, decimal Balance)> ExtractAccountBalances(XLWorkbook workbook, DateTime endDate)
    {
        var accounts = new Dictionary<string, (BalanceSheetGroup Group, decimal Balance, DateTime Date)>();

        foreach (var ws in workbook.Worksheets)
        {
            // پیدا کردن ستون‌های مورد نیاز
            var dateCol = FindDateColumn(ws);
            var codeKolCol = FindColumn(ws, "كد كل");
            var balanceCol = FindColumn(ws, "مانده در خط");
            var detailCol = FindColumn(ws, "كد تفصيلي");
            var moeinCol = FindColumn(ws, "كد معين");

            if (!dateCol.HasValue || !codeKolCol.HasValue || !balanceCol.HasValue)
                continue;

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                var dateStr = row.Cell(dateCol.Value).GetString().Trim();
                if (!TryParseShamsiDate(dateStr, out var dt) || dt > endDate)
                    continue;

                var codeKol = row.Cell(codeKolCol.Value).GetString().Trim();
                if (string.IsNullOrEmpty(codeKol) || !AccountGroupMap.TryGetValue(codeKol, out var group))
                    continue;

                var balanceStr = row.Cell(balanceCol.Value).GetString().Replace(",", "").Trim();
                if (!decimal.TryParse(balanceStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var balance))
                    continue;

                // تعیین کلید منحصربه‌فرد حساب
                var detail = detailCol.HasValue ? row.Cell(detailCol.Value).GetString().Trim() : "";
                var moein = moeinCol.HasValue ? row.Cell(moeinCol.Value).GetString().Trim() : "";
                var key = !string.IsNullOrEmpty(detail) ? detail :
                          !string.IsNullOrEmpty(moein) ? moein :
                          codeKol;

                // تنظیم علامت برای استهلاک
                var adjusted = codeKol == "1213" ? -balance : balance;

                // نگهداری آخرین مانده
                if (!accounts.TryGetValue(key, out var existing) || dt > existing.Date)
                {
                    accounts[key] = (group, adjusted, dt);
                }
            }
        }

        return accounts.Values.Select(x => (x.Group, x.Balance)).ToList();
    }

    private static bool TryParseShamsiDate(string input, out DateTime result)
    {
        result = default;

        if (string.IsNullOrWhiteSpace(input))
            return false;

        // حذف کلمات غیرتاریخی
        input = input.Split(' ').FirstOrDefault()?.Trim();
        if (string.IsNullOrEmpty(input))
            return false;

        // نرمال‌سازی اعداد فارسی/عربی
        input = input
            .Replace("۰", "0").Replace("۱", "1").Replace("۲", "2").Replace("۳", "3")
            .Replace("۴", "4").Replace("۵", "5").Replace("۶", "6").Replace("۷", "7")
            .Replace("۸", "8").Replace("۹", "9");

        // جدا کردن قسمت تاریخ (اولین بخش عددی با / یا -)
        foreach (var part in input.Split(' ', '/', '\\', '-', '.'))
        {
            if (part.Contains("/") || part.Contains("-"))
            {
                input = part;
                break;
            }
        }

        var parts = input.Split('/', '\\', '-', '.');
        if (parts.Length < 3) return false;

        if (int.TryParse(parts[0], out int y) &&
            int.TryParse(parts[1], out int m) &&
            int.TryParse(parts[2].Split(' ').FirstOrDefault(), out int d))
        {
            // تبدیل تقریبی شمسی → میلادی (برای مقایسه کافی است)
            try
            {
                var gregorianYear = y + 621;
                var baseDate = new DateTime(gregorianYear, 3, 21);
                var dayOfYear = (m - 1) * 31 + d;
                if (m > 6) dayOfYear -= (m - 6); // برای ماه‌های 7+ (۳۰ روزه)
                result = baseDate.AddDays(dayOfYear - 1);
                return true;
            }
            catch
            {
                // fallback: فقط سال/ماه/روز برای مقایسه
                result = new DateTime(y, Math.Max(1, m), Math.Min(d, 29));
                return true;
            }
        }

        return false;
    }

    private static string FormatCurrency(decimal value)
    {
        return value.ToString("#,##0", CultureInfo.InvariantCulture);
    }

    private enum BalanceSheetGroup
    {
        CurrentAssets,
        NonCurrentAssets,
        CurrentLiabilities,
        NonCurrentLiabilities,
        Equity
    }
}
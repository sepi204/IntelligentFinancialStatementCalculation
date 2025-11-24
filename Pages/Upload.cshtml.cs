using IntelligentFinancialStatementCalculation.Models;
using IntelligentFinancialStatementCalculation.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace IntelligentFinancialStatementCalculation.Pages;

[IgnoreAntiforgeryToken] // آپلود با AJAX انجام می‌شود
public class UploadModel : PageModel
{
    private readonly ReportCoordinator _coordinator;
    private readonly ILogger<UploadModel> _logger;

    public UploadModel(ReportCoordinator coordinator, ILogger<UploadModel> logger)
    {
        _coordinator = coordinator;
        _logger = logger;
    }

    [BindProperty]
    public IFormFile? UploadedFile { get; set; }

    public void OnGet()
    {
    }

    public async Task<IActionResult> OnPostUploadAsync(CancellationToken cancellationToken)
    {
        if (UploadedFile == null || UploadedFile.Length == 0)
        {
            return BadRequest("فایل معتبر انتخاب نشده است.");
        }

        var extension = Path.GetExtension(UploadedFile.FileName).ToLowerInvariant();
        if (extension is not ".xlsx" and not ".xls")
        {
            return BadRequest("لطفاً تنها فایل اکسل آپلود کنید.");
        }

        await using var buffer = new MemoryStream();
        await UploadedFile.CopyToAsync(buffer, cancellationToken);
        buffer.Position = 0;

        var workbookInput = new WorkbookInput(UploadedFile.FileName, buffer);

        try
        {
            var outputStream = await _coordinator
                .GenerateCombinedAsync(workbookInput, cancellationToken);

            var bytes = outputStream
                .ToArray();

            var fileName = $"کزارش_صورت_مالی{DateTime.Now.ToString("yyyy-MM-dd")}.xlsx";

            return File(
                bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName);
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("تولید گزارش توسط کاربر لغو شد.");
            return StatusCode(StatusCodes.Status499ClientClosedRequest);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "خطا در تولید گزارش");
            return StatusCode(StatusCodes.Status500InternalServerError, "خطا در پردازش فایل.");
        }
    }
}


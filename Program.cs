using IntelligentFinancialStatementCalculation.Services;
using IntelligentFinancialStatementCalculation.Services.Generators;
using Microsoft.AspNetCore.Http.Features;

var builder = WebApplication.CreateBuilder(args);

// افزودن Razor Pages
builder.Services.AddRazorPages();

// ثبت سرویس‌های گزارش
builder.Services.AddScoped<IReportGenerator, SampleSummaryReport>();
builder.Services.AddScoped<IReportGenerator, SampleDetailReport>();
builder.Services.AddScoped<IReportGenerator, FinancialPositionReportGenerator>();
builder.Services.AddScoped<ReportCoordinator>();

// حذف محدودیت‌های پیش‌فرض آپلود
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = long.MaxValue;
    options.ValueLengthLimit = int.MaxValue;
    options.MemoryBufferThreshold = int.MaxValue;
});

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = long.MaxValue;
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Upload");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.MapRazorPages();

app.Run();

using Microsoft.AspNetCore.Http.Features;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// Set the EPPlus license context for non-commercial use
ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial for commercial use

// Add services to the container.
builder.Services.AddControllersWithViews();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Qr}/{action=Index}/{id?}");

app.Run();

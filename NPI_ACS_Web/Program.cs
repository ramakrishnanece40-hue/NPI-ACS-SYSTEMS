using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// EPPlus license
ExcelPackage.License.SetNonCommercialPersonal("NPI ACS Tracking");

// Add services
builder.Services.AddControllersWithViews();

// Database connection (disabled for cloud deploy)
builder.Services.AddDbContext<ApplicationDbContext>(options =>
{
    // options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
}

app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Dashboard}/{action=Index}/{id?}");

app.Run();
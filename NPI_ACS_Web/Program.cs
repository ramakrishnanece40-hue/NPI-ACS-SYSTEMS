using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;
using OfficeOpenXml;

var builder = WebApplication.CreateBuilder(args);

// =============================
// EPPLUS LICENSE (EPPlus 8 FIX)
// =============================
ExcelPackage.License.SetNonCommercialPersonal("NPI ACS System");

// =============================
// Add MVC Services
// =============================
builder.Services.AddControllersWithViews();

// =============================
// DATABASE CONNECTION
// =============================

var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

// Railway DATABASE_URL support
var databaseUrl = Environment.GetEnvironmentVariable("DATABASE_URL");

if (!string.IsNullOrEmpty(databaseUrl))
{
var uri = new Uri(databaseUrl);
var userInfo = uri.UserInfo.Split(':');
connectionString =
    $"Host={uri.Host};Port={uri.Port};Database={uri.AbsolutePath.Trim('/')};Username={userInfo[0]};Password={userInfo[1]};SSL Mode=Require;Trust Server Certificate=true";


}

builder.Services.AddDbContext<ApplicationDbContext>(options =>
options.UseNpgsql(connectionString));

// =============================
// BUILD APP
// =============================
var app = builder.Build();

// =============================
// ERROR HANDLING
// =============================
if (!app.Environment.IsDevelopment())
{
app.UseExceptionHandler("/Home/Error");
}

// =============================
// MIDDLEWARE
// =============================
app.UseStaticFiles();
app.UseRouting();
app.UseAuthorization();

// =============================
// ROUTES
// =============================
app.MapControllerRoute(
name: "default",
pattern: "{controller=ACSTasks}/{action=Index}/{id?}");

// =============================
// DATABASE MIGRATION (SAFE)
// =============================
try
{
using (var scope = app.Services.CreateScope())
{
var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
db.Database.Migrate();
}
}
catch (Exception ex)
{
Console.WriteLine("Migration error: " + ex.Message);
}

// =============================
app.Run();

using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();

string connectionString;

// Railway PostgreSQL variables (automatic)
var host = Environment.GetEnvironmentVariable("PGHOST");
var port = Environment.GetEnvironmentVariable("PGPORT");
var db = Environment.GetEnvironmentVariable("PGDATABASE");
var user = Environment.GetEnvironmentVariable("PGUSER");
var password = Environment.GetEnvironmentVariable("PGPASSWORD");

// If Railway variables exist → use them
if (!string.IsNullOrEmpty(host))
{
    connectionString =
        $"Host={host};Port={port};Database={db};Username={user};Password={password};SSL Mode=Require;Trust Server Certificate=true";
}
else
{
    // fallback for local development
    connectionString = builder.Configuration.GetConnectionString("DefaultConnection");
}

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseNpgsql(connectionString));

builder.Services.AddDefaultIdentity<IdentityUser>(options =>
{
    options.SignIn.RequireConfirmedAccount = false;
})
.AddEntityFrameworkStores<ApplicationDbContext>();

var app = builder.Build();


// Auto run migrations safely
using (var scope = app.Services.CreateScope())
{
    var dbContext = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    dbContext.Database.Migrate();
}


// Railway port binding
var portEnv = Environment.GetEnvironmentVariable("PORT") ?? "8080";
app.Urls.Add($"http://0.0.0.0:{portEnv}");

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
}

app.UseStaticFiles();
app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=ACSTasks}/{action=Index}/{id?}");

app.MapRazorPages();

app.Run();
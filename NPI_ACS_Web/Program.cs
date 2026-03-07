using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();

var databaseUrl = Environment.GetEnvironmentVariable("DATABASE_URL");

string connectionString;

if (!string.IsNullOrEmpty(databaseUrl))
{
    var uri = new Uri(databaseUrl);
    var userInfo = uri.UserInfo.Split(':');

    connectionString =
        $"Host={uri.Host};" +
        $"Port={uri.Port};" +
        $"Database={uri.AbsolutePath.Trim('/')};" +
        $"Username={userInfo[0]};" +
        $"Password={userInfo[1]};" +
        $"SSL Mode=Require;" +
        $"Trust Server Certificate=true";
}
else
{
    throw new Exception("DATABASE_URL not found");
}

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseNpgsql(connectionString));

builder.Services.AddDefaultIdentity<IdentityUser>(options =>
{
    options.SignIn.RequireConfirmedAccount = false;
})
.AddEntityFrameworkStores<ApplicationDbContext>();

var app = builder.Build();


// Wait for DB to be ready (important for Railway)
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();

    var retries = 10;

    while (retries > 0)
    {
        try
        {
            db.Database.Migrate();
            break;
        }
        catch
        {
            retries--;
            Thread.Sleep(5000);
        }
    }
}


// Railway port
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
app.Urls.Add($"http://0.0.0.0:{port}");

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
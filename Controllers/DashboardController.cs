using Microsoft.AspNetCore.Mvc;
using NPI_ACS_Web.Data;
using System.Linq;

namespace NPI_ACS_Web.Controllers
{
    public class DashboardController : Controller
    {
        private readonly ApplicationDbContext _context;

        public DashboardController(ApplicationDbContext context)
        {
            _context = context;
        }

        public IActionResult Index()
        {
            var tasks = _context.ACSTasks.ToList();

            ViewBag.Total = tasks.Count;

            ViewBag.Open = tasks
                .Count(x => x.Status != null && x.Status.ToLower() == "open");

            ViewBag.Ongoing = tasks
                .Count(x => x.Status != null && x.Status.ToLower() == "ongoing");

            ViewBag.Closed = tasks
                .Count(x => x.Status != null && x.Status.ToLower() == "closed");

            return View(tasks);
        }
    }
}
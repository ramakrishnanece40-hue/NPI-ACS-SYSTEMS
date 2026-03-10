using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;
using NPI_ACS_Web.Models;
using OfficeOpenXml;

namespace NPI_ACS_Web.Controllers
{
    public class ACSTasksController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ACSTasksController(ApplicationDbContext context)
        {
            _context = context;
        }

        // =============================
        // DASHBOARD
        // =============================

        public IActionResult Index()
        {
            var tasks = _context.ACSTasks.OrderByDescending(x => x.Id).ToList();
            return View(tasks);
        }

        // =============================
        // CREATE PAGE
        // =============================

        public IActionResult Create()
        {
            return View();
        }

        // =============================
        // SAVE TASK
        // =============================

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ACSTask task, IFormFile Attachment)
        {
            try
            {
                if (Attachment != null && Attachment.Length > 0)
                {
                    var uploads = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads");

                    if (!Directory.Exists(uploads))
                        Directory.CreateDirectory(uploads);

                    var fileName = Guid.NewGuid().ToString() + Path.GetExtension(Attachment.FileName);
                    var filePath = Path.Combine(uploads, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await Attachment.CopyToAsync(stream);
                    }

                    task.AttachmentPath = "/uploads/" + fileName;
                }

                _context.Add(task);
                await _context.SaveChangesAsync();

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                ViewBag.Error = ex.Message;
                return View(task);
            }
        }

        // =============================
        // EXPORT EXCEL
        // =============================

        public IActionResult ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var tasks = _context.ACSTasks.ToList();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ACS Tasks");

                worksheet.Cells[1, 1].Value = "Project";
                worksheet.Cells[1, 2].Value = "ODM";
                worksheet.Cells[1, 3].Value = "Priority";
                worksheet.Cells[1, 4].Value = "Status";
                worksheet.Cells[1, 5].Value = "Start Date";
                worksheet.Cells[1, 6].Value = "Due Date";
                worksheet.Cells[1, 7].Value = "Remarks";

                int row = 2;

                foreach (var task in tasks)
                {
                    worksheet.Cells[row, 1].Value = task.Project;
                    worksheet.Cells[row, 2].Value = task.ODM;
                    worksheet.Cells[row, 3].Value = task.Priority;
                    worksheet.Cells[row, 4].Value = task.Status;
                    worksheet.Cells[row, 5].Value = task.StartDate;
                    worksheet.Cells[row, 6].Value = task.DueDate;
                    worksheet.Cells[row, 7].Value = task.Remarks;
                    row++;
                }

                var stream = new MemoryStream(package.GetAsByteArray());

                return File(stream,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "ACS_Task_Report.xlsx");
            }
        }
    }
}
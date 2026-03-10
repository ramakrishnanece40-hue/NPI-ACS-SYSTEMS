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
        private readonly IWebHostEnvironment _env;

        public ACSTasksController(ApplicationDbContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        // ======================
        // LIST
        // ======================
        public async Task<IActionResult> Index()
        {
            var tasks = await _context.ACSTasks.ToListAsync();
            return View(tasks);
        }

        // ======================
        // CREATE GET
        // ======================
        public IActionResult Create()
        {
            return View();
        }

        // ======================
        // CREATE POST
        // ======================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ACSTask task, IFormFile? AttachmentFile)
        {
            try
            {
                ModelState.Remove("AttachmentPath");

                if (AttachmentFile != null && AttachmentFile.Length > 0)
                {
                    string uploads = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

                    if (!Directory.Exists(uploads))
                        Directory.CreateDirectory(uploads);

                    string fileName = Guid.NewGuid() + Path.GetExtension(AttachmentFile.FileName);
                    string path = Path.Combine(uploads, fileName);

                    using (var stream = new FileStream(path, FileMode.Create))
                    {
                        await AttachmentFile.CopyToAsync(stream);
                    }

                    task.AttachmentPath = "/uploads/" + fileName;
                }

                _context.ACSTasks.Add(task);
                await _context.SaveChangesAsync();

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return View(task);
            }
        }

        // ======================
        // EDIT GET
        // ======================
        public async Task<IActionResult> Edit(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        // ======================
        // EDIT POST
        // ======================
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ACSTask task, IFormFile? AttachmentFile)
        {
            if (id != task.Id)
                return NotFound();

            try
            {
                var existing = await _context.ACSTasks.FindAsync(id);

                if (existing == null)
                    return NotFound();

                ModelState.Remove("AttachmentPath");

                existing.Project = task.Project;
                existing.ODM = task.ODM;
                existing.Product = task.Product;
                existing.Model = task.Model;
                existing.Question = task.Question;
                existing.ActionDetail = task.ActionDetail;
                existing.FourM = task.FourM;
                existing.NeolyncPIC = task.NeolyncPIC;
                existing.CustomerPIC = task.CustomerPIC;
                existing.Priority = task.Priority;
                existing.Status = task.Status;
                existing.StartDate = task.StartDate;
                existing.DueDate = task.DueDate;
                existing.ActualCloseDate = task.ActualCloseDate;
                existing.Remarks = task.Remarks;

                if (AttachmentFile != null && AttachmentFile.Length > 0)
                {
                    string uploads = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

                    if (!Directory.Exists(uploads))
                        Directory.CreateDirectory(uploads);

                    string fileName = Guid.NewGuid() + Path.GetExtension(AttachmentFile.FileName);
                    string path = Path.Combine(uploads, fileName);

                    using (var stream = new FileStream(path, FileMode.Create))
                    {
                        await AttachmentFile.CopyToAsync(stream);
                    }

                    existing.AttachmentPath = "/uploads/" + fileName;
                }

                _context.Update(existing);
                await _context.SaveChangesAsync();

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return View(task);
            }
        }

        // ======================
        // DELETE
        // ======================
        public async Task<IActionResult> Delete(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        [HttpPost, ActionName("Delete")]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task != null)
            {
                _context.ACSTasks.Remove(task);
                await _context.SaveChangesAsync();
            }

            return RedirectToAction("Index");
        }

        // ======================
        // EXPORT EXCEL
        // ======================
        public IActionResult ExportToExcel()
        {
            ExcelPackage.License.SetNonCommercialPersonal("NPI ACS");

            var tasks = _context.ACSTasks.ToList();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("ACS Tasks");

            ws.Cells[1, 1].Value = "Project";
            ws.Cells[1, 2].Value = "ODM";
            ws.Cells[1, 3].Value = "Product";
            ws.Cells[1, 4].Value = "Model";
            ws.Cells[1, 5].Value = "Question";
            ws.Cells[1, 6].Value = "Action Detail";
            ws.Cells[1, 7].Value = "4M";
            ws.Cells[1, 8].Value = "Neolync PIC";
            ws.Cells[1, 9].Value = "Customer PIC";
            ws.Cells[1, 10].Value = "Priority";
            ws.Cells[1, 11].Value = "Status";
            ws.Cells[1, 12].Value = "Start Date";
            ws.Cells[1, 13].Value = "Due Date";
            ws.Cells[1, 14].Value = "Actual Close Date";
            ws.Cells[1, 15].Value = "Remarks";
            ws.Cells[1, 16].Value = "Attachment";

            int row = 2;

            foreach (var t in tasks)
            {
                ws.Cells[row, 1].Value = t.Project;
                ws.Cells[row, 2].Value = t.ODM;
                ws.Cells[row, 3].Value = t.Product;
                ws.Cells[row, 4].Value = t.Model;
                ws.Cells[row, 5].Value = t.Question;
                ws.Cells[row, 6].Value = t.ActionDetail;
                ws.Cells[row, 7].Value = t.FourM;
                ws.Cells[row, 8].Value = t.NeolyncPIC;
                ws.Cells[row, 9].Value = t.CustomerPIC;
                ws.Cells[row, 10].Value = t.Priority;
                ws.Cells[row, 11].Value = t.Status;
                ws.Cells[row, 12].Value = t.StartDate;
                ws.Cells[row, 13].Value = t.DueDate;
                ws.Cells[row, 14].Value = t.ActualCloseDate;
                ws.Cells[row, 15].Value = t.Remarks;
                ws.Cells[row, 16].Value = t.AttachmentPath;

                row++;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream(package.GetAsByteArray());

            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ACS_Tasks.xlsx");
        }
    }
}
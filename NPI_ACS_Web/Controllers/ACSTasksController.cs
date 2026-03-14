using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc.Rendering;
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

        // =====================
        // INDEX
        // =====================
        public IActionResult Index()
        {
            var tasks = _context.ACSTasks.OrderByDescending(x => x.Id).ToList();
            return View(tasks);
        }

        // =====================
        // CREATE
        // =====================
        public IActionResult Create()
        {
            LoadDropdowns();
            return View();
        }

       [HttpPost]
[ValidateAntiForgeryToken]
public async Task<IActionResult> Create(ACSTask task, IFormFile? Attachment)
{
    if (ModelState.IsValid)
    {
        // Attachment is OPTIONAL
        if (Attachment != null && Attachment.Length > 0)
        {
            var folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads");

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            var fileName = Guid.NewGuid().ToString() + Path.GetExtension(Attachment.FileName);
            var path = Path.Combine(folder, fileName);

            using (var stream = new FileStream(path, FileMode.Create))
            {
                await Attachment.CopyToAsync(stream);
            }

            task.AttachmentPath = "/uploads/" + fileName;
        }

        // Save task even if no file uploaded
        _context.ACSTasks.Add(task);
        await _context.SaveChangesAsync();

        return RedirectToAction(nameof(Index));
    }

    LoadDropdowns();
    return View(task);
}

        // =====================
        // DROPDOWN LIST
        // =====================
        private void LoadDropdowns()
        {
            ViewBag.PriorityList = new List<SelectListItem>
            {
                new SelectListItem{ Value="High", Text="High"},
                new SelectListItem{ Value="Medium", Text="Medium"},
                new SelectListItem{ Value="Low", Text="Low"}
            };

            ViewBag.StatusList = new List<SelectListItem>
            {
                new SelectListItem{ Value="Open", Text="Open"},
                new SelectListItem{ Value="In Progress", Text="In Progress"},
                new SelectListItem{ Value="Closed", Text="Closed"}
            };
        }

        // =====================
        // DETAILS
        // =====================
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
                return NotFound();

            var task = await _context.ACSTasks.FirstOrDefaultAsync(m => m.Id == id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        // =====================
        // EDIT
        // =====================
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
                return NotFound();

            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            LoadDropdowns();
            return View(task);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ACSTask task)
        {
            if (id != task.Id)
                return NotFound();

            if (ModelState.IsValid)
            {
                _context.Update(task);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }

            LoadDropdowns();
            return View(task);
        }

        // =====================
        // DELETE
        // =====================
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
                return NotFound();

            var task = await _context.ACSTasks.FirstOrDefaultAsync(m => m.Id == id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task != null)
                _context.ACSTasks.Remove(task);

            await _context.SaveChangesAsync();

            return RedirectToAction(nameof(Index));
        }

        // =====================
        // EXPORT EXCEL
        // =====================
        public IActionResult ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var tasks = _context.ACSTasks.ToList();

            using var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("ACS Tasks");

            sheet.Cells[1, 1].Value = "Project";
            sheet.Cells[1, 2].Value = "Product";
            sheet.Cells[1, 3].Value = "Model";
            sheet.Cells[1, 4].Value = "Question";
            sheet.Cells[1, 5].Value = "Action Detail";
            sheet.Cells[1, 6].Value = "4M";
            sheet.Cells[1, 7].Value = "Neolync PIC";
            sheet.Cells[1, 8].Value = "Customer PIC";
            sheet.Cells[1, 9].Value = "Priority";
            sheet.Cells[1, 10].Value = "Status";

            int row = 2;

            foreach (var t in tasks)
            {
                sheet.Cells[row, 1].Value = t.Project;
                sheet.Cells[row, 2].Value = t.Product;
                sheet.Cells[row, 3].Value = t.Model;
                sheet.Cells[row, 4].Value = t.Question;
                sheet.Cells[row, 5].Value = t.ActionDetail;
                sheet.Cells[row, 6].Value = t.FourM;
                sheet.Cells[row, 7].Value = t.NeolyncPIC;
                sheet.Cells[row, 8].Value = t.CustomerPIC;
                sheet.Cells[row, 9].Value = t.Priority;
                sheet.Cells[row, 10].Value = t.Status;

                row++;
            }

            var stream = new MemoryStream(package.GetAsByteArray());

            return File(stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ACS_Tasks.xlsx");
        }
    }
}
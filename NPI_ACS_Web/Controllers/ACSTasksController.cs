using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPI_ACS_Web.Data;
using NPI_ACS_Web.Models;
using OfficeOpenXml;

namespace NPI_ACS_Web.Controllers
{
    [Route("ACSTasks")]
    public class ACSTasksController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IWebHostEnvironment _environment;

        public ACSTasksController(ApplicationDbContext context, IWebHostEnvironment environment)
        {
            _context = context;
            _environment = environment;
        }

        // ===============================
        // LIST PAGE
        // ===============================
        [HttpGet("")]
        [HttpGet("Index")]
        public async Task<IActionResult> Index()
        {
            var tasks = await _context.ACSTasks.ToListAsync();
            return View(tasks);
        }

        // ===============================
        // CREATE GET
        // ===============================
        [HttpGet("Create")]
        public IActionResult Create()
        {
            return View();
        }

        // ===============================
        // CREATE POST
        // ===============================
        [HttpPost("Create")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ACSTask task, IFormFile? AttachmentFile)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    if (AttachmentFile != null && AttachmentFile.Length > 0)
                    {
                        var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

                        if (!Directory.Exists(uploadsFolder))
                            Directory.CreateDirectory(uploadsFolder);

                        var fileName = Guid.NewGuid().ToString() + Path.GetExtension(AttachmentFile.FileName);
                        var filePath = Path.Combine(uploadsFolder, fileName);

                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            await AttachmentFile.CopyToAsync(stream);
                        }

                        task.AttachmentPath = "/uploads/" + fileName;
                    }

                    _context.ACSTasks.Add(task);
                    await _context.SaveChangesAsync();

                    return RedirectToAction(nameof(Index));
                }

                return View(task);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return View(task);
            }
        }

        // ===============================
        // EDIT GET
        // ===============================
        [HttpGet("Edit/{id}")]
        public async Task<IActionResult> Edit(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        // ===============================
        // EDIT POST
        // ===============================
        [HttpPost("Edit/{id}")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ACSTask task, IFormFile? AttachmentFile)
        {
            if (id != task.Id)
                return NotFound();

            var existingTask = await _context.ACSTasks.FindAsync(id);

            if (existingTask == null)
                return NotFound();

            try
            {
                existingTask.Project = task.Project;
                existingTask.ODM = task.ODM;
                existingTask.Product = task.Product;
                existingTask.Model = task.Model;
                existingTask.Question = task.Question;
                existingTask.ActionDetail = task.ActionDetail;
                existingTask.FourM = task.FourM;
                existingTask.NeolyncPIC = task.NeolyncPIC;
                existingTask.CustomerPIC = task.CustomerPIC;
                existingTask.Priority = task.Priority;
                existingTask.Status = task.Status;
                existingTask.StartDate = task.StartDate;
                existingTask.DueDate = task.DueDate;
                existingTask.ActualCloseDate = task.ActualCloseDate;
                existingTask.Remarks = task.Remarks;

                if (AttachmentFile != null && AttachmentFile.Length > 0)
                {
                    var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");

                    if (!Directory.Exists(uploadsFolder))
                        Directory.CreateDirectory(uploadsFolder);

                    var fileName = Guid.NewGuid().ToString() + Path.GetExtension(AttachmentFile.FileName);
                    var filePath = Path.Combine(uploadsFolder, fileName);

                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await AttachmentFile.CopyToAsync(stream);
                    }

                    existingTask.AttachmentPath = "/uploads/" + fileName;
                }

                _context.Update(existingTask);
                await _context.SaveChangesAsync();

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return View(task);
            }
        }

        // ===============================
        // DETAILS
        // ===============================
        [HttpGet("Details/{id}")]
        public async Task<IActionResult> Details(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        // ===============================
        // DELETE GET
        // ===============================
        [HttpGet("Delete/{id}")]
        public async Task<IActionResult> Delete(int id)
        {
            var task = await _context.ACSTasks.FindAsync(id);

            if (task == null)
                return NotFound();

            return View(task);
        }

        // ===============================
        // DELETE POST
        // ===============================
        [HttpPost("Delete/{id}")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            try
            {
                var task = await _context.ACSTasks.FindAsync(id);

                if (task != null)
                {
                    _context.ACSTasks.Remove(task);
                    await _context.SaveChangesAsync();
                }

                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return RedirectToAction(nameof(Index));
            }
        }

        // ===============================
        // EXPORT EXCEL
        // ===============================
        [HttpGet("ExportToExcel")]
        public IActionResult ExportToExcel()
        {
            ExcelPackage.License.SetNonCommercialPersonal("NPI ACS System");

            var tasks = _context.ACSTasks.ToList();

            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("ACS Task List");

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

            foreach (var item in tasks)
            {
                ws.Cells[row, 1].Value = item.Project;
                ws.Cells[row, 2].Value = item.ODM;
                ws.Cells[row, 3].Value = item.Product;
                ws.Cells[row, 4].Value = item.Model;
                ws.Cells[row, 5].Value = item.Question;
                ws.Cells[row, 6].Value = item.ActionDetail;
                ws.Cells[row, 7].Value = item.FourM;
                ws.Cells[row, 8].Value = item.NeolyncPIC;
                ws.Cells[row, 9].Value = item.CustomerPIC;
                ws.Cells[row, 10].Value = item.Priority;
                ws.Cells[row, 11].Value = item.Status;
                ws.Cells[row, 12].Value = item.StartDate;
                ws.Cells[row, 13].Value = item.DueDate;
                ws.Cells[row, 14].Value = item.ActualCloseDate;
                ws.Cells[row, 15].Value = item.Remarks;
                ws.Cells[row, 16].Value = item.AttachmentPath;

                row++;
            }

            ws.Cells.AutoFitColumns();

            var stream = new MemoryStream(package.GetAsByteArray());

            return File(
                stream,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ACS_Task_List.xlsx"
            );
        }
    }
}
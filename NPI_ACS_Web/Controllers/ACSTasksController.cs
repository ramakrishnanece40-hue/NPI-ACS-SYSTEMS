using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc.Rendering;
using NPI_ACS_Web.Data;
using NPI_ACS_Web.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace NPI_ACS_Web.Controllers
{
public class ACSTasksController : Controller
{
private readonly ApplicationDbContext _context;

    public ACSTasksController(ApplicationDbContext context)
    {
        _context = context;

        // EPPlus License (VERY IMPORTANT)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    // INDEX
    public IActionResult Index()
    {
        var tasks = _context.ACSTasks.OrderByDescending(x => x.Id).ToList();
        return View(tasks);
    }

    // CREATE
    public IActionResult Create()
    {
        LoadDropdowns();
        return View();
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(ACSTask task, IFormFile? Attachment)
    {
        try
        {
            if (Attachment != null && Attachment.Length > 0)
            {
                var uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads");

                if (!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                var fileName = Guid.NewGuid() + Path.GetExtension(Attachment.FileName);
                var path = Path.Combine(uploadPath, fileName);

                using var stream = new FileStream(path, FileMode.Create);
                await Attachment.CopyToAsync(stream);

                task.AttachmentPath = "/uploads/" + fileName;
            }

            _context.ACSTasks.Add(task);
            await _context.SaveChangesAsync();

            return RedirectToAction(nameof(Index));
        }
        catch (Exception ex)
        {
            var message = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
            return Content("DATABASE ERROR: " + message);
        }
    }

    // DROPDOWNS
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

    // DETAILS
    public async Task<IActionResult> Details(int? id)
    {
        if (id == null) return NotFound();

        var task = await _context.ACSTasks.FirstOrDefaultAsync(m => m.Id == id);
        if (task == null) return NotFound();

        return View(task);
    }

    // EDIT
    public async Task<IActionResult> Edit(int? id)
    {
        if (id == null) return NotFound();

        var task = await _context.ACSTasks.FindAsync(id);
        if (task == null) return NotFound();

        LoadDropdowns();
        return View(task);
    }

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(int id, ACSTask task)
    {
        if (id != task.Id) return NotFound();

        try
        {
            if (ModelState.IsValid)
            {
                _context.Update(task);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
        }
        catch (Exception ex)
        {
            var message = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
            return Content("DATABASE ERROR: " + message);
        }

        LoadDropdowns();
        return View(task);
    }

    // DELETE
    public async Task<IActionResult> Delete(int? id)
    {
        if (id == null) return NotFound();

        var task = await _context.ACSTasks.FirstOrDefaultAsync(m => m.Id == id);
        if (task == null) return NotFound();

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

    // EXPORT TO EXCEL
    public IActionResult ExportToExcel()
    {
        var tasks = _context.ACSTasks.ToList();

        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("ACS Tasks");

        string[] headers =
        {
            "Project","ODM","Product","Model","Question","Action Detail","4M",
            "Neolync PIC","Customer PIC","Priority","Status",
            "Start Date","Due Date","Actual Close Date","Remarks","Attachment"
        };

        for (int i = 0; i < headers.Length; i++)
        {
            sheet.Cells[1, i + 1].Value = headers[i];
            sheet.Cells[1, i + 1].Style.Font.Bold = true;
            sheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        }

        int row = 2;

        foreach (var t in tasks)
        {
            sheet.Cells[row,1].Value = t.Project;
            sheet.Cells[row,2].Value = t.ODM;
            sheet.Cells[row,3].Value = t.Product;
            sheet.Cells[row,4].Value = t.Model;
            sheet.Cells[row,5].Value = t.Question;
            sheet.Cells[row,6].Value = t.ActionDetail;
            sheet.Cells[row,7].Value = t.FourM;
            sheet.Cells[row,8].Value = t.NeolyncPIC;
            sheet.Cells[row,9].Value = t.CustomerPIC;
            sheet.Cells[row,10].Value = t.Priority;
            sheet.Cells[row,11].Value = t.Status;
            sheet.Cells[row,12].Value = t.StartDate;
            sheet.Cells[row,13].Value = t.DueDate;
            sheet.Cells[row,14].Value = t.ActualCloseDate;
            sheet.Cells[row,15].Value = t.Remarks;
            sheet.Cells[row,16].Value = t.AttachmentPath;

            sheet.Cells[row,11].Style.Fill.PatternType = ExcelFillStyle.Solid;

            if (t.Status == "Open")
                sheet.Cells[row,11].Style.Fill.BackgroundColor.SetColor(Color.Red);

            if (t.Status == "In Progress")
                sheet.Cells[row,11].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

            if (t.Status == "Closed")
                sheet.Cells[row,11].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            row++;
        }

        sheet.Cells.AutoFitColumns();

        var stream = new MemoryStream(package.GetAsByteArray());

        return File(stream,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "ACS_Tasks.xlsx");
    }
}


}

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Mvc.Rendering;
using NPI_ACS_Web.Data;
using NPI_ACS_Web.Models;
using System.Text;

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

    // =====================
    // DROPDOWNS
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
    // EXPORT TO EXCEL (CSV FORMAT)
    // =====================
    public IActionResult ExportToExcel()
    {
        var tasks = _context.ACSTasks.ToList();

        var builder = new StringBuilder();

        builder.AppendLine("Project,ODM,Product,Model,Question,Action Detail,4M,Neolync PIC,Customer PIC,Priority,Status");

        foreach (var t in tasks)
        {
            builder.AppendLine($"{t.Project},{t.ODM},{t.Product},{t.Model},{t.Question},{t.ActionDetail},{t.FourM},{t.NeolyncPIC},{t.CustomerPIC},{t.Priority},{t.Status}");
        }

        return File(
            Encoding.UTF8.GetBytes(builder.ToString()),
            "text/csv",
            "ACS_Tasks.csv"
        );
    }
}


}

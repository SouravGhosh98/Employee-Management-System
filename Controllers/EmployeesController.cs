using Employee_Management_System.Data;
using Employee_Management_System.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;


namespace Employee_Management_System.Controllers
{
    public class EmployeesController : Controller
    {
        private readonly AppDbContext _context;
        public EmployeesController(AppDbContext context)
        {
            _context = context;
        }
        public IActionResult Create()
        {
            ViewBag.Designations = new List<SelectListItem>
    {
        new SelectListItem { Value = "Software Engineer", Text = "Software Engineer" },
        new SelectListItem { Value = "Team Lead", Text = "Team Lead" },
        new SelectListItem { Value = "HR", Text = "HR" },
        new SelectListItem { Value = "Manager", Text = "Manager" },
        new SelectListItem { Value = "Intern", Text = "Intern" }
    };
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(Employee employee)
        {
            ViewBag.Designations = new List<SelectListItem>
            {
                new SelectListItem { Text = "Software Engineer", Value = "Software Engineer" },
                new SelectListItem { Text = "Team Lead", Value = "Team Lead" },
                new SelectListItem { Text = "HR", Value = "HR" },
                new SelectListItem { Text = "Manager", Value = "Manager" },
                new SelectListItem { Text = "Intern", Value = "Intern" }
            };
            try
            {
                _context.Employees.Add(employee);
                await _context.SaveChangesAsync();
                TempData["Success"] = "Employee added successfully!";
                return Ok("Save");
               // return View(employee);

            }
            catch (Exception)
            {
                ModelState.AddModelError("", "Something went wrong while saving data.");
                return View(employee);
            }
        }

        // List all employees
        public async Task<IActionResult> Index()
        {
            var employees = await _context.Employees.ToListAsync();
            return View(employees);
        }

        // Edit page
        public async Task<IActionResult> Edit(int id)
        {
            var employee = await _context.Employees.FindAsync(id);
            if (employee == null) return NotFound();

            ViewBag.Designations = GetDesignations();
            return View(employee);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Employee employee)
        {
            //if (!ModelState.IsValid)
            //{
                ViewBag.Designations = GetDesignations();
               // return View(employee);
            //}

            _context.Update(employee);
            await _context.SaveChangesAsync();
            TempData["Success"] = "Employee updated successfully!";
            return RedirectToAction(nameof(Index));
        }

        // Delete
        public async Task<IActionResult> Delete(int id)
        {
            var employee = await _context.Employees.FindAsync(id);
            if (employee != null)
            {
                _context.Employees.Remove(employee);
                await _context.SaveChangesAsync();
                TempData["Success"] = "Employee deleted successfully!";
            }
            else
            {
                TempData["Error"] = "Employee not found.";
            }

            return RedirectToAction(nameof(Index));
        }

        // Upload Excel
        [HttpPost]
        public async Task<IActionResult> UploadExcel(IFormFile excelFile)
        {
            if (excelFile != null && excelFile.Length > 0)
            {
                using var stream = new MemoryStream();
                await excelFile.CopyToAsync(stream);

                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    var employee = new Employee
                    {
                        Name = worksheet.Cells[row, 1].Text,
                        ContactNo = worksheet.Cells[row, 2].Text,
                        Email = worksheet.Cells[row, 3].Text,
                        City = worksheet.Cells[row, 4].Text,
                        State = worksheet.Cells[row, 5].Text,
                        Pincode = worksheet.Cells[row, 6].Text,
                        Designation = worksheet.Cells[row, 7].Text
                    };

                    _context.Employees.Add(employee);
                }

                await _context.SaveChangesAsync();
                TempData["Success"] = "Employees uploaded successfully!";
            }
            else
            {
                TempData["Error"] = "Invalid file.";
            }

            return RedirectToAction(nameof(Index));
        }

        // Helper method
        private List<SelectListItem> GetDesignations()
        {
            return new List<SelectListItem>
    {
        new SelectListItem { Text = "Software Engineer", Value = "Software Engineer" },
        new SelectListItem { Text = "Team Lead", Value = "Team Lead" },
        new SelectListItem { Text = "HR", Value = "HR" },
        new SelectListItem { Text = "Manager", Value = "Manager" },
        new SelectListItem { Text = "Intern", Value = "Intern" }
    };
        }

        public async Task<IActionResult> DownloadExcel()
        {
            // ✅ Correct license setup for EPPlus 8 (still valid)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var employees = await _context.Employees.ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Employees");

                // Headers
                worksheet.Cells[1, 1].Value = "Name";
                worksheet.Cells[1, 2].Value = "Contact No";
                worksheet.Cells[1, 3].Value = "Email";
                worksheet.Cells[1, 4].Value = "City";
                worksheet.Cells[1, 5].Value = "State";
                worksheet.Cells[1, 6].Value = "Pincode";
                worksheet.Cells[1, 7].Value = "Designation";

                // Data
                int row = 2;
                foreach (var emp in employees)
                {
                    worksheet.Cells[row, 1].Value = emp.Name;
                    worksheet.Cells[row, 2].Value = emp.ContactNo;
                    worksheet.Cells[row, 3].Value = emp.Email;
                    worksheet.Cells[row, 4].Value = emp.City;
                    worksheet.Cells[row, 5].Value = emp.State;
                    worksheet.Cells[row, 6].Value = emp.Pincode;
                    worksheet.Cells[row, 7].Value = emp.Designation;
                    row++;
                }

                var stream = new MemoryStream(package.GetAsByteArray());
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Employees.xlsx");
            }
        }

    }

}


using Employee_Management_System.Data;
using Employee_Management_System.Models;
using Microsoft.AspNetCore.Mvc;

namespace Employee_Management_System.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EmployeeApiController : ControllerBase
    {
        private readonly AppDbContext _context;

        public EmployeeApiController(AppDbContext context)
        {
            _context = context;
        }
        [HttpPost("upload")]
        public async Task<IActionResult> Upload([FromBody] List<EmployeeDto> employees)
        {
            if (employees == null || !employees.Any())
                return BadRequest(new { message = "No data received." });
            var employeeEntities = employees.Select(e => new Employee
            {
                Name = e.Name,
                Email = e.Email,
                ContactNo = e.ContactNo,
                City = e.City,
                State = e.State,
                Pincode = e.Pincode,
                Designation = e.Designation,
                Dependents = new List<Dependent>()
            }).ToList();

            await _context.Employees.AddRangeAsync(employeeEntities);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Employees uploaded successfully!" });
        }
    }
}


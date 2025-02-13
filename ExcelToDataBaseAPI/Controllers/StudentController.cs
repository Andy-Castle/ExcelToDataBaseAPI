using ExcelDataReader;
using ExcelToDataBaseAPI.Data;
using ExcelToDataBaseAPI.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ExcelToDataBaseAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;

        public StudentController(ApplicationDbContext context)
        {
            _context = context;
        }

        [HttpPost("ImportarExcelFile")]
        public IActionResult ImportarExcelFile(IFormFile file)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                if (file == null || file.Length == 0)
                {
                    return BadRequest("No hay archivo importado");
                }

                var uploadsFolder = $"{Directory.GetCurrentDirectory()}\\Uploads";

                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                var filePath = Path.Combine(uploadsFolder, file.FileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        bool isHeaderSkipped = false;
                        
                        do
                        {
                            while (reader.Read())
                            {
                                if (!isHeaderSkipped)
                                {
                                    isHeaderSkipped = true;
                                    continue;
                                }


                                Student student = new Student();
                                student.Name = reader.GetValue(1).ToString()!;
                                student.Marks = Convert.ToInt32(reader.GetValue(2).ToString());

                                _context.Add(student);

                                _context.SaveChanges();
                            }
                        } while (reader.NextResult());

                  
                    }
                }

                return Ok("Datos importados correctamente");

            }
            catch (System.Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }
    }
}

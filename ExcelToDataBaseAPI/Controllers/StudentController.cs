using ExcelDataReader; //Libreria para leer archivos de Excel
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
                // Registra el proveedor de codificación necesario para leer archivos Excel en ciertos formatos
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                // Verifica si el archivo es nulo o está vacío
                if (file == null || file.Length == 0)
                {
                    return BadRequest("No hay archivo importado");
                }

                // Define la carpeta donde se guardarán los archivos subidos
                var uploadsFolder = $"{Directory.GetCurrentDirectory()}\\Uploads";

                // Crea la carpeta si no existe
                if (!Directory.Exists(uploadsFolder))
                {
                    Directory.CreateDirectory(uploadsFolder);
                }

                // Construye la ruta del archivo donde se almacenará temporalmente
                var filePath = Path.Combine(uploadsFolder, file.FileName);

                // Guarda el archivo en la carpeta Uploads
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                // Abre el archivo para lectura
                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
                {

                    // Crea un lector de Excel
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        bool isHeaderSkipped = false; // Variable para omitir la primera fila (encabezados)

                        do
                        {
                            while (reader.Read()) // Lee cada fila del archivo Excel
                            {
                                if (!isHeaderSkipped) // Si aún no se ha saltado la cabecera, lo hace
                                {
                                    isHeaderSkipped = true;
                                    continue;
                                }

                                // Crea un objeto Student y asigna valores de las celdas
                                Student student = new Student();

                                var nombre = reader.GetValue(1)?.ToString();

                                if (string.IsNullOrEmpty(nombre))
                                {
                                    return BadRequest("No puede haber datos vacios en los nombres");
                                }

                                student.Name = nombre;
                                student.Marks = Convert.ToInt32(reader.GetValue(2).ToString());

                                _context.Add(student);

                                _context.SaveChanges();
                            }
                        } while (reader.NextResult()); // Pasa a la siguiente hoja del archivo Excel si existe


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

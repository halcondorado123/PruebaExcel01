using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using PruebaExcel01.Models;

namespace PruebaExcel01.Controllers
{
    public class AsignaturasController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LlenarCeldas()
        {
            // Obtener el archivo Excel (puedes ajustar esto según tu implementación)
            using (var package = new ExcelPackage(new FileInfo("ruta/a/tu/archivo.xlsx")))
            {
                // Obtener la hoja de trabajo (ajusta "NombreDeTuHoja" al nombre real)
                var worksheet = package.Workbook.Worksheets["NombreDeTuHoja"];

                // Lista de asignaturas
                List<AsignaturasME> asignaturas = datosArrayTemp();

                // Definir la celda de inicio (por ejemplo, A1)
                int startRow = 1;
                int column = 1; // Columna A


                // Iterar sobre las asignaturas y llenar el rango de celdas
                for (int i = 0; i < asignaturas.Count; i++)
                {
                    worksheet.Cells[startRow + i, column].Value = asignaturas[i].Asignatura;
                    // Llena otras celdas según sea necesario
                }
            }

            return RedirectToAction("Index"); // O redirige a donde sea necesario
        }



        [HttpPost]
        private List<AsignaturasME> datosArrayTemp()
        {
            // Declaración y creación de un array de objetos Persona
            List<AsignaturasME> asignaturas = new List<AsignaturasME>();

            // Inicializando objetos en el array
            

            return asignaturas;
        }
    }
}

using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using PruebaExcel01.Models;

namespace PruebaExcel01.Controllers
{
    public class TinyFormatController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private void ContentTable(ExcelWorksheet worksheet)
        {
            AsignaturasME asignaturasME = new AsignaturasME();
            List<AsignaturasME> subjects = GetSubjects.SubjectGenerator();

            int row = 33;
            int column = 1;
            int subjectsPerColumn = 10; // Cambiar de columna después de 30 elementos
            int subjectCount = 0; // Contador para llevar el seguimiento de los elementos en una columna

            // Llena el archivo de Excel con los datos
            foreach (var subject in subjects)
            {
                if (subjectCount >= subjectsPerColumn)
                {
                    // Cambiar de columna después de 30 elementos
                    column += 7; // Cambia a la siguiente columna
                    row = 33; // Reinicia la fila
                    subjectCount = 0; // Reiniciar el contador
                }

                //if (row == 32 && column == 12)
                //{
                //    // Mueve el valor que debería estar en la columna 12 a la columna 13
                //    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                //}
                //else if (row == 32 && column == 16)
                //{
                //    // Mueve el valor que debería estar en la columna 16 a la columna 17
                //    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                //}
                //else
                //{

                //}



                row++; // Avanzar una fila

                subjectCount++;
            }
        }
    }
}

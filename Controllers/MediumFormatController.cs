using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using PruebaExcel01.Models;
using System.Collections.Generic;

namespace PruebaExcel01.Controllers
{
    public class MediumFormatController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private void ContentTable(ExcelWorksheet worksheet)
        {
            // ajuste de asignaturas por 54 celdas (medium) ----------------------------------------- //
            AsignaturasME asignaturasME = new AsignaturasME();
            List<AsignaturasME> subjects = GetSubjects.SubjectGenerator();

            int row = 33;
            int column = 1;
            int subjectspercolumn = 18; // cambiar de columna después de 30 elementos
            int subjectcount = 0; // contador para llevar el seguimiento de los elementos en una columna

            // llena el archivo de excel con los datos
            foreach (var subject in subjects)
            {
                if (subjectcount >= subjectspercolumn)
                {
                    // cambiar de columna después de 30 elementos
                    column += 7; // cambia a la siguiente columna
                    row = 33; // reinicia la fila
                    subjectcount = 0; // reiniciar el contador
                }

                if (row == 32 && column == 12)
                {
                    // mueve el valor que debería estar en la columna 12 a la columna 13
                    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                }
                else if (row == 32 && column == 16)
                {
                    // mueve el valor que debería estar en la columna 16 a la columna 17
                    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                }
                else
                {
                    worksheet.Cells[row, column].Value = subject.Numero[0];
                }

                worksheet.Cells[row, column + 1].Value = subject.Asignatura[0];
                worksheet.Cells[row, column + 2].Value = subject.Creditos[0];
                worksheet.Cells[row, column + 3].Value = subject.Semestre[0];
                worksheet.Cells[row, column + 4].Value = subject.CalificacionNumerica[0];
                worksheet.Cells[row, column + 5].Value = subject.CalificacionLiteral[0];
                worksheet.Cells[row, column + 6].Value = subject.Nivel[0];

                row++; // avanzar una fila

                subjectcount++;
            }
        }
    }
}

using Microsoft.AspNetCore.Mvc;
using PruebaExcel01.Models;

namespace PruebaExcel01.Controllers
{
    public class AsignaturasController1 : Controller
    {
        public IActionResult Index()
        {

        // Declaración y creación de un array de objetos Persona
        List<AsignaturasME> asignaturas = new List<AsignaturasME>();

            // Inicializando objetos en el array
            asignaturas[0] = new AsignaturasME
            {
                Numero = 1,
                Asignatura = "LÓGICA Y PENSAMIENTO MATEMÁTICO",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[1] = new AsignaturasME
            {
                Numero = 2,
                Asignatura = "PROYECTO DE VIDA",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[2] = new AsignaturasME
            {
                Numero = 3,
                Asignatura = "INFORMATICA Y CONVERGENCIA TECNOLÓGICA",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[3] = new AsignaturasME
            {
                Numero = 4,
                Asignatura = "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[4] = new AsignaturasME
            {
                Numero = 5,
                Asignatura = "HABILIDADES COMUNICATIVAS",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[5] = new AsignaturasME
            {
                Numero = 6,
                Asignatura = "FUNDAMENTOS DE ADMINISTRACIÓN",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[6] = new AsignaturasME
            {
                Numero = 7,
                Asignatura = "ATENCIÓN AL USUARIO",
                Creditos = 3,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[7] = new AsignaturasME
            {
                Numero = 8,
                Asignatura = "CONTABILIDAD BÁSICA",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };

            asignaturas[8] = new AsignaturasME
            {
                Numero = 9,
                Asignatura = "EXPLORAR PARA INVESTIGAR",
                Creditos = 2,
                Semestre = 1,
                CalificacionNumerica = 4.5,
                CalificacionLiteral = "CUATRO CINCO",
                Nivel = "Técnico Profesional"
            };


            return View(asignaturas);
        }
    }
}

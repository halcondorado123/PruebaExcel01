using OfficeOpenXml;
using System.Reflection.Emit;
using System.Security.Cryptography.Xml;

namespace PruebaExcel01.Models
{
    public class AsignaturasME
    {
        public int?[] Numero { get; set; } 
        public string[] Asignatura { get; set; }
        public int?[] Creditos { get; set; }
        public int?[] Semestre { get; set; }
        public double?[] CalificacionNumerica { get; set; }
        public string[] CalificacionLiteral { get; set; }
        public string[] Nivel { get; set; }

    }

    public class GetSubjects
    {
        public static void SubjectsGen()
        {
            // Llamamos al método SubjectGenerator y almacenamos el resultado en una lista
            var subjects = SubjectGenerator();

            // Hacer algo con la lista de asignaturas, como exportar a Excel
        }

        public static List<AsignaturasME> SubjectGenerator()
        {
            var subjects = new List<AsignaturasME>
            // Crear una lista de asignaturas
            {
                new AsignaturasME
                {
                    Numero = new int?[]{ 1 },
                    Asignatura = new string[] {"LÓGICA Y PENSAMIENTO MATEMÁTICO"},
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new AsignaturasME
                {
                    Numero = new int?[] { 2 },
                    Asignatura = new string[] { "PROYECTO DE VIDA" },
                    Creditos =  new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

                new AsignaturasME
                {
                    Numero = new int?[] { 3 },
                    Asignatura = new string[] { "INFORMATICA Y CONVERGENCIA TECNOLÓGICA" },
                    Creditos = new int?[] { 2 },
                    Semestre = new int?[] { 1 },
                    CalificacionNumerica = new double?[] { 4.5 },
                    CalificacionLiteral = new string[] { "CUATRO CINCO" },
                    Nivel = new string[] { "Técnico Profesional" }
                },

            new AsignaturasME
            {
                Numero = new int?[] { 4 },
                Asignatura = new string[] { "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 5 },
                Asignatura = new string[] { "HABILIDADES COMUNICATIVAS" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 6 },
                Asignatura = new string[] { "FUNDAMENTOS DE ADMINISTRACIÓN" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 7 },
                Asignatura = new string[] { "ATENCIÓN AL USUARIO" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 8 },
                Asignatura = new string[] { "CONTABILIDAD BÁSICA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 1 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 9 },
                Asignatura = new string[] { "EXPLORAR PARA INVESTIGAR" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 10 },
                Asignatura = new string[] { "MATEMATICA 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 11 },
                Asignatura = new string[] { "CONTABILIDAD DE ENTIDADES DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 12 },
                Asignatura = new string[] { "ADMINISTRACIÓN PÚBLICA DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 13 },
                Asignatura = new string[] { "SISTEMAS DE INFORMACIÓN EN SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 14 },
                Asignatura = new string[] { "FUNDAMENTOS DE ECONOMIA" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 15 },
                Asignatura = new string[] { "ELECTIVA DE FORMACION INTEGRAL I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 16 },
                Asignatura = new string[] { "MERCADEO EN EMPRESAS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 2 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 17 },
                Asignatura = new string[] { "SALUD Y SEGURIDAD SOCIAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 18 },
                Asignatura = new string[] { "ESTADISTICA I" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 19 },
                Asignatura = new string[] { "ANÁLISIS DE LA SALUD" },
                Creditos = new int?[] { 3 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 20 },
                Asignatura = new string[] { "LEGISLACIÓN LABORAL" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 21 },
                Asignatura = new string[] { "OPCIÓN DE GRADO TÉCNICO" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 22 },
                Asignatura = new string[] { "PRÁCTICA TÉCNICO PROFESIONAL" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 23 },
                Asignatura = new string[] { "ELECTIVA DE PROFUNDIZACIÓN 1" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 24 },
                Asignatura = new string[] { "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD" },
                Creditos = new int?[] { 2 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 25 },
                Asignatura = new string[] { "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)" },
                Creditos = new int?[] { 0 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            },

            new AsignaturasME
            {
                Numero = new int?[] { 26 },
                Asignatura = new string[] { "CÁTEDRA DE PENSAMIENTO CUNISTA I" },
                Creditos = new int?[] { 1 },
                Semestre = new int?[] { 3 },
                CalificacionNumerica = new double?[] { 4.5 },
                CalificacionLiteral = new string[] { "CUATRO CINCO" },
                Nivel = new string[] { "Técnico Profesional" }
            }

        };


            return subjects;
        }
    }

}



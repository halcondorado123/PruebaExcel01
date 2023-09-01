using OfficeOpenXml;

namespace PruebaExcel01.Models
{
    public class AsignaturasME
    {
        public int? Numero { get; set; } 
        public string Asignatura { get; set; }
        public int? Creditos { get; set; }
        public int? Semestre { get; set; }
        public double? CalificacionNumerica { get; set; }
        public string CalificacionLiteral { get; set; }
        public string Nivel { get; set; }

    }

    public class SubjectGenerator
    {
        private static int[] Numero = { 1, 2, 3, 4, 5, 6, 7, 8, 9 ,10 ,
                                        11, 12, 13, 14 , 15, 16, 17, 18,
                                        19, 20, 21, 22, 23, 24, 25, 26};

        private static string[] Asignatura = { "LÓGICA Y PENSAMIENTO MATEMÁTICO", "PROYECTO DE VIDA", "INFORMATICA Y CONVERGENCIA TECNOLÓGICA",
                                                "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL", "HABILIDADES COMUNICATIVAS", "FUNDAMENTOS DE ADMINISTRACIÓN",
                                                "ATENCIÓN AL USUARIO", "CONTABILIDAD BÁSICA" , "EXPLORAR PARA INVESTIGAR", 
                                                "MATEMATICA 1", "CONTABILIDAD DE ENTIDADES DE SALUD", "ADMINISTRACIÓN PÚBLICA DE SERVICIOS DE SALUD",
                                                "SISTEMAS DE INFORMACIÓN EN SALUD", "FUNDAMENTOS DE ECONOMIA", "ELECTIVA DE FORMACION INTEGRAL I",
                                                "MERCADEO EN EMPRESAS DE SALUD", "SALUD Y SEGURIDAD SOCIAL", "ESTADISTICA I",
                                                "ANÁLISIS DE LA SALUD", "LEGISLACIÓN LABORAL", "OPCIÓN DE GRADO TÉCNICO", "PRÁCTICA TÉCNICO PROFESIONAL",
                                                "ELECTIVA DE PROFUNDIZACIÓN 1", "GESTIÓN DE LA PRODUCCIÓN DE SERVICIOS DE SALUD", 
                                                "GESTIÓN DE SUMINISTROS EN SALUD (OPCIÓN PROP)", "CÁTEDRA DE PENSAMIENTO CUNISTA I"};

        private static int[] Creditos = { 2, 2, 2, 2, 2, 2, 3, 2, 2,
                                          2, 2, 3, 2, 2, 2, 2, 2, 2,
                                          3, 2, 2, 1, 2, 2, 0, 1 };

        private static int[] Semestre = {1, 1, 1, 1, 1, 1, 1, 1, 2,
                                         2, 2, 2, 2, 2, 2, 2, 3, 3,
                                         3, 3, 3, 3, 3, 3, 3, 3};


        // LIMITAR QUE LAS NOTAS NO SUPEREN EL 5.0 O SEAN INFERIORES A 0.0
        private static double[] CalificacionNumerica = { 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 
                                                        4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5,
                                                        4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5 }; // Este valor por defecto va a ser de 4.5

        private static string[] CalificacionLiteral = { "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO",
                                                        "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", 
                                                        "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO",
                                                        "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO",
                                                        "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO",
                                                        "CUATRO CINCO", "CUATRO CINCO" };

        private static string[] Nivel = {"Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional" };

        public AsignaturasME[] GenerateSubjects()
        {
            AsignaturasME[] materias = new AsignaturasME[Numero.Length];

            for (int i = 0; i < materias.Length; i++)
            {
                if (i < Numero.Length)
                { 
                    materias[i] = new AsignaturasME
                    {
                    Numero = Numero[i],
                    Asignatura = Asignatura[i],
                    Creditos = Creditos[i],
                    Semestre = Semestre[i],
                    CalificacionNumerica = CalificacionNumerica[i],
                    CalificacionLiteral = CalificacionLiteral[i],
                    Nivel = Nivel[i]
                    };
            }
                else
                {
                    // Asigna null a todas las propiedades si no hay más datos disponibles
                    materias[i] = new AsignaturasME
                    {
                        Numero = null,
                        Asignatura = null,
                        Creditos = null,
                        Semestre = null,
                        CalificacionNumerica = null,
                        CalificacionLiteral = null,
                        Nivel = null

                    };
                }
            }

            return materias;

        }
    }
}

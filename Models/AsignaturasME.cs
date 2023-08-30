using OfficeOpenXml;

namespace PruebaExcel01.Models
{
    public class AsignaturasME
    {
        public int Numero { get; set; } 
        public string Asignatura { get; set; }
        public int Creditos { get; set; }
        public int Semestre { get; set; }
        public double CalificacionNumerica { get; set; }
        public string CalificacionLiteral { get; set; }
        public string Nivel { get; set; }

    }

    public class SubjectGenerator
    {
        private static int[] Numero = { 1, 2, 3, 4, 5, 6, 7, 8, 9 };

        private static string[] Asignatura = { "LÓGICA Y PENSAMIENTO MATEMÁTICO", "PROYECTO DE VIDA", "INFORMATICA Y CONVERGENCIA TECNOLÓGICA",
                                                "TERMINOLOGÍA DE LA SEGURIDAD SOCIAL", "HABILIDADES COMUNICATIVAS", "FUNDAMENTOS DE ADMINISTRACIÓN",
                                                "ATENCIÓN AL USUARIO", "CONTABILIDAD BÁSICA" , "EXPLORAR PARA INVESTIGAR", "PEPITO JUAREZ" };

        private static int[] Creditos = { 2, 2, 2, 2, 2, 2, 3, 2, 2, 99};

        private static int[] Semestre = {1, 1, 1, 1, 1, 1, 1, 1, 2, 99};


        // LIMITAR QUE LAS NOTAS NO SUPEREN EL 5.0 O SEAN INFERIORES A 0.0
        private static double[] CalificacionNumerica = { 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 4.5, 5.0}; // Este valor por defecto va a ser de 4.5

        private static string[] CalificacionLiteral = { "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO",
                                                        "CUATRO CINCO", "CUATRO CINCO", "CUATRO CINCO", "NO SE"};
        
        private static string[] Nivel = { "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional",
                                        "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "Técnico Profesional", "MASTER" };

        public AsignaturasME[] GenerateSubjects()
        {
            AsignaturasME[] materias = new AsignaturasME[Numero.Length];

            for (int i = 0; i < materias.Length; i++)
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

            return materias;

        }
    }
}

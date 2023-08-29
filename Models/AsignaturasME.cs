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
}

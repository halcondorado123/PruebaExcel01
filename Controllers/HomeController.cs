using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using PruebaExcel01.Models;
using System;
using System.Diagnostics;


namespace PruebaExcel01.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        public ActionResult ExportToExcel()
        {
            return View();
        }


        // Todos los bordes
        public static void AplicarBordes(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
            CentrarContenido(range);
        }

        public static void AplicarBordesIzquierda(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
            IzquierdaContenido(range);
        }

        public static void AplicarBordeTipoFirma(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Bottom.Style = borderStyle;
        }


        public static void CentrarContenido(ExcelRange range)
        { 
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        public static void DerechaContenido(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        public static void IzquierdaContenido(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
        public static void AplicarNegrita(ExcelRange range, bool negrita = true)
        {
            range.Style.Font.Bold = negrita;
        }

        public static void FormatoGeneralTexto(ExcelRange range)
        {
            CentrarContenido(range);
            AplicarNegrita(range);
        }

        public static void FormatoTotalesLabel(ExcelRange range)
        {
            DerechaContenido(range);
            AplicarNegrita(range);
        }

        public static void UnirCeldas(ExcelRange range, bool merge = true)
        {
            range.Merge = merge;
        }

        public static ExcelRange GetExcelRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            return worksheet.Cells[startRow, startColumn, endRow, endColumn];
        }


        [HttpPost]
        public ActionResult Exportar()
        {


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var fechaHoraActual = DateTime.Now;
                string formatoFecha = "yyyyMMdd_HHmmss";
                string nombreArchivo = "Datos_Docs_Est_" + fechaHoraActual.ToString(formatoFecha) + ".xlsx";
                string nombreHoja = "docs_titulos_01" + fechaHoraActual;

                // Crear Hoja xlsx
                var worksheet = package.Workbook.Worksheets.Add("nombreHoja");

                //------------------- SHEET SIZE--------------------//
                //int columnaIndex = 1; // Índice de la columna que deseas modificar (empezando desde 1)
                //double nuevoAncho = 20; // Nuevo ancho de la columna en puntos

                //worksheet.Column(columnaIndex).Width = nuevoAncho;00
                // -------------------------------------------------- //

                // Labels Registro Fecha y Hora
                string fechaRegistro = "Fecha Registro:";
                string HoraRegistro = "Hora Registro:";

                // Agrega encabezados de columnas
                var celdaFecha = worksheet.Cells[1, 1];
                celdaFecha.Value = fechaRegistro;
                celdaFecha.Style.Font.Bold = true; // Aplica formato negrita
                worksheet.Column(1).Width = 5.89; // Cambia el valor 15 según tus necesidades
                //worksheet.Row(2).Height = 20; // Cambia el valor 20 según tus necesidades

                var celdaHora = worksheet.Cells[1, 2];
                celdaFecha.Value = fechaRegistro;
                celdaFecha.Style.Font.Bold = true; // Aplica formato negrita
                worksheet.Column(2).Width = 29.78; // Cambia el valor 15 según tus necesidades
                
                
                // Ajuste de Row para la firma
                var filaFirma = worksheet.Row(66);
                filaFirma.Height = 93.60;

                celdaHora.Value = HoraRegistro;
                celdaHora.Style.Font.Bold = true; // Aplica formato negrita


                //worksheet.Cells[2, 1].Value = fechaHoraActual.ToString("yyyy-MM-dd");
                //worksheet.Cells[2, 2].Value = fechaHoraActual.ToString("HH:mm:ss");
                //worksheet.Cells[5, 1].Value = "Nombre";
                //worksheet.Cells[5, 2].Value = "Correo electrónico";
                //worksheet.Cells[6, 1].Value = "Bruce Banner";
                //worksheet.Cells[6, 2].Value = "elvengadormasfuerte@gmail.com";

                //var logoCelda = worksheet.Cells[1, 2];
                //celdaFecha.Value = fechaRegistro;
                //celdaFecha.Style.Font.Bold = true; // Aplica formato negrita
                //worksheet.Column(1).Width = 29.78;


                // LogoCelda Square (Por parametrizar)
                ExcelRange logoSquare = GetExcelRange(worksheet, 1, 1, 5, 2);
                logoSquare.Value = "LOGOTIPO SAMPLE";
                UnirCeldas(logoSquare);
                CentrarContenido(logoSquare);

                // Header Rectangle
                ExcelRange headerRectangle = GetExcelRange(worksheet, 1, 3, 3, 24);
                headerRectangle.Value = "ACTA DE RECONOCIMIENTO DE TITULO";
                UnirCeldas(headerRectangle);
                FormatoGeneralTexto(headerRectangle);


                ExcelRange labelFormA = GetExcelRange(worksheet, 5, 8, 5, 10);
                labelFormA.Value = "INTERNA";
                UnirCeldas(labelFormA);
                FormatoGeneralTexto(labelFormA);

                ExcelRange txtbox1 = GetExcelRange(worksheet, 5, 11, 5, 12);
                // TARGET //
                UnirCeldas(txtbox1);
                AplicarBordes(txtbox1);

                ExcelRange labelFormB = GetExcelRange(worksheet, 5, 14, 5, 15);
                labelFormB.Value = "EXTERNA";
                UnirCeldas(labelFormB);
                FormatoGeneralTexto(labelFormB);

                ExcelRange txtbox2 = GetExcelRange(worksheet, 5, 16, 5, 17);
                // TARGET //               
                UnirCeldas(txtbox2);
                AplicarBordes(txtbox2);

                ExcelRange labelFormC = GetExcelRange(worksheet, 5, 20, 5, 22);
                labelFormC.Value = "PERIODO ACADÉMICO";
                UnirCeldas(labelFormC);
                FormatoGeneralTexto(labelFormC);

                var txtbox3 = worksheet.Cells[5, 23];
                // TARGET //               
                AplicarBordes(txtbox3);

                var labelFormD = worksheet.Cells[8, 7];
                labelFormD.Value = "MODALIDAD:";
                FormatoGeneralTexto(labelFormD);

                ExcelRange labelFormE = GetExcelRange(worksheet, 8, 8, 8, 9);
                labelFormE.Value = "DISTANCIA";
                UnirCeldas(labelFormE);
                FormatoGeneralTexto(labelFormE);

                var txtbox4 = worksheet.Cells[8, 10];
                // TARGET //               
                AplicarBordes(txtbox4);

                ExcelRange labelFormF = GetExcelRange(worksheet, 8, 14, 8, 15);
                labelFormF.Value = "PRESENCIAL";
                UnirCeldas(labelFormF);
                FormatoGeneralTexto(labelFormF);

                ExcelRange txtbox5 = GetExcelRange(worksheet, 8, 16, 8, 17);
                // TARGET //               
                UnirCeldas(txtbox5);
                AplicarBordes(txtbox5);

                ExcelRange labelFormG = GetExcelRange(worksheet, 8, 19, 8, 20);
                labelFormG.Value = "VIRTUAL";
                UnirCeldas(labelFormG);
                FormatoGeneralTexto(labelFormG);

                var txtbox6 = worksheet.Cells[8, 21];
                // TARGET //               
                AplicarBordes(txtbox6);

                // ------------------------------------------------------------------------------------- //

                ExcelRange txtbox7 = GetExcelRange(worksheet, 10, 3, 10, 5);
                // TARGET //               
                UnirCeldas(txtbox7);
                AplicarBordeTipoFirma(txtbox7);

                ExcelRange labelFormH = GetExcelRange(worksheet, 11, 3, 12, 5);
                labelFormH.Value = "REGIONAL - SEDE O CUNAD";
                UnirCeldas(labelFormH);
                FormatoGeneralTexto(labelFormH);

                ExcelRange txtbox8 = GetExcelRange(worksheet, 10, 8, 10, 10);
                // TARGET //               
                UnirCeldas(txtbox8);
                AplicarBordeTipoFirma(txtbox8);

                ExcelRange labelFormI = GetExcelRange(worksheet, 11, 8, 11, 10);
                labelFormI.Value = "FECHA DEL RECONOCIMIENTO";
                UnirCeldas(labelFormI);
                FormatoGeneralTexto(labelFormI);

                ExcelRange txtbox9 = GetExcelRange(worksheet, 10, 15, 10, 16);
                // TARGET //               
                UnirCeldas(txtbox9);
                AplicarBordeTipoFirma(txtbox9);

                ExcelRange labelFormJ = GetExcelRange(worksheet, 11, 15, 11, 16);
                labelFormJ.Value = "PLAN DE ESTUDIOS A APLICAR";
                UnirCeldas(labelFormJ);
                FormatoGeneralTexto(labelFormJ);

                ExcelRange txtbox10 = GetExcelRange(worksheet, 10, 21, 10, 22);
                // TARGET //               
                UnirCeldas(txtbox10);
                AplicarBordeTipoFirma(txtbox10);

                ExcelRange labelFormK = GetExcelRange(worksheet,11, 21, 11, 22);
                labelFormK.Value = "CÓDIGO DEL PLAN DE ESTUDIOS A APLICAR";
                UnirCeldas(labelFormK);
                FormatoGeneralTexto(labelFormK);

                // ------------------------------------------------------------------------------------- //


                ExcelRange txtbox11 = GetExcelRange(worksheet, 14, 3, 14, 5);
                // TARGET //               
                UnirCeldas(txtbox11);
                AplicarBordeTipoFirma(txtbox11);

                ExcelRange labelFormL = GetExcelRange(worksheet, 15, 3, 15, 5);
                labelFormL.Value = "APELLIDOS Y NOMBRES DEL ESTUDIANTE";
                UnirCeldas(labelFormL);
                FormatoGeneralTexto(labelFormL);

                ExcelRange txtbox12 = GetExcelRange(worksheet, 14, 8, 14, 10);
                // TARGET //               
                UnirCeldas(txtbox12);
                AplicarBordeTipoFirma(txtbox12);

                ExcelRange labelFormM = GetExcelRange(worksheet, 15, 8, 15, 10);
                labelFormM.Value = "DOCUMENTO DE IDENTIDAD";
                UnirCeldas(labelFormM);
                FormatoGeneralTexto(labelFormM);

                ExcelRange txtbox13 = GetExcelRange(worksheet, 14, 15, 14, 16);
                // TARGET //               
                UnirCeldas(txtbox13);
                AplicarBordeTipoFirma(txtbox13);

                ExcelRange labelFormO = GetExcelRange(worksheet, 15, 15, 15, 16);
                labelFormO.Value = "CORREO ELECTRONICO";
                UnirCeldas(labelFormO);
                FormatoGeneralTexto(labelFormO);

                ExcelRange txtbox14 = GetExcelRange(worksheet, 14, 21, 14, 22);
                // TARGET //               
                UnirCeldas(txtbox14);
                AplicarBordeTipoFirma(txtbox14);

                ExcelRange labelFormP = GetExcelRange(worksheet, 15, 21, 15, 22);
                labelFormP.Value = "TELEFONO FIJO - CELULAR";
                UnirCeldas(labelFormP);
                FormatoGeneralTexto(labelFormP);

                // ------------------------------------------------------------------------------------- //

                ExcelRange txtbox15 = GetExcelRange(worksheet, 18, 3, 18, 5);
                // TARGET //               
                UnirCeldas(txtbox15);
                AplicarBordeTipoFirma(txtbox15);

                ExcelRange labelFormQ = GetExcelRange(worksheet, 19, 3, 19, 5);
                labelFormQ.Value = "INSTITUCIÓN DE DONDE PROVIENE";
                UnirCeldas(labelFormQ);
                FormatoGeneralTexto(labelFormQ);

                ExcelRange txtbox16 = GetExcelRange(worksheet, 18, 8, 18, 10);
                // TARGET //               
                UnirCeldas(txtbox16);
                AplicarBordeTipoFirma(txtbox16);

                ExcelRange labelFormR = GetExcelRange(worksheet, 19, 8, 19, 10);
                labelFormR.Value = "PROGRAMA CURSADO";
                UnirCeldas(labelFormR);
                FormatoGeneralTexto(labelFormR);

                ExcelRange txtbox17 = GetExcelRange(worksheet, 18, 14, 18, 16);
                // TARGET //               
                UnirCeldas(txtbox17);
                AplicarBordeTipoFirma(txtbox17);

                ExcelRange labelFormS = GetExcelRange(worksheet, 19, 14, 19, 16);
                labelFormS.Value = "PROGRAMA A CURSAR";
                UnirCeldas(labelFormS);
                FormatoGeneralTexto(labelFormS);


                // ------------------------------------------------------------------------------------- //

                var numbSena = worksheet.Cells[22, 1];
                int startRow = 22; // Fila inicial
                for (int i = 1; i <= 7; i++)
                {
                    var cell = worksheet.Cells[startRow + i - 1, 1];
                    cell.Value = i;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                ExcelRange labelTitle1 = GetExcelRange(worksheet, 21, 1, 21, 24);
                labelTitle1.Value = "ESTRUCTURA CURRICULAR SENA";
                UnirCeldas(labelTitle1);
                AplicarBordes(labelTitle1);
                FormatoGeneralTexto(labelTitle1);

                ExcelRange Context1 = GetExcelRange(worksheet, 22, 2, 22, 24);
                Context1.Value = "ADMITIR AL USUARIO EN LA RED DE SERVICIOS DE SALUD SEGÚN NIVELES DE ATENCIÓN Y NORMATIVA VIGENTE.";
                Context1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                UnirCeldas(Context1);

                ExcelRange Context2 = GetExcelRange(worksheet, 23, 2, 23, 24);
                UnirCeldas(Context2);
                Context2.Value = "AFILIAR A LA POBLACIÓN AL SISTEMA GENERAL DE SEGURIDAD SOCIAL EN SALUD SEGÚN NORMATIVIDAD VIGENTE.";
                Context2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ExcelRange Context3 = GetExcelRange(worksheet, 24, 2, 24, 24);
                UnirCeldas(Context3);
                Context3.Value = "FACTURAR LA PRESTACIÓN DE LOS SERVICIOS DE SALUD SEGÚN NORMATIVIDAD Y CONTRATACIÓN";
                Context3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ExcelRange Context4 = GetExcelRange(worksheet, 25, 2, 25, 24);
                UnirCeldas(Context4);
                Context4.Value = "MANEJAR VALORES E INGRESOS RELACIONADOS CON LA OPERACIÓN DEL ESTABLECIMIENTO. (EQUIVALE A LA NORMA NTS 005 DEL MINCOMERCIO, INDUSTRIA Y TURISMO)";
                Context4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ExcelRange Context5 = GetExcelRange(worksheet, 26, 2, 26, 24);
                UnirCeldas(Context5);
                Context5.Value = "ORIENTAR AL USUARIO EN RELACIÓN CON SUS NECESIDADES Y EXPECTATIVAS DE ACUERDO CON POLÍTICAS INSTITUCIONALES Y NORMAS DE SALUD VIGENTES.";
                Context5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ExcelRange Context6 = GetExcelRange(worksheet, 27, 2, 27, 24);
                UnirCeldas(Context6);
                Context6.Value = "PROMOVER LA INTERACCION IDONEA CONSIGO MISMO, CON LOS DEMAS Y CON LA NATURALEZA EN LOS CONTEXTOS LABORAL Y SOCIAL.";
                Context6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                ExcelRange Context7 = GetExcelRange(worksheet, 28, 2, 28, 24);
                UnirCeldas(Context7);
                Context7.Value = "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA";
                Context7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                // ----------------------- INSCRIPCION DE MATERIAS (TABLA)------------------------------ //

                ExcelRange Celda1 = GetExcelRange(worksheet, 31, 1, 32, 1); ;
                Celda1.Value = "No";
                UnirCeldas(Celda1);
                AplicarBordes(Celda1);
                FormatoGeneralTexto(Celda1);

                ExcelRange Celda2 = GetExcelRange(worksheet, 31, 2, 32, 2);
                Celda2.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                UnirCeldas(Celda2);
                AplicarBordes(Celda2);
                FormatoGeneralTexto(Celda2);

                ExcelRange Celda3 = GetExcelRange(worksheet, 31, 3, 31, 4);
                Celda3.Value = "SISTEMA";
                UnirCeldas(Celda3);
                AplicarBordes(Celda3);
                FormatoGeneralTexto(Celda3);

                    var SubCelda1 = worksheet.Cells[32, 3];
                    SubCelda1.Value = "Créditos";
                    UnirCeldas(SubCelda1);
                    AplicarBordes(SubCelda1);
                    FormatoGeneralTexto(SubCelda1);

                    var SubCelda2 = worksheet.Cells[32, 4];
                    SubCelda2.Value = "Semestre";
                    UnirCeldas(SubCelda2);
                    AplicarBordes(SubCelda2);
                    FormatoGeneralTexto(SubCelda2);

                ExcelRange Celda4 = GetExcelRange(worksheet, 31, 5, 32, 5);
                Celda4.Value = "CALIFICACIÓN NUMERICA";
                UnirCeldas(Celda4);
                AplicarBordes(Celda4);
                FormatoGeneralTexto(Celda4);

                ExcelRange Celda5 = GetExcelRange(worksheet, 31, 6, 32, 6);
                Celda5.Value = "CALIFICACION LITERAL";
                UnirCeldas(Celda5);
                AplicarBordes(Celda5);
                FormatoGeneralTexto(Celda5);

                ExcelRange Celda6 = GetExcelRange(worksheet, 31, 7, 32, 7);
                Celda6.Value = "NIVEL";
                UnirCeldas(Celda6);
                AplicarBordes(Celda6);
                FormatoGeneralTexto(Celda6);

                ExcelRange Celda7 = GetExcelRange(worksheet, 31, 8, 32, 8);
                Celda7.Value = "No";
                UnirCeldas(Celda7);
                AplicarBordes(Celda7);
                FormatoGeneralTexto(Celda7);

                ExcelRange Celda8 = GetExcelRange(worksheet, 31, 9, 32, 10);
                Celda8.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                UnirCeldas(Celda8);
                AplicarBordes(Celda8);
                FormatoGeneralTexto(Celda8);

                ExcelRange Celda9 = GetExcelRange(worksheet, 31, 11, 31, 12);
                Celda9.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                UnirCeldas(Celda9);
                AplicarBordes(Celda9);
                FormatoGeneralTexto(Celda9);

                    var SubCelda3 = worksheet.Cells[32, 11];
                    SubCelda3.Value = "Créditos";
                    UnirCeldas(SubCelda3);
                    AplicarBordes(SubCelda3);
                    FormatoGeneralTexto(SubCelda3);

                    var SubCelda4 = worksheet.Cells[32, 12];
                    SubCelda4.Value = "Semestre";
                    UnirCeldas(SubCelda4);
                    AplicarBordes(SubCelda4);
                    FormatoGeneralTexto(SubCelda4);

                // ------------------ //

                ExcelRange Celda10 = GetExcelRange(worksheet, 31, 13, 32, 13);
                Celda10.Value = "CALIFICACIÓN NUMERICA";
                UnirCeldas(Celda10);
                AplicarBordes(Celda10);
                FormatoGeneralTexto(Celda10);

                ExcelRange Celda11 = GetExcelRange(worksheet, 31, 14, 32, 14);
                Celda11.Value = "CALIFICACION LITERAL";
                UnirCeldas(Celda11);
                AplicarBordes(Celda11);
                FormatoGeneralTexto(Celda11);

                ExcelRange Celda12 = GetExcelRange(worksheet, 31, 15, 32, 15);
                Celda12.Value = "NIVEL";
                UnirCeldas(Celda12);
                AplicarBordes(Celda12);
                FormatoGeneralTexto(Celda12);

                ExcelRange Celda13 = GetExcelRange(worksheet, 31, 16, 32, 16);
                Celda13.Value = "No";
                UnirCeldas(Celda13);
                AplicarBordes(Celda13);
                FormatoGeneralTexto(Celda13);

                // ------------------------------------------------------------------------------------- //

                ExcelRange Celda14 = GetExcelRange(worksheet, 31, 17, 32, 19);
                Celda14.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                UnirCeldas(Celda14);
                AplicarBordes(Celda14);
                FormatoGeneralTexto(Celda14);

                ExcelRange Celda15 = GetExcelRange(worksheet, 31, 20, 31, 21);
                Celda15.Value = "SISTEMA";
                UnirCeldas(Celda15);
                AplicarBordes(Celda15);
                FormatoGeneralTexto(Celda15);

                    var SubCelda5 = worksheet.Cells[32, 20];
                    SubCelda5.Value = "Créditos";
                    UnirCeldas(SubCelda5);
                    AplicarBordes(SubCelda5);
                    FormatoGeneralTexto(SubCelda5);

                    var SubCelda6 = worksheet.Cells[32, 21];
                    SubCelda6.Value = "Semestre";
                    UnirCeldas(SubCelda6);
                    AplicarBordes(SubCelda6);
                    FormatoGeneralTexto(SubCelda6);


                ExcelRange Celda16 = GetExcelRange(worksheet, 31, 22, 32, 22);
                Celda16.Value = "CALIFICACIÓN NUMERICA";
                UnirCeldas(Celda16);
                AplicarBordes(Celda16);
                FormatoGeneralTexto(Celda16);

                ExcelRange Celda17 = GetExcelRange(worksheet, 31, 23, 32, 23);
                Celda17.Value = "CALIFICACION LITERAL";
                UnirCeldas(Celda17);
                AplicarBordes(Celda17);
                FormatoGeneralTexto(Celda17);

                ExcelRange Celda18 = GetExcelRange(worksheet, 31, 24, 32, 24);
                Celda18.Value = "NIVEL";
                UnirCeldas(Celda18);
                AplicarBordes(Celda18);
                FormatoGeneralTexto(Celda18);

                ExcelRange celdasMaterias = GetExcelRange(worksheet, 33, 1, 41, 24);
                AplicarBordes(celdasMaterias);

                var numbCeldasMaterias = worksheet.Cells[33, 1];
                int celdaInicial= 33; // Fila inicial
                for (int i = 1; i <= 9; i++)
                {
                    var cell = worksheet.Cells[celdaInicial + i - 1, 1];
                    cell.Value = i;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                // ------------------------------  TOTALES MATERIAS ------------------------------------------ //

                ExcelRange LabelTotal = GetExcelRange(worksheet, 42, 1, 42, 2);
                LabelTotal.Value = "TOTALES";
                UnirCeldas(LabelTotal);
                AplicarBordes(LabelTotal);
                FormatoGeneralTexto(LabelTotal);

                var rangeToSum1 = worksheet.Cells["C33:C41"]; // Cambia esto al rango que necesitas
                // Suma los valores en el rango
                var sumCell1 = worksheet.Cells["C42"]; // Cambia esto a la celda donde deseas mostrar la suma

                var totalMaterias1 = worksheet.Cells[42, 3];
                sumCell1.Formula = $"SUM({rangeToSum1.Address})";
                AplicarBordes(totalMaterias1);
                FormatoGeneralTexto(totalMaterias1);

                var rangeToSum2 = worksheet.Cells["K33:K41"]; // Cambia esto al rango que necesitas
                // Suma los valores en el rango
                var sumCell2 = worksheet.Cells["K42"]; // Cambia esto a la celda donde deseas mostrar la suma

                var totalMaterias2 = worksheet.Cells[42, 11];
                sumCell2.Formula = $"SUM({rangeToSum2.Address})";
                AplicarBordes(totalMaterias2);
                FormatoGeneralTexto(totalMaterias2);

                var rangeToSum3 = worksheet.Cells["T33:T41"]; // Cambia esto al rango que necesitas
                // Suma los valores en el rango
                var sumCell3 = worksheet.Cells["T42"]; // Cambia esto a la celda donde deseas mostrar la suma

                var totalMaterias3 = worksheet.Cells[42, 20];
                sumCell3.Formula = $"SUM({rangeToSum3.Address})";
                AplicarBordes(totalMaterias3);
                FormatoGeneralTexto(totalMaterias3);


                    ExcelRange LabelTotal1 = GetExcelRange(worksheet, 46, 1, 46, 7);
                    LabelTotal1.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TÉCNICO PROFESIONAL";
                    UnirCeldas(LabelTotal1);
                    AplicarBordes(LabelTotal1);
                    FormatoTotalesLabel(LabelTotal1);

                    ExcelRange LabelTotal2 = GetExcelRange(worksheet, 47, 1, 47, 7);
                    LabelTotal2.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TECNOLÓGICO ";
                    UnirCeldas(LabelTotal2);
                    AplicarBordes(LabelTotal2);
                    FormatoTotalesLabel(LabelTotal2);

                    ExcelRange LabelTotal3 = GetExcelRange(worksheet, 48, 1, 48, 7);
                    LabelTotal3.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL PROFESIONAL";
                    UnirCeldas(LabelTotal3);
                    AplicarBordes(LabelTotal3);
                    FormatoTotalesLabel(LabelTotal3);


                ExcelRange LabelTotalMaterias = GetExcelRange(worksheet, 44, 8, 44, 10);
                UnirCeldas(LabelTotalMaterias);
                AplicarBordes(LabelTotalMaterias);
                FormatoGeneralTexto(LabelTotalMaterias);

                var cellTotal1 = worksheet.Cells["C42"];
                var cellTotal2 = worksheet.Cells["K42"];
                var cellTotal3 = worksheet.Cells["T42"];

                // Define la celda donde mostrarás el resultado de la suma
                var sumCellTotales = worksheet.Cells["H44"];

                // Aplica la fórmula de suma a la celda de resultado
                sumCellTotales.Formula = $"SUM({cellTotal1.Address},{cellTotal2.Address},{cellTotal3.Address})";


                var LabelAprobado = worksheet.Cells[45, 8];
                LabelAprobado.Value = "APRO";
                UnirCeldas(LabelAprobado);
                AplicarBordes(LabelAprobado);
                FormatoGeneralTexto(LabelAprobado);

                var LabelPendiente = worksheet.Cells[45, 9];
                LabelPendiente.Value = "PEN";
                UnirCeldas(LabelPendiente);
                AplicarBordes(LabelPendiente);
                FormatoGeneralTexto(LabelPendiente);

                var LabelTotalCreditos = worksheet.Cells[45, 10];
                LabelTotalCreditos.Value = "TOTAL CRED";
                UnirCeldas(LabelTotalCreditos);
                AplicarBordes(LabelTotalCreditos);
                FormatoGeneralTexto(LabelTotalCreditos);

                ExcelRange celdasMateriasTotales = GetExcelRange(worksheet, 46, 8, 48, 10);
                AplicarBordes(celdasMateriasTotales);

                // CREDITOS TECNICO PROFESIONAL
                var TotalCreditosTecnico = worksheet.Cells["J46"];
                var TotalAprobadoTecnico = worksheet.Cells["H46"];
                var TotalPendienteTecnico = worksheet.Cells["I46"];
                TotalPendienteTecnico.Formula = $"({TotalCreditosTecnico.Address}) - ({TotalAprobadoTecnico.Address})";

                // TOTAL CREDITOS APROBADOS
                var TotalCreditosTecnologico = worksheet.Cells["J47"];
                var TotalAprobadoTecnologico = worksheet.Cells["H47"];
                var TotalPendienteTecnologico = worksheet.Cells["I47"];
                TotalPendienteTecnologico.Formula = $"({TotalCreditosTecnologico.Address}) - ({TotalAprobadoTecnologico.Address})";

                // TOTAL CREDITOS PENDIENTES
                var TotalCreditosProfesional = worksheet.Cells["J48"];
                var TotalAprobadoProfesional = worksheet.Cells["H48"];
                var TotalPendienteProfesional = worksheet.Cells["I48"];
                TotalPendienteProfesional.Formula = $"({TotalCreditosProfesional.Address}) - ({TotalAprobadoProfesional.Address})";


                // IMPORTANTE: CODIGO PROPENSO A SER MODIFICADO //
                var cellsToSetZero = new List<ExcelRange>
                {
                    worksheet.Cells["H46"],
                    worksheet.Cells["J46"],
                    worksheet.Cells["H47"],
                    worksheet.Cells["J47"],
                    worksheet.Cells["H48"],
                    worksheet.Cells["J48"]
                };

                // Asigna el valor cero a todas las celdas usando un bucle
                foreach (var cell in cellsToSetZero)
                {
                    cell.Value = 0;
                }

                // --------------------------------------------------------------------- //

                // --------------------- PARRAFO ALCARACIONES --------------------------- //

                // Unir celdas horizontalmente y verticalmente
                ExcelRange labelTitle2 = worksheet.Cells[50, 1, 50, 2];
                labelTitle2.Value = "Aclaraciones";
                AplicarNegrita(labelTitle2);
                UnirCeldas(labelTitle2);
                IzquierdaContenido(labelTitle2);

                ExcelRange ContextA = worksheet.Cells[51, 1, 51, 24];
                ContextA.Value = "Los créditos académicos faltantes para cumplir a cabalidad con la oferta académica del nivel técnico profesional y/o tecnológico y/o  profesional deben ser cursados y aprobados conforme las reglamentaciones institucionales vigentes.";
                UnirCeldas(ContextA);
                IzquierdaContenido(ContextA);

                ExcelRange ContextB = worksheet.Cells[52, 1, 52, 24];
                ContextB.Value = "La Escuela de Ciencias Administrativas de la Corporación Unificada Nacional -CUN, reconoce las asignaturas del programa de  Administración de Servicios de Salud   del nivel técnico ( 1 - 3  semestre) para que dar continuidad a su proceso de formación académica a partir de las asignaturas de nivel tecnológico correspondientes a (  4 - 5  semestre) Y profesional correspondientes a ( 6 - 9  semestre)";
                UnirCeldas(ContextB);
                IzquierdaContenido(ContextB);

                ExcelRange ContextC = worksheet.Cells[53, 1, 53, 24];
                ContextC.Value = "La prueba TyT del ciclo técnico es homologable en la institución. El estudiante deberá presentar la prueba saber TyT para el ciclo tecnológico y Saber PRO para el ciclo profesional.";
                UnirCeldas(ContextC);
                IzquierdaContenido(ContextC);

                ExcelRange ContextD = worksheet.Cells[54, 1, 54, 24];
                ContextD.Value = "Teniendo en cuenta que el plan de estudios vigente del programa  Administración de Servicios de Salud   no incluye los respectivos niveles de inglés requeridos para obtener las diferentes tituluaciones, el estudiante deberá garantizar lo pertinente al momento de radicar su solilctud de grado, para ello se cuenta con la oferta del centro de Idiomas de la intitución.";
                UnirCeldas(ContextD);
                IzquierdaContenido(ContextD);

                ExcelRange labelTitle3 = worksheet.Cells[56, 1, 56, 2];
                labelTitle3.Value = "Manifestación expresa del estudiante";
                AplicarNegrita(labelTitle3);
                UnirCeldas(labelTitle3);
                IzquierdaContenido(labelTitle3);

                ExcelRange finalParraf = worksheet.Cells[57, 1, 57, 24];
                finalParraf.Value = "Con el presente documento manifiesto expresamente y sin que medie ninguna clase de vicio o limitación a mi consentimiento, mi plena conformidad con las asignaturas y/o créditos reconocidos u homologados para mi ingreso al nivel técnico profesional y/o tecnológico y/o profesional del programa   Administración de Servicios de Salud    de la Corporación Unificada Nacional de Educación Superior CUN. Las competencias que considere me hagan falta, del ciclo técnico, las podré realizar voluntariamente a través de tutorías en cada área transversal o del programa, talleres nivelatorios y/o participando como asistente a clases sin que estos generen nota alguna y solicitando previamente el ingreso a la clase o tutoría.";
                UnirCeldas(finalParraf);
                IzquierdaContenido(finalParraf);

                var labelTitle4 = worksheet.Cells[63, 1];
                labelTitle4.Value = "En Constancia de lo anterior firman:";
                AplicarNegrita(labelTitle3);
                UnirCeldas(labelTitle3);
                IzquierdaContenido(labelTitle3);

                // -------------------------------------------------------------------------- //

                // ----------------------------- SIGNATURE INFO ----------------------------- //

                string imagePath = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\lennon_signature.jpg";
                int widthInPixels = 230;
                int heightInPixels = 70;

                var picture = worksheet.Drawings.AddPicture("Firma", new FileInfo(imagePath));

                picture.SetPosition(66, -80, 1, -30);
                picture.SetSize(widthInPixels, heightInPixels);
                picture.Locked = true;

                ExcelRange cellSignatureLiderPrograma = worksheet.Cells[66, 1, 66, 2];
                AplicarBordeTipoFirma(cellSignatureLiderPrograma);

                ExcelRange labelSignature1 = worksheet.Cells[67, 1, 67, 3];
                labelSignature1.Value = "Líder de Programa";
                UnirCeldas(labelSignature1);
                IzquierdaContenido(labelSignature1);

                ExcelRange labelSignature2 = worksheet.Cells[68, 1, 68, 3];
                labelSignature2.Value = "Nombre: SAMPLE NAME";  // Convertir y generar valor dinámico
                UnirCeldas(labelSignature2);
                IzquierdaContenido(labelSignature2);

                var parrafSquare = worksheet.Cells[66, 7];
                parrafSquare.Value = "finalParraf";
                CentrarContenido(parrafSquare);
                UnirCeldas(parrafSquare);

                // Asignar un valor dinámico para la firma (ESTEBAN)

                ExcelRange cellSignatureStudent = worksheet.Cells[66, 9, 66, 12];
                AplicarBordeTipoFirma(cellSignatureStudent);

                ExcelRange labelSignature3 = worksheet.Cells[67, 9, 67, 12];
                labelSignature3.Value = "Estudiante: "; // Convertir y generar valor dinámico
                UnirCeldas(labelSignature3);
                IzquierdaContenido(labelSignature3);

                ExcelRange labelSignature4 = worksheet.Cells[68, 9, 68, 12];
                labelSignature4.Value = "Nombre: ";  // Convertir y generar valor dinámico
                UnirCeldas(labelSignature4);
                IzquierdaContenido(labelSignature4);

                ExcelRange labelSignature5 = worksheet.Cells[69, 9, 69, 12];
                labelSignature5.Value = "Doc de Identidad: "; // Convertir y generar valor dinámico
                UnirCeldas(labelSignature5);
                IzquierdaContenido(labelSignature5);

                // ------------------------------ FOOTER -----------------------------------//
                ExcelRange CellFooter1 = worksheet.Cells[71, 1, 71, 4];
                CellFooter1.Value = "ELABORÓ: "; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter1);
                UnirCeldas(CellFooter1);

                ExcelRange CellFooter2 = worksheet.Cells[71, 5, 71, 10];
                CellFooter2.Value = "FECHA"; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter2);
                UnirCeldas(CellFooter2);

                ExcelRange CellFooter3 = worksheet.Cells[71, 11, 71, 15];
                CellFooter3.Value = "REVISÓ: "; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter3);
                UnirCeldas(CellFooter3);

                ExcelRange CellFooter4 = worksheet.Cells[71, 11, 71, 15];
                CellFooter4.Value = "FECHA"; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter4);
                UnirCeldas(CellFooter4);

                ExcelRange CellFooter5 = worksheet.Cells[69, 9, 69, 12];
                CellFooter5.Value = "APROBÓ: "; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter5);
                UnirCeldas(CellFooter5);

                ExcelRange CellFooter6 = worksheet.Cells[71, 11, 71, 15];
                CellFooter6.Value = "FECHA: "; // Convertir y generar valor dinámico
                AplicarBordesIzquierda(CellFooter6);
                UnirCeldas(CellFooter6);




                // ------------------------------ END FOOTER --------------------------------//


                var filePath = @"C:\Users\Jhonattan_Casallas\Downloads\" + nombreArchivo;
                package.SaveAs(new System.IO.FileInfo(filePath));




                return File(filePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);
            }
        }
    }
}
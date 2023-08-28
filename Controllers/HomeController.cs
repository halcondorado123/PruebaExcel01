using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using PruebaExcel01.Models;
using System;
using System.Diagnostics;
using System.Drawing;

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

        // Propiedades generales del texto
        public static void ContentCenter(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // DerechaContenido
        public static void ContentRight(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // IzquierdaContenido
        public static void ContentLeft (ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // Aplicar Negrilla
        public static void FontWeightBold (ExcelRange range, bool bold = true)
        {
            range.Style.Font.Bold = bold;
        }

        public static void WhiteColor(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Font.Color.SetColor(System.Drawing.Color.White);
        }


        // Propiedades de celda

        // Bordes Tabla (4 bordes)
        public static void ApplyBorders(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Top.Style = borderStyle;
            range.Style.Border.Left.Style = borderStyle;
            range.Style.Border.Right.Style = borderStyle;
            range.Style.Border.Bottom.Style = borderStyle;
        }

        // Bordes Tipo Firma (Borde inferior)
        public static void ApplySignatureBorders(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            range.Style.Border.Bottom.Style = borderStyle;
        }

        // Unir celdas
        public static void MergedCells(ExcelRange range, bool merge = true)
        {
            range.Merge = merge;
        }

        // Aplicar bordes + Centrado Texto
        public static void CellCenter(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentCenter(range); 
        }

        // Aplicar bordes + Texto Derecha
        public static void CellRight(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentRight(range);
        }

        // Aplicar bordes + Texto Izquierda
        public static void CellLeft(ExcelRange range, ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin)
        {
            ApplyBorders(range);
            ContentLeft(range);
        }

        
        // Seleccion de celdas por rango
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

                #region infoDownloadDoc

                // Crea el nombre de mi documento 
                string nombreArchivo = "Datos_Docs_Est_" + fechaHoraActual.ToString(formatoFecha) + ".xlsx";
                string nombreHoja = "docs_titulos_01" + fechaHoraActual;

                // Crear Hoja xlsx

                #endregion
               
                var worksheet = package.Workbook.Worksheets.Add("nombreHoja");

                // Manejo de tamaños de celdas [Rows][Columns] //
                // Índices de las columnas con diferentes anchos
                int[] columnIndices = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24};

                // Anchura respectiva para cada columna
                double[] columnWidths = { 5.89, 29.78, 10.78, 10.22, 14.56, 14.78, 14.89, 7.56, 12.56, 14.89,
                                        9.56, 10.78, 14.33, 14.22, 14.22, 6.33, 9.89, 9.89, 7.22, 8.33,
                                        13.56, 16.56, 14.89, 13.56,};

                // Ajustar los anchos de las columnas en función de los índices y los anchos
                for (int i = 0; i < columnIndices.Length; i++)
                {
                    int columnIndex = columnIndices[i];
                    double columnWidth = columnWidths[i];
                    worksheet.Column(columnIndex).Width = columnWidth;
                }

                //// Agrega encabezados de columnas
                //var celdaFecha = worksheet.Cells[1, 1];
                //celdaFecha.Value = fechaRegistro;
                //celdaFecha.Style.Font.Bold = true; // Aplica formato negrita
                //worksheet.Column(1).Width = 5.89; // Cambia el valor 15 según tus necesidades
                ////worksheet.Row(2).Height = 20; // Cambia el valor 20 según tus necesidades

                //var celdaHora = worksheet.Cells[1, 2];
                //celdaFecha.Value = fechaRegistro;
                //celdaFecha.Style.Font.Bold = true; // Aplica formato negrita
                //worksheet.Column(2).Width = 29.78; // Cambia el valor 15 según tus necesidades


                //// Ajuste de Row para la firma
                //var filaFirma = worksheet.Row(66);
                //filaFirma.Height = 93.60;

                //celdaHora.Value = HoraRegistro;
                //celdaHora.Style.Font.Bold = true; // Aplica formato negrita


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

                string imagePathLogo = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\log1.png";
                int widthLogoInPixels = 255;
                int heightLogoInPixels = 98;

                var pictureLogo = worksheet.Drawings.AddPicture("Logo", new FileInfo(imagePathLogo));

                pictureLogo.SetPosition(1, -20, 1, -50);
                pictureLogo.SetSize(widthLogoInPixels, heightLogoInPixels);
                pictureLogo.Locked = true;

                string fechaRegistro = "Fecha Registro / Hora Registro";
                
                ExcelRange logoSpace = GetExcelRange(worksheet, 1, 1, 1, 2);
                logoSpace.Value = fechaRegistro;
                MergedCells(logoSpace);
                ContentCenter(logoSpace);
                WhiteColor(logoSpace);

                ExcelRange logoDate = GetExcelRange(worksheet, 2, 1, 5, 2);
                logoDate.Value = fechaHoraActual.ToString("yyyy-MM-dd \n" + fechaHoraActual.ToString("HH:mm:ss"));
                MergedCells(logoDate);
                ContentCenter(logoDate);
                WhiteColor(logoDate);


                //ExcelRange logoTime = GetExcelRange(worksheet, 3, 1, 3, 2);
                //logoTime.Value = fechaHoraActual.ToString("yyyy-MM-dd");
                //MergedCells(logoTime);
                ////CellCenter(logoTime);


                // Header Rectangle
                ExcelRange headerRectangle = GetExcelRange(worksheet, 1, 3, 3, 24);
                headerRectangle.Value = "ACTA DE RECONOCIMIENTO DE TITULO";
                CellCenter(headerRectangle);
                MergedCells(headerRectangle);
                FontWeightBold(headerRectangle);


                ExcelRange labelFormA = GetExcelRange(worksheet, 5, 8, 5, 10);
                labelFormA.Value = "INTERNA";
                MergedCells(labelFormA);
                FontWeightBold(labelFormA);
                ContentCenter(labelFormA);

                ExcelRange txtbox1 = GetExcelRange(worksheet, 5, 11, 5, 12);
                // TARGET //
                MergedCells(txtbox1);
                ApplyBorders(txtbox1);

                ExcelRange labelFormB = GetExcelRange(worksheet, 5, 14, 5, 15);
                labelFormB.Value = "EXTERNA";
                MergedCells(labelFormB);
                FontWeightBold(labelFormB);
                ContentCenter(labelFormB);

                ExcelRange txtbox2 = GetExcelRange(worksheet, 5, 16, 5, 17);
                // TARGET //               
                MergedCells(txtbox2);
                ApplyBorders(txtbox2);

                ExcelRange labelFormC = GetExcelRange(worksheet, 5, 20, 5, 22);
                labelFormC.Value = "PERIODO ACADÉMICO";
                MergedCells(labelFormC);
                FontWeightBold(labelFormC);
                ContentCenter(labelFormC);

                var txtbox3 = worksheet.Cells[5, 23];
                // TARGET //               
                ApplyBorders(txtbox3);

                var labelFormD = worksheet.Cells[8, 7];
                labelFormD.Value = "MODALIDAD:";
                FontWeightBold(labelFormD);
                ContentCenter(labelFormD);

                ExcelRange labelFormE = GetExcelRange(worksheet, 8, 8, 8, 9);
                labelFormE.Value = "DISTANCIA";
                MergedCells(labelFormE);
                FontWeightBold(labelFormE);
                ContentCenter(labelFormE);

                var txtbox4 = worksheet.Cells[8, 10];
                // TARGET //               
                ApplyBorders(txtbox4);

                ExcelRange labelFormF = GetExcelRange(worksheet, 8, 14, 8, 15);
                labelFormF.Value = "PRESENCIAL";
                MergedCells(labelFormF);
                FontWeightBold(labelFormF);
                ContentCenter(labelFormF);

                ExcelRange txtbox5 = GetExcelRange(worksheet, 8, 16, 8, 17);
                // TARGET //               
                MergedCells(txtbox5);
                ApplyBorders(txtbox5);

                ExcelRange labelFormG = GetExcelRange(worksheet, 8, 19, 8, 20);
                labelFormG.Value = "VIRTUAL";
                MergedCells(labelFormG);
                FontWeightBold(labelFormG);
                ContentCenter(labelFormG);

                var txtbox6 = worksheet.Cells[8, 21];
                // TARGET //               
                ApplyBorders(txtbox6);

                // ------------------------------------------------------------------------------------- //

                ExcelRange txtbox7 = GetExcelRange(worksheet, 10, 3, 10, 5);
                // TARGET //               
                MergedCells(txtbox7);
                ApplySignatureBorders(txtbox7);

                ExcelRange labelFormH = GetExcelRange(worksheet, 11, 3, 12, 5);
                labelFormH.Value = "REGIONAL - SEDE O CUNAD";
                MergedCells(labelFormH);
                FontWeightBold(labelFormH);
                ContentCenter(labelFormH);

                ExcelRange txtbox8 = GetExcelRange(worksheet, 10, 8, 10, 10);
                // TARGET //               
                MergedCells(txtbox8);
                ApplySignatureBorders(txtbox8);

                ExcelRange labelFormI = GetExcelRange(worksheet, 11, 8, 11, 10);
                labelFormI.Value = "FECHA DEL RECONOCIMIENTO";
                MergedCells(labelFormI);
                FontWeightBold(labelFormI);
                ContentCenter(labelFormI);

                ExcelRange txtbox9 = GetExcelRange(worksheet, 10, 15, 10, 16);
                // TARGET //               
                MergedCells(txtbox9);
                ApplySignatureBorders(txtbox9);

                ExcelRange labelFormJ = GetExcelRange(worksheet, 11, 15, 11, 16);
                labelFormJ.Value = "PLAN DE ESTUDIOS A APLICAR";
                MergedCells(labelFormJ);
                FontWeightBold(labelFormJ);
                ContentCenter(labelFormJ);

                ExcelRange txtbox10 = GetExcelRange(worksheet, 10, 21, 10, 22);
                // TARGET //               
                MergedCells(txtbox10);
                ApplySignatureBorders(txtbox10);

                ExcelRange labelFormK = GetExcelRange(worksheet,11, 21, 11, 22);
                labelFormK.Value = "CÓDIGO DEL PLAN DE ESTUDIOS A APLICAR";
                MergedCells(labelFormK);
                FontWeightBold(labelFormK);
                ContentCenter(labelFormK);

                // ------------------------------------------------------------------------------------- //


                ExcelRange txtbox11 = GetExcelRange(worksheet, 14, 3, 14, 5);
                // TARGET //               
                MergedCells(txtbox11);
                ApplySignatureBorders(txtbox11);

                ExcelRange labelFormL = GetExcelRange(worksheet, 15, 3, 15, 5);
                labelFormL.Value = "APELLIDOS Y NOMBRES DEL ESTUDIANTE";
                MergedCells(labelFormL);
                FontWeightBold(labelFormL);
                ContentCenter(labelFormL);

                ExcelRange txtbox12 = GetExcelRange(worksheet, 14, 8, 14, 10);
                // TARGET //               
                MergedCells(txtbox12);
                ApplySignatureBorders(txtbox12);

                ExcelRange labelFormM = GetExcelRange(worksheet, 15, 8, 15, 10);
                labelFormM.Value = "DOCUMENTO DE IDENTIDAD";
                MergedCells(labelFormM);
                FontWeightBold(labelFormM);
                ContentCenter(labelFormM);

                ExcelRange txtbox13 = GetExcelRange(worksheet, 14, 15, 14, 16);
                // TARGET //               
                MergedCells(txtbox13);
                ApplySignatureBorders(txtbox13);

                ExcelRange labelFormO = GetExcelRange(worksheet, 15, 15, 15, 16);
                labelFormO.Value = "CORREO ELECTRONICO";
                MergedCells(labelFormO);
                FontWeightBold(labelFormO);
                ContentCenter(labelFormO);

                ExcelRange txtbox14 = GetExcelRange(worksheet, 14, 21, 14, 22);
                // TARGET //               
                MergedCells(txtbox14);
                ApplySignatureBorders(txtbox14);

                ExcelRange labelFormP = GetExcelRange(worksheet, 15, 21, 15, 22);
                labelFormP.Value = "TELEFONO FIJO - CELULAR";
                MergedCells(labelFormP);
                FontWeightBold(labelFormP);
                ContentCenter(labelFormP);

                // ------------------------------------------------------------------------------------- //

                ExcelRange txtbox15 = GetExcelRange(worksheet, 18, 3, 18, 5);
                // TARGET //               
                MergedCells(txtbox15);
                ApplySignatureBorders(txtbox15);

                ExcelRange labelFormQ = GetExcelRange(worksheet, 19, 3, 19, 5);
                labelFormQ.Value = "INSTITUCIÓN DE DONDE PROVIENE";
                MergedCells(labelFormQ);
                FontWeightBold(labelFormQ);
                ContentCenter(labelFormQ);

                ExcelRange txtbox16 = GetExcelRange(worksheet, 18, 8, 18, 10);
                // TARGET //               
                MergedCells(txtbox16);
                ApplySignatureBorders(txtbox16);

                ExcelRange labelFormR = GetExcelRange(worksheet, 19, 8, 19, 10);
                labelFormR.Value = "PROGRAMA CURSADO";
                MergedCells(labelFormR);
                FontWeightBold(labelFormR);
                ContentCenter(labelFormR);

                ExcelRange txtbox17 = GetExcelRange(worksheet, 18, 14, 18, 16);
                // TARGET //               
                MergedCells(txtbox17);
                ApplySignatureBorders(txtbox17);

                ExcelRange labelFormS = GetExcelRange(worksheet, 19, 14, 19, 16);
                labelFormS.Value = "PROGRAMA A CURSAR";
                MergedCells(labelFormS);
                FontWeightBold(labelFormS);
                ContentCenter(labelFormS);

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
                CellCenter(labelTitle1);
                MergedCells(labelTitle1);
                FontWeightBold(labelTitle1);

                // Evaluar la posibilidad de unir las filas y enviarlas a la izquierda en una sola formula

                ExcelRange Context1 = GetExcelRange(worksheet, 22, 2, 22, 24);
                Context1.Value = "ADMITIR AL USUARIO EN LA RED DE SERVICIOS DE SALUD SEGÚN NIVELES DE ATENCIÓN Y NORMATIVA VIGENTE.";
                CellLeft(Context1);
                MergedCells(Context1);

                ExcelRange Context2 = GetExcelRange(worksheet, 23, 2, 23, 24);
                Context2.Value = "AFILIAR A LA POBLACIÓN AL SISTEMA GENERAL DE SEGURIDAD SOCIAL EN SALUD SEGÚN NORMATIVIDAD VIGENTE.";
                MergedCells(Context2);
                ContentLeft(Context2);

                ExcelRange Context3 = GetExcelRange(worksheet, 24, 2, 24, 24);
                Context3.Value = "FACTURAR LA PRESTACIÓN DE LOS SERVICIOS DE SALUD SEGÚN NORMATIVIDAD Y CONTRATACIÓN";
                MergedCells(Context3);
                ContentLeft(Context3);

                ExcelRange Context4 = GetExcelRange(worksheet, 25, 2, 25, 24);
                Context4.Value = "MANEJAR VALORES E INGRESOS RELACIONADOS CON LA OPERACIÓN DEL ESTABLECIMIENTO. (EQUIVALE A LA NORMA NTS 005 DEL MINCOMERCIO, INDUSTRIA Y TURISMO)";
                MergedCells(Context4);
                ContentLeft(Context4);

                ExcelRange Context5 = GetExcelRange(worksheet, 26, 2, 26, 24);
                Context5.Value = "ORIENTAR AL USUARIO EN RELACIÓN CON SUS NECESIDADES Y EXPECTATIVAS DE ACUERDO CON POLÍTICAS INSTITUCIONALES Y NORMAS DE SALUD VIGENTES.";
                MergedCells(Context5);
                ContentLeft(Context5);

                ExcelRange Context6 = GetExcelRange(worksheet, 27, 2, 27, 24);
                Context6.Value = "PROMOVER LA INTERACCION IDONEA CONSIGO MISMO, CON LOS DEMAS Y CON LA NATURALEZA EN LOS CONTEXTOS LABORAL Y SOCIAL.";
                MergedCells(Context6);
                ContentLeft(Context6);

                ExcelRange Context7 = GetExcelRange(worksheet, 28, 2, 28, 24);
                Context7.Value = "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA";
                MergedCells(Context7);
                ContentLeft(Context7);

                // ----------------------- INSCRIPCION DE MATERIAS (Main Table)------------------------------ //

                ExcelRange CellMT1 = GetExcelRange(worksheet, 31, 1, 32, 1); ;
                CellMT1.Value = "No";
                MergedCells(CellMT1);
                CellCenter(CellMT1);
                FontWeightBold(CellMT1);

                ExcelRange CellMT2 = GetExcelRange(worksheet, 31, 2, 32, 2);
                CellMT2.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                MergedCells(CellMT2);
                CellCenter(CellMT2);
                FontWeightBold(CellMT2);

                ExcelRange CellMT3 = GetExcelRange(worksheet, 31, 3, 31, 4);
                CellMT3.Value = "SISTEMA";
                MergedCells(CellMT3);
                CellCenter(CellMT3);
                FontWeightBold(CellMT3);

                    var subCellMT1 = worksheet.Cells[32, 3];
                    subCellMT1.Value = "Créditos";
                    MergedCells(subCellMT1);
                    CellCenter(subCellMT1);
                    FontWeightBold(subCellMT1);

                    var subCellMT2 = worksheet.Cells[32, 4];
                    subCellMT2.Value = "Semestre";
                    MergedCells(subCellMT2);
                    CellCenter(subCellMT2);
                    FontWeightBold(subCellMT2);

                ExcelRange CellMT4 = GetExcelRange(worksheet, 31, 5, 32, 5);
                CellMT4.Value = "CALIFICACIÓN NUMERICA";
                MergedCells(CellMT4);
                CellCenter(CellMT4);
                FontWeightBold(CellMT4);

                ExcelRange CellMT5 = GetExcelRange(worksheet, 31, 6, 32, 6);
                CellMT5.Value = "CALIFICACION LITERAL";
                MergedCells(CellMT5);
                CellCenter(CellMT5);
                FontWeightBold(CellMT5);

                ExcelRange CellMT6 = GetExcelRange(worksheet, 31, 7, 32, 7);
                CellMT6.Value = "NIVEL";
                MergedCells(CellMT6);
                CellCenter(CellMT6);
                FontWeightBold(CellMT6);

                ExcelRange CellMT7 = GetExcelRange(worksheet, 31, 8, 32, 8);
                CellMT7.Value = "No";
                MergedCells(CellMT7);
                CellCenter(CellMT7);
                FontWeightBold(CellMT7);

                ExcelRange CellMT8 = GetExcelRange(worksheet, 31, 9, 32, 10);
                CellMT8.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                MergedCells(CellMT8);
                CellCenter(CellMT8);
                FontWeightBold(CellMT8);

                ExcelRange CellMT9 = GetExcelRange(worksheet, 31, 11, 31, 12);
                CellMT9.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                MergedCells(CellMT9);
                CellCenter(CellMT9);
                FontWeightBold(CellMT9);

                    var subCellMT3 = worksheet.Cells[32, 11];
                    MergedCells(subCellMT3);
                    CellCenter(subCellMT3);
                    FontWeightBold(subCellMT3);

                    var subCellMT4 = worksheet.Cells[32, 12];
                    subCellMT4.Value = "Semestre";
                    MergedCells(subCellMT4);
                    CellCenter(subCellMT4);
                    FontWeightBold(subCellMT4);

                ExcelRange CellMT10 = GetExcelRange(worksheet, 31, 13, 32, 13);
                CellMT10.Value = "CALIFICACIÓN NUMERICA";
                MergedCells(CellMT10);
                CellCenter(CellMT10);
                FontWeightBold(CellMT10);

                ExcelRange CellMT11 = GetExcelRange(worksheet, 31, 14, 32, 14);
                CellMT11.Value = "CALIFICACION LITERAL";
                MergedCells(CellMT11);
                CellCenter(CellMT11);
                FontWeightBold(CellMT11);

                ExcelRange CellMT12 = GetExcelRange(worksheet, 31, 15, 32, 15);
                CellMT12.Value = "NIVEL";
                MergedCells(CellMT12);
                CellCenter(CellMT12);
                FontWeightBold(CellMT12);

                ExcelRange CellMT13 = GetExcelRange(worksheet, 31, 16, 32, 16);
                CellMT13.Value = "No";
                MergedCells(CellMT13);
                CellCenter(CellMT13);
                FontWeightBold(CellMT13);


                ExcelRange CellMT14 = GetExcelRange(worksheet, 31, 17, 32, 19);
                CellMT14.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
                MergedCells(CellMT14);
                CellCenter(CellMT14);
                FontWeightBold(CellMT14);

                ExcelRange CellMT15 = GetExcelRange(worksheet, 31, 20, 31, 21);
                CellMT15.Value = "SISTEMA";
                MergedCells(CellMT15);
                CellCenter(CellMT15);
                FontWeightBold(CellMT15);

                    var subCellMT5 = worksheet.Cells[32, 20];
                    subCellMT5.Value = "Créditos";
                    MergedCells(subCellMT5);
                    CellCenter(subCellMT5);
                    FontWeightBold(subCellMT5);

                    var subCellMT6 = worksheet.Cells[32, 21];
                    subCellMT6.Value = "Semestre";
                    MergedCells(subCellMT6);
                    CellCenter(subCellMT6);
                    FontWeightBold(subCellMT6);


                ExcelRange CellMT16 = GetExcelRange(worksheet, 31, 22, 32, 22);
                CellMT16.Value = "CALIFICACIÓN NUMERICA";
                MergedCells(CellMT16);
                CellCenter(CellMT16);
                FontWeightBold(CellMT16);

                ExcelRange CellMT17 = GetExcelRange(worksheet, 31, 23, 32, 23);
                CellMT17.Value = "CALIFICACION LITERAL";
                MergedCells(CellMT17);
                CellCenter(CellMT17);
                FontWeightBold(CellMT17);

                ExcelRange CellMT18 = GetExcelRange(worksheet, 31, 24, 32, 24);
                CellMT18.Value = "NIVEL";
                MergedCells(CellMT18);
                CellCenter(CellMT18);
                FontWeightBold(CellMT18);

                ExcelRange celdasMaterias = GetExcelRange(worksheet, 33, 1, 41, 24);
                CellCenter(celdasMaterias);

                var numbCeldasMaterias = worksheet.Cells[33, 1];
                int initCell= 33;
                for (int i = 1; i <= 9; i++)
                {
                    var firstCellColumn = worksheet.Cells[initCell + i - 1, 1];
                    firstCellColumn.Value = i;
                    CellCenter(firstCellColumn);
                }

                // ------------------------------  TOTALES MATERIAS ------------------------------------------ //

                ExcelRange LabelTotals = GetExcelRange(worksheet, 42, 1, 42, 2);
                LabelTotals.Value = "TOTALES";
                MergedCells(LabelTotals);
                CellCenter(LabelTotals);
                FontWeightBold(LabelTotals);

                var rangeToSum1 = worksheet.Cells["C33:C41"];
                var result1 = worksheet.Cells["C42"];
                var totalSubjects1 = worksheet.Cells[42, 3];

                result1.Formula = $"SUM({rangeToSum1.Address})";
                CellCenter(totalSubjects1);
                FontWeightBold(totalSubjects1);

                var rangeToSum2 = worksheet.Cells["K33:K41"];
                var result2 = worksheet.Cells["K42"];
                var totalSubjects2 = worksheet.Cells[42, 11];

                result2.Formula = $"SUM({rangeToSum2.Address})";
                CellCenter(totalSubjects2);
                FontWeightBold(totalSubjects2);

                var rangeToSum3 = worksheet.Cells["T33:T41"];
                var result3 = worksheet.Cells["T42"];
                var totalSubjects3 = worksheet.Cells[42, 20];

                result3.Formula = $"SUM({rangeToSum3.Address})";
                CellCenter(totalSubjects3);
                FontWeightBold(totalSubjects3);


                    ExcelRange LabelTotal1 = GetExcelRange(worksheet, 46, 1, 46, 7);
                    LabelTotal1.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TÉCNICO PROFESIONAL";
                    MergedCells(LabelTotal1);
                    CellRight(LabelTotal1); 

                    ExcelRange LabelTotal2 = GetExcelRange(worksheet, 47, 1, 47, 7);
                    LabelTotal2.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TECNOLÓGICO ";
                    MergedCells(LabelTotal2);
                    CellRight(LabelTotal2);

                    ExcelRange LabelTotal3 = GetExcelRange(worksheet, 48, 1, 48, 7);
                    LabelTotal3.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL PROFESIONAL";
                    MergedCells(LabelTotal3);
                    CellRight(LabelTotal3);


                ExcelRange LabelTotalMaterias = GetExcelRange(worksheet, 44, 8, 44, 10);
                MergedCells(LabelTotalMaterias);
                CellCenter(LabelTotalMaterias);
                FontWeightBold(LabelTotalMaterias);

                //var cellTotal1 = worksheet.Cells["C42"];
                //var cellTotal2 = worksheet.Cells["K42"];
                //var cellTotal3 = worksheet.Cells["T42"];

                // Define la celda donde mostrarás el resultado de la suma
                var sumCellTotals = worksheet.Cells["H44"];

                // Aplica la fórmula de suma a la celda de resultado
                sumCellTotals.Formula = $"SUM({totalSubjects1.Address},{totalSubjects2.Address},{totalSubjects3.Address})";


                var LabelAprobado = worksheet.Cells[45, 8];
                LabelAprobado.Value = "APRO";
                CellCenter(LabelAprobado);
                FontWeightBold(LabelAprobado);

                var LabelPending = worksheet.Cells[45, 9];
                LabelPending.Value = "PEN";
                CellCenter(LabelPending);
                FontWeightBold(LabelPending);

                var LabelTotalCredits = worksheet.Cells[45, 10];
                LabelTotalCredits.Value = "TOTAL CRED";
                CellCenter(LabelTotalCredits);
                FontWeightBold(LabelTotalCredits);

                ExcelRange celdasMateriasTotales = GetExcelRange(worksheet, 46, 8, 48, 10);
                CellCenter(celdasMateriasTotales);

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


                // --------------------- PARRAF "ACLARACIONES" --------------------------- //

                // Unir celdas horizontalmente y verticalmente
                ExcelRange labelTitle2 = worksheet.Cells[50, 1, 50, 2];
                labelTitle2.Value = "Aclaraciones";
                FontWeightBold(labelTitle2);
                MergedCells(labelTitle2);
                ContentLeft(labelTitle2);

                ExcelRange ContextA = worksheet.Cells[51, 1, 51, 24];
                ContextA.Value = "Los créditos académicos faltantes para cumplir a cabalidad con la oferta académica del nivel técnico profesional y/o tecnológico y/o  profesional deben ser cursados y aprobados conforme las reglamentaciones institucionales vigentes.";
                MergedCells(ContextA);
                ContentLeft(ContextA);

                ExcelRange ContextB = worksheet.Cells[52, 1, 52, 24];
                ContextB.Value = "La Escuela de Ciencias Administrativas de la Corporación Unificada Nacional -CUN, reconoce las asignaturas del programa de  Administración de Servicios de Salud   del nivel técnico ( 1 - 3  semestre) para que dar continuidad a su proceso de formación académica a partir de las asignaturas de nivel tecnológico correspondientes a (  4 - 5  semestre) Y profesional correspondientes a ( 6 - 9  semestre)";
                MergedCells(ContextB);
                ContentLeft(ContextB);

                ExcelRange ContextC = worksheet.Cells[53, 1, 53, 24];
                ContextC.Value = "La prueba TyT del ciclo técnico es homologable en la institución. El estudiante deberá presentar la prueba saber TyT para el ciclo tecnológico y Saber PRO para el ciclo profesional.";
                MergedCells(ContextC);
                ContentLeft(ContextC);

                ExcelRange ContextD = worksheet.Cells[54, 1, 54, 24];
                ContextD.Value = "Teniendo en cuenta que el plan de estudios vigente del programa  Administración de Servicios de Salud   no incluye los respectivos niveles de inglés requeridos para obtener las diferentes tituluaciones, el estudiante deberá garantizar lo pertinente al momento de radicar su solilctud de grado, para ello se cuenta con la oferta del centro de Idiomas de la intitución.";
                MergedCells(ContextD);
                ContentLeft(ContextD);

                ExcelRange labelTitle3 = worksheet.Cells[56, 1, 56, 2];
                labelTitle3.Value = "Manifestación expresa del estudiante";
                FontWeightBold(labelTitle3);
                MergedCells(labelTitle3);
                ContentLeft(labelTitle3);

                ExcelRange finalParraf = worksheet.Cells[57, 1, 57, 24];
                finalParraf.Value = "Con el presente documento manifiesto expresamente y sin que medie ninguna clase de vicio o limitación a mi consentimiento, mi plena conformidad con las asignaturas y/o créditos reconocidos u homologados para mi ingreso al nivel técnico profesional y/o tecnológico y/o profesional del programa   Administración de Servicios de Salud    de la Corporación Unificada Nacional de Educación Superior CUN. Las competencias que considere me hagan falta, del ciclo técnico, las podré realizar voluntariamente a través de tutorías en cada área transversal o del programa, talleres nivelatorios y/o participando como asistente a clases sin que estos generen nota alguna y solicitando previamente el ingreso a la clase o tutoría.";
                MergedCells(finalParraf);
                ContentLeft(finalParraf);

                var labelTitle4 = worksheet.Cells[63, 1];
                labelTitle4.Value = "En Constancia de lo anterior firman:";
                FontWeightBold(labelTitle3);
                MergedCells(labelTitle3);
                ContentLeft(labelTitle3);

                // -------------------------------------------------------------------------- //

                // ----------------------------- SIGNATURE INFO ----------------------------- //

                string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\lennon_signature.jpg";
                int widthSignatureInPixels = 230;
                int heightSignatureInPixels = 70;

                var pictureSignature = worksheet.Drawings.AddPicture("Firma", new FileInfo(imagePathSignature));

                pictureSignature.SetPosition(66, -80, 1, -30);
                pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);
                pictureSignature.Locked = true;

                ExcelRange cellSignatureLiderPrograma = worksheet.Cells[66, 1, 66, 2];
                ApplySignatureBorders(cellSignatureLiderPrograma);

                ExcelRange labelSignature1 = worksheet.Cells[67, 1, 67, 3];
                labelSignature1.Value = "Líder de Programa";
                MergedCells(labelSignature1);
                ContentLeft(labelSignature1);

                ExcelRange labelSignature2 = worksheet.Cells[68, 1, 68, 3];
                labelSignature2.Value = "Nombre: SAMPLE NAME";  // Convertir y generar valor dinámico
                MergedCells(labelSignature2);
                ContentLeft(labelSignature2);

                var parrafSquare = worksheet.Cells[66, 7];
                parrafSquare.Value = "finalParraf";
                ContentCenter(parrafSquare);
                MergedCells(parrafSquare);

                // Asignar un valor dinámico para la firma electrónica, como en formatos anteriores

                ExcelRange cellSignatureStudent = worksheet.Cells[66, 9, 66, 12];
                ApplySignatureBorders(cellSignatureStudent);

                ExcelRange labelSignature3 = worksheet.Cells[67, 9, 67, 12];
                labelSignature3.Value = "Estudiante: "; // Convertir y generar valor dinámico
                MergedCells(labelSignature3);
                ContentLeft(labelSignature3);

                ExcelRange labelSignature4 = worksheet.Cells[68, 9, 68, 12];
                labelSignature4.Value = "Nombre: ";  // Convertir y generar valor dinámico
                MergedCells(labelSignature4);
                ContentLeft(labelSignature4);

                ExcelRange labelSignature5 = worksheet.Cells[69, 9, 69, 12];
                labelSignature5.Value = "Doc de Identidad: "; // Convertir y generar valor dinámico
                MergedCells(labelSignature5);
                ContentLeft(labelSignature5);


                // ------------------------------ FOOTER -----------------------------------//
                ExcelRange tableFooter = worksheet.Cells[71, 1, 72, 15];
                CellLeft(tableFooter);

                ExcelRange CellFooter1 = worksheet.Cells[71, 1, 71, 4];
                CellFooter1.Value = "ELABORÓ: "; // Convertir y generar valor dinámico
                //CellLeft(CellFooter1);
                MergedCells(CellFooter1);

                ExcelRange CellFooter2 = worksheet.Cells[72, 1, 72, 4];
                CellFooter2.Value = "FECHA: "; // Convertir y generar valor dinámico
                //CellLeft(CellFooter2);
                MergedCells(CellFooter2);

                ExcelRange CellFooter3 = worksheet.Cells[71, 5, 71, 10];
                CellFooter3.Value = "REVISÓ: "; // Convertir y generar valor dinámico
                //CellLeft(CellFooter3);
                MergedCells(CellFooter3);

                ExcelRange CellFooter4 = worksheet.Cells[72, 5, 72, 10];
                CellFooter4.Value = "FECHA:"; // Convertir y generar valor dinámico
                //CellLeft(CellFooter4);
                MergedCells(CellFooter4);

                ExcelRange CellFooter5 = worksheet.Cells[71, 11, 71, 15];
                CellFooter5.Value = "APROBÓ:"; // Convertir y generar valor dinámico
                //CellLeft(CellFooter5);
                MergedCells(CellFooter5);

                ExcelRange CellFooter6 = worksheet.Cells[72, 11, 72, 15];
                CellFooter6.Value = "FECHA:"; // Convertir y generar valor dinámico
                //ApplyBorders(CellFooter6);
                MergedCells(CellFooter6);


                // ------------------------------ END FOOTER --------------------------------//


                var filePath = @"C:\Users\Jhonattan_Casallas\Downloads\" + nombreArchivo;
                package.SaveAs(new System.IO.FileInfo(filePath));




                return File(filePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);
            }
        }
    }
}
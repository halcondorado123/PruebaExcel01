using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using PruebaExcel01.Models;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;
using System.Xml.Linq;
using System.Security.Cryptography.Xml;
using System.Collections.Generic;
using System.Data;
using System.Reflection.Emit;

namespace PruebaExcel01.Controllers
{

    // Implementar bloqueo de la hoja de Excel para evitar que la hoja se modifique, alterando la información del estudiante
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

                #region Cellsizes
                var worksheet = package.Workbook.Worksheets.Add("SummaryDocs");

                ApplyBackgroundColorToRange(worksheet, 1, 1, 72, 24, System.Drawing.Color.White);

                int[] columnIndices = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
                                       11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                                       21, 22, 23, 24};

                double[] columnWidths = { 5.89, 29.78, 10.78, 10.22, 14.56, 14.78, 14.89, 7.56, 12.56, 14.89,
                                        9.56, 10.78, 14.33, 14.22, 14.22, 6.33, 9.89, 9.89, 7.22, 8.33,
                                        13.56, 16.56, 14.89, 13.56,};

                for (int i = 0; i < columnIndices.Length; i++)
                {
                    int columnIndex = columnIndices[i];
                    double columnWidth = columnWidths[i];
                    worksheet.Column(columnIndex).Width = columnWidth;
                }

                int[] rowIndices = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
                                    11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                                    21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
                                    31, 32, 33, 34, 35, 36, 37, 38, 39, 40,
                                    41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
                                    51, 52, 53, 54, 55, 56, 57, 58, 59, 60,
                                    61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72};

                double[] rowHeights = {12.60, 13.20, 12.60, 12.60, 25.20, 12.60, 12.60, 25.20, 12.60, 38.40,
                                        37.20, 12.60, 12.60, 21.00, 24.00, 15.60, 12.60, 33.60, 12.60, 25.20,
                                        12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60,
                                        51.00, 12.60, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20,
                                        25.20, 12.60, 13.20, 12.60, 12.60, 12.60, 12.60, 13.20, 12.60, 12.60,
                                        15.60, 51.60, 16.20, 39.00, 24.00, 12.60, 55.80, 18.60, 24.60, 19.20,
                                        12.60, 12.60, 12.60, 12.60, 12.60, 93.60, 12.60, 12.60, 12.60, 12.60,
                                        12.60, 12.60};

                for (int i = 0; i < rowIndices.Length; i++)
                {
                    int rowIndex = rowIndices[i];
                    double rowHeight = rowHeights[i];
                    worksheet.Row(rowIndex).Height = rowHeight;
                }

                #endregion

                #region DocumentContent

                List<AsignaturasME> subjects = new List<AsignaturasME>();

                AddLogo(worksheet);
                AddHeaderInfo(worksheet);

                AddFormUserContent(worksheet);
                AddSenaStructureInformation(worksheet);
                ConstructionHeaderTable(worksheet);
                ContentTable(worksheet);

                RecognizedCredits(worksheet);
                ConditionLabels(worksheet);
                StudentStatement(worksheet);
                ContractArea(worksheet);
                ApplyCellsFooter(worksheet);

                #endregion


                var filePath = @"C:\Users\Jhonattan_Casallas\Downloads\" + nombreArchivo;
                package.SaveAs(new System.IO.FileInfo(filePath));

                return File(filePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);

            }
        }




        // Propiedades generales del texto
        public static void ContentCenter(ExcelRange range)
        {

            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // DerechaContenido
        public static void ContentRight(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // IzquierdaContenido
        public static void ContentLeft(ExcelRange range)
        {
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.WrapText = true;
        }

        // Aplicar Negrilla
        public static void FontWeightBold(ExcelRange range, bool bold = true)
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

        public static void MergedCellsHorizontally(ExcelWorksheet worksheet, int startRow, int endRow, int startColumn, int endColumn, bool merge = true)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                ExcelRange rangeToMerge = worksheet.Cells[row, startColumn, row, endColumn];
                rangeToMerge.Style.WrapText = true;
                rangeToMerge.Merge = merge;
            }
        }
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
        public static void ApplyBackgroundColorToRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn, System.Drawing.Color color)
        {
            var range = worksheet.Cells[startRow, startColumn, endRow, endColumn];
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        #region GetCellAddition
        private ExcelRange GetTotalSubjects1(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[42, 3];
        }

        private ExcelRange GetTotalSubjects2(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[42, 11];
        }

        private ExcelRange GetTotalSubjects3(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[42, 20];
        }
        #endregion


        // Seleccion de celdas por rango
        public static ExcelRange GetExcelRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            return worksheet.Cells[startRow, startColumn, endRow, endColumn];
        }



        private void AddLogo(ExcelWorksheet worksheet)
        {
            string imagePathLogo = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\log1.png";
            int widthLogoInPixels = 255;
            int heightLogoInPixels = 98;

            var pictureLogo = worksheet.Drawings.AddPicture("Logo", new FileInfo(imagePathLogo));

            pictureLogo.SetPosition(1, -20, 1, -50);
            pictureLogo.SetSize(widthLogoInPixels, heightLogoInPixels);
            pictureLogo.Locked = true;
        }

        private void AddHeaderInfo(ExcelWorksheet worksheet)
        {
            string fechaRegistro = "Fecha Registro / Hora Registro";
            DateTime fechaHoraActual = DateTime.Now;

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

            ExcelRange headerRectangle = GetExcelRange(worksheet, 1, 3, 3, 24);
            headerRectangle.Value = "ACTA DE RECONOCIMIENTO DE TITULO";
            CellCenter(headerRectangle);
            MergedCells(headerRectangle);
            FontWeightBold(headerRectangle);
        }

        private void AddFormUserContent(ExcelWorksheet worksheet)
        {
            #region FirstDivision

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
            #endregion      

            #region SecondDivision
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

            #endregion

            #region ThirdDivision

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

            ExcelRange labelFormK = GetExcelRange(worksheet, 11, 21, 11, 22);
            labelFormK.Value = "CÓDIGO DEL PLAN DE ESTUDIOS A APLICAR";
            MergedCells(labelFormK);
            FontWeightBold(labelFormK);
            ContentCenter(labelFormK);


            ExcelRange txtbox11 = GetExcelRange(worksheet, 14, 3, 14, 5);
            // TARGET //               
            MergedCells(txtbox11);
            ApplySignatureBorders(txtbox11);

            #endregion

            #region FourthDivision

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

            ExcelRange txtbox15 = GetExcelRange(worksheet, 18, 3, 18, 5);
            // TARGET //               
            MergedCells(txtbox15);
            ApplySignatureBorders(txtbox15);

            #endregion

            #region FifthDivision
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

            #endregion
        }

        private void AddSenaStructureInformation(ExcelWorksheet worksheet)
        {
            var numbSena = worksheet.Cells[22, 1];
            int startRow = 22;
            for (int i = 1; i <= 7; i++)
            {
                var cell = worksheet.Cells[startRow + i - 1, 1];
                cell.Value = i;
                ContentRight(cell);
            }

            var numbSenaLastCell = worksheet.Cells[28, 1];
            ApplySignatureBorders(numbSenaLastCell);


            ExcelRange labelTitle1 = GetExcelRange(worksheet, 21, 1, 21, 24);
            labelTitle1.Value = "ESTRUCTURA CURRICULAR SENA";
            CellCenter(labelTitle1);
            MergedCells(labelTitle1);
            FontWeightBold(labelTitle1);

            ExcelRange context1 = GetExcelRange(worksheet, 22, 2, 22, 24);
            context1.Value = "ADMITIR AL USUARIO EN LA RED DE SERVICIOS DE SALUD SEGÚN NIVELES DE ATENCIÓN Y NORMATIVA VIGENTE.";
            ContentLeft(context1);
            MergedCells(context1);

            ExcelRange context2 = GetExcelRange(worksheet, 23, 2, 23, 24);
            context2.Value = "AFILIAR A LA POBLACIÓN AL SISTEMA GENERAL DE SEGURIDAD SOCIAL EN SALUD SEGÚN NORMATIVIDAD VIGENTE.";
            MergedCells(context2);
            ContentLeft(context2);

            ExcelRange context3 = GetExcelRange(worksheet, 24, 2, 24, 24);
            context3.Value = "FACTURAR LA PRESTACIÓN DE LOS SERVICIOS DE SALUD SEGÚN NORMATIVIDAD Y CONTRATACIÓN";
            MergedCells(context3);
            ContentLeft(context3);

            ExcelRange context4 = GetExcelRange(worksheet, 25, 2, 25, 24);
            context4.Value = "MANEJAR VALORES E INGRESOS RELACIONADOS CON LA OPERACIÓN DEL ESTABLECIMIENTO. (EQUIVALE A LA NORMA NTS 005 DEL MINCOMERCIO, INDUSTRIA Y TURISMO)";
            MergedCells(context4);
            ContentLeft(context4);

            ExcelRange context5 = GetExcelRange(worksheet, 26, 2, 26, 24);
            context5.Value = "ORIENTAR AL USUARIO EN RELACIÓN CON SUS NECESIDADES Y EXPECTATIVAS DE ACUERDO CON POLÍTICAS INSTITUCIONALES Y NORMAS DE SALUD VIGENTES.";
            MergedCells(context5);
            ContentLeft(context5);

            ExcelRange context6 = GetExcelRange(worksheet, 27, 2, 27, 24);
            context6.Value = "PROMOVER LA INTERACCION IDONEA CONSIGO MISMO, CON LOS DEMAS Y CON LA NATURALEZA EN LOS CONTEXTOS LABORAL Y SOCIAL.";
            MergedCells(context6);
            ContentLeft(context6);

            ExcelRange context7 = GetExcelRange(worksheet, 28, 2, 28, 24);
            context7.Value = "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA";
            ApplySignatureBorders(context7);
            MergedCells(context7);
            ContentLeft(context7);

        }

        private void ConstructionHeaderTable(ExcelWorksheet worksheet)
        {
            ExcelRange HeaderBar = GetExcelRange(worksheet, 31, 1, 32, 24);
            CellCenter(HeaderBar);
            FontWeightBold(HeaderBar);

            #region CellsMainTable

            ExcelRange CellMT1 = GetExcelRange(worksheet, 31, 1, 32, 1);
            CellMT1.Value = "No";
            MergedCells(CellMT1);

            ExcelRange CellMT2 = GetExcelRange(worksheet, 31, 2, 32, 2);
            CellMT2.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT2);

            ExcelRange CellMT3 = GetExcelRange(worksheet, 31, 3, 31, 4);
            CellMT3.Value = "SISTEMA";
            MergedCells(CellMT3);

            ExcelRange CellMT4 = GetExcelRange(worksheet, 31, 5, 32, 5);
            CellMT4.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT4);

            ExcelRange CellMT5 = GetExcelRange(worksheet, 31, 6, 32, 6);
            CellMT5.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT5);

            ExcelRange CellMT6 = GetExcelRange(worksheet, 31, 7, 32, 7);
            CellMT6.Value = "NIVEL";
            MergedCells(CellMT6);

            ExcelRange CellMT7 = GetExcelRange(worksheet, 31, 8, 32, 8);
            CellMT7.Value = "No";
            MergedCells(CellMT7);

            ExcelRange CellMT8 = GetExcelRange(worksheet, 31, 9, 32, 10);
            CellMT8.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT8);

            ExcelRange CellMT9 = GetExcelRange(worksheet, 31, 11, 31, 12);
            CellMT9.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT9);

            ExcelRange CellMT10 = GetExcelRange(worksheet, 31, 13, 32, 13);
            CellMT10.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT10);

            ExcelRange CellMT11 = GetExcelRange(worksheet, 31, 14, 32, 14);
            CellMT11.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT11);

            ExcelRange CellMT12 = GetExcelRange(worksheet, 31, 15, 32, 15);
            CellMT12.Value = "NIVEL";
            MergedCells(CellMT12);

            ExcelRange CellMT13 = GetExcelRange(worksheet, 31, 16, 32, 16);
            CellMT13.Value = "No";
            MergedCells(CellMT13);

            ExcelRange CellMT14 = GetExcelRange(worksheet, 31, 17, 32, 19);
            CellMT14.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT14);

            ExcelRange CellMT15 = GetExcelRange(worksheet, 31, 20, 31, 21);
            CellMT15.Value = "SISTEMA";
            MergedCells(CellMT15);

            ExcelRange CellMT16 = GetExcelRange(worksheet, 31, 22, 32, 22);
            CellMT16.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT16);

            ExcelRange CellMT17 = GetExcelRange(worksheet, 31, 23, 32, 23);
            CellMT17.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT17);

            ExcelRange CellMT18 = GetExcelRange(worksheet, 31, 24, 32, 24);
            CellMT18.Value = "NIVEL";
            MergedCells(CellMT18);

            #endregion

            #region SubCellsMainTable

            var subCellMT1 = worksheet.Cells[32, 3];
            subCellMT1.Value = "Créditos";

            var subCellMT2 = worksheet.Cells[32, 4];
            subCellMT2.Value = "Semestre";

            var subCellMT3 = worksheet.Cells[32, 11];
            subCellMT3.Value = "Créditos";

            var subCellMT4 = worksheet.Cells[32, 12];
            subCellMT4.Value = "Semestre";

            var subCellMT5 = worksheet.Cells[32, 20];
            subCellMT5.Value = "Créditos";

            var subCellMT6 = worksheet.Cells[32, 21];
            subCellMT6.Value = "Semestre";

            #endregion

        }


        //public void InsertSubjectColumn2(ExcelWorksheet worksheet, int row, int column, int i, List<AsignaturasME> listSubjects)
        //{
        //    SubjectGenerator subjectGenerator = new SubjectGenerator();
        //    AsignaturasME[] subjects = subjectGenerator.GenerateSubjects(listSubjects);
        //    MergedCellsHorizontally(worksheet, 33, 41, 9, 10);

        //    worksheet.Cells[row, column].Value = subjects[i].Numero;

        //    //int asignaturaIndex = 9; // Variable para rastrear el índice de la asignatura

        //    // Itera a través de las filas fusionadas y establece el valor de la asignatura en cada una de ellas
        //    for (int rowIndex = 33; rowIndex <= 41; rowIndex++)
        //    {
        //        int asignaturaIndex = (rowIndex - 24) % 26; ; // Calcula el índice relativo a la fila actual
        //        asignaturaIndex %= subjects.Length;  // Asegura que el índice no supere el número de elementos en subjects

        //        worksheet.Cells[rowIndex, 9].Value = subjects[asignaturaIndex].Asignatura;
        //        worksheet.Cells[rowIndex, 10].Value = subjects[asignaturaIndex].Asignatura;

        //        // Aumenta el índice de la asignatura para la próxima fila
        //        asignaturaIndex++;
        //    }

        //    worksheet.Cells[row, column + 3].Value = subjects[i].Creditos;
        //    worksheet.Cells[row, column + 4].Value = subjects[i].Semestre;
        //    worksheet.Cells[row, column + 5].Value = subjects[i].CalificacionNumerica;
        //    worksheet.Cells[row, column + 6].Value = subjects[i].CalificacionLiteral;
        //    worksheet.Cells[row, column + 7].Value = subjects[i].Nivel;
        //}


        //public void InsertSubjectColumn3(ExcelWorksheet worksheet, int row, int column, int i, List<AsignaturasME> listSubjects)
        //{
        //    SubjectGenerator subjectGenerator = new SubjectGenerator();
        //    AsignaturasME[] subjects = subjectGenerator.GenerateSubjects(listSubjects);
        //    MergedCellsHorizontally(worksheet, 33, 41, 17, 19);


        //    //int asignaturaIndex = 18; // Variable para rastrear el índice de la asignatura


        //    // Itera a través de las filas fusionadas y establece el valor de la asignatura en cada una de ellas
        //    for (int rowIndex = 33; rowIndex <= 41; rowIndex++)
        //    {
        //        int asignaturaIndex = rowIndex - 15; // Calcula el índice relativo a la fila actual
        //        asignaturaIndex %= subjects.Length;  // Asegura que el índice no supere el número de elementos en subjects

        //        if (asignaturaIndex < subjects.Length)
        //        {
        //            worksheet.Cells[row, 16].Value = subjects[i].Numero;
        //            worksheet.Cells[rowIndex, 17].Value = subjects[asignaturaIndex].Asignatura;
        //            worksheet.Cells[rowIndex, 18].Value = subjects[asignaturaIndex].Asignatura;
        //            worksheet.Cells[rowIndex, 19].Value = subjects[asignaturaIndex].Asignatura;

        //            // Aumenta el índice de la asignatura para la próxima fila
        //            asignaturaIndex++;
        //        }
        //    }

        //    worksheet.Cells[row, column + 5].Value = subjects[i].Creditos;
        //    worksheet.Cells[row, column + 6].Value = subjects[i].Semestre;
        //    worksheet.Cells[row, column + 7].Value = subjects[i].CalificacionNumerica;
        //    worksheet.Cells[row, column + 8].Value = subjects[i].CalificacionLiteral;
        //    worksheet.Cells[row, column + 9].Value = subjects[i].Nivel;
        //}


        private void ContentTable(ExcelWorksheet worksheet)
        {
            AsignaturasME asignaturasME = new AsignaturasME();
            List<AsignaturasME> subjects = GetSubjects.SubjectGenerator();

            int row = 33;
            int column = 1;
            int subjectsPerColumn = 9; // Cambiar de columna después de 9 elementos
            bool ignoreMergedCells = false; // Variable para ignorar celdas fusionadas en la fila 32

            MergedCellsHorizontally(worksheet, 33, 41, 9, 10);
            MergedCellsHorizontally(worksheet, 33, 41, 17, 19);

            int subjectCount = 0; // Contador para llevar el seguimiento de los elementos en una columna

            // Llena el archivo de Excel con los datos
            foreach (var subject in subjects)
            {
                if (subjectCount >= subjectsPerColumn)
                {
                    // Cambiar de columna después de 9 elementos (cuenta 8 y 9 como una sola celda)
                    row = 33;
                    column += 7; // Cambia a la siguiente columna

                    // Si la columna actual es la 15 o 16, aumenta el desplazamiento en 2
                    if (column == 9 || column == 10)
                    {
                        column += 2;
                    }

                    subjectCount = 0; // Reiniciar el contador
                }

                if (row == 32 && column == 12)
                {
                    // Mueve el valor que debería estar en la columna 12 a la columna 13
                    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                }
                else if (row == 32 && column == 16)
                {
                    // Mueve el valor que debería estar en la columna 16 a la columna 17
                    worksheet.Cells[row, column + 1].Value = subject.Numero[0];
                }
                else
                {
                    // Considerar celdas fusionadas en otras filas
                    worksheet.Cells[row, column].Value = subject.Numero[0];
                }

                worksheet.Cells[row, column + 1].Value = subject.Asignatura[0];
                worksheet.Cells[row, column + 2].Value = subject.Creditos[0];
                worksheet.Cells[row, column + 3].Value = subject.Semestre[0];
                worksheet.Cells[row, column + 4].Value = subject.CalificacionNumerica[0];
                worksheet.Cells[row, column + 5].Value = subject.CalificacionLiteral[0];
                worksheet.Cells[row, column + 6].Value = subject.Nivel[0];

                if (row == 33 && worksheet.Cells[row, column, row + 1, column].Merge && !ignoreMergedCells)
                {
                    row += 3; // Si las celdas 17 a 19 están fusionadas y no estamos en la fila 32, avanzar tres filas
                }
                else
                {
                    row++; // Si no están fusionadas o estamos en la fila 32, avanzar una fila
                }

                subjectCount++;

                // Verificar si estamos en las columnas 9 y 10 y marcar la fusión si es necesario
                if (column == 9 || column == 10)
                {
                    ignoreMergedCells = true; // Ignorar la fusión en las columnas 9 y 10
                }
            }



            ExcelRange celdasMaterias = GetExcelRange(worksheet, 33, 1, 41, 24);
            CellCenter(celdasMaterias);

            //    SubjectGenerator subjectGenerator = new SubjectGenerator();
            //    AsignaturasME[] subjects = subjectGenerator.GenerateSubjects(listSubjects);

            //    int initRow = 33;
            //    int materiasPerColumn = 9;

            //    try
            //    {
            //        for (int i = 0; i < subjects.Length; i++)
            //        {
            //            int rowNumber = initRow + (i % materiasPerColumn);

            //            if (materiasPerColumn == 8)
            //            {
            //                int columnNumber = (i / materiasPerColumn) * 7 + 1;

            //                if (i <= 8)
            //                {
            //                    subjects[i].InsertSubject(worksheet, rowNumber, columnNumber); // Llama a InsertSubject en el objeto AsignaturasME
            //                }
            //                else if (i >= 9 && i < 18)
            //                {
            //                    subjects[i].InsertSubjectColumn2(worksheet, rowNumber, columnNumber);
            //                }
            //                else if (i >= 18 && i < 27)
            //                {
            //                    subjects[i].InsertSubjectColumn3(worksheet, rowNumber, columnNumber);
            //                }
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        ex.Message.ToString();
            //    }




            //HASTA ACA SE BORRA

            #region ResultSubjects

            ExcelRange LabelTotals = GetExcelRange(worksheet, 42, 1, 42, 2);
            LabelTotals.Value = "TOTALES";
            MergedCells(LabelTotals);
            CellCenter(LabelTotals);
            FontWeightBold(LabelTotals);

            var rangeToSum1 = worksheet.Cells["C33:C41"];
            var result1 = worksheet.Cells["C42"];
            ExcelRange totalSubjects1 = GetTotalSubjects1(worksheet);

            result1.Formula = $"SUM({rangeToSum1.Address})";
            CellCenter(totalSubjects1);
            FontWeightBold(totalSubjects1);

            var rangeToSum2 = worksheet.Cells["K33:K41"];
            var result2 = worksheet.Cells["K42"];
            ExcelRange totalSubjects2 = GetTotalSubjects2(worksheet);

            result2.Formula = $"SUM({rangeToSum2.Address})";
            CellCenter(totalSubjects2);
            FontWeightBold(totalSubjects2);

            var rangeToSum3 = worksheet.Cells["T33:T41"];
            var result3 = worksheet.Cells["T42"];
            ExcelRange totalSubjects3 = GetTotalSubjects3(worksheet);

            result3.Formula = $"SUM({rangeToSum3.Address})";
            CellCenter(totalSubjects3);
            FontWeightBold(totalSubjects3);

            #endregion

        }

        private void RecognizedCredits(ExcelWorksheet worksheet)
        {
            // EVALUAR CRITERIOS DE LA CARRERA (1)TECNICO (2)TECNOLOGICO (3)PROFESIONAL
            // IDENTIFICAR LOS CREDITOS QUE SE APRUEBAN, PENDIENTES Y EL TOTAL DE CREDITOS(POR LO CUAL SE DEPENDE DEL PUNTO ANTERIOR

            ExcelRange totalSubjects1 = GetTotalSubjects1(worksheet);
            ExcelRange totalSubjects2 = GetTotalSubjects2(worksheet);
            ExcelRange totalSubjects3 = GetTotalSubjects3(worksheet);

            ExcelRange cellLabelsRecCred = GetExcelRange(worksheet, 46, 1, 48, 7);
            CellRight(cellLabelsRecCred);
            FontWeightBold(cellLabelsRecCred);

            ExcelRange CellLabelRc1 = GetExcelRange(worksheet, 46, 1, 46, 7);
            MergedCells(CellLabelRc1);
            CellLabelRc1.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TÉCNICO PROFESIONAL";

            ExcelRange CellLabelRc2 = GetExcelRange(worksheet, 47, 1, 47, 7);
            MergedCells(CellLabelRc2);
            CellLabelRc2.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TECNOLÓGICO ";


            ExcelRange CellLabelRc3 = GetExcelRange(worksheet, 48, 1, 48, 7);
            MergedCells(CellLabelRc3);
            CellLabelRc3.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL PROFESIONAL";

            ExcelRange LabelTotalMaterias = GetExcelRange(worksheet, 44, 8, 44, 10);
            MergedCells(LabelTotalMaterias);
            CellCenter(LabelTotalMaterias);
            FontWeightBold(LabelTotalMaterias);


            var sumCellTotals = worksheet.Cells["H44"];
            sumCellTotals.Formula = $"SUM({totalSubjects1.Address},{totalSubjects2.Address},{totalSubjects3.Address})";


            var labelApproved = worksheet.Cells[45, 8];
            labelApproved.Value = "APRO";
            CellCenter(labelApproved);
            FontWeightBold(labelApproved);

            var labelPending = worksheet.Cells[45, 9];
            labelPending.Value = "PEN";
            CellCenter(labelPending);
            FontWeightBold(labelPending);

            var labelTotalCredits = worksheet.Cells[45, 10];
            labelTotalCredits.Value = "TOTAL CRED";
            CellCenter(labelTotalCredits);
            FontWeightBold(labelTotalCredits);

            ExcelRange celdasMateriasTotales = GetExcelRange(worksheet, 46, 8, 48, 10);
            CellCenter(celdasMateriasTotales);

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

            foreach (var cell in cellsToSetZero)
            {
                cell.Value = 0;
            }

            // FORMULAS TO DETERMINE CREDITS

            // CREDITOS TECNICO PROFESIONAL (1)
            var totalTechnicalCredits = worksheet.Cells["J46"];
            var totalTechnicalApproved = worksheet.Cells["H46"];
            var totalTechnicalPending = worksheet.Cells["I46"];
            totalTechnicalPending.Formula = $"({totalTechnicalCredits.Address}) - ({totalTechnicalApproved.Address})";

            // TOTAL CREDITOS APROBADOS (2)
            var totalTechnologistCredits = worksheet.Cells["J47"];
            var totalTechnologistApproved = worksheet.Cells["H47"];
            var totalTechnologistPending = worksheet.Cells["I47"];
            totalTechnologistPending.Formula = $"({totalTechnologistCredits.Address}) - ({totalTechnologistApproved.Address})";

            // TOTAL CREDITOS PENDIENTES (3)
            var totalProfessionalCredits = worksheet.Cells["J48"];
            var totalProfessionalApproved = worksheet.Cells["H48"];
            var totalProfessionalPending = worksheet.Cells["I48"];
            totalProfessionalPending.Formula = $"({totalProfessionalCredits.Address}) - ({totalProfessionalApproved.Address})";

        }

        private void ConditionLabels(ExcelWorksheet worksheet)
        {
            ExcelRange conditionText = worksheet.Cells[50, 1, 54, 24];
            ContentLeft(conditionText);

            ExcelRange labelTitle2 = worksheet.Cells[50, 1, 50, 2];
            labelTitle2.Value = "Aclaraciones";
            FontWeightBold(labelTitle2);
            MergedCells(labelTitle2);

            ExcelRange contextA = worksheet.Cells[51, 1, 51, 24];
            contextA.Value = "Los créditos académicos faltantes para cumplir a cabalidad con la oferta académica del nivel técnico profesional y/o tecnológico y/o  profesional deben ser cursados y aprobados conforme las reglamentaciones institucionales vigentes.";
            MergedCells(contextA);

            ExcelRange contextB = worksheet.Cells[52, 1, 52, 24];
            contextB.Value = "La Escuela de Ciencias Administrativas de la Corporación Unificada Nacional -CUN, reconoce las asignaturas del programa de  Administración de Servicios de Salud   del nivel técnico ( 1 - 3  semestre) para que dar continuidad a su proceso de formación académica a partir de las asignaturas de nivel tecnológico correspondientes a (  4 - 5  semestre) Y profesional correspondientes a ( 6 - 9  semestre)";
            MergedCells(contextB);

            ExcelRange contextC = worksheet.Cells[53, 1, 53, 24];
            contextC.Value = "La prueba TyT del ciclo técnico es homologable en la institución. El estudiante deberá presentar la prueba saber TyT para el ciclo tecnológico y Saber PRO para el ciclo profesional.";
            MergedCells(contextC);

            ExcelRange contextD = worksheet.Cells[54, 1, 54, 24];
            contextD.Value = "Teniendo en cuenta que el plan de estudios vigente del programa  Administración de Servicios de Salud   no incluye los respectivos niveles de inglés requeridos para obtener las diferentes tituluaciones, el estudiante deberá garantizar lo pertinente al momento de radicar su solilctud de grado, para ello se cuenta con la oferta del centro de Idiomas de la intitución.";
            MergedCells(contextD);

        }

        private void StudentStatement(ExcelWorksheet worksheet)
        {
            ExcelRange studentStatementArea = worksheet.Cells[56, 1, 57, 24];
            ContentLeft(studentStatementArea);

            ExcelRange labelTitle3 = worksheet.Cells[56, 1, 56, 2];
            labelTitle3.Value = "Manifestación expresa del estudiante";
            FontWeightBold(labelTitle3);
            MergedCells(labelTitle3);

            ExcelRange contextE = worksheet.Cells[57, 1, 57, 24];
            contextE.Value = "Con el presente documento manifiesto expresamente y sin que medie ninguna clase de vicio o limitación a mi consentimiento, mi plena conformidad con las asignaturas y/o créditos reconocidos u homologados para mi ingreso al nivel técnico profesional y/o tecnológico y/o profesional del programa   Administración de Servicios de Salud    de la Corporación Unificada Nacional de Educación Superior CUN. Las competencias que considere me hagan falta, del ciclo técnico, las podré realizar voluntariamente a través de tutorías en cada área transversal o del programa, talleres nivelatorios y/o participando como asistente a clases sin que estos generen nota alguna y solicitando previamente el ingreso a la clase o tutoría.";
            MergedCells(contextE);

        }

        private void ProgramManagerSignature(ExcelWorksheet worksheet)
        {
            // Por asignar una biblioteca de imagenes de firmas de jefes de programa
            // Valor quemado
            string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\lennon_signature.jpg";
            int widthSignatureInPixels = 230;
            int heightSignatureInPixels = 70;

            var pictureSignature = worksheet.Drawings.AddPicture("Firma", new FileInfo(imagePathSignature));

            pictureSignature.SetPosition(66, -80, 1, -30);
            pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);
            pictureSignature.Locked = true;

        }

        private void StudentSignature(ExcelWorksheet worksheet)
        {
            // Valor quemado, por evaluar la opcion de firma generada por parte del estudiante
            string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\2560px-Freddie_Mercury_signature.svg.png";
            int widthSignatureInPixels = 230;
            int heightSignatureInPixels = 70;

            var pictureSignature = worksheet.Drawings.AddPicture("FirmaStu", new FileInfo(imagePathSignature));

            pictureSignature.SetPosition(66, -80, 9, -30);
            pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);
            pictureSignature.Locked = true;

        }


        private void ContractArea(ExcelWorksheet worksheet)
        {
            var parrafSquare = worksheet.Cells[66, 7];
            parrafSquare.Value = "finalParraf";
            ContentCenter(parrafSquare);
            MergedCells(parrafSquare);

            ExcelRange studentStatementArea = worksheet.Cells[63, 1, 70, 24];
            ContentLeft(studentStatementArea);


            ExcelRange labelTitle4 = worksheet.Cells[63, 1, 63, 2];
            labelTitle4.Value = "En Constancia de lo anterior firman:";
            FontWeightBold(labelTitle4);
            MergedCells(labelTitle4);

            var labelPending = worksheet.Cells[45, 9];

            ExcelRange cellSignatureProgramManager = worksheet.Cells[66, 1, 66, 3];
            MergedCells(cellSignatureProgramManager);
            ApplySignatureBorders(cellSignatureProgramManager);
            ProgramManagerSignature(worksheet);

            ExcelRange labelSignature1 = worksheet.Cells[67, 1, 67, 3];
            labelSignature1.Value = "Líder de Programa";
            MergedCells(labelSignature1);

            ExcelRange labelSignature2 = worksheet.Cells[68, 1, 68, 3];
            labelSignature2.Value = "Nombre: SAMPLE NAME";  // Convertir y generar valor dinámico
            MergedCells(labelSignature2);

            ExcelRange cellSignatureStudent = worksheet.Cells[66, 9, 66, 12];
            MergedCells(cellSignatureStudent);
            ApplySignatureBorders(cellSignatureStudent);
            StudentSignature(worksheet);

            ExcelRange labelSignature3 = worksheet.Cells[67, 9, 67, 12];
            labelSignature3.Value = "Estudiante: "; // Convertir y generar valor dinámico
            MergedCells(labelSignature3);

            ExcelRange labelSignature4 = worksheet.Cells[68, 9, 68, 12];
            labelSignature4.Value = "Nombre: ";  // Convertir y generar valor dinámico
            MergedCells(labelSignature4);
            ContentLeft(labelSignature4);

            ExcelRange labelSignature5 = worksheet.Cells[69, 9, 69, 12];
            labelSignature5.Value = "Doc de Identidad: "; // Convertir y generar valor dinámico
            MergedCells(labelSignature5);
            ContentLeft(labelSignature5);

        }

        private void ApplyCellsFooter(ExcelWorksheet worksheet)
        {
            ExcelRange tableFooter = worksheet.Cells[71, 1, 72, 15];
            CellLeft(tableFooter);

            ExcelRange CellFooter1 = worksheet.Cells[71, 1, 71, 4];
            CellFooter1.Value = "ELABORÓ: "; // Convertir y generar valor dinámico
            MergedCells(CellFooter1);
            //CellLeft(CellFooter1);


            ExcelRange CellFooter2 = worksheet.Cells[72, 1, 72, 4];
            CellFooter2.Value = "FECHA: "; // Convertir y generar valor dinámico
            MergedCells(CellFooter2);
            //CellLeft(CellFooter2);

            ExcelRange CellFooter3 = worksheet.Cells[71, 5, 71, 10];
            CellFooter3.Value = "REVISÓ: "; // Convertir y generar valor dinámico
            MergedCells(CellFooter3);
            //CellLeft(CellFooter3);

            ExcelRange CellFooter4 = worksheet.Cells[72, 5, 72, 10];
            CellFooter4.Value = "FECHA:"; // Convertir y generar valor dinámico
            MergedCells(CellFooter4);
            //CellLeft(CellFooter4);

            ExcelRange CellFooter5 = worksheet.Cells[71, 11, 71, 15];
            CellFooter5.Value = "APROBÓ:"; // Convertir y generar valor dinámico
            MergedCells(CellFooter5);
            //CellLeft(CellFooter5);

            ExcelRange CellFooter6 = worksheet.Cells[72, 11, 72, 15];
            CellFooter6.Value = "FECHA:"; // Convertir y generar valor dinámico
            MergedCells(CellFooter6);
            //ApplyBorders(CellFooter6);
        }
    }
}
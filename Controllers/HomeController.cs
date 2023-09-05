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

                double[] columnWidths = { 11.00, 35.00, 10.00, 10.00, 15.00, 15.00, 14.89, 11.00, 35.00, 14.89,
                                        13.00, 15.00, 15.00, 14.22, 11.00, 35.00, 14.00, 14.00, 15.00, 15.00,
                                        18.00, 16.56, 14.89, 13.56,};

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
                                    61, 62, 63, 64, 65, 66, 67, 68, 69, 70,
                                    71, 72, 73, 74, 75, 76, 77, 78, 79, 80};

                double[] rowHeights = {12.60, 13.20, 12.60, 12.60, 25.20, 12.60, 12.60, 25.20, 12.60, 38.40,
                                        37.20, 12.60, 12.60, 21.00, 24.00, 15.60, 12.60, 33.60, 12.60, 25.20,
                                        12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60,
                                        51.00, 12.60, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20,
                                        25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 25.20, 12.60, 12.60,
                                        15.60, 15.60, 16.20, 16.20, 16.20, 12.60, 55.80, 18.60, 24.60, 19.20,
                                        12.60, 12.60, 12.60, 12.60, 12.60, 93.60, 12.60, 12.60, 12.60, 12.60,
                                        12.60, 12.60, 93.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60, 12.60};

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

        public static void WrapText(ExcelRange range)
        {
            range.Style.WrapText = true;
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
            return worksheet.Cells[48, 3];
        }

        private ExcelRange GetTotalSubjects2(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[48, 10];
        }

        private ExcelRange GetTotalSubjects3(ExcelWorksheet worksheet)
        {
            return worksheet.Cells[48, 17];
        }
        #endregion


        // Seleccion de celdas por rango
        public static ExcelRange GetExcelRange(ExcelWorksheet worksheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            return worksheet.Cells[startRow, startColumn, endRow, endColumn];
        }


        #region ImagesDocument
        private void AddLogo(ExcelWorksheet worksheet)
        {
            string imagePathLogo = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\log1.png";
            int widthLogoInPixels = 300;
            int heightLogoInPixels = 110;

            var pictureLogo = worksheet.Drawings.AddPicture("Logo", new FileInfo(imagePathLogo));

            pictureLogo.SetPosition(1, -20, 1, -50);
            pictureLogo.SetSize(widthLogoInPixels, heightLogoInPixels);
            pictureLogo.Locked = true;
        }

        private void ProgramManagerSignature(ExcelWorksheet worksheet)
        {
            // Por asignar una biblioteca de imagenes de firmas de jefes de programa
            // Valor quemado
            string imagePathSignature = "C:\\Users\\Jhonattan_Casallas\\Desktop\\EnsayoExcel\\PruebaExcel_Version02\\Img_sample\\lennon_signature.jpg";
            int widthSignatureInPixels = 230;
            int heightSignatureInPixels = 70;

            var pictureSignature = worksheet.Drawings.AddPicture("Firma", new FileInfo(imagePathSignature));

            pictureSignature.SetPosition(73, -80, 1, -30);
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

            pictureSignature.SetPosition(73, -80, 8, -30);
            pictureSignature.SetSize(widthSignatureInPixels, heightSignatureInPixels);
            pictureSignature.Locked = true;

        }

        #endregion



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

            ExcelRange headerRectangle = GetExcelRange(worksheet, 1, 3, 3, 21);
            headerRectangle.Value = "ACTA DE RECONOCIMIENTO DE TITULO";
            CellCenter(headerRectangle);
            MergedCells(headerRectangle);
            FontWeightBold(headerRectangle);
        }

        private void AddFormUserContent(ExcelWorksheet worksheet)
        {
            ExcelRange formMainUser = GetExcelRange(worksheet, 4, 1, 20, 21);
            ContentCenter(formMainUser);
            WrapText(formMainUser);

            #region FirstDivision

            var labelFormA = worksheet.Cells[5, 6];
            labelFormA.Value = "INTERNA";
            MergedCells(labelFormA);
            FontWeightBold(labelFormA);
            ContentCenter(labelFormA);

            var txtbox1 = worksheet.Cells[5, 7];
            ApplyBorders(txtbox1);
            // TARGET //


            var labelFormB = worksheet.Cells[5, 11];
            labelFormB.Value = "EXTERNA";
            MergedCells(labelFormB);
            FontWeightBold(labelFormB);
            ContentCenter(labelFormB);

            var txtbox2 = worksheet.Cells[5, 12];
            ApplyBorders(txtbox2);
            // TARGET //               

            var labelFormC = worksheet.Cells[5, 16];
            labelFormC.Value = "PERIODO ACADÉMICO";
            MergedCells(labelFormC);
            FontWeightBold(labelFormC);
            ContentCenter(labelFormC);

            var txtbox3 = worksheet.Cells[5, 17];
            ApplyBorders(txtbox3);
            // TARGET //               


            #endregion

            #region SecondDivision

            ExcelRange labelFormD = GetExcelRange(worksheet, 8, 3, 8, 4);
            labelFormD.Value = "MODALIDAD:";
            MergedCells(labelFormD);
            FontWeightBold(labelFormD);
            ContentCenter(labelFormD);

            var labelFormE = worksheet.Cells[8, 6];
            labelFormE.Value = "DISTANCIA";
            MergedCells(labelFormE);
            FontWeightBold(labelFormE);
            ContentCenter(labelFormE);

            var txtbox4 = worksheet.Cells[8, 7];
            ApplyBorders(txtbox4);
            // TARGET //               

            var labelFormF = worksheet.Cells[8, 11];
            labelFormF.Value = "PRESENCIAL";
            MergedCells(labelFormF);
            FontWeightBold(labelFormF);
            ContentCenter(labelFormF);

            var txtbox5 = worksheet.Cells[8, 12];
            ApplyBorders(txtbox5);
            // TARGET //               

            var labelFormG = worksheet.Cells[8, 16];
            labelFormG.Value = "VIRTUAL";
            MergedCells(labelFormG);
            FontWeightBold(labelFormG);
            ContentCenter(labelFormG);

            var txtbox6 = worksheet.Cells[8, 17];
            ApplyBorders(txtbox6);
            // TARGET //               

            #endregion

            #region ThirdDivision

            var txtbox7 = worksheet.Cells[10, 2];
            ApplySignatureBorders(txtbox7);
            // TARGET               

            ExcelRange labelFormH = GetExcelRange(worksheet, 11, 2, 12, 2);
            labelFormH.Value = "REGIONAL - SEDE O CUNAD";
            MergedCells(labelFormH);
            FontWeightBold(labelFormH);
            ContentCenter(labelFormH);

            var txtbox8 = worksheet.Cells[10, 9];
            ApplySignatureBorders(txtbox8);
            // TARGET //

            ExcelRange labelFormI = GetExcelRange(worksheet, 11, 9, 12, 9);
            labelFormI.Value = "FECHA DEL RECONOCIMIENTO";
            MergedCells(labelFormI);
            FontWeightBold(labelFormI);
            ContentCenter(labelFormI);

            var txtbox9 = worksheet.Cells[10, 16];
            ApplySignatureBorders(txtbox9);
            // TARGET //               

            ExcelRange labelFormJ = GetExcelRange(worksheet, 11, 16, 12, 16);
            labelFormJ.Value = "PLAN DE ESTUDIOS A APLICAR";
            MergedCells(labelFormJ);
            FontWeightBold(labelFormJ);
            ContentCenter(labelFormJ);

            var txtbox10 = worksheet.Cells[10, 21];
            ApplySignatureBorders(txtbox10);
            // TARGET //               

            ExcelRange labelFormK = GetExcelRange(worksheet, 11, 21, 12, 21);
            labelFormK.Value = "CÓDIGO DEL PLAN DE ESTUDIOS A APLICAR";
            MergedCells(labelFormK);
            FontWeightBold(labelFormK);
            ContentCenter(labelFormK);

            #endregion

            #region FourthDivision

            var txtbox11 = worksheet.Cells[14, 2];
            ApplySignatureBorders(txtbox11);
            // TARGET //               

            ExcelRange labelFormL = GetExcelRange(worksheet, 15, 2, 16, 2);
            labelFormL.Value = "APELLIDOS Y NOMBRES DEL ESTUDIANTE";
            MergedCells(labelFormL);
            FontWeightBold(labelFormL);
            ContentCenter(labelFormL);

            var txtbox12 = worksheet.Cells[14, 9];
            ApplySignatureBorders(txtbox12);
            // TARGET //               

            ExcelRange labelFormM = GetExcelRange(worksheet, 15, 9, 16, 9);
            labelFormM.Value = "DOCUMENTO DE IDENTIDAD";
            MergedCells(labelFormM);
            FontWeightBold(labelFormM);
            ContentCenter(labelFormM);

            var txtbox13 = worksheet.Cells[14, 16];
            ApplySignatureBorders(txtbox13);
            // TARGET //               

            ExcelRange labelFormO = GetExcelRange(worksheet, 15, 16, 16, 16);
            labelFormO.Value = "CORREO ELECTRONICO";
            MergedCells(labelFormO);
            FontWeightBold(labelFormO);
            ContentCenter(labelFormO);

            var txtbox14 = worksheet.Cells[14, 21];
            ApplySignatureBorders(txtbox14);
            // TARGET //               

            ExcelRange labelFormP = GetExcelRange(worksheet, 15, 21, 16, 21);
            labelFormP.Value = "TELEFONO FIJO - CELULAR";
            MergedCells(labelFormP);
            FontWeightBold(labelFormP);
            ContentCenter(labelFormP);


            #endregion

            #region FifthDivision

            var txtbox15 = worksheet.Cells[18, 2];
            ApplySignatureBorders(txtbox15);
            // TARGET //               

            ExcelRange labelFormQ = GetExcelRange(worksheet, 19, 2, 20, 2);
            labelFormQ.Value = "INSTITUCIÓN DE DONDE PROVIENE";
            MergedCells(labelFormQ);
            FontWeightBold(labelFormQ);
            ContentCenter(labelFormQ);

            var txtbox16 = worksheet.Cells[18, 9];
            ApplySignatureBorders(txtbox16);
            // TARGET //               

            ExcelRange labelFormR = GetExcelRange(worksheet, 19, 9, 20, 9);
            labelFormR.Value = "PROGRAMA CURSADO";
            MergedCells(labelFormR);
            FontWeightBold(labelFormR);
            ContentCenter(labelFormR);

            var txtbox17 = worksheet.Cells[18, 16];
            ApplySignatureBorders(txtbox17);
            // TARGET //               

            ExcelRange labelFormS = GetExcelRange(worksheet, 19, 16, 20, 16);
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


            ExcelRange labelTitle1 = GetExcelRange(worksheet, 21, 1, 21, 21);
            labelTitle1.Value = "ESTRUCTURA CURRICULAR SENA";
            CellCenter(labelTitle1);
            MergedCells(labelTitle1);
            FontWeightBold(labelTitle1);

            ExcelRange context1 = GetExcelRange(worksheet, 22, 2, 22, 21);
            context1.Value = "ADMITIR AL USUARIO EN LA RED DE SERVICIOS DE SALUD SEGÚN NIVELES DE ATENCIÓN Y NORMATIVA VIGENTE.";
            ContentLeft(context1);
            MergedCells(context1);

            ExcelRange context2 = GetExcelRange(worksheet, 23, 2, 23, 21);
            context2.Value = "AFILIAR A LA POBLACIÓN AL SISTEMA GENERAL DE SEGURIDAD SOCIAL EN SALUD SEGÚN NORMATIVIDAD VIGENTE.";
            MergedCells(context2);
            ContentLeft(context2);

            ExcelRange context3 = GetExcelRange(worksheet, 24, 2, 24, 21);
            context3.Value = "FACTURAR LA PRESTACIÓN DE LOS SERVICIOS DE SALUD SEGÚN NORMATIVIDAD Y CONTRATACIÓN";
            MergedCells(context3);
            ContentLeft(context3);

            ExcelRange context4 = GetExcelRange(worksheet, 25, 2, 25, 21);
            context4.Value = "MANEJAR VALORES E INGRESOS RELACIONADOS CON LA OPERACIÓN DEL ESTABLECIMIENTO. (EQUIVALE A LA NORMA NTS 005 DEL MINCOMERCIO, INDUSTRIA Y TURISMO)";
            MergedCells(context4);
            ContentLeft(context4);

            ExcelRange context5 = GetExcelRange(worksheet, 26, 2, 26, 21);
            context5.Value = "ORIENTAR AL USUARIO EN RELACIÓN CON SUS NECESIDADES Y EXPECTATIVAS DE ACUERDO CON POLÍTICAS INSTITUCIONALES Y NORMAS DE SALUD VIGENTES.";
            MergedCells(context5);
            ContentLeft(context5);

            ExcelRange context6 = GetExcelRange(worksheet, 27, 2, 27, 21);
            context6.Value = "PROMOVER LA INTERACCION IDONEA CONSIGO MISMO, CON LOS DEMAS Y CON LA NATURALEZA EN LOS CONTEXTOS LABORAL Y SOCIAL.";
            MergedCells(context6);
            ContentLeft(context6);

            ExcelRange context7 = GetExcelRange(worksheet, 28, 2, 28, 21);
            context7.Value = "RESULTADOS DE APRENDIZAJE ETAPA PRACTICA";
            ApplySignatureBorders(context7);
            MergedCells(context7);
            ContentLeft(context7);

        }

        private void ConstructionHeaderTable(ExcelWorksheet worksheet)
        {
            ExcelRange HeaderBar = GetExcelRange(worksheet, 31, 1, 32, 21);
            CellCenter(HeaderBar);
            WrapText(HeaderBar);
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

            ExcelRange CellMT6 = GetExcelRange(worksheet, 31, 7, 32, 7);    // OK
            CellMT6.Value = "NIVEL";
            MergedCells(CellMT6);

            ExcelRange CellMT7 = GetExcelRange(worksheet, 31, 8, 32, 8);
            CellMT7.Value = "No";
            MergedCells(CellMT7);

            ExcelRange CellMT8 = GetExcelRange(worksheet, 31, 9, 32, 9);
            CellMT8.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT8);

            ExcelRange CellMT9 = GetExcelRange(worksheet, 31, 10, 31, 11);
            CellMT9.Value = "SISTEMA";
            MergedCells(CellMT9);

            ExcelRange CellMT10 = GetExcelRange(worksheet, 31, 12, 32, 12);
            CellMT10.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT10);

            ExcelRange CellMT11 = GetExcelRange(worksheet, 31, 13, 32, 13);
            CellMT11.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT11);

            ExcelRange CellMT12 = GetExcelRange(worksheet, 31, 14, 32, 14);     // OK
            CellMT12.Value = "NIVEL";
            MergedCells(CellMT12);

            ExcelRange CellMT13 = GetExcelRange(worksheet, 31, 15, 32, 15);
            CellMT13.Value = "No";
            MergedCells(CellMT13);

            ExcelRange CellMT14 = GetExcelRange(worksheet, 31, 16, 32, 16);
            CellMT14.Value = "ASIGNATURA Y/O CRÉDITO HOMOLOGADO";
            MergedCells(CellMT14);

            ExcelRange CellMT15 = GetExcelRange(worksheet, 31, 17, 31, 18);
            CellMT15.Value = "SISTEMA";
            MergedCells(CellMT15);

            ExcelRange CellMT16 = GetExcelRange(worksheet, 31, 19, 32, 19);
            CellMT16.Value = "CALIFICACIÓN NUMERICA";
            MergedCells(CellMT16);

            ExcelRange CellMT17 = GetExcelRange(worksheet, 31, 20, 32, 20);
            CellMT17.Value = "CALIFICACION LITERAL";
            MergedCells(CellMT17);

            ExcelRange CellMT18 = GetExcelRange(worksheet, 31, 21, 32, 21);
            CellMT18.Value = "NIVEL";
            MergedCells(CellMT18);

            #endregion

            #region SubCellsMainTable

            var subCellMT1 = worksheet.Cells[32, 3];
            subCellMT1.Value = "Créditos";

            var subCellMT2 = worksheet.Cells[32, 4];
            subCellMT2.Value = "Semestre";

            var subCellMT3 = worksheet.Cells[32, 10];
            subCellMT3.Value = "Créditos";

            var subCellMT4 = worksheet.Cells[32, 11];
            subCellMT4.Value = "Semestre";

            var subCellMT5 = worksheet.Cells[32, 17];
            subCellMT5.Value = "Créditos";

            var subCellMT6 = worksheet.Cells[32, 18];
            subCellMT6.Value = "Semestre";

            #endregion

        }


        private void ContentTable(ExcelWorksheet worksheet)
        {
            AsignaturasME asignaturasME = new AsignaturasME();
            List<AsignaturasME> subjects = GetSubjects.SubjectGenerator();

            int row = 33;
            int column = 1;
            int subjectsPerColumn = 15; // Cambiar de columna después de 9 elementos
            int subjectCount = 0; // Contador para llevar el seguimiento de los elementos en una columna

            // Llena el archivo de Excel con los datos
            foreach (var subject in subjects)
            {
                if (subjectCount >= subjectsPerColumn)
                {
                    // Cambiar de columna después de 9 elementos (cuenta 8 y 9 como una sola celda)
                    row = 33;
                    column += 7; // Cambia a la siguiente columna

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
                    worksheet.Cells[row, column].Value = subject.Numero[0];
                }

                worksheet.Cells[row, column + 1].Value = subject.Asignatura[0];
                worksheet.Cells[row, column + 2].Value = subject.Creditos[0];
                worksheet.Cells[row, column + 3].Value = subject.Semestre[0];
                worksheet.Cells[row, column + 4].Value = subject.CalificacionNumerica[0];
                worksheet.Cells[row, column + 5].Value = subject.CalificacionLiteral[0];
                worksheet.Cells[row, column + 6].Value = subject.Nivel[0];

                row++; // Avanzar una fila

                subjectCount++;
            }
        
            ExcelRange celdasMaterias = GetExcelRange(worksheet, 33, 1, 47, 21);
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

            ExcelRange LabelTotals = GetExcelRange(worksheet, 48, 1, 48, 2);
            LabelTotals.Value = "TOTALES";
            MergedCells(LabelTotals);
            CellCenter(LabelTotals);
            FontWeightBold(LabelTotals);

            var rangeToSum1 = worksheet.Cells["C33:C47"];
            var result1 = worksheet.Cells["C48"];
            ExcelRange totalSubjects1 = GetTotalSubjects1(worksheet);

            result1.Formula = $"SUM({rangeToSum1.Address})";
            CellCenter(result1);
            FontWeightBold(result1);

            var rangeToSum2 = worksheet.Cells["J33:J47"];
            var result2 = worksheet.Cells["J48"];
            ExcelRange totalSubjects2 = GetTotalSubjects2(worksheet);

            result2.Formula = $"SUM({rangeToSum2.Address})";
            CellCenter(result2);
            FontWeightBold(result2);

            var rangeToSum3 = worksheet.Cells["Q33:Q47"];
            var result3 = worksheet.Cells["Q48"];
            ExcelRange totalSubjects3 = GetTotalSubjects3(worksheet);

            result3.Formula = $"SUM({rangeToSum3.Address})";
            CellCenter(result3);
            FontWeightBold(result3);

            #endregion

        }

        private void RecognizedCredits(ExcelWorksheet worksheet)
        {
            // EVALUAR CRITERIOS DE LA CARRERA (1)TECNICO (2)TECNOLOGICO (3)PROFESIONAL
            // IDENTIFICAR LOS CREDITOS QUE SE APRUEBAN, PENDIENTES Y EL TOTAL DE CREDITOS(POR LO CUAL SE DEPENDE DEL PUNTO ANTERIOR

            ExcelRange totalSubjects1 = GetTotalSubjects1(worksheet);
            ExcelRange totalSubjects2 = GetTotalSubjects2(worksheet);
            ExcelRange totalSubjects3 = GetTotalSubjects3(worksheet);

            ExcelRange cellLabelsRecCred = GetExcelRange(worksheet, 53, 1, 55, 4);
            CellRight(cellLabelsRecCred);
            FontWeightBold(cellLabelsRecCred);

            ExcelRange CellLabelRc1 = GetExcelRange(worksheet, 53, 1, 53, 4);
            MergedCells(CellLabelRc1);
            CellLabelRc1.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TÉCNICO PROFESIONAL";

            ExcelRange CellLabelRc2 = GetExcelRange(worksheet, 54, 1, 54, 4);
            MergedCells(CellLabelRc2);
            CellLabelRc2.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL TECNOLÓGICO ";

            ExcelRange CellLabelRc3 = GetExcelRange(worksheet, 55, 1, 55, 4);
            MergedCells(CellLabelRc3);
            CellLabelRc3.Value = "TOTAL CRÉDITOS RECONOCIDOS PARA EL NIVEL PROFESIONAL";

            ExcelRange LabelTotalMaterias = GetExcelRange(worksheet, 51, 5, 51, 7);
            MergedCells(LabelTotalMaterias);
            CellCenter(LabelTotalMaterias);
            FontWeightBold(LabelTotalMaterias);


            var sumCellTotals = worksheet.Cells["E51"];
            sumCellTotals.Formula = $"SUM({totalSubjects1.Address},{totalSubjects2.Address},{totalSubjects3.Address})";


            var labelApproved = worksheet.Cells[52, 5];
            labelApproved.Value = "APRO";
            CellCenter(labelApproved);
            FontWeightBold(labelApproved);

            var labelPending = worksheet.Cells[52, 6];
            labelPending.Value = "PEN";
            CellCenter(labelPending);
            FontWeightBold(labelPending);

            var labelTotalCredits = worksheet.Cells[52, 7];
            labelTotalCredits.Value = "TOTAL CRED";
            CellCenter(labelTotalCredits);
            FontWeightBold(labelTotalCredits);

            ExcelRange celdasMateriasTotales = GetExcelRange(worksheet, 53, 5, 55, 7);
            CellCenter(celdasMateriasTotales);

            // IMPORTANTE: CODIGO PROPENSO A SER MODIFICADO //

            var cellsToSetZero = new List<ExcelRange>
                {
                    worksheet.Cells["G53"],
                    worksheet.Cells["E53"],
                    worksheet.Cells["G54"],
                    worksheet.Cells["E54"],
                    worksheet.Cells["G55"],
                    worksheet.Cells["E55"]
                };

            foreach (var cell in cellsToSetZero)
            {
                cell.Value = 0;
            }

            // FORMULAS TO DETERMINE CREDITS

            // CREDITOS TECNICO PROFESIONAL (1)
            var totalTechnicalCredits = worksheet.Cells["G53"]; // 3
            var totalTechnicalApproved = worksheet.Cells["E53"]; // 1
            var totalTechnicalPending = worksheet.Cells["F53"]; // 2
            totalTechnicalPending.Formula = $"({totalTechnicalCredits.Address}) - ({totalTechnicalApproved.Address})";

            // TOTAL CREDITOS APROBADOS (2)
            var totalTechnologistCredits = worksheet.Cells["G54"];
            var totalTechnologistApproved = worksheet.Cells["E54"];
            var totalTechnologistPending = worksheet.Cells["F54"];
            totalTechnologistPending.Formula = $"({totalTechnologistCredits.Address}) - ({totalTechnologistApproved.Address})";

            // TOTAL CREDITOS PENDIENTES (3)
            var totalProfessionalCredits = worksheet.Cells["G55"];
            var totalProfessionalApproved = worksheet.Cells["E55"];
            var totalProfessionalPending = worksheet.Cells["F55"];
            totalProfessionalPending.Formula = $"({totalProfessionalCredits.Address}) - ({totalProfessionalApproved.Address})";
        }

        private void ConditionLabels(ExcelWorksheet worksheet)
        {
            ExcelRange conditionText = worksheet.Cells[57, 1, 61, 21];
            ContentLeft(conditionText);

            ExcelRange labelTitle2 = worksheet.Cells[57, 1, 57, 2];
            labelTitle2.Value = "Aclaraciones";
            FontWeightBold(labelTitle2);
            MergedCells(labelTitle2);

            ExcelRange contextA = worksheet.Cells[58, 1, 58, 21];
            MergedCells(contextA);
            WrapText(contextA);
            contextA.Value = "Los créditos académicos faltantes para cumplir a cabalidad con la oferta académica del nivel técnico profesional y/o tecnológico y/o " +
                "profesional deben ser cursados y aprobados conforme las reglamentaciones institucionales vigentes.";

            ExcelRange contextB = worksheet.Cells[59, 1, 59, 21];
            MergedCells(contextB);
            WrapText(contextB);
            contextB.Value = "La Escuela de Ciencias Administrativas de la Corporación Unificada Nacional -CUN, reconoce las asignaturas del programa de  Administración " +
                "de Servicios de Salud   del nivel técnico ( 1 - 3  semestre) para que dar continuidad a su proceso de formación académica a partir de las asignaturas de" +
                " nivel tecnológico correspondientes a (  4 - 5  semestre) Y profesional correspondientes a ( 6 - 9  semestre)";

            ExcelRange contextC = worksheet.Cells[60, 1, 60, 21];
            MergedCells(contextC);
            WrapText(contextC);
            contextC.Value = "La prueba TyT del ciclo técnico es homologable en la institución. El estudiante deberá presentar la prueba saber TyT para el ciclo " +
                "tecnológico y Saber PRO para el ciclo profesional.";

            ExcelRange contextD = worksheet.Cells[61, 1, 61, 21];
            MergedCells(contextD);
            WrapText(contextD);
            contextD.Value = "Teniendo en cuenta que el plan de estudios vigente del programa  Administración de Servicios de Salud  no incluye los respectivos niveles " +
                "de inglés requeridos para obtener las diferentes titulaciones, el estudiante deberá garantizar lo pertinente al momento de radicar su solicitud de " +
                "grado, para ello se cuenta con la oferta del centro de Idiomas de la institución.";
        }

        private void StudentStatement(ExcelWorksheet worksheet)
        {
            ExcelRange studentStatementArea = worksheet.Cells[56, 1, 57, 24];
            ContentLeft(studentStatementArea);

            ExcelRange labelTitle3 = worksheet.Cells[63, 1, 63, 2];
            labelTitle3.Value = "Manifestación expresa del estudiante";
            FontWeightBold(labelTitle3);
            MergedCells(labelTitle3);

            ExcelRange contextE = worksheet.Cells[64, 1, 64, 21];
            contextE.Value = "Con el presente documento manifiesto expresamente y sin que medie ninguna clase de vicio o limitación a mi consentimiento, mi plena conformidad con las asignaturas y/o créditos reconocidos u homologados para mi ingreso al nivel técnico profesional y/o tecnológico y/o profesional del programa   Administración de Servicios de Salud    de la Corporación Unificada Nacional de Educación Superior CUN. Las competencias que considere me hagan falta, del ciclo técnico, las podré realizar voluntariamente a través de tutorías en cada área transversal o del programa, talleres nivelatorios y/o participando como asistente a clases sin que estos generen nota alguna y solicitando previamente el ingreso a la clase o tutoría.";
            MergedCells(contextE);
            WrapText(contextE);
        }

        
            //var labelPending = worksheet.Cells[45, 9];

        private void ContractArea(ExcelWorksheet worksheet)
        {
            ExcelRange studentStatementArea = worksheet.Cells[70, 1, 79, 21];
            ContentLeft(studentStatementArea);

            ExcelRange labelTitle4 = worksheet.Cells[70, 1, 70, 5]; // OK 
            labelTitle4.Value = "En Constancia de lo anterior firman:";
            FontWeightBold(labelTitle4);
            MergedCells(labelTitle4);

            var parrafSquare = worksheet.Cells[73, 7];
            parrafSquare.Value = "finalParraf";
            ContentLeft(parrafSquare);


            // INFORMATION: JEFE DE PROGRAMA
            ExcelRange cellSignatureProgramManager = worksheet.Cells[73, 1, 73, 2]; // OK 
            MergedCells(cellSignatureProgramManager);
            ApplySignatureBorders(cellSignatureProgramManager);
            ProgramManagerSignature(worksheet);

            var labelSignature1 = worksheet.Cells[74, 1]; // OK 
            labelSignature1.Value = "Líder de Programa: ";

            var labelSignature2 = worksheet.Cells[75, 1]; // OK 
            labelSignature2.Value = "Nombre:";
            var nameProgramLeader = worksheet.Cells[75, 2];
            FontWeightBold(nameProgramLeader);
            // TARGET

            // INFORMATION: ESTUDIANTE
            ExcelRange cellSignatureStudent = worksheet.Cells[73, 8, 73, 9]; // OK
            MergedCells(cellSignatureStudent);
            ApplySignatureBorders(cellSignatureStudent);
            StudentSignature(worksheet);

            var labelSignature3 = worksheet.Cells[74, 8]; // OK 
            labelSignature3.Value = "Estudiante: "; 

            var labelSignature4 = worksheet.Cells[75, 8]; // OK 
            labelSignature4.Value = "Nombre: "; 
            var nameStudent = worksheet.Cells[75, 9];
            FontWeightBold(nameStudent);
            // TARGET

            var labelSignature5 = worksheet.Cells[76, 8]; // OK 
            labelSignature5.Value = "Doc de Identidad: ";
            var docNumber = worksheet.Cells[76, 9];
            FontWeightBold(docNumber);
            // TARGET

        }

        private void ApplyCellsFooter(ExcelWorksheet worksheet)
        {
            ExcelRange tableFooter = worksheet.Cells[78, 1, 79, 14];
            CellLeft(tableFooter);

            var cellFooter1 = worksheet.Cells[78, 1];
            cellFooter1.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter2 = worksheet.Cells[79, 1];
            cellFooter2.Value = "FECHA: "; // Convertir y generar valor dinámico

                ExcelRange cellSpaceFooter1 = worksheet.Cells[78, 2, 78, 4];
                MergedCells(cellSpaceFooter1);
                // TARGET

                ExcelRange cellSpaceFooter2 = worksheet.Cells[79, 2, 79, 4];
                MergedCells(cellSpaceFooter2);
                // TARGET

            var cellFooter3 = worksheet.Cells[78, 5];
            cellFooter3.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter4 = worksheet.Cells[79, 5];
            cellFooter4.Value = "FECHA: "; // Convertir y generar valor dinámico

                ExcelRange cellSpaceFooter3 = worksheet.Cells[78, 6, 78, 9];
                MergedCells(cellSpaceFooter3);
                // TARGET

                ExcelRange cellSpaceFooter4 = worksheet.Cells[79, 6, 79, 9];
                MergedCells(cellSpaceFooter4);
                // TARGET

            var cellFooter5 = worksheet.Cells[78, 10];
            cellFooter5.Value = "ELABORÓ: "; // Convertir y generar valor dinámico

            var cellFooter6 = worksheet.Cells[79, 10];
            cellFooter6.Value = "FECHA: "; // Convertir y generar valor dinámico

                ExcelRange cellSpaceFooter5 = worksheet.Cells[78, 11, 78, 14];
                MergedCells(cellSpaceFooter5);
                // TARGET

                ExcelRange cellSpaceFooter6 = worksheet.Cells[79, 11, 79, 14];
                MergedCells(cellSpaceFooter6);
                // TARGET
        }
    }
}
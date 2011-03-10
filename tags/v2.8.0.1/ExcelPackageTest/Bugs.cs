using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace ExcelPackageTest
{
    public static class ex
    {
        public static void PintarNota(this ExcelRange rango, decimal? nota)
        {
            if (!nota.HasValue) return;
            rango.Value = nota;
            if (nota < 4)
            {
                rango.Style.Font.Color.SetColor(Color.Red);
                rango.Style.Font.Bold = true;
            }
        }
    }
    [TestClass]
    public class Bugs
    {
        private static readonly string[] GlosaMeses = new string[] { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };


        [TestMethod]
        public void TestMethod1()
        {
	    using (ExcelPackage pck = new ExcelPackage())
        {
                    ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("Detalle IGS");

                    //worksheet.Cells[2, 2].Value = "RESULTADOS " + GlosaMeses[informe.Fecha.Month-1].ToUpper() + "  IGS ZONALES";

                    using (ExcelRange rng = worksheet.Cells[2, 2, 4, 8])
                    {
                        rng.Style.Font.Bold = true;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid; 
                        rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#99CCFF"));
                        rng.Style.Font.Color.SetColor(Color.Black);
                    }
                
                    const int filaBase = 5;
                    int filaActual = filaBase;

                    using (ExcelRange rng = worksheet.Cells[filaActual, 2, filaActual, 8])
                    {
                        rng.Style.Font.Bold = true;
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#969696")); 
                        rng.Style.Font.Color.SetColor(Color.White);
                    }

                    worksheet.Cells[filaActual, 2].Value = "Jefe Zonal";
                    worksheet.Cells[filaActual, 3].Value = "N° Local";
                    worksheet.Cells[filaActual, 4].Value = "Cuenta de IGS";
                    worksheet.Cells[filaActual, 5].Value = "Promedio de IGS2";
                    worksheet.Cells[filaActual, 5].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF00"));
                    worksheet.Cells[filaActual, 5].Style.Font.Color.SetColor(Color.Black);
                    worksheet.Cells[filaActual, 6].Value = "Promedio de I4P";
                    worksheet.Cells[filaActual, 6].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#000080"));
                    worksheet.Cells[filaActual, 7].Value = "Promedio de IA";
                    worksheet.Cells[filaActual, 7].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#008080"));
                    worksheet.Cells[filaActual, 8].Value = "Promedio de IP";
                    worksheet.Cells[filaActual, 8].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FF0000"));

                    using (ExcelRange rng = worksheet.Cells[filaBase, 5, filaActual, 8])
                    {
                        rng.Style.Numberformat.Format = "#.0";
                        rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }

                    using (ExcelRange rng = worksheet.Cells[filaActual, 2, filaActual, 8])
                    {
                        rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rng.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF99"));
                    }

                    pck.SaveAs(new FileInfo("c:\\temp\\bug1.xlsx"));
            }
        }
    }
}

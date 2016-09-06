using System;
using System.Collections.Generic;
using System.Linq;

namespace Creatidea.Library.Office.MsOffice
{
    using System.Diagnostics;
    using System.IO;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Interop.PowerPoint;

    /// <summary>
    /// Microsoft Office OfficeConverter.
    /// </summary>
    public class OfficeConverter
    {
        /// <summary>
        /// Words to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input file path.</param>
        /// <returns>Pdf檔案路徑.</returns>
        [STAThread]
        public static string WordToPdf(string inputFilePath)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".doc") && !ext.Equals(".docx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.doc, *.docx", nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            var wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(inputFilePath);

            outputPath = Path.Combine(outputDir, outputFileName + ".pdf");
            doc.SaveAs(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
            wordApp.Visible = false;
            wordApp.Quit();

            return outputPath;
        }

        /// <summary>
        /// Excels to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input file path.</param>
        /// <param name="paperSize">Size of the paper.</param>
        /// <param name="orientation">The paper orientation.</param>
        /// <returns>Pdf檔案路徑.</returns>
        [STAThread]
        public static string ExcelToPdf(string inputFilePath, XlPaperSize paperSize = XlPaperSize.xlPaperA4, XlPageOrientation orientation = XlPageOrientation.xlLandscape)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".xls") && !ext.Equals(".xlsx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.xls, *.xlsx", nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook book = excelApp.Workbooks.Open(inputFilePath);

            foreach (Worksheet sheet in book.Sheets)
            {
                sheet.PageSetup.PaperSize = paperSize;
                sheet.PageSetup.Orientation = orientation;

                // Fit Sheet on One Page 
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;
            }

            outputPath = Path.Combine(outputDir, outputFileName + ".pdf");
            book.SaveAs(outputPath, (XlFileFormat)57);
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            excelApp.Quit();

            return outputPath;
        }

        /// <summary>
        /// PPTs to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input file path.</param>
        /// <returns>Pdf檔案路徑.</returns>
        public static string PptToPdf(string inputFilePath)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".ppt") && !ext.Equals(".pptx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.ppt, *.pptx", nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation presentation = pptApp.Presentations.Open(
                inputFilePath,
                MsoTriState.msoTrue,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse);

            outputPath = Path.Combine(outputDir, outputFileName + ".pdf");
            presentation.SaveAs(outputPath, PpSaveAsFileType.ppSaveAsPDF);
            pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            pptApp.Quit();

            return outputPath;
        }
    }
}

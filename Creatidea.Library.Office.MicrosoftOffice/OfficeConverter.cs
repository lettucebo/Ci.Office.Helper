using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.MsOffice
{
    using System.Diagnostics;
    using System.IO;

    using Creatidea.Library.Results;

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
        /// <returns>CiResult&lt;System.String&gt;Pdf檔案路徑.</returns>
        [STAThread]
        public static CiResult<string> WordToPdf(string inputFilePath)
        {
            var result = new CiResult<string>() { Message = "Word轉Pdf失敗！" };

            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".doc") && !ext.Equals(".docx")))
            {
                result.Message += "副檔名錯誤！應為*.doc, *.docx";
                return result;
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                result.Message += string.Format("找不到檔案：{0}！", inputFilePath);
                return result;
            }

            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(inputFilePath);

                outputPath = Path.Combine(outputDir, outputFileName + ".pdf");
                doc.SaveAs(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                wordApp.Visible = false;
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                result.Message += ex.ToString();
                return result;
            }

            result.Success = true;
            result.Message = "Word轉Pdf成功";
            result.Data = outputPath;

            return result;
        }

        /// <summary>
        /// Excels to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input file path.</param>
        /// <param name="paperSize">Size of the paper.</param>
        /// <param name="orientation">The paper orientation.</param>
        /// <returns>CiResult&lt;System.String&gt;Pdf檔案路徑.</returns>
        [STAThread]
        public static CiResult<string> ExcelToPdf(string inputFilePath, XlPaperSize paperSize = XlPaperSize.xlPaperA4, XlPageOrientation orientation = XlPageOrientation.xlLandscape)
        {
            var result = new CiResult<string>() { Message = "Excel轉Pdf失敗！" };

            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".xls") && !ext.Equals(".xlsx")))
            {
                result.Message += "副檔名錯誤！應為*.xls, *.xlsx";
                return result;
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                result.Message += string.Format("找不到檔案：{0}！", inputFilePath);
                return result;
            }

            try
            {
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
            }
            catch (Exception ex)
            {
                result.Message += ex.ToString();
                return result;
            }

            result.Success = true;
            result.Message = "Excel轉Pdf成功";
            result.Data = outputPath;

            return result;
        }

        public static CiResult<string> PptToPdf(string inputFilePath)
        {
            var result = new CiResult<string>() { Message = "Ppt轉Pdf失敗！" };

            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".ppt") && !ext.Equals(".pptx")))
            {
                result.Message += "副檔名錯誤！應為*.ppt, *.pptx";
                return result;
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Guid.NewGuid().ToString().ToUpper();
            string outputPath = string.Empty;

            if (!File.Exists(inputFilePath))
            {
                result.Message += string.Format("找不到檔案：{0}！", inputFilePath);
                return result;
            }

            try
            {
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
            }
            catch (Exception ex)
            {
                result.Message += ex.ToString();
                return result;
            }

            result.Success = true;
            result.Message = "Ppt轉Pdf成功";
            result.Data = outputPath;

            return result;
        }
    }
}

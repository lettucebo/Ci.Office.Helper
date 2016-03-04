using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.Example
{
    using System.IO;

    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// 範例程式
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        public static void Main(string[] args)
        {
            string appDirectory = Directory.GetCurrentDirectory();
            string docPath = Path.Combine(appDirectory, "Demo\\Word", "Demo.doc");
            string docxPath = Path.Combine(appDirectory, "Demo\\Word", "Demo.docx");
            string xlsPath = Path.Combine(appDirectory, "Demo\\Excel", "Demo.xls");
            string xlsxPath = Path.Combine(appDirectory, "Demo\\Excel", "Demo.xlsx");

            DemoLibreOffice(docPath, docxPath);

            DemoMicrosoftOffice(docPath, docxPath, xlsPath, xlsxPath);
        }

        /// <summary>
        /// Demoes the microsoft office.
        /// </summary>
        /// <param name="docPath">The doc path.</param>
        /// <param name="docxPath">The docx path.</param>
        /// <param name="xlsPath">The XLS path.</param>
        /// <param name="xlsxPath">The XLSX path.</param>
        private static void DemoMicrosoftOffice(string docPath, string docxPath, string xlsPath, string xlsxPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("示範 MicrosoftOffice");

            Console.WriteLine();
            Console.WriteLine("doc 轉為 pdf：");
            var docResult = MsOffice.OfficeConverter.WordToPdf(docPath);
            if (!docResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", docResult.Message);
            }
            else
            {
                var link = SaveFile(docResult.Data, "msdoc.pdf");
                Console.WriteLine("Show docResult: {0}", link);
            }

            Console.WriteLine();
            Console.WriteLine("docx 轉為 pdf：");
            var docxResult = MsOffice.OfficeConverter.WordToPdf(docxPath);
            if (!docxResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", docxResult.Message);
            }
            else
            {
                var link = SaveFile(docxResult.Data, "msdocx.pdf");
                Console.WriteLine("Show docxResult: {0}", link);
            }

            Console.WriteLine();
            Console.WriteLine("xls 轉為 pdf：");
            var xlsResult = MsOffice.OfficeConverter.ExcelToPdf(xlsPath);
            if (!xlsResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", xlsResult.Message);
            }
            else
            {
                var link = SaveFile(xlsResult.Data, "msxls.pdf");
                Console.WriteLine("Show xlsResult: {0}", link);
            }

            Console.WriteLine();
            Console.WriteLine("xlsx 轉為 pdf：");
            var xlsxResult = MsOffice.OfficeConverter.ExcelToPdf(xlsxPath, XlPaperSize.xlPaperB5, XlPageOrientation.xlLandscape);
            if (!xlsxResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", xlsxResult.Message);
            }
            else
            {
                var link = SaveFile(xlsxResult.Data, "msxlsx.pdf");
                Console.WriteLine("Show xlsxResult: {0}", link);
            }
        }

        /// <summary>
        /// Demoes the libre office.
        /// </summary>
        /// <param name="docPath">The doc path.</param>
        /// <param name="docxPath">The docx path.</param>
        [STAThread]
        private static void DemoLibreOffice(string docPath, string docxPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("示範 LibreOffice");

            Console.WriteLine();
            Console.WriteLine("doc 轉為 pdf：");
            var docResult = LibreOffice.OfficeConverter.WordToPdf(docPath);
            if (!docResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", docResult.Message);
            }
            else
            {
                var link = SaveFile(docResult.Data, "libredoc.pdf");
                Console.WriteLine("Show docResult: {0}", link);
            }

            Console.WriteLine();
            Console.WriteLine("docx 轉為 pdf：");
            var docxResult = LibreOffice.OfficeConverter.WordToPdf(docxPath);
            if (!docxResult.Success)
            {
                Console.WriteLine("發生錯誤：{0}", docxResult.Message);
            }
            else
            {
                var link = SaveFile(docxResult.Data, "libredocx.pdf");
                Console.WriteLine("Show docxResult: {0}", link);
            }
        }

        /// <summary>
        /// Saves the file.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <returns>System.String.</returns>
        private static string SaveFile(string path, string fileName)
        {
            string appDirectory = Directory.GetCurrentDirectory();
            string docPath = Path.Combine(appDirectory, "Temp", fileName);

            FileInfo file = new FileInfo(docPath);
            // If the directory already exists, this method does nothing.
            file.Directory.Create();

            File.Copy(path, docPath, true);

            return fileName;
        }
    }
}

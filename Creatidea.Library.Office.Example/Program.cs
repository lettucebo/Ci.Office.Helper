using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.Example
{
    using System.IO;

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

            DemoLibreOffice(docPath, docxPath);
        }

        /// <summary>
        /// Demoes the libre office.
        /// </summary>
        /// <param name="docPath">The document path.</param>
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
                var link = SaveFile(docResult.Data, "doc.pdf");
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
                var link = SaveFile(docxResult.Data, "docx.pdf");
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

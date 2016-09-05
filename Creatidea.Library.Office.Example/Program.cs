using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.Example
{
    using System.IO;

    using Ci.Extensions;

    using Creatidea.Library.Office.LibreOffice.Enums;
    using Creatidea.Library.Office.OpenXml;
    using Creatidea.Library.Office.OpenXml.Enums;
    using Creatidea.Library.Office.OpenXml.Manager;
    using Creatidea.Library.Office.OpenXml.Models;

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
            string pptPath = Path.Combine(appDirectory, "Demo\\Ppt", "Demo.ppt");
            string pptxPath = Path.Combine(appDirectory, "Demo\\Ppt", "Demo.pptx");
            string docxTemplatePath = Path.Combine(appDirectory, "Demo\\Word", "Template.docx");

            Console.WriteLine("Office Relate Library Demo:");
            Console.WriteLine("1. Use LibreOffice convert ms document type to Open Document Format(odf)");
            Console.WriteLine("2. Use LibreOffice convert document to pdf");
            Console.WriteLine("3. Use Microsoft Office convert document to pdf");
            Console.WriteLine("4. Use Open Xml template to docx");

            Console.Write("Please enter the option: ");
            var chooese = Console.Read();

            switch (chooese)
            {
                case 49:
                    UseLibreOfficeFromMsToOdf(docPath, docxPath, xlsPath, xlsxPath, pptPath, pptxPath);
                    break;
                case 50:
                    UseLibreOfficeFromMsToPdf(docPath, docxPath, xlsPath, xlsxPath, pptPath, pptxPath);
                    break;
                case 51:
                    DemoMicrosoftOffice(docPath, docxPath, xlsPath, xlsxPath, pptPath, pptxPath);
                    break;
                case 52:
                    UseOpenXmlToDocx(docxTemplatePath);
                    break;
                default:
                    Console.WriteLine("Unknow option.");
                    break;
            }
        }

        private static void UseOpenXmlToDocx(string docxTemplatePath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("Open Xml template to docx");

            Console.WriteLine();
            Console.WriteLine("template to docx：");

            #region parameter dictionary

            var textDict = new Dictionary<string, OpenXmlTextInfo>
             {
               {"SDT01",new OpenXmlTextInfo(){Text = "12345678", IsInnerXml = false}},
               {"SDT02",new OpenXmlTextInfo(){Text = "Money Yu", IsInnerXml = false}},
               {"SDT03",new OpenXmlTextInfo(){Text = "Creatidea", IsInnerXml = false}},
               {"SDT04",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},

               {"SDT05",new OpenXmlTextInfo(){Text = "Abc12207" , IsInnerXml = false}},
               {"SDT06",new OpenXmlTextInfo(){Text = "+886912345678", IsInnerXml = false}},

                // start demo: for parameter initialize, will update parameter later
               {"SDT07",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT08",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT09",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT10",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT11",new OpenXmlTextInfo(){Text = string.Empty, IsInnerXml = false}},

               {"SDT12",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT13",new OpenXmlTextInfo(){Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true}},
               {"SDT14",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
               {"SDT15",new OpenXmlTextInfo(){Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true}},
               {"SDT16",new OpenXmlTextInfo(){Text = WordSymbols.UnChecked.GetDescription(), IsInnerXml = true}},
                // end demo: for parameter initialize, will update parameter later

               {"SDT17",new OpenXmlTextInfo(){Text = "9F, No.1, Sec.4, Nanjing E. Rd., Songshan Dist., Taipei City, Taiwan ", IsInnerXml = false}},
               {"SDT18",new OpenXmlTextInfo(){Text = "2806-1479", IsInnerXml = false}},
               {"SDT19",new OpenXmlTextInfo(){Text = "For Test", IsInnerXml = false}},
               {"SDT20",new OpenXmlTextInfo(){Text = "100", IsInnerXml = false}},
               {"SDT21",new OpenXmlTextInfo(){Text = "200", IsInnerXml = false}},
               {"SDT22",new OpenXmlTextInfo(){Text = "300", IsInnerXml = false}},
               {"SDT23",new OpenXmlTextInfo(){Text = "3", IsInnerXml = false}},
               {"SDT24",new OpenXmlTextInfo(){Text = "2", IsInnerXml = false}},
               {"SDT25",new OpenXmlTextInfo(){Text = "1", IsInnerXml = false}},
               {"SDT26",new OpenXmlTextInfo(){Text = "1", IsInnerXml = false}},
               {"SDT27",new OpenXmlTextInfo(){Text = "abc12207@gmail.com", IsInnerXml = false}}
            };

            var imageDict = new Dictionary<string, MemoryStream>();

            imageDict.Add("SDTI01", WordManager.GetStreamFromImagePath(Path.Combine(Directory.GetCurrentDirectory(), "Demo\\Image", "Creatidea.png")));
            imageDict.Add("SDTI02", WordManager.GetStreamFromImagePath("C:/NotExistImage.jpg"));

            // start demo: use dynamic parameter, if random%2 == 0 ,then set to checked
            Random rand = new Random(Guid.NewGuid().GetHashCode());
            if (rand.Next(1, 3) % 2 == 0)
            {
                textDict["SDT07"] = new OpenXmlTextInfo()
                { Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true };
            }

            if (rand.Next(1, 3) % 2 == 0)
            {
                textDict["SDT08"] = new OpenXmlTextInfo()
                { Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true };
            }

            if (rand.Next(1, 3) % 2 == 0)
            {
                textDict["SDT09"] = new OpenXmlTextInfo()
                { Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true };
            }

            if (rand.Next(1, 3) % 2 == 0)
            {
                textDict["SDT10"] = new OpenXmlTextInfo()
                { Text = WordSymbols.Checked.GetDescription(), IsInnerXml = true };
            }
            // end demo: use dynamic parameter

            #endregion

            var template = new Template();
            var filePath = template.DocxMaker(docxTemplatePath, textDict, imageDict);
            var linkdocx = SaveFile(filePath, "templatedocx.docx");
            Console.WriteLine("Show docResult: {0}", linkdocx);
            Console.WriteLine(Enum.GetName(typeof(WordSymbols), WordSymbols.UnChecked));
        }

        /// <summary>
        /// Demoes the microsoft office.
        /// </summary>
        /// <param name="docPath">The doc path.</param>
        /// <param name="docxPath">The docx path.</param>
        /// <param name="xlsPath">The XLS path.</param>
        /// <param name="xlsxPath">The XLSX path.</param>
        /// <param name="pptPath">The PPT path.</param>
        /// <param name="pptxPath">The PPTX path.</param>
        private static void DemoMicrosoftOffice(
            string docPath,
            string docxPath,
            string xlsPath,
            string xlsxPath,
            string pptPath,
            string pptxPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("Microsoft Office to pdf");

            Console.WriteLine();
            Console.WriteLine("doc 轉為 pdf：");
            var docResult = MsOffice.OfficeConverter.WordToPdf(docPath);
            var linkdoc = SaveFile(docResult, "msdoc.pdf");
            Console.WriteLine("Show docResult: {0}", linkdoc);

            Console.WriteLine();
            Console.WriteLine("docx 轉為 pdf：");
            var docxResult = MsOffice.OfficeConverter.WordToPdf(docxPath);
            var linkdocx = SaveFile(docxResult, "msdocx.pdf");
            Console.WriteLine("Show docxResult: {0}", linkdocx);

            Console.WriteLine();
            Console.WriteLine("xls 轉為 pdf：");
            var xlsResult = MsOffice.OfficeConverter.ExcelToPdf(xlsPath);
            var linkxls = SaveFile(xlsResult, "msxls.pdf");
            Console.WriteLine("Show xlsResult: {0}", linkxls);

            Console.WriteLine();
            Console.WriteLine("xlsx 轉為 pdf：");
            // 一定使用輸出為整頁
            var xlsxResult = MsOffice.OfficeConverter.ExcelToPdf(xlsxPath);
            // 提供尺寸與方向選項
            var xlsxResult2 = MsOffice.OfficeConverter.ExcelToPdf(
                xlsxPath,
                XlPaperSize.xlPaperB4,
                XlPageOrientation.xlPortrait);
            var linkxlsx = SaveFile(xlsxResult, "msxlsx.pdf");
            Console.WriteLine("Show xlsxResult: {0}", linkxlsx);

            Console.WriteLine();
            Console.WriteLine("ppt 轉為 pdf：");
            var pptResult = MsOffice.OfficeConverter.PptToPdf(pptPath);
            var linkppt = SaveFile(pptResult, "msppt.pdf");
            Console.WriteLine("Show pptResult: {0}", linkppt);

            Console.WriteLine();
            Console.WriteLine("pptx 轉為 pdf：");
            var pptxResult = MsOffice.OfficeConverter.PptToPdf(pptxPath);
            var linkpptx = SaveFile(pptxResult, "mspptx.pdf");
            Console.WriteLine("Show pptxResult: {0}", linkpptx);
        }

        /// <summary>
        /// Demoes the libre office.
        /// </summary>
        /// <param name="docPath">The doc path.</param>
        /// <param name="docxPath">The docx path.</param>
        /// <param name="xlsPath">The XLS path.</param>
        /// <param name="xlsxPath">The XLSX path.</param>
        /// <param name="pptPath">The PPT path.</param>
        /// <param name="pptxPath">The PPTX path.</param>
        [STAThread]
        private static void UseLibreOfficeFromMsToOdf(
            string docPath,
            string docxPath,
            string xlsPath,
            string xlsxPath,
            string pptPath,
            string pptxPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("LibreOffice to odf");

            Console.WriteLine();
            Console.WriteLine("doc 轉為 odt：");
            var odtResult = LibreOffice.OfficeConverter.ConvertDocument(docPath, OutputExtensionType.odt);
            var linkodt = SaveFile(odtResult, "libredoc.odt");
            Console.WriteLine("Show odtResult: {0}", linkodt);

            Console.WriteLine();
            Console.WriteLine("xls 轉為 ods：");
            var odsResult = LibreOffice.OfficeConverter.ConvertDocument(xlsPath, OutputExtensionType.ods);
            var linkods = SaveFile(odsResult, "librexls.ods");
            Console.WriteLine("Show odsResult: {0}", linkods);

            Console.WriteLine();
            Console.WriteLine("ppt 轉為 odp：");
            var odpResult = LibreOffice.OfficeConverter.ConvertDocument(pptPath, OutputExtensionType.odp);
            var linkodp = SaveFile(odpResult, "libreppt.odp");
            Console.WriteLine("Show odpResult: {0}", linkodp);

            Console.WriteLine();
            Console.WriteLine("docx 轉為 odt：");
            var docxResult = LibreOffice.OfficeConverter.ConvertDocument(docxPath, OutputExtensionType.odt);
            var linkdocx = SaveFile(docxResult, "libredocx.odt");
            Console.WriteLine("Show odtResult: {0}", linkdocx);

            Console.WriteLine();
            Console.WriteLine("xlsx 轉為 ods：");
            var xlsxResult = LibreOffice.OfficeConverter.ConvertDocument(xlsxPath, OutputExtensionType.ods);
            var linkxlsx = SaveFile(xlsxResult, "librexlsx.ods");
            Console.WriteLine("Show xlsResult: {0}", linkxlsx);

            Console.WriteLine();
            Console.WriteLine("pptx 轉為 pdf：");
            var pptxResult = LibreOffice.OfficeConverter.ConvertDocument(pptxPath, OutputExtensionType.odp);
            var linkpptx = SaveFile(pptxResult, "librepptx.odp");
            Console.WriteLine("Show pptxResult: {0}", linkpptx);
        }

        private static void UseLibreOfficeFromMsToPdf(
            string docPath,
            string docxPath,
            string xlsPath,
            string xlsxPath,
            string pptPath,
            string pptxPath)
        {
            Console.WriteLine("========================================");
            Console.WriteLine("LibreOffice to pdf");

            Console.WriteLine();
            Console.WriteLine("doc 轉為 pdf：");
            var docResult = LibreOffice.OfficeConverter.ConvertDocument(docPath, OutputExtensionType.pdf);
            var linkdoc = SaveFile(docResult, "libredoc.pdf");
            Console.WriteLine("Show docResult: {0}", linkdoc);

            Console.WriteLine();
            Console.WriteLine("docx 轉為 pdf：");
            var docxResult = LibreOffice.OfficeConverter.ConvertDocument(docxPath, OutputExtensionType.pdf);
            var linkdocx = SaveFile(docxResult, "libredocx.pdf");
            Console.WriteLine("Show docxResult: {0}", linkdocx);

            Console.WriteLine();
            Console.WriteLine("xls 轉為 pdf：");
            var xlsResult = LibreOffice.OfficeConverter.ConvertDocument(xlsPath, OutputExtensionType.pdf);
            var linkxls = SaveFile(xlsResult, "librexls.pdf");
            Console.WriteLine("Show xlsResult: {0}", linkxls);

            Console.WriteLine();
            Console.WriteLine("xls 轉為 pdf：");
            var xlsxResult = LibreOffice.OfficeConverter.ConvertDocument(xlsxPath, OutputExtensionType.pdf);
            var linkxlsx = SaveFile(xlsxResult, "librexlsx.pdf");
            Console.WriteLine("Show xlsResult: {0}", linkxlsx);

            Console.WriteLine();
            Console.WriteLine("ppt 轉為 pdf：");
            var pptResult = LibreOffice.OfficeConverter.ConvertDocument(pptPath, OutputExtensionType.pdf);
            var linkppt = SaveFile(pptResult, "libreppt.pdf");
            Console.WriteLine("Show pptResult: {0}", linkppt);

            Console.WriteLine();
            Console.WriteLine("pptx 轉為 pdf：");
            var pptxResult = LibreOffice.OfficeConverter.ConvertDocument(pptxPath, OutputExtensionType.pdf);
            var linkpptx = SaveFile(pptxResult, "librepptx.pdf");
            Console.WriteLine("Show pptxResult: {0}", linkpptx);
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
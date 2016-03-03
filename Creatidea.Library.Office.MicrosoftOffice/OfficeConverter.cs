using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.MicrosoftOffice
{
    using System.Diagnostics;
    using System.IO;

    using Creatidea.Library.Results;

    /// <summary>
    /// Microsoft Office OfficeConverter.
    /// </summary>
    public class OfficeConverter
    {
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
            string outputFileName = string.Empty;
            string outputPath = string.Empty;

            if (File.Exists(inputFilePath))
            {
                outputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
            }
            else
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
    }
}

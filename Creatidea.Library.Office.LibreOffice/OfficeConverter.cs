namespace Creatidea.Library.Office.LibreOffice
{
    using System;
    using System.Diagnostics;
    using System.IO;

    using Creatidea.Library.Configs;
    using Creatidea.Library.Results;

    /// <summary>
    /// LibreOffice 轉檔器
    /// </summary>
    public class OfficeConverter
    {
        /// <summary>
        /// Words to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input.</param>
        /// <returns><see cref="CiResult{T}"/> Data為轉檔後之PDF路徑</returns>
        public static CiResult<string> WordToPdf(string inputFilePath)
        {
            var result = new CiResult<string>() { Message = "Word轉Pdf失敗！" };

            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".doc") && !ext.Equals(".docx")))
            {
                result.Message += "副檔名錯誤！，應為*.doc, *.docx";
                return result;
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = string.Empty;
            string outputPath = string.Empty;

            try
            {
                if (File.Exists(inputFilePath))
                {
                    outputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
                }

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.CreateNoWindow = false;
                startInfo.UseShellExecute = false;

                // 設定執行檔路徑
                startInfo.FileName = CiConfig.Global.CiLibreOffice.BinPath;
                startInfo.WorkingDirectory = outputDir;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Arguments = string.Format(" -headless -convert-to pdf {0}", inputFilePath);

                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                }

                outputPath = Path.Combine(outputDir, outputFileName + ".pdf");
            }
            catch (Exception ex)
            {
                result.Message += ex.ToString();
                return result;
            }

            result.Success = true;
            result.Message = "Word轉Pdf失敗成功";
            result.Data = outputPath;

            return result;
        }
    }
}

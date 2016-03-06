namespace Creatidea.Library.Office.LibreOffice
{
    using System;
    using System.Diagnostics;
    using System.IO;

    using Creatidea.Library.Configs;

    /// <summary>
    /// LibreOffice 轉檔器
    /// </summary>
    public class OfficeConverter
    {
        /// <summary>
        /// Words to PDF.
        /// </summary>
        /// <param name="inputFilePath">The input.</param>
        /// <returns>Data為轉檔後之PDF路徑</returns>
        [STAThread]
        public static string WordToPdf(string inputFilePath)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".doc") && !ext.Equals(".docx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.doc, *.docx", nameof(inputFilePath));
            }

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
            string outputPath = string.Empty;

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

            return outputPath;
        }

        [STAThread]
        public static string ExcelToPdf(string inputFilePath)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".xls") && !ext.Equals(".xlsx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.xls, *.xlsx", nameof(inputFilePath));
            }

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
            string outputPath = string.Empty;

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

            return outputPath;
        }

        [STAThread]
        public static string PptToPdf(string inputFilePath)
        {
            var ext = Path.GetExtension(inputFilePath);
            if (string.IsNullOrEmpty(ext) || (!ext.Equals(".ppt") && !ext.Equals(".pptx")))
            {
                throw new ArgumentException("副檔名錯誤！應為*.ppt, *.pptx", nameof(inputFilePath));
            }

            if (!File.Exists(inputFilePath))
            {
                throw new ArgumentException(string.Format("找不到檔案：{0}！", inputFilePath), nameof(inputFilePath));
            }

            string outputDir = Path.GetTempPath();
            string outputFileName = Path.GetFileNameWithoutExtension(inputFilePath);
            string outputPath = string.Empty;

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

            return outputPath;
        }
    }
}

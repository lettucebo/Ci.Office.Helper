using System;
using System.Diagnostics;
using System.IO;
using Creatidea.Library.Configs;
using Creatidea.Library.Office.LibreOffice.Enums;

namespace Creatidea.Library.Office.LibreOffice
{
    /// <summary>
    /// LibreOffice 轉檔器
    /// </summary>
    public class OfficeConverter
    {
        /// <summary>
        /// LibreOffice通用型轉檔功能
        /// </summary>
        /// <param name="inputFilePath">The input.</param>
        /// <returns>Data為轉檔後之PDF路徑</returns>
        [STAThread]
        public static string ConvertDocument(string inputFilePath, OutputExtensionType outputExtensionType)
        {
            // 副得副檔名對應的參數
            string extensionParameter = ConvertExtensionToArg(Path.GetExtension(inputFilePath), outputExtensionType);

            if (string.IsNullOrWhiteSpace(extensionParameter))
                throw new InvalidProgramException("Unknown file type for LibreOffice. File = " + inputFilePath);

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
            startInfo.Arguments = string.Format(" -headless -convert-to {1} {0}", inputFilePath, extensionParameter);

            using (Process exeProcess = Process.Start(startInfo))
            {
                exeProcess.WaitForExit();
            }
            string[] ext = extensionParameter.Split(':');

            outputPath = Path.Combine(outputDir, outputFileName + "." + ext[0]);

            return outputPath;
        }

        /// <summary>
        /// 依據副檔名轉換成相對應的參數
        /// </summary>
        /// <param name="inputExtension"></param>
        /// <param name="outputExtensionType"></param>
        /// <returns></returns>
        private static string ConvertExtensionToArg(string inputExtension, OutputExtensionType outputExtensionType)
        {
            switch (inputExtension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    if (outputExtensionType == OutputExtensionType.pdf)
                    {
                        return "pdf:writer_pdf_Export";
                    }
                    else if (outputExtensionType == OutputExtensionType.odt)
                    {
                        return "odt:writer8";
                    }
                    return null;
                case ".xls":
                case ".xlsx":
                case ".xlsb":
                case ".ods":
                    if (outputExtensionType == OutputExtensionType.pdf)
                    {
                        return "pdf:calc_pdf_Export";
                    }
                    else if (outputExtensionType == OutputExtensionType.ods)
                    {
                        return "ods:calc8";
                    }
                    return null;
                case ".ppt":
                case ".pptx":
                case ".odp":
                    if (outputExtensionType == OutputExtensionType.pdf)
                    {
                        return "pdf:impress_pdf_Export";
                    }
                    else if (outputExtensionType == OutputExtensionType.odp)
                    {
                        return "odp:impress8";
                    }
                    return null;
                default:
                    return null;
            }
        }
    }
}

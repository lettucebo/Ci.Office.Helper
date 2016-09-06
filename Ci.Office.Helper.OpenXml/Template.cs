namespace Ci.Office.Helper.OpenXml
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using Ci.Office.Helper.OpenXml.Manager;
    using Ci.Office.Helper.OpenXml.Models;

    using DocumentFormat.OpenXml.Wordprocessing;

    public class Template
    {
        /// <summary>
        /// 將字典檔丟附openxmlmanager
        /// </summary>
        /// <param name="templateFilePath">template file path</param>
        /// <param name="textDict">text value pair dicitionary</param>
        /// <param name="imageDict">image value pair dicitionary</param>
        /// <param name="tableDict">table value pair dicitionary</param>
        /// <returns>complete template docx file path</returns>
        public string DocxMaker(
            string templateFilePath,
            Dictionary<string, OpenXmlTextInfo> textDict = null,
            Dictionary<string, MemoryStream> imageDict = null,
            Dictionary<string, Table> tableDict = null)
        {
            string templateDocx = templateFilePath;
            string tempDocx = Path.GetTempPath() + Guid.NewGuid() + ".docx";

            // copy the word doc so you can see the difference between the two
            File.Delete(tempDocx);
            File.Copy(templateDocx, tempDocx);

            var wordManager = new WordManager();
            wordManager.OpenDocuemnt(tempDocx);

            if (textDict != null && textDict.Any())
            {
                wordManager.UpdateText(textDict);
            }

            if (tableDict != null && tableDict.Any())
            {
                wordManager.UpdateTable(tableDict);
            }

            if (imageDict != null && imageDict.Any())
            {
                wordManager.UpdateImage(imageDict);
            }

            wordManager.CloseDocument();

            return tempDocx;
        }
    }
}

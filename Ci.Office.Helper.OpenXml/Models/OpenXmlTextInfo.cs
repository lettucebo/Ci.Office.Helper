using DocumentFormat.OpenXml.Wordprocessing;

namespace Ci.Office.Helper.OpenXml.Models
{
    public class OpenXmlTextInfo
    {
        public string Text;

        /// <summary>
        /// Is replace all inner xml
        /// </summary>
        public bool IsInnerXml;

        public string Separator;

        public RunFonts RunFonts;

    }
}

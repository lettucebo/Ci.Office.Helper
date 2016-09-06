namespace Ci.Office.Helper.OpenXml.Enums
{
    using System.ComponentModel;

    public enum WordSymbols
    {
        /// <summary>
        /// 未勾選
        /// </summary>
        [Description("<w:sym w:char=\"00A3\" w:font=\"Wingdings 2\"/>")]
        UnChecked = 1,
        /// <summary>
        /// 己勾選
        /// </summary>
        [Description("<w:sym w:char=\"0052\" w:font=\"Wingdings 2\"/>")]
        Checked = 2,

        /// <summary>
        /// 段落符號
        /// </summary>
        [Description("<w:br />")]
        Paragraph = 3
    }
}

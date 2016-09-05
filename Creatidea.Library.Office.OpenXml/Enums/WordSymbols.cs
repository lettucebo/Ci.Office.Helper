using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Creatidea.Library.Office.OpenXml.Enums
{
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;

    public enum WordSymbols
    {
        /// <summary>
        /// 未勾選
        /// </summary>
        [Display(Name = "<w:sym w:char=\"00A3\" w:font=\"Wingdings 2\"/>")]
        UnChecked = 1,
        /// <summary>
        /// 己勾選
        /// </summary>
        [Display(Name = "<w:sym w:char=\"0052\" w:font=\"Wingdings 2\"/>")]
        Checked = 2,

        /// <summary>
        /// 段落符號
        /// </summary>
        [Display(Name = "<w:br />")]
        Paragraph = 3
    }
}

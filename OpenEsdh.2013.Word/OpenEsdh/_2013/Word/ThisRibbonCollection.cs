namespace OpenEsdh._2013.Word
{
    using Microsoft.Office.Tools.Ribbon;
    using System;
    using System.CodeDom.Compiler;
    using System.Diagnostics;

    internal sealed partial class ThisRibbonCollection : RibbonCollectionBase
    {
        internal OpenEsdh._2013.Word.OpenESDHRibbon OpenESDHRibbon
        {
            get
            {
                return base.GetRibbon<OpenEsdh._2013.Word.OpenESDHRibbon>();
            }
        }
    }
}


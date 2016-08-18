namespace OpenEsdh._2013.Excel
{
    using Microsoft.Office.Tools.Ribbon;
    using System;
    using System.CodeDom.Compiler;
    using System.Diagnostics;

    [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
    internal sealed class ThisRibbonCollection : RibbonCollectionBase
    {
        internal ThisRibbonCollection(Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : base((Microsoft.Office.Tools.Ribbon.RibbonFactory) factory)
        {
        }

        internal OpenEsdh._2013.Excel.OpenESDHRibbon OpenESDHRibbon
        {
            get
            {
                return base.GetRibbon<OpenEsdh._2013.Excel.OpenESDHRibbon>();
            }
        }
    }
}


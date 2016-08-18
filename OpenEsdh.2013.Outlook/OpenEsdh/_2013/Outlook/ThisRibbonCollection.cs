namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Ribbon;
    using System;
    using System.CodeDom.Compiler;
    using System.Diagnostics;
    using System.Reflection;

    [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
    internal sealed class ThisRibbonCollection : RibbonCollectionBase
    {
        internal ThisRibbonCollection(Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : base((Microsoft.Office.Tools.Ribbon.RibbonFactory) factory)
        {
        }

        internal ThisRibbonCollection this[Microsoft.Office.Interop.Outlook.Inspector inspector]
        {
            get
            {
                return base.GetRibbonContextCollection<ThisRibbonCollection>(inspector);
            }
        }

        internal ThisRibbonCollection this[Microsoft.Office.Interop.Outlook.Explorer explorer]
        {
            get
            {
                return base.GetRibbonContextCollection<ThisRibbonCollection>(explorer);
            }
        }

        internal OpenEsdh._2013.Outlook.OpenESDHRibbon OpenESDHRibbon
        {
            get
            {
                return base.GetRibbon<OpenEsdh._2013.Outlook.OpenESDHRibbon>();
            }
        }
    }
}


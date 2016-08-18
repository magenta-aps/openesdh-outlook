namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Outlook;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;

    [DebuggerNonUserCode]
    internal sealed class ThisFormRegionCollection : FormRegionCollectionBase
    {
        public ThisFormRegionCollection(IList<Microsoft.Office.Tools.Outlook.IFormRegion> list) : base((IList<Microsoft.Office.Tools.Outlook.IFormRegion>) list)
        {
        }

        internal WindowFormRegionCollection this[Microsoft.Office.Interop.Outlook.Explorer explorer]
        {
            get
            {
                return (WindowFormRegionCollection) Globals.ThisAddIn.GetFormRegions((Microsoft.Office.Interop.Outlook.Explorer) explorer, typeof(WindowFormRegionCollection));
            }
        }

        internal WindowFormRegionCollection this[Microsoft.Office.Interop.Outlook.Inspector inspector]
        {
            get
            {
                return (WindowFormRegionCollection) Globals.ThisAddIn.GetFormRegions((Microsoft.Office.Interop.Outlook.Inspector) inspector, typeof(WindowFormRegionCollection));
            }
        }
    }
}


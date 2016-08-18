namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Tools.Outlook;
    using System;
    using System.CodeDom.Compiler;
    using System.Collections.Generic;
    using System.Diagnostics;

    [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
    internal sealed class Globals
    {
        private static Microsoft.Office.Tools.Outlook.Factory _factory;
        private static OpenEsdh._2013.Outlook.ThisAddIn _ThisAddIn;
        private static ThisFormRegionCollection _ThisFormRegionCollection;
        private static ThisRibbonCollection _ThisRibbonCollection;

        private Globals()
        {
        }

        internal static Microsoft.Office.Tools.Outlook.Factory Factory
        {
            get
            {
                return _factory;
            }
            set
            {
                if (_factory != null)
                {
                    throw new NotSupportedException();
                }
                _factory = value;
            }
        }

        internal static ThisFormRegionCollection FormRegions
        {
            get
            {
                if (_ThisFormRegionCollection == null)
                {
                    _ThisFormRegionCollection = new ThisFormRegionCollection((IList<Microsoft.Office.Tools.Outlook.IFormRegion>) ThisAddIn.GetFormRegions());
                }
                return _ThisFormRegionCollection;
            }
        }

        internal static ThisRibbonCollection Ribbons
        {
            get
            {
                if (_ThisRibbonCollection == null)
                {
                    _ThisRibbonCollection = new ThisRibbonCollection(_factory.GetRibbonFactory());
                }
                return _ThisRibbonCollection;
            }
        }

        internal static OpenEsdh._2013.Outlook.ThisAddIn ThisAddIn
        {
            get
            {
                return _ThisAddIn;
            }
            set
            {
                if (_ThisAddIn != null)
                {
                    throw new NotSupportedException();
                }
                _ThisAddIn = value;
            }
        }
    }
}


namespace OpenEsdh._2013.Excel
{
    using Microsoft.Office.Tools.Excel;
    using System;
    using System.CodeDom.Compiler;
    using System.Diagnostics;

    [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
    internal sealed class Globals
    {
        private static ApplicationFactory _factory;
        private static OpenEsdh._2013.Excel.ThisAddIn _ThisAddIn;
        private static ThisRibbonCollection _ThisRibbonCollection;

        private Globals()
        {
        }

        internal static ApplicationFactory Factory
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

        internal static OpenEsdh._2013.Excel.ThisAddIn ThisAddIn
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


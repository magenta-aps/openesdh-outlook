namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Tools.Outlook;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;

    [DebuggerNonUserCode]
    internal sealed class WindowFormRegionCollection : FormRegionCollectionBase
    {
        public WindowFormRegionCollection(IList<Microsoft.Office.Tools.Outlook.IFormRegion> list) : base((IList<Microsoft.Office.Tools.Outlook.IFormRegion>) list)
        {
        }

        internal OpenEsdh._2013.Outlook.OpenESDHIcon OpenESDHIcon
        {
            get
            {
                foreach (Microsoft.Office.Tools.Outlook.IFormRegion region in this)
                {
                    if (region.GetType() == typeof(OpenEsdh._2013.Outlook.OpenESDHIcon))
                    {
                        return (OpenEsdh._2013.Outlook.OpenESDHIcon) region;
                    }
                }
                return null;
            }
        }

        internal OpenEsdh._2013.Outlook.OpenESDHRegion OpenESDHRegion
        {
            get
            {
                foreach (Microsoft.Office.Tools.Outlook.IFormRegion region in this)
                {
                    if (region.GetType() == typeof(OpenEsdh._2013.Outlook.OpenESDHRegion))
                    {
                        return (OpenEsdh._2013.Outlook.OpenESDHRegion) region;
                    }
                }
                return null;
            }
        }
    }
}


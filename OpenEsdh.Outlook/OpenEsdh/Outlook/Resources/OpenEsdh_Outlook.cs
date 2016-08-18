namespace OpenEsdh.Outlook.Resources
{
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Globalization;
    using System.Resources;
    using System.Runtime.CompilerServices;

    [DebuggerNonUserCode, CompilerGenerated, GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    internal class OpenEsdh_Outlook
    {
        private static CultureInfo resourceCulture;
        private static System.Resources.ResourceManager resourceMan;

        internal OpenEsdh_Outlook()
        {
        }

        internal static string ApplicationSave
        {
            get
            {
                return ResourceManager.GetString("ApplicationSave", resourceCulture);
            }
        }

        internal static string ApplicationSaveAs
        {
            get
            {
                return ResourceManager.GetString("ApplicationSaveAs", resourceCulture);
            }
        }

        internal static string AttachFileBtn
        {
            get
            {
                return ResourceManager.GetString("AttachFileBtn", resourceCulture);
            }
        }

        internal static string AttachFileDialogTitle
        {
            get
            {
                return ResourceManager.GetString("AttachFileDialogTitle", resourceCulture);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        internal static CultureInfo Culture
        {
            get
            {
                return resourceCulture;
            }
            set
            {
                resourceCulture = value;
            }
        }

        internal static string OpenESDHAppGroup
        {
            get
            {
                return ResourceManager.GetString("OpenESDHAppGroup", resourceCulture);
            }
        }

        internal static string OpenESDHButtonGroup
        {
            get
            {
                return ResourceManager.GetString("OpenESDHButtonGroup", resourceCulture);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        internal static System.Resources.ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    System.Resources.ResourceManager manager = new System.Resources.ResourceManager("OpenEsdh.Outlook.Resources.OpenEsdh.Outlook", typeof(OpenEsdh_Outlook).Assembly);
                    resourceMan = manager;
                }
                return resourceMan;
            }
        }

        internal static string SaveAsDialogSaveAndSend
        {
            get
            {
                return ResourceManager.GetString("SaveAsDialogSaveAndSend", resourceCulture);
            }
        }

        internal static string SaveAsDialogTitle
        {
            get
            {
                return ResourceManager.GetString("SaveAsDialogTitle", resourceCulture);
            }
        }

        internal static string SaveBtn
        {
            get
            {
                return ResourceManager.GetString("SaveBtn", resourceCulture);
            }
        }

        internal static string SaveSendBtn
        {
            get
            {
                return ResourceManager.GetString("SaveSendBtn", resourceCulture);
            }
        }

        internal static string ViewRegionTitle
        {
            get
            {
                return ResourceManager.GetString("ViewRegionTitle", resourceCulture);
            }
        }
    }
}


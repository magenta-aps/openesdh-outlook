namespace OpenEsdh._2013.Excel.Properties
{
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.Resources;
    using System.Runtime.CompilerServices;

    [GeneratedCode("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"), DebuggerNonUserCode, CompilerGenerated]
    internal class Resources
    {
        private static CultureInfo resourceCulture;
        private static System.Resources.ResourceManager resourceMan;

        internal Resources()
        {
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

        internal static Bitmap openesdh_16x16
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("openesdh_16x16", resourceCulture);
            }
        }

        internal static Bitmap openesdh_32x32
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("openesdh_32x32", resourceCulture);
            }
        }

        internal static Bitmap openesdh_logo_16
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("openesdh-logo-16", resourceCulture);
            }
        }

        internal static Bitmap openesdh_logo_32
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("openesdh-logo-32", resourceCulture);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Advanced)]
        internal static System.Resources.ResourceManager ResourceManager
        {
            get
            {
                if (object.ReferenceEquals(resourceMan, null))
                {
                    System.Resources.ResourceManager manager = new System.Resources.ResourceManager("OpenEsdh._2013.Excel.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = manager;
                }
                return resourceMan;
            }
        }

        internal static Bitmap VismaCase32x32
        {
            get
            {
                return (Bitmap) ResourceManager.GetObject("VismaCase32x32", resourceCulture);
            }
        }
    }
}


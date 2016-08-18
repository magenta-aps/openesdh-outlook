namespace OpenEsdh.Outlook.Model
{
    using Microsoft.Win32;
    using System;
    using System.IO;
    using System.Runtime.CompilerServices;

    public class Attachment
    {
        public Attachment()
        {
            this.Name = "";
            this.MimeType = "";
            this.ForceImport = false;
        }

        public Attachment(string name, string mimeType) : this()
        {
            this.Name = name;
            this.MimeType = mimeType;
            this.ForceImport = false;
        }

        public Attachment(string name, string mimeType, bool forceImport) : this()
        {
            this.Name = name;
            this.MimeType = mimeType;
            this.ForceImport = forceImport;
        }

        public static string GetMimeType(string fileName)
        {
            string str = "application/unknown";
            string name = Path.GetExtension(fileName).ToLower();
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(name);
            if ((key != null) && (key.GetValue("Content Type") != null))
            {
                str = key.GetValue("Content Type").ToString();
            }
            return str;
        }

        public bool ForceImport { get; set; }

        public string MimeType { get; set; }

        public string Name { get; set; }
    }
}


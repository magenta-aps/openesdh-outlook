namespace OpenEsdh.Outlook.Model.Alfresco
{
    using System;
    using System.Runtime.CompilerServices;

    public class DocumentIdPackage
    {
        public DocumentIdPackage()
        {
            this.workspace = "";
            this.spacestore = "";
            this.id = "";
        }

        public DocumentIdPackage(string nodeRef) : this()
        {
            if (!string.IsNullOrEmpty(nodeRef))
            {
                string str = nodeRef;
                if (str.Contains("://"))
                {
                    this.workspace = str.Substring(0, str.IndexOf("://"));
                    str = str.Replace(this.workspace + "://", "");
                }
                if (str.Contains("/"))
                {
                    string[] strArray = str.Split(new char[] { '/' });
                    this.spacestore = strArray[0];
                    this.id = strArray[1];
                }
            }
        }

        public string id { get; set; }

        public string spacestore { get; set; }

        public string workspace { get; set; }
    }
}


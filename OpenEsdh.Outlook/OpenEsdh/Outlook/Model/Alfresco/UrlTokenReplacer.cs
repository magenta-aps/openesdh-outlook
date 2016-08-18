namespace OpenEsdh.Outlook.Model.Alfresco
{
    using System;
    using System.Web;
    using System.Web.Script.Serialization;

    public class UrlTokenReplacer
    {
        private ApplicationPackage _package;
        private string _url;

        public UrlTokenReplacer(string url, ApplicationPackage package)
        {
            this._package = null;
            this._url = "";
            this._url = url;
            this._package = package;
        }

        public UrlTokenReplacer(string url, string package)
        {
            this._package = null;
            this._url = "";
            this._url = url;
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            try
            {
                ApplicationPackage package2 = serializer.Deserialize<ApplicationPackage>(package);
                if (package2 != null)
                {
                    this._package = package2;
                }
            }
            catch
            {
                this._package = null;
            }
        }

        public string OriginalUrl
        {
            get
            {
                return this._url;
            }
        }

        public string Url
        {
            get
            {
                DocumentIdPackage package;
                string str = this._url;
                if (this._package != null)
                {
                    str = !string.IsNullOrEmpty(this._package.caseId) ? str.Replace("[@CaseId]", HttpUtility.UrlEncode(this._package.caseId)) : str;
                    str = !string.IsNullOrEmpty(this._package.documentName) ? str.Replace("[@Name]", HttpUtility.UrlEncode(this._package.documentName)) : str;
                    str = !string.IsNullOrEmpty(this._package.ticket) ? str.Replace("[@Ticket]", HttpUtility.UrlEncode(this._package.ticket)) : str;
                    str = !string.IsNullOrEmpty(this._package.nodeRef) ? str.Replace("[@NodeRef]", HttpUtility.UrlEncode(this._package.nodeRef)) : str;
                    str = str.Replace("[@Ticks]", DateTime.Now.Ticks.ToString());
                }
                if (!string.IsNullOrEmpty(this._package.nodeRef))
                {
                    package = new DocumentIdPackage(this._package.nodeRef);
                    str = !string.IsNullOrEmpty(package.workspace) ? str.Replace("[@WorkSpace]", HttpUtility.UrlEncode(package.workspace)) : str;
                    str = !string.IsNullOrEmpty(package.spacestore) ? str.Replace("[@SpaceStore]", HttpUtility.UrlEncode(package.spacestore)) : str;
                    return (!string.IsNullOrEmpty(package.id) ? str.Replace("[@Id]", HttpUtility.UrlEncode(package.id)) : str);
                }
                if (!string.IsNullOrEmpty(this._package.docType))
                {
                    package = new DocumentIdPackage(this._package.docType);
                    str = !string.IsNullOrEmpty(package.workspace) ? str.Replace("[@WorkSpace]", HttpUtility.UrlEncode(package.workspace)) : str;
                    str = !string.IsNullOrEmpty(package.spacestore) ? str.Replace("[@SpaceStore]", HttpUtility.UrlEncode(package.spacestore)) : str;
                    str = !string.IsNullOrEmpty(package.id) ? str.Replace("[@Id]", HttpUtility.UrlEncode(package.id)) : str;
                }
                return str;
            }
        }
    }
}


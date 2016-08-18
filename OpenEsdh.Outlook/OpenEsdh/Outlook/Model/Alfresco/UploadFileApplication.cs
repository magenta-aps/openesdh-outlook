namespace OpenEsdh.Outlook.Model.Alfresco
{
    using System;
    using System.Web.Script.Serialization;

    public class UploadFileApplication : AlfrescoFilePost
    {
        private ApplicationPackage _package;
        private bool _sendPackage;

        public UploadFileApplication(bool sendPackage)
        {
            this._sendPackage = false;
            this._package = null;
            this._sendPackage = sendPackage;
        }

        public UploadFileApplication(string url, bool sendPackage) : base(url)
        {
            this._sendPackage = false;
            this._package = null;
            this._sendPackage = sendPackage;
        }

        public override string UploadFile(string unknownJson, string fileName, string Name)
        {
            try
            {
                ApplicationPackage package = new JavaScriptSerializer().Deserialize<ApplicationPackage>(unknownJson);
                if (package != null)
                {
                    this._package = package;
                }
            }
            catch
            {
            }
            string str = "";
            if (this._sendPackage)
            {
                str = base.UploadFile(unknownJson, fileName, Name);
            }
            else
            {
                str = base.UploadFile("", fileName, Name);
            }
            this._package = null;
            return str;
        }

        protected override string Url
        {
            get
            {
                UrlTokenReplacer replacer = new UrlTokenReplacer(base.Url, this._package);
                return replacer.Url;
            }
        }
    }
}


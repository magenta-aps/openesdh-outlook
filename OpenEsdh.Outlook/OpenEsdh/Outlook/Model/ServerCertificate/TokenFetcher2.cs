namespace OpenEsdh.Outlook.Model.ServerCertificate
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using System;
    using System.Collections.Specialized;
    using System.IO;
    using System.Net;
    using System.Text;

    public class TokenFetcher2
    {
        private NameValueCollection _additionalParameters = null;

        private void DoWebRequest(WebRequest webRequest)
        {
            webRequest.Credentials = CredentialCache.DefaultNetworkCredentials;
            webRequest.PreAuthenticate = false;
            WebResponse response = webRequest.GetResponse();
            using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                string str = reader.ReadToEnd();
                foreach (string str2 in response.Headers.AllKeys)
                {
                    switch (str2.ToLower())
                    {
                        case "authorization":
                            this._additionalParameters.Add("Authorization", response.Headers[str]);
                            break;

                        case "set-cookie":
                            this._additionalParameters.Add("Cookie", response.Headers[str]);
                            break;
                    }
                }
            }
        }

        private void GenerateRequest()
        {
            this._additionalParameters = new NameValueCollection();
            IOutlookConfiguration configuration = TypeResolver.Current.Create<IOutlookConfiguration>();
            WebRequest webRequest = WebRequest.Create(configuration.SaveAsDialogUrl);
            if (configuration.PreAuthentication.UseConfigCredentials)
            {
                using (new ImpersonationContext(configuration.PreAuthentication.Username, configuration.PreAuthentication.Password, configuration.PreAuthentication.Domain))
                {
                    string str = new WebClient { UseDefaultCredentials = true }.DownloadString(configuration.SaveAsDialogUrl);
                    this.DoWebRequest(webRequest);
                }
            }
            else
            {
                this.DoWebRequest(webRequest);
            }
        }

        public string GetParameters()
        {
            if (this._additionalParameters == null)
            {
                this.GenerateRequest();
            }
            string str = "";
            string str2 = "";
            for (int i = 0; i < this._additionalParameters.Count; i++)
            {
                string str4 = str2;
                str2 = str4 + str + this._additionalParameters.GetKey(i) + "=" + this._additionalParameters[this._additionalParameters.GetKey(i)];
                str = "&";
            }
            return str2;
        }

        public NameValueCollection AdditionalParameters
        {
            get
            {
                if (this._additionalParameters == null)
                {
                    this.GenerateRequest();
                }
                return this._additionalParameters;
            }
        }
    }
}


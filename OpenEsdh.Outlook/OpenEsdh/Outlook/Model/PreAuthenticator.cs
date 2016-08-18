namespace OpenEsdh.Outlook.Model
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using System;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.IO;
    using System.Net;
    using System.Runtime.CompilerServices;
    using System.Text;

    public class PreAuthenticator : IPreAuthenticator
    {
        private NameValueCollection _additionalParameters = null;
        private IPreAuthenticateConfiguration _configuration;

        public PreAuthenticator(IPreAuthenticateConfiguration configuration)
        {
            this._configuration = configuration;
        }

        public string AddAuthentificationUrl(string url)
        {
            this.ClearAuthentication();
            string str = url;
            if (this.AdditionalParameters.Count >= 0)
            {
                string str2 = "?";
                if (str.Contains("?"))
                {
                    str2 = "&";
                }
                foreach (string str3 in this.AdditionalParameters)
                {
                    if (str.Contains("[@" + str3 + "]"))
                    {
                        str = str.Replace("[@" + str3 + "]", this.AdditionalParameters[str3]);
                    }
                    else
                    {
                        if (str.Contains(str3 + "="))
                        {
                            Debug.WriteLine("ticket already exists" + str);
                        }
                        string str5 = str;
                        str = str5 + str2 + str3 + "=" + this.AdditionalParameters[str3];
                        str2 = "&";
                    }
                }
            }
            return str;
        }

        private void ClearAuthentication()
        {
            if (this._configuration.ReauthenticateOnEachRequest)
            {
                this._additionalParameters = null;
            }
        }

        public string GetParameters()
        {
            if (this._additionalParameters == null)
            {
                this.PreAuthenticate();
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

        private void PreAuthenticate()
        {
            try
            {
                WebRequest request;
                string str4;
                string str5;
                WebResponse response;
                StreamReader reader;
                string[] strArray;
                ICookieJar jar;
                this._additionalParameters = new NameValueCollection();
                AttachEmail.TimeOutWebClient client = new AttachEmail.TimeOutWebClient {
                    TimeOut = 0x7d0
                };
                string newValue = "";
                string password = "";
                if (this._configuration.UseConfigCredentials)
                {
                    newValue = this._configuration.Username;
                    password = this._configuration.Password;
                    string s = this._configuration.AuthenticationPackageFormat.Replace("[@username]", newValue).Replace("[@password]", password);
                    request = WebRequest.Create(this._configuration.AuthenticationUrl);
                    request.Method = "POST";
                    byte[] bytes = Encoding.ASCII.GetBytes(s);
                    request.ContentLength = bytes.Length;
                    using (Stream stream = request.GetRequestStream())
                    {
                        stream.Write(bytes, 0, bytes.Length);
                        stream.Close();
                    }
                    str4 = "";
                    str5 = "";
                    using (response = request.GetResponse())
                    {
                        using (reader = new StreamReader(response.GetResponseStream()))
                        {
                            str4 = reader.ReadToEnd();
                            if (!string.IsNullOrEmpty(response.Headers["Set-Cookie"]))
                            {
                                str5 = response.Headers["Set-Cookie"];
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(str5) && str5.Contains(";"))
                    {
                        str5 = str5.Split(new char[] { ';' })[0];
                        jar = TypeResolver.Current.Create<ICookieJar>();
                        strArray = str5.Split(new char[] { '=' });
                        if (!jar.Cookies.ContainsKey(strArray[0]))
                        {
                            jar.Cookies.Add(strArray[0], strArray[1]);
                        }
                    }
                }
                else
                {
                    Uri uri = new Uri(this._configuration.AuthenticationUrl);
                    request = WebRequest.Create(uri.Scheme + "://" + uri.DnsSafeHost);
                    request.Method = "GET";
                    request.ContentLength = 0L;
                    str4 = "";
                    str5 = "";
                    using (response = request.GetResponse())
                    {
                        using (reader = new StreamReader(response.GetResponseStream()))
                        {
                            str4 = reader.ReadToEnd();
                            if (!string.IsNullOrEmpty(response.Headers["Set-Cookie"]))
                            {
                                str5 = response.Headers["Set-Cookie"];
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(str5) && str5.Contains(";"))
                    {
                        str5 = str5.Split(new char[] { ';' })[0];
                        jar = TypeResolver.Current.Create<ICookieJar>();
                        strArray = str5.Split(new char[] { '=' });
                        if (!jar.Cookies.ContainsKey(strArray[0]))
                        {
                            jar.Cookies.Add(strArray[0], strArray[1]);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        public NameValueCollection AdditionalParameters
        {
            get
            {
                if (this._additionalParameters == null)
                {
                    this.PreAuthenticate();
                }
                return this._additionalParameters;
            }
        }

        public class Data
        {
            public PreAuthenticator.DataTicket data { get; set; }
        }

        public class DataTicket
        {
            public string ticket { get; set; }
        }
    }
}


namespace OpenEsdh.Outlook.Model
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using System;
    using System.IO;
    using System.Linq;
    using System.Net;

    public class AttachEmail : IAttachEmail
    {
        public AddFileDelegate AddFile;
        public AddMailPropertyDelegate AddProperty;
        public SetMailPropertyDelegate SetProperty;

        public void AddAttachmentConfiguration(string[] ConfigurationSettings, SetMailPropertyDelegate setProperty, AddMailPropertyDelegate addProperty, AddFileDelegate addFile)
        {
            this.SetProperty = setProperty;
            this.AddProperty = addProperty;
            this.AddFile = addFile;
            IWordConfiguration configuration = TypeResolver.Current.Create<IWordConfiguration>();
            foreach (string str in ConfigurationSettings)
            {
                if (string.IsNullOrEmpty(str.Trim()))
                {
                    continue;
                }
                bool flag = false;
                if (str.Contains<char>(':') && !str.Contains(@":\"))
                {
                    string str2 = str.Substring(0, str.IndexOf(':'));
                    string str3 = str.Substring(str.IndexOf(':') + 1);
                    switch (str2.ToLower())
                    {
                        case "to":
                            this.SetPropertyInternal("TO,", str3);
                            flag = true;
                            break;

                        case "cc":
                            this.SetPropertyInternal("CC,", str3);
                            flag = true;
                            break;

                        case "bcc":
                            this.SetPropertyInternal("BCC,", str3);
                            flag = true;
                            break;

                        case "subject":
                            this.SetPropertyInternal("SUBJECT,", str3);
                            flag = true;
                            break;

                        case "htmlbody":
                            flag = true;
                            this.SetPropertyInternal("HTMLBODY", str3);
                            break;

                        case "htmlbody+":
                            flag = true;
                            this.AddPropertyInternal("HTMLBODY", str3);
                            break;

                        case "body":
                            flag = true;
                            this.SetPropertyInternal("BODY", str3);
                            break;

                        case "body+":
                            flag = true;
                            this.AddPropertyInternal("BODY", str3);
                            break;
                    }
                }
                else
                {
                    flag = false;
                }
                if (!flag)
                {
                    Exception exception;
                    if (System.IO.File.Exists(str))
                    {
                        try
                        {
                            this.AddFile(str);
                        }
                        catch (Exception exception1)
                        {
                            exception = exception1;
                            Logger.Current.LogException(exception, "");
                        }
                    }
                    else
                    {
                        try
                        {
                            string[] strArray = str.Replace("://", "/").Split(new char[] { '/' });
                            if (strArray.Length == 3)
                            {
                                string getFileEndPoint = configuration.GetFileEndPoint;
                                Uri uri = new Uri(getFileEndPoint);
                                getFileEndPoint = getFileEndPoint.Replace("{@store-protocol}", strArray[0]).Replace("{@store-identifier}", strArray[1]);
                                string[] strArray2 = strArray[2].Split(new char[] { '(' });
                                getFileEndPoint = getFileEndPoint.Replace("{@node-identifier}", strArray2[0]);
                                string tempPath = Path.GetTempPath();
                                string str7 = strArray2[1].Replace(")", "");
                                CookieWebClient client = new CookieWebClient();
                                if (configuration.PreAuthentication.UseConfigCredentials)
                                {
                                    client.Credentials = new NetworkCredential(configuration.PreAuthentication.Username, configuration.PreAuthentication.Password);
                                }
                                else
                                {
                                    client.Credentials = CredentialCache.DefaultNetworkCredentials;
                                }
                                ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
                                if (jar.Cookies.Count > 0)
                                {
                                    CookieContainer cookies = new CookieContainer();
                                    foreach (string str8 in jar.Cookies.Keys)
                                    {
                                        cookies.Add(new Cookie(str8, jar.Cookies[str8], "/", uri.DnsSafeHost));
                                    }
                                    client.SetCookies(cookies);
                                }
                                client.BaseAddress = getFileEndPoint;
                                int num = 0;
                                str7 = Path.Combine(tempPath, str7);
                                string path = str7;
                                while (System.IO.File.Exists(str7))
                                {
                                    num++;
                                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);
                                    fileNameWithoutExtension = fileNameWithoutExtension + "(" + num.ToString() + ")" + Path.GetExtension(path);
                                    str7 = Path.Combine(tempPath, fileNameWithoutExtension);
                                }
                                try
                                {
                                    client.DownloadFile(getFileEndPoint, str7);
                                    if (System.IO.File.Exists(str7))
                                    {
                                        addFile(str7);
                                    }
                                }
                                catch (Exception exception2)
                                {
                                    exception = exception2;
                                    Logger.Current.LogException(exception, "");
                                }
                            }
                        }
                        catch (Exception exception3)
                        {
                            exception = exception3;
                            Logger.Current.LogException(exception, "");
                        }
                    }
                }
            }
        }

        private void AddPropertyInternal(string name, string value)
        {
            if (this.SetProperty != null)
            {
                this.AddProperty(name, value);
            }
        }

        private void AttachFileInternal(string filename)
        {
            if (this.AddFile != null)
            {
                this.AddFile(filename);
            }
        }

        private void SetPropertyInternal(string name, string value)
        {
            if (this.SetProperty != null)
            {
                this.SetProperty(name, value);
            }
        }

        public class CookieWebClient : AttachEmail.TimeOutWebClient
        {
            private CookieContainer _cookies = new CookieContainer();

            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest webRequest = base.GetWebRequest(address);
                HttpWebRequest request2 = webRequest as HttpWebRequest;
                if (request2 != null)
                {
                    request2.CookieContainer = this._cookies;
                }
                return webRequest;
            }

            public void SetCookies(CookieContainer cookies)
            {
                this._cookies = cookies;
            }
        }

        public class TimeOutWebClient : WebClient
        {
            private int _timeOut = 0x7d0;

            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest webRequest = base.GetWebRequest(address);
                HttpWebRequest request2 = webRequest as HttpWebRequest;
                if (request2 != null)
                {
                    request2.Timeout = this.TimeOut;
                }
                return webRequest;
            }

            public int TimeOut
            {
                get
                {
                    return this._timeOut;
                }
                set
                {
                    this._timeOut = value;
                }
            }
        }
    }
}


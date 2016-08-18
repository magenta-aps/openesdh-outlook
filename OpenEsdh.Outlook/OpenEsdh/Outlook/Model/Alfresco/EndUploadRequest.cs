namespace OpenEsdh.Outlook.Model.Alfresco
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using System;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Security;
    using System.Text;

    public class EndUploadRequest
    {
        private IConfiguration _configuration = null;

        public EndUploadRequest(IConfiguration configuration)
        {
            this._configuration = configuration;
        }

        public void EndUpload(string UnknownJson)
        {
            if (!string.IsNullOrEmpty(this._configuration.EndUploadEndpoint))
            {
                Exception exception;
                UrlTokenReplacer replacer = new UrlTokenReplacer(this._configuration.EndUploadEndpoint, UnknownJson);
                string url = replacer.Url;
                if (this._configuration.PreAuthenticate)
                {
                    url = TypeResolver.Current.Create<IPreAuthenticator>().AddAuthentificationUrl(url);
                }
                ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.AllowAutoRedirect = true;
                request.AuthenticationLevel = AuthenticationLevel.MutualAuthRequested;
                request.PreAuthenticate = true;
                request.UseDefaultCredentials = true;
                request.ContentType = "application/json;charset=UTF-8";
                if (!this._configuration.PreAuthentication.UseConfigCredentials)
                {
                    request.Credentials = CredentialCache.DefaultNetworkCredentials;
                }
                try
                {
                    CookieContainer uriCookieContainerEx = new CookieJar().GetUriCookieContainerEx(new Uri(url));
                    if ((uriCookieContainerEx != null) && (uriCookieContainerEx.Count > 0))
                    {
                        foreach (object obj2 in uriCookieContainerEx.GetCookies(new Uri(url)))
                        {
                            string[] strArray = obj2.ToString().Split(new char[] { '=' });
                            if (strArray.Length > 0)
                            {
                                string s = strArray[0];
                                if ((from c in jar.Cookies.Keys
                                    where c.Contains(s)
                                    select c).FirstOrDefault<string>() == null)
                                {
                                    jar.Add(obj2.ToString());
                                }
                            }
                        }
                    }
                    request.CookieContainer = new CookieContainer();
                    if (jar.Cookies.Count > 0)
                    {
                        foreach (string str3 in jar.Cookies.Keys)
                        {
                            request.CookieContainer.Add(new Cookie(str3, jar.Cookies[str3], "/", request.RequestUri.DnsSafeHost));
                        }
                    }
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                }
                request.Method = "POST";
                using (Stream stream = request.GetRequestStream())
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(this._configuration.EndUploadPackage + Environment.NewLine);
                    stream.Write(bytes, 0, bytes.Length);
                }
                try
                {
                    using (WebResponse response = request.GetResponse())
                    {
                        using (Stream stream2 = response.GetResponseStream())
                        {
                            using (MemoryStream stream3 = new MemoryStream())
                            {
                                stream2.CopyTo(stream3);
                                string str4 = Encoding.ASCII.GetString(stream3.ToArray());
                            }
                        }
                    }
                }
                catch (Exception exception2)
                {
                    exception = exception2;
                    Logger.Current.LogException(exception, "");
                }
            }
        }
    }
}


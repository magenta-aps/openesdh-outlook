namespace OpenEsdh.Outlook.Model.Alfresco
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Security;
    using System.Text;

    public class AlfrescoFilePost : IAlfrescoFilePost
    {
        protected string _url;

        public AlfrescoFilePost()
        {
            this._url = "";
        }

        public AlfrescoFilePost(string Url)
        {
            this._url = "";
            this._url = Url;
        }

        public virtual string UploadFile(string unknownJson, string fileName, string Name)
        {
            return this.UploadFile(this.Url, unknownJson, fileName, Name);
        }

        public string UploadFile(string address, string unknownJson, string fileName, string name)
        {
            byte[] bytes = null;
            using (FileStream stream = System.IO.File.Open(fileName, FileMode.Open))
            {
                NameValueCollection values = new NameValueCollection();
                if (!string.IsNullOrEmpty(unknownJson))
                {
                    values.Add("metadata", unknownJson);
                }
                OpenEsdh.Outlook.Model.Alfresco.UploadFile[] fileArray2 = new OpenEsdh.Outlook.Model.Alfresco.UploadFile[1];
                OpenEsdh.Outlook.Model.Alfresco.UploadFile file = new OpenEsdh.Outlook.Model.Alfresco.UploadFile {
                    Name = name,
                    Filename = Path.GetFileName(fileName),
                    ContentType = Attachment.GetMimeType(name),
                    Stream = stream
                };
                fileArray2[0] = file;
                OpenEsdh.Outlook.Model.Alfresco.UploadFile[] files = fileArray2;
                bytes = this.UploadFiles(address, files, values);
                stream.Close();
                stream.Dispose();
            }
            if (bytes != null)
            {
                return Encoding.ASCII.GetString(bytes);
            }
            return "";
        }

        private byte[] UploadFiles(string address, IEnumerable<OpenEsdh.Outlook.Model.Alfresco.UploadFile> files, NameValueCollection values)
        {
            byte[] buffer3;
            ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
            HttpWebRequest request = WebRequest.Create(address) as HttpWebRequest;
            request.AllowAutoRedirect = true;
            request.AuthenticationLevel = AuthenticationLevel.MutualAuthRequested;
            request.PreAuthenticate = true;
            request.UseDefaultCredentials = true;
            try
            {
                CookieContainer uriCookieContainerEx = new CookieJar().GetUriCookieContainerEx(new Uri(address));
                if ((uriCookieContainerEx != null) && (uriCookieContainerEx.Count > 0))
                {
                    foreach (object obj2 in uriCookieContainerEx.GetCookies(new Uri(address)))
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
                    foreach (string str2 in jar.Cookies.Keys)
                    {
                        request.CookieContainer.Add(new Cookie(str2, jar.Cookies[str2], "/", request.RequestUri.DnsSafeHost));
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
            request.Method = "POST";
            string str3 = "---------------------------" + DateTime.Now.Ticks.ToString("x", NumberFormatInfo.InvariantInfo);
            if ((values != null) && (values.Count > 0))
            {
                request.ContentType = "multipart/form-data; boundary=" + str3;
                str3 = "--" + str3;
            }
            using (Stream stream = request.GetRequestStream())
            {
                byte[] bytes;
                if ((values != null) && (values.Count > 0))
                {
                    foreach (string str4 in values.Keys)
                    {
                        bytes = Encoding.ASCII.GetBytes(str3 + Environment.NewLine);
                        stream.Write(bytes, 0, bytes.Length);
                        bytes = Encoding.ASCII.GetBytes(string.Format("Content-Disposition: form-data; name=\"{0}\"{1}{1}", str4, Environment.NewLine));
                        stream.Write(bytes, 0, bytes.Length);
                        bytes = Encoding.UTF8.GetBytes(values[str4] + Environment.NewLine);
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
                foreach (OpenEsdh.Outlook.Model.Alfresco.UploadFile file in files)
                {
                    if ((values != null) && (values.Count > 0))
                    {
                        bytes = Encoding.ASCII.GetBytes(str3 + Environment.NewLine);
                        stream.Write(bytes, 0, bytes.Length);
                        bytes = Encoding.UTF8.GetBytes(string.Format("Content-Disposition: form-data; name=\"filedata\"; filename=\"{0}\"{1}", file.Name, Environment.NewLine));
                        stream.Write(bytes, 0, bytes.Length);
                        bytes = Encoding.ASCII.GetBytes(string.Format("Content-Type: {0}{1}{1}", file.ContentType, Environment.NewLine));
                        stream.Write(bytes, 0, bytes.Length);
                        file.Stream.CopyTo(stream);
                        bytes = Encoding.ASCII.GetBytes(Environment.NewLine);
                        stream.Write(bytes, 0, bytes.Length);
                    }
                    else
                    {
                        bytes = Encoding.ASCII.GetBytes(string.Format("Content-Type: {0}{1}{1}", file.ContentType, Environment.NewLine));
                        stream.Write(bytes, 0, bytes.Length);
                        file.Stream.CopyTo(stream);
                        bytes = Encoding.ASCII.GetBytes(Environment.NewLine);
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
                if ((values != null) && (values.Count > 0))
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(str3 + "--");
                    stream.Write(buffer, 0, buffer.Length);
                }
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
                            buffer3 = stream3.ToArray();
                        }
                    }
                }
            }
            catch
            {
                buffer3 = null;
            }
            return buffer3;
        }

        protected virtual string Url
        {
            get
            {
                return this._url;
            }
        }
    }
}


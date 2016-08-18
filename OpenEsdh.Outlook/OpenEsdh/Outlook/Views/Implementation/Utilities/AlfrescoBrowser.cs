namespace OpenEsdh.Outlook.Views.Implementation.Utilities
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using OpenEsdh.Outlook.Views.ServerCertificate;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;

    [ComVisible(true)]
    public class AlfrescoBrowser : UserControl
    {
        private IConfiguration _configuration = null;
        private Dictionary<string, string> _cookies = new Dictionary<string, string>();
        private bool _doneLogin = false;
        private string _payload1 = "";
        private string _payload2 = "";
        private int _redirectRetry = 0;
        private Uri _uri;
        private IContainer components = null;
        private const int INTERNET_COOKIE_HTTPONLY = 0x2000;
        private WebBrowser webBrowser;

        public event CancelDelegate OnCancel;

        public event SaveDelegate OnSave;

        public event SetSizeDelegate OnSetSize;

        public AlfrescoBrowser()
        {
            this.InitializeComponent();
            this.webBrowser.ObjectForScripting = this;
        }

        public void CancelOpenEsdh()
        {
            if (this.OnCancel != null)
            {
                this.OnCancel(this, new OpenEsdh.Outlook.Views.Implementation.Utilities.CancelEventArgs());
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this._redirectRetry <= this._configuration.MaxRedirectRetries)
            {
                this._redirectRetry++;
                if ((this._redirectRetry < this._configuration.MaxRedirectRetries) && !this._uri.AbsoluteUri.Contains(this.webBrowser.Url.AbsoluteUri))
                {
                    if (!this._configuration.UseRedirectJavascript)
                    {
                        this.RunRequests();
                    }
                    else
                    {
                        this.webBrowser.Document.InvokeScript("eval", new object[] { "document.location='" + this._uri.AbsoluteUri + "'" });
                    }
                }
                else
                {
                    Rectangle offsetRectangle = this.webBrowser.Document.GetElementsByTagName("body")[0].OffsetRectangle;
                    int num = Math.Max(base.Height, offsetRectangle.Height + this._configuration.DialogExtend.Y);
                    int num2 = Math.Max(base.Width, offsetRectangle.Width + this._configuration.DialogExtend.X);
                    num = Math.Min(base.Height, this._configuration.DialogExtend.MaxHeight);
                    num2 = Math.Min(base.Width, this._configuration.DialogExtend.MaxWidth);
                    if (this.OnSetSize != null)
                    {
                        SetSizeEventArgs size = new SetSizeEventArgs {
                            Height = num,
                            Width = num2
                        };
                        this.OnSetSize(this, size);
                    }
                    if (this._configuration.CommunicationConfiguration.SendMethod == SendDataMethod.JavascriptMethod)
                    {
                        this.SendJSONPayload();
                    }
                }
            }
        }

        public string getParameter1()
        {
            return this._payload1;
        }

        public string getParameter2()
        {
            return this._payload2;
        }

        public void Initialize(IConfiguration configuration, Uri uri, string payload1)
        {
            this.Initialize(configuration, uri, payload1, "");
        }

        public void Initialize(IConfiguration configuration, Uri uri, string payload1, string payload2)
        {
            this._redirectRetry = 0;
            this._configuration = configuration;
            Debug.WriteLine(string.Concat(new object[] { "Open browser:", uri, ":", payload1, ":", payload2 }));
            this._payload1 = payload1;
            this._payload2 = payload2;
            this._uri = uri;
            try
            {
                if (configuration.IgnoreCertificateErrors)
                {
                    WindowsInterop.SecurityAlertDialogWillBeShown += new GenericDelegate<bool, bool>(this.WindowsInterop_SecurityAlertDialogWillBeShown);
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        private void InitializeComponent()
        {
            this.webBrowser = new WebBrowser();
            base.SuspendLayout();
            this.webBrowser.Dock = DockStyle.Fill;
            this.webBrowser.Location = new Point(0, 0);
            this.webBrowser.MinimumSize = new Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.ScriptErrorsSuppressed = true;
            this.webBrowser.Size = new Size(0x428, 0x25f);
            this.webBrowser.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.Controls.Add(this.webBrowser);
            base.Name = "AlfrescoBrowser";
            base.Size = new Size(0x428, 0x25f);
            base.ResumeLayout(false);
        }

        [DllImport("wininet.dll")]
        private static extern InternetCookieState InternetSetCookieEx(string lpszURL, string lpszCookieName, string lpszCookieData, int dwFlags, IntPtr dwReserved);
        public void RunRequests()
        {
            if (this._redirectRetry == 0)
            {
                WebBrowserDocumentCompletedEventHandler handler = this.webBrowser.GetType().GetField("DocumentCompleted", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(this.webBrowser) as WebBrowserDocumentCompletedEventHandler;
                if (handler == null)
                {
                    this.webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(this.DocumentCompleted);
                }
            }
            if (this._redirectRetry <= this._configuration.MaxRedirectRetries)
            {
                ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
                if ((jar != null) && (jar.Cookies.Count > 0))
                {
                    foreach (KeyValuePair<string, string> pair in jar.Cookies)
                    {
                        this.SetCookie(this._uri.Scheme + "://" + this._uri.DnsSafeHost, pair.Key, pair.Value);
                    }
                }
                if (this._configuration.CommunicationConfiguration.SendMethod == SendDataMethod.JavascriptMethod)
                {
                    this.webBrowser.Url = this._uri;
                }
                else
                {
                    JavaScriptSerializer serializer = new JavaScriptSerializer();
                    string s = "";
                    if (!string.IsNullOrEmpty(this._payload1))
                    {
                        s = serializer.Serialize(this._payload1);
                        s = this._configuration.CommunicationConfiguration.PostMethodName + "=" + s;
                    }
                    if (!string.IsNullOrEmpty(this._payload2))
                    {
                        s = "&" + this._payload2;
                    }
                    this.webBrowser.Navigate(this._uri, "_top", Encoding.ASCII.GetBytes(s), "");
                }
            }
        }

        public void RunRequests(IConfiguration configuration, Uri uri, string payload1)
        {
            this.RunRequests(configuration, uri, payload1, "");
        }

        public void RunRequests(IConfiguration configuration, Uri uri, string payload1, string payload2)
        {
            this.Initialize(configuration, uri, payload1, payload2);
            this.RunRequests();
        }

        public void SaveAsOpenEsdh(string unknownJson, string attachmentSelectedJson)
        {
            if (this.OnSave != null)
            {
                SaveEventArgs args = new SaveEventArgs {
                    ReturnValues1 = unknownJson,
                    ReturnValues2 = attachmentSelectedJson
                };
                this.OnSave(this, args);
            }
        }

        public void SaveAsOpenEsdhApplication(string unknownJson)
        {
            if (this.OnSave != null)
            {
                SaveEventArgs args = new SaveEventArgs {
                    ReturnValues1 = unknownJson
                };
                this.OnSave(this, args);
            }
        }

        private void SendJSONPayload()
        {
            new Thread(new ThreadStart(this.ThreadInvokeJsonPayload)).Start();
        }

        private InternetCookieState SetCookie(string url, string name, string data)
        {
            return InternetSetCookieEx(url, name, data, 0x2000, IntPtr.Zero);
        }

        public void ThreadInvokeJsonPayload()
        {
            MethodInvoker method = null;
            try
            {
                Thread.Sleep(this._configuration.CommunicationConfiguration.DelayUntilJavaMethodCall);
                if (method == null)
                {
                    method = delegate {
                        if (this._uri.AbsoluteUri.Contains(this.webBrowser.Url.AbsoluteUri))
                        {
                            List<object> list = new List<object>();
                            if (!string.IsNullOrEmpty(this._payload1))
                            {
                                list.Add(this._payload1);
                            }
                            if (!string.IsNullOrEmpty(this._payload2))
                            {
                                list.Add(this._payload2);
                            }
                            bool flag = true;
                            this.webBrowser.Document.InvokeScript(this._configuration.CommunicationConfiguration.JavaScriptMethodName, list.ToArray());
                            Debug.WriteLine(this._configuration.CommunicationConfiguration.JavaScriptMethodName + "(" + string.Join(",", list.ToArray()) + ")");
                            HtmlElementCollection elementsByTagName = this.webBrowser.Document.GetElementsByTagName(this._configuration.LoginTagToFind);
                            foreach (HtmlElement element in elementsByTagName)
                            {
                                if ((!string.IsNullOrEmpty(element.Id) && element.Id.ToLower().Contains(this._configuration.LoginIdToFind.ToLower())) || (!string.IsNullOrEmpty(element.InnerHtml) && element.InnerHtml.ToLower().Contains(this._configuration.LoginIdToFind.ToLower())))
                                {
                                    Debug.WriteLine(this._configuration.CommunicationConfiguration.JavaScriptMethodName + "(" + string.Join(",", list.ToArray()) + ")");
                                    this.webBrowser.Document.InvokeScript(this._configuration.CommunicationConfiguration.JavaScriptMethodName, list.ToArray());
                                    flag = false;
                                    this.RunRequests();
                                    break;
                                }
                            }
                            if (flag)
                            {
                                this.webBrowser.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(this.DocumentCompleted);
                            }
                        }
                    };
                }
                base.Invoke(method);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        private bool WindowsInterop_SecurityAlertDialogWillBeShown(bool IsSSLDialog)
        {
            return true;
        }

        private Uri Url
        {
            get
            {
                return this._uri;
            }
            set
            {
                this._uri = value;
            }
        }

        public enum InternetCookieState
        {
            COOKIE_STATE_ACCEPT = 1,
            COOKIE_STATE_DOWNGRADE = 4,
            COOKIE_STATE_LEASH = 3,
            COOKIE_STATE_MAX = 5,
            COOKIE_STATE_PROMPT = 2,
            COOKIE_STATE_REJECT = 5,
            COOKIE_STATE_UNKNOWN = 0
        }
    }
}


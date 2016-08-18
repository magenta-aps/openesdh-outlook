namespace OpenEsdh.Outlook.Views.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using OpenEsdh.Outlook.Views.Interface;
    using OpenEsdh.Outlook.Views.ServerCertificate;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using System.Windows.Forms;

    [ComVisible(true)]
    public class DisplayRegionControl : UserControl, IDisplayRegionControl
    {
        private bool _doneLogin = false;
        private int _redirectRetry = 0;
        private string _startUrl = "";
        private IContainer components = null;
        private IOutlookConfiguration config = null;
        private WebBrowser webBrowser1;

        public DisplayRegionControl()
        {
            this.InitializeComponent();
            try
            {
                this.webBrowser1.ObjectForScripting = this;
                this.config = TypeResolver.Current.Create<IOutlookConfiguration>();
                if (this.config.IgnoreCertificateErrors)
                {
                    WindowsInterop.SecurityAlertDialogWillBeShown += new GenericDelegate<bool, bool>(this.WindowsInterop_SecurityAlertDialogWillBeShown);
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
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

        public void Initialize()
        {
            MethodInvoker method = null;
            try
            {
                Thread.Sleep(this.config.CommunicationConfiguration.DelayUntilJavaMethodCall);
                if (method == null)
                {
                    method = delegate {
                        if (this._startUrl.Contains(this.webBrowser1.Url.AbsoluteUri))
                        {
                            bool flag = true;
                            HtmlElementCollection elementsByTagName = this.webBrowser1.Document.GetElementsByTagName(this.config.LoginTagToFind);
                            foreach (HtmlElement element in elementsByTagName)
                            {
                                if (!string.IsNullOrEmpty(element.Id) && element.Id.ToLower().Contains(this.config.LoginIdToFind.ToLower()))
                                {
                                    flag = false;
                                    if (((this.config.PreAuthentication.UseConfigCredentials && !this._doneLogin) && !string.IsNullOrEmpty(this.config.PreAuthentication.Username)) && !string.IsNullOrEmpty(this.config.PreAuthentication.Password))
                                    {
                                        string urlString = this.webBrowser1.Url.AbsoluteUri;
                                        if (!string.IsNullOrEmpty(this.config.PreAuthentication.AuthenticationUrl))
                                        {
                                            urlString = this.config.PreAuthentication.AuthenticationUrl;
                                        }
                                        string newValue = this.config.PreAuthentication.Username;
                                        string password = this.config.PreAuthentication.Password;
                                        string s = this.config.PreAuthentication.AuthenticationPackageFormat.Replace("[@username]", newValue).Replace("[@password]", password);
                                        string additionalHeaders = "Referer: " + this._startUrl;
                                        string[] strArray = this.config.PreAuthentication.AdditionalRequestHeaders.Split(new char[] { ';' });
                                        string str6 = "\r\n";
                                        foreach (string str7 in strArray)
                                        {
                                            additionalHeaders = additionalHeaders + str6 + str7;
                                        }
                                        this.webBrowser1.Navigate(urlString, "_top", Encoding.ASCII.GetBytes(s), additionalHeaders);
                                        this._doneLogin = true;
                                    }
                                    break;
                                }
                            }
                            if (flag)
                            {
                                this.webBrowser1.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);
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

        private void InitializeComponent()
        {
            this.webBrowser1 = new WebBrowser();
            base.SuspendLayout();
            this.webBrowser1.Dock = DockStyle.Fill;
            this.webBrowser1.Location = new Point(0, 0);
            this.webBrowser1.MinimumSize = new Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new Size(0x5d6, 0xf7);
            this.webBrowser1.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.Controls.Add(this.webBrowser1);
            base.Name = "DisplayRegion";
            base.Size = new Size(0x5d6, 0xf7);
            base.ResumeLayout(false);
        }

        public void InitializeOpenEsdh(IOutlookConfiguration config)
        {
            new Thread(new ThreadStart(this.Initialize)).Start();
        }

        public void Show(string url)
        {
            this._redirectRetry = 0;
            this._startUrl = url;
            if (this._redirectRetry == 0)
            {
                WebBrowserDocumentCompletedEventHandler handler = this.webBrowser1.GetType().GetField("DocumentCompleted", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(this.webBrowser1) as WebBrowserDocumentCompletedEventHandler;
                if (handler == null)
                {
                    this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);
                }
            }
            this.webBrowser1.Navigate(url);
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            IOutlookConfiguration config = TypeResolver.Current.Create<IOutlookConfiguration>();
            if (this._redirectRetry >= config.MaxRedirectRetries)
            {
                try
                {
                    this.webBrowser1.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(this.webBrowser1_DocumentCompleted);
                }
                catch
                {
                }
            }
            else if ((this._redirectRetry < config.MaxRedirectRetries) && !this._startUrl.Contains(this.webBrowser1.Url.AbsoluteUri))
            {
                this._redirectRetry++;
                if (!config.UseRedirectJavascript)
                {
                    this.Show(this._startUrl);
                }
                else
                {
                    this.webBrowser1.Document.InvokeScript("eval", new object[] { "document.location='" + this._startUrl + "'" });
                }
            }
            else
            {
                ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
                jar.Add(this.webBrowser1.Document.Cookie);
                jar.AddCookiesForUri(this.webBrowser1.Url);
                this.InitializeOpenEsdh(config);
            }
        }

        private bool WindowsInterop_SecurityAlertDialogWillBeShown(bool IsSSLDialog)
        {
            return true;
        }
    }
}


namespace OpenEsdh.Outlook.Views.Implementation
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Interface;
    using OpenEsdh.Outlook.Views.ServerCertificate;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;

    [ComVisible(true)]
    public class AttachFileView : Form, IAttachFileView
    {
        private bool _closePressed = false;
        private bool _doneLogin = false;
        private IAttachFilePresenter _presenter = null;
        private int _redirectRetry = 0;
        private string _startUrl = "";
        private IContainer components = null;
        private IOutlookConfiguration config = null;
        private WebBrowser OpenEsdhBrowser;

        public AttachFileView()
        {
            this._closePressed = false;
            this.InitializeComponent();
            this.OpenEsdhBrowser.ObjectForScripting = this;
            this.Text = ResourceResolver.Current.GetString("AttachFileDialogTitle");
            try
            {
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

        public void AttachFile(string AttachmentConfiguration)
        {
            this._presenter.AttachFile(AttachmentConfiguration);
            base.Close();
        }

        public void Cancel()
        {
            this._presenter.Cancel();
            if (!this._closePressed)
            {
                this._closePressed = true;
                base.Close();
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
            IOutlookConfiguration config = TypeResolver.Current.Create<IOutlookConfiguration>();
            if ((this._redirectRetry < config.MaxRedirectRetries) && !this._startUrl.Contains(this.OpenEsdhBrowser.Url.AbsoluteUri))
            {
                this._redirectRetry++;
                if (!config.UseRedirectJavascript)
                {
                    this.Initialize(this._startUrl);
                }
                else
                {
                    this.OpenEsdhBrowser.Document.InvokeScript("eval", new object[] { "document.location='" + this._startUrl + "'" });
                }
            }
            else
            {
                this._redirectRetry = 0;
                ICookieJar jar = TypeResolver.Current.Create<ICookieJar>();
                jar.Add(this.OpenEsdhBrowser.Document.Cookie);
                jar.AddCookiesForUri(this.OpenEsdhBrowser.Url);
                Rectangle offsetRectangle = this.OpenEsdhBrowser.Document.GetElementsByTagName("body")[0].OffsetRectangle;
                base.Height = Math.Max(base.Height, offsetRectangle.Height + config.DialogExtend.Y);
                base.Width = Math.Max(base.Width, offsetRectangle.Width + config.DialogExtend.X);
                base.Height = Math.Min(base.Height, config.DialogExtend.MaxHeight);
                base.Width = Math.Min(base.Width, config.DialogExtend.MaxWidth);
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                if (this.OpenEsdhBrowser.Url.AbsoluteUri.ToLower().Contains(this._startUrl.ToLower()))
                {
                    if (config.CommunicationConfiguration.SendMethod == SendDataMethod.JavascriptMethod)
                    {
                        this.InitializeOpenEsdh(config);
                    }
                    else
                    {
                        this.InitializeOpendEsdhPost(config);
                    }
                }
            }
        }

        public void Initialize()
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            this.OpenEsdhBrowser.Document.InvokeScript(this.config.CommunicationConfiguration.JavaScriptMethodName);
        }

        public void Initialize(object o)
        {
            MethodInvoker method = null;
            try
            {
                Thread.Sleep(this.config.CommunicationConfiguration.DelayUntilJavaMethodCall);
                if (method == null)
                {
                    method = delegate {
                        if (this._startUrl.Contains(this.OpenEsdhBrowser.Url.AbsoluteUri))
                        {
                            object[] objArray1 = o as object[];
                            bool flag = true;
                            HtmlElementCollection elementsByTagName = this.OpenEsdhBrowser.Document.GetElementsByTagName(this.config.LoginTagToFind);
                            foreach (HtmlElement element in elementsByTagName)
                            {
                                if (!string.IsNullOrEmpty(element.Id) && element.Id.ToLower().Contains(this.config.LoginIdToFind.ToLower()))
                                {
                                    flag = false;
                                    if (((this.config.PreAuthentication.UseConfigCredentials && !this._doneLogin) && !string.IsNullOrEmpty(this.config.PreAuthentication.Username)) && !string.IsNullOrEmpty(this.config.PreAuthentication.Password))
                                    {
                                        string urlString = this.OpenEsdhBrowser.Url.AbsoluteUri;
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
                                        this.OpenEsdhBrowser.Navigate(urlString, "_top", Encoding.ASCII.GetBytes(s), additionalHeaders);
                                        this._doneLogin = true;
                                    }
                                    break;
                                }
                            }
                            if (flag)
                            {
                                this.OpenEsdhBrowser.DocumentCompleted -= new WebBrowserDocumentCompletedEventHandler(this.DocumentCompleted);
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

        public void Initialize(string uri, EmailDescriptor Email)
        {
            this._startUrl = uri;
            if (this._redirectRetry == 0)
            {
                this.OpenEsdhBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(this.DocumentCompleted);
            }
            if (this.config.CommunicationConfiguration.SendMethod == SendDataMethod.JavascriptMethod)
            {
                this.OpenEsdhBrowser.Url = new Uri(uri);
            }
            else
            {
                this.OpenEsdhBrowser.Navigate(new Uri(uri), "_top");
            }
            base.ShowDialog();
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(AttachFileView));
            this.OpenEsdhBrowser = new WebBrowser();
            base.SuspendLayout();
            this.OpenEsdhBrowser.Dock = DockStyle.Fill;
            this.OpenEsdhBrowser.Location = new Point(0, 0);
            this.OpenEsdhBrowser.MinimumSize = new Size(20, 20);
            this.OpenEsdhBrowser.Name = "OpenEsdhBrowser";
            this.OpenEsdhBrowser.Size = new Size(0x538, 0x21d);
            this.OpenEsdhBrowser.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x538, 0x21d);
            base.Controls.Add(this.OpenEsdhBrowser);
            base.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.MaximizeBox = false;
            base.MinimizeBox = false;
            base.Name = "AttachFileView";
            base.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "AttachFileView";
            base.ResumeLayout(false);
        }

        public void InitializeOpendEsdhPost(IOutlookConfiguration config)
        {
        }

        public void InitializeOpenEsdh(IOutlookConfiguration config)
        {
            new Thread(new ParameterizedThreadStart(this.Initialize)).Start();
        }

        private void SaveAsView_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Cancel();
        }

        public void ShowView()
        {
            base.ShowDialog();
        }

        private bool WindowsInterop_SecurityAlertDialogWillBeShown(bool IsSSLDialog)
        {
            return true;
        }

        public IAttachFilePresenter Presenter
        {
            get
            {
                if (this._presenter == null)
                {
                    this._presenter = TypeResolver.Current.Create<IAttachFilePresenter>();
                }
                return this._presenter;
            }
            set
            {
                this._presenter = value;
            }
        }
    }
}


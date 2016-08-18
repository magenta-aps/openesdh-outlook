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
    public class ApplicationSaveAsView : Form, IApplicationSaveAsView
    {
        private bool _closePressed;
        private ApplicationDescriptor _document;
        private bool _doneLogin;
        private IApplicationSaveAsPresenter _presenter;
        private int _redirectRetry;
        private string _startUrl;
        private IContainer components;
        private IWordConfiguration config;
        private WebBrowser OpenEsdhBrowser;

        public ApplicationSaveAsView()
        {
            this._document = null;
            this._presenter = null;
            this.config = null;
            this._startUrl = "";
            this._redirectRetry = 0;
            this._doneLogin = false;
            this._closePressed = false;
            this.components = null;
            this._closePressed = false;
            this.InitializeComponent();
            this.OpenEsdhBrowser.ObjectForScripting = this;
            this.Text = ResourceResolver.Current.GetString("SaveAsDialogTitle");
            try
            {
                this.config = TypeResolver.Current.Create<IWordConfiguration>();
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

        public ApplicationSaveAsView(IApplicationSaveAsPresenter presenter) : this()
        {
            this._presenter = presenter;
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

        public void CancelOpenEsdh()
        {
            this.Cancel();
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
            IWordConfiguration config = TypeResolver.Current.Create<IWordConfiguration>();
            if ((this._redirectRetry < config.MaxRedirectRetries) && !this._startUrl.Contains(this.OpenEsdhBrowser.Url.AbsoluteUri))
            {
                this._redirectRetry++;
                if (!config.UseRedirectJavascript)
                {
                    this.Initialize(this._startUrl, this._document);
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
                string jsonEmail = new JavaScriptSerializer().Serialize(this._document);
                if (this.OpenEsdhBrowser.Url.AbsoluteUri.ToLower().Contains(this._startUrl.ToLower()))
                {
                    if (config.CommunicationConfiguration.SendMethod == SendDataMethod.JavascriptMethod)
                    {
                        this.InitializeOpenEsdh(config, jsonEmail);
                    }
                    else
                    {
                        this.InitializeOpendEsdhPost(config, jsonEmail);
                    }
                }
            }
        }

        public void Initialize()
        {
            string str = new JavaScriptSerializer().Serialize(this._document);
            object[] args = new string[] { str };
            this.OpenEsdhBrowser.Document.InvokeScript(this.config.CommunicationConfiguration.JavaScriptMethodName, args);
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
                                if (!string.IsNullOrEmpty(element.Id) && element.Id.ToLower().Contains(this.config.LoginTagToFind.ToLower()))
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

        public void Initialize(string uri, ApplicationDescriptor document)
        {
            this._startUrl = uri;
            this._document = document;
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
                string s = new JavaScriptSerializer().Serialize(this._document);
                s = this.config.CommunicationConfiguration.PostMethodName + "=" + s;
                this.OpenEsdhBrowser.Navigate(new Uri(uri), "_top", Encoding.ASCII.GetBytes(s), "");
            }
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(ApplicationSaveAsView));
            this.OpenEsdhBrowser = new WebBrowser();
            base.SuspendLayout();
            this.OpenEsdhBrowser.CausesValidation = false;
            this.OpenEsdhBrowser.Dock = DockStyle.Fill;
            this.OpenEsdhBrowser.Location = new Point(0, 0);
            this.OpenEsdhBrowser.MinimumSize = new Size(20, 20);
            this.OpenEsdhBrowser.Name = "OpenEsdhBrowser";
            this.OpenEsdhBrowser.Size = new Size(0x49a, 0x218);
            this.OpenEsdhBrowser.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x3ab, 0x10a);
            base.Controls.Add(this.OpenEsdhBrowser);
            base.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.MaximizeBox = false;
            base.MinimizeBox = false;
            base.Name = "ApplicationSaveAsView";
            base.StartPosition = FormStartPosition.CenterParent;
            this.Text = "ApplicationSaveAsView";
            base.FormClosing += new FormClosingEventHandler(this.SaveAsView_FormClosing);
            base.ResumeLayout(false);
        }

        public void InitializeOpendEsdhPost(IWordConfiguration config, string payload)
        {
        }

        public void InitializeOpenEsdh(IWordConfiguration config, string jsonEmail)
        {
            new Thread(new ParameterizedThreadStart(this.Initialize)).Start(new object[] { jsonEmail });
        }

        public void SaveAs(string unknown)
        {
            try
            {
                Logger.Current.LogInformation("SaveAs(" + unknown + ")", "");
                this._presenter.SaveAs(unknown);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public void SaveAsOpenEsdh(string unknownJson, string test)
        {
            this.SaveAs(unknownJson);
            base.Close();
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

        public IApplicationSaveAsPresenter Presenter
        {
            get
            {
                if (this._presenter == null)
                {
                    this._presenter = TypeResolver.Current.Create<IApplicationSaveAsPresenter>();
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


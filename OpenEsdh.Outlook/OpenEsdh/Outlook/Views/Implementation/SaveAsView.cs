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
    public class SaveAsView : Form, ISaveAsBrowserView, ISaveAsView
    {
        private bool _doneLogin;
        private EmailDescriptor _Email;
        private ISaveAsPresenter _presenter;
        private int _redirectRetry;
        private string _startUrl;
        private IContainer components;
        private IOutlookConfiguration config;
        private WebBrowser OpenEsdhBrowser;

        public SaveAsView()
        {
            this._Email = null;
            this._presenter = null;
            this.config = null;
            this._startUrl = "";
            this._redirectRetry = 0;
            this._doneLogin = false;
            this.components = null;
            this.InitializeComponent();
            this.OpenEsdhBrowser.ObjectForScripting = this;
            this.Text = ResourceResolver.Current.GetString("SaveAsDialogTitle");
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

        public SaveAsView(ISaveAsPresenter presenter) : this()
        {
            this._presenter = presenter;
        }

        public void Cancel()
        {
            this._presenter.Cancel();
            base.Close();
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
            IOutlookConfiguration config = TypeResolver.Current.Create<IOutlookConfiguration>();
            if ((this._redirectRetry < config.MaxRedirectRetries) && !this._startUrl.Contains(this.OpenEsdhBrowser.Url.AbsoluteUri))
            {
                this._redirectRetry++;
                if (!config.UseRedirectJavascript)
                {
                    this.Initialize(this._startUrl, this._Email);
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
                string jsonEmail = new JavaScriptSerializer().Serialize(this._Email);
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

        public void Initialize()
        {
            string str = new JavaScriptSerializer().Serialize(this._Email);
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
                            object[] args = o as object[];
                            bool flag = true;
                            this.OpenEsdhBrowser.Document.InvokeScript(this.config.CommunicationConfiguration.JavaScriptMethodName, args);
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
            this._Email = Email;
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
                string s = new JavaScriptSerializer().Serialize(this._Email);
                s = this.config.CommunicationConfiguration.PostMethodName + "=" + s;
                this.OpenEsdhBrowser.Navigate(new Uri(uri), "_top", Encoding.ASCII.GetBytes(s), "");
            }
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(SaveAsView));
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
            base.ClientSize = new Size(0x49a, 0x218);
            base.Controls.Add(this.OpenEsdhBrowser);
            base.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.Name = "SaveAsView";
            base.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Visma Case";
            base.FormClosing += new FormClosingEventHandler(this.SaveAsView_FormClosing);
            base.ResumeLayout(false);
        }

        public void InitializeOpendEsdhPost(IOutlookConfiguration config, string payload)
        {
        }

        public void InitializeOpenEsdh(IOutlookConfiguration config, string jsonEmail)
        {
            new Thread(new ParameterizedThreadStart(this.Initialize)).Start(new object[] { jsonEmail });
        }

        public void SaveAs(string unknown, SelectableAttachment[] SelectedAttachments)
        {
            try
            {
                Logger.Current.LogInformation("SaveAs(" + unknown + "," + ((SelectedAttachments != null) ? SelectedAttachments.Length.ToString() : "null") + ")", "");
                this._presenter.SaveAs(unknown, SelectedAttachments);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public void SaveAsOpenEsdh(string unknownJson, string attachmentSelectedJson)
        {
            SelectableAttachment[] selectedAttachments = new JavaScriptSerializer().Deserialize<SelectableAttachment[]>(attachmentSelectedJson);
            this.SaveAs(unknownJson, selectedAttachments);
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

        public ISaveAsPresenter Presenter
        {
            get
            {
                if (this._presenter == null)
                {
                    this._presenter = TypeResolver.Current.Create<ISaveAsPresenter>();
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


namespace OpenEsdh.Outlook.Views.Implementation.OfficeApplications
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Implementation.Utilities;
    using OpenEsdh.Outlook.Views.Interface;
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;

    public class SaveAs : Form, IApplicationSaveAsView
    {
        private IWordConfiguration _config;
        private ApplicationDescriptor _document;
        private IApplicationSaveAsPresenter _presenter;
        private AlfrescoBrowser alfrescoBrowser;
        private IContainer components;

        public SaveAs()
        {
            this._presenter = null;
            this._config = null;
            this._document = null;
            this.components = null;
            this.InitializeComponent();
            this._config = TypeResolver.Current.Create<IWordConfiguration>();
            this.Text = ResourceResolver.Current.GetString("SaveAsDialogTitle");
            this.alfrescoBrowser.OnCancel += new CancelDelegate(this.alfrescoBrowser_OnCancel);
            this.alfrescoBrowser.OnSave += new SaveDelegate(this.alfrescoBrowser_OnSave);
            this.alfrescoBrowser.OnSetSize += new SetSizeDelegate(this.alfrescoBrowser_OnSetSize);
        }

        public SaveAs(IApplicationSaveAsPresenter presenter) : this()
        {
            this._presenter = presenter;
        }

        private void alfrescoBrowser_OnCancel(object sender, OpenEsdh.Outlook.Views.Implementation.Utilities.CancelEventArgs args)
        {
            base.Close();
        }

        private void alfrescoBrowser_OnSave(object sender, SaveEventArgs args)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            this.DoSaveAs(args.ReturnValues1);
            base.Close();
        }

        private void alfrescoBrowser_OnSetSize(object Sender, SetSizeEventArgs Size)
        {
            base.Height = Size.Height;
            base.Width = Size.Width;
        }

        public void Cancel()
        {
            base.Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void DoSaveAs(string unknown)
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

        public void Initialize(string uri, ApplicationDescriptor document)
        {
            this._document = document;
            string str = new JavaScriptSerializer().Serialize(this._document);
            Debug.WriteLine("Url:" + uri);
            this.alfrescoBrowser.RunRequests(this._config, new Uri(uri), str);
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(SaveAs));
            this.alfrescoBrowser = new AlfrescoBrowser();
            base.SuspendLayout();
            this.alfrescoBrowser.Dock = DockStyle.Fill;
            this.alfrescoBrowser.Location = new Point(0, 0);
            this.alfrescoBrowser.Name = "alfrescoBrowser";
            this.alfrescoBrowser.Size = new Size(0x3ea, 0x19c);
            this.alfrescoBrowser.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x3ea, 0x19c);
            base.Controls.Add(this.alfrescoBrowser);
            base.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.MaximizeBox = false;
            base.MinimizeBox = false;
            base.Name = "SaveAs";
            base.FormClosing += new FormClosingEventHandler(this.SaveAs_FormClosing);
            base.ResumeLayout(false);
        }

        private void SaveAs_FormClosing(object sender, FormClosingEventArgs e)
        {
            this._presenter.Cancel();
        }

        public void ShowView()
        {
            base.ShowDialog();
        }

        public IApplicationSaveAsPresenter Presenter
        {
            get
            {
                return this._presenter;
            }
            set
            {
                this._presenter = value;
            }
        }
    }
}


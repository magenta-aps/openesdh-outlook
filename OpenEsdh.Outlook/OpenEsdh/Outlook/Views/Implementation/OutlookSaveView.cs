namespace OpenEsdh.Outlook.Views.Implementation
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

    public class OutlookSaveView : Form, ISaveAsView
    {
        private IOutlookConfiguration _config;
        private EmailDescriptor _Email;
        private ISaveAsPresenter _presenter;
        private AlfrescoBrowser alfrescoBrowser1;
        private IContainer components;

        public OutlookSaveView()
        {
            this._Email = null;
            this._presenter = null;
            this._config = null;
            this.components = null;
            this.InitializeComponent();
            this._config = TypeResolver.Current.Create<IOutlookConfiguration>();
            this.Text = ResourceResolver.Current.GetString("SaveAsDialogTitle");
            this.alfrescoBrowser1.OnCancel += new CancelDelegate(this.alfrescoBrowser1_OnCancel);
            this.alfrescoBrowser1.OnSave += new SaveDelegate(this.alfrescoBrowser1_OnSave);
            this.alfrescoBrowser1.OnSetSize += new SetSizeDelegate(this.alfrescoBrowser1_OnSetSize);
        }

        public OutlookSaveView(ISaveAsPresenter presenter) : this()
        {
            this._presenter = presenter;
        }

        private void alfrescoBrowser1_OnCancel(object sender, OpenEsdh.Outlook.Views.Implementation.Utilities.CancelEventArgs args)
        {
            base.Close();
        }

        private void alfrescoBrowser1_OnSave(object sender, SaveEventArgs args)
        {
            SelectableAttachment[] selectedAttachments = new JavaScriptSerializer().Deserialize<SelectableAttachment[]>(args.ReturnValues2);
            this.SaveAs(args.ReturnValues1, selectedAttachments);
            base.Close();
        }

        private void alfrescoBrowser1_OnSetSize(object Sender, SetSizeEventArgs Size)
        {
            base.Height = Size.Height;
            base.Width = base.Width;
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

        public void Initialize(string uri, EmailDescriptor Email)
        {
            this._Email = Email;
            string str = new JavaScriptSerializer().Serialize(this._Email);
            Debug.WriteLine("Url:" + uri);
            this.alfrescoBrowser1.RunRequests(this._config, new Uri(uri), str);
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(OutlookSaveView));
            this.alfrescoBrowser1 = new AlfrescoBrowser();
            base.SuspendLayout();
            this.alfrescoBrowser1.AutoSize = true;
            this.alfrescoBrowser1.Dock = DockStyle.Fill;
            this.alfrescoBrowser1.Location = new Point(0, 0);
            this.alfrescoBrowser1.Name = "alfrescoBrowser1";
            this.alfrescoBrowser1.Size = new Size(0x49a, 0x218);
            this.alfrescoBrowser1.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x49a, 0x218);
            base.Controls.Add(this.alfrescoBrowser1);
            base.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.Name = "OutlookSaveView";
            base.StartPosition = FormStartPosition.CenterParent;
            this.Text = "Visma Case";
            base.FormClosing += new FormClosingEventHandler(this.OnFormClosing);
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        public void OnFormClosing(object sender, EventArgs args)
        {
            this._presenter.Cancel();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void SaveAs(string unknown, SelectableAttachment[] SelectedAttachments)
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

        public void ShowView()
        {
            base.ShowDialog();
        }

        public ISaveAsPresenter Presenter
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


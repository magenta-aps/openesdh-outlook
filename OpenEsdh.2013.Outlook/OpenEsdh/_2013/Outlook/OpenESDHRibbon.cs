namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Ribbon;
    using OpenEsdh._2013.Outlook.Presentation.Interface;
    using OpenEsdh._2013.Outlook.Properties;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using System;
    using System.ComponentModel;
    using System.Reflection;

    public class OpenESDHRibbon : RibbonBase, ISaveEmailButtonView
    {
        private ISaveEmailPresenter _presenter;
        internal RibbonButton btnSaveAsSend;
        internal RibbonButton btnSaveFile;
        private IContainer components;
        internal RibbonGroup group1;
        internal RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;

        public OpenESDHRibbon() : base((Microsoft.Office.Tools.Ribbon.RibbonFactory) Globals.Factory.GetRibbonFactory())
        {
            this.components = null;
            this.InitializeComponent();
        }

        private void btnSaveAsSend_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                object context = base.Context;
                MailItem item = ((dynamic) e.Control.Context).CurrentItem as MailItem;
                if (item != null)
                {
                    this._presenter.SaveEmailClick(item);
                    try
                    {
                        Microsoft.Office.Interop.Outlook.Inspector getInspector = item.GetInspector;
                        if (getInspector != null)
                        {
                            getInspector.ShowFormPage(ResourceResolver.Current.GetString("ViewRegionTitle"));
                            getInspector.Display(Missing.Value);
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        private void btnSaveFile_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MailItem item = ((dynamic) e.Control.Context).CurrentItem as MailItem;
                if (this._presenter.SaveEmailAndSend(item, () => item.Send()))
                {
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void Initialize()
        {
            try
            {
                this._presenter = TypeResolver.Current.Create<ISaveEmailPresenter>();
                this._presenter.View = this;
                this._presenter.Load(base.Context);
                this.group1.Label = ResourceResolver.Current.GetString("OpenESDHButtonGroup");
                this.group2.Label = ResourceResolver.Current.GetString("OpenESDHButtonGroup");
                this.btnSaveAsSend.Label = ResourceResolver.Current.GetString("SaveSendBtn");
                this.btnSaveFile.Label = ResourceResolver.Current.GetString("SaveBtn");
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        private void InitializeComponent()
        {
            this.tab1 = base.Factory.CreateRibbonTab();
            this.group1 = base.Factory.CreateRibbonGroup();
            this.tab2 = base.Factory.CreateRibbonTab();
            this.group2 = base.Factory.CreateRibbonGroup();
            this.btnSaveFile = base.Factory.CreateRibbonButton();
            this.btnSaveAsSend = base.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.tab1.ControlId.ControlIdType = RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabNewMailMessage";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabNewMailMessage";
            this.tab1.Name = "tab1";
            this.group1.Items.Add(this.btnSaveFile);
            this.group1.Label = "Visma Case";
            this.group1.Name = "group1";
            this.group1.Position = base.Factory.RibbonPosition.BeforeOfficeId("GroupSend");
            this.tab2.ControlId.ControlIdType = RibbonControlIdType.Office;
            this.tab2.ControlId.OfficeId = "TabReadMessage";
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "TabReadMessage";
            this.tab2.Name = "tab2";
            this.group2.Items.Add(this.btnSaveAsSend);
            this.group2.Label = "Visma Case";
            this.group2.Name = "group2";
            this.group2.Position = base.Factory.RibbonPosition.BeforeOfficeId("GroupRespond");
            this.btnSaveFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveFile.Image = Resources.openesdh_logo_32;
            this.btnSaveFile.Label = "Journalis\x00e9r";
            this.btnSaveFile.Name = "btnSaveFile";
            this.btnSaveFile.ShowImage = true;
            this.btnSaveFile.Click += new RibbonControlEventHandler(this.btnSaveFile_Click);
            this.btnSaveAsSend.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveAsSend.Image = Resources.openesdh_logo_32;
            this.btnSaveAsSend.Label = "Journalis\x00e9r";
            this.btnSaveAsSend.Name = "btnSaveAsSend";
            this.btnSaveAsSend.ShowImage = true;
            this.btnSaveAsSend.Click += new RibbonControlEventHandler(this.btnSaveAsSend_Click);
            base.Name = "OpenESDHRibbon";
            base.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mail.Read";
            base.Tabs.Add(this.tab1);
            base.Tabs.Add(this.tab2);
            base.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OpenESDHRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
        }

        private void OpenESDHRibbon_Load(object sender, Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs e)
        {
            if (this._presenter == null)
            {
                this.Initialize();
            }
        }

        public bool Visible
        {
            get
            {
                return this.group1.Visible;
            }
            set
            {
                this.group1.Visible = value;
                this.group2.Visible = value;
            }
        }
    }
}


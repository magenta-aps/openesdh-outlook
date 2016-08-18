namespace OpenEsdh._2013.Powerpoint
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.PowerPoint;
    using Microsoft.Office.Tools.Ribbon;
    using OpenEsdh._2013.Powerpoint.Presentation.Interface;
    using OpenEsdh._2013.Powerpoint.Properties;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class OpenESDHRibbon : RibbonBase, IPowerpointView
    {
        private IPowerpointPresenter _presenter;
        private bool _viewIsLocked;
        private IContainer components;
        internal RibbonGroup group1;
        internal RibbonButton Save;
        internal RibbonButton SaveAs;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;

        public OpenESDHRibbon() : base((Microsoft.Office.Tools.Ribbon.RibbonFactory) Globals.Factory.GetRibbonFactory())
        {
            this._presenter = null;
            this._viewIsLocked = false;
            this.components = null;
            this.InitializeComponent();
        }

        private void Application_WindowActivate(Presentation Pres, DocumentWindow Wn)
        {
            this._presenter.Load(Pres);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string sClassName, string sAppName);
        [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)]
        private static extern IntPtr GetDesktopWindow();
        public void Initialize()
        {
            try
            {
                this._presenter = TypeResolver.Current.Create<IPowerpointPresenter>(this);
                this.group1.Label = ResourceResolver.Current.GetString("OpenESDHAppGroup");
                this.Save.Label = ResourceResolver.Current.GetString("ApplicationSave");
                this.SaveAs.Label = ResourceResolver.Current.GetString("ApplicationSaveAs");
                new ComAwareEventInfo(typeof(EApplication_Event), "WindowActivate").AddEventHandler(Globals.ThisAddIn.Application, new EApplication_WindowActivateEventHandler(this.Application_WindowActivate));
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        private void InitializeComponent()
        {
            this.tab1 = base.Factory.CreateRibbonTab();
            this.group1 = base.Factory.CreateRibbonGroup();
            this.SaveAs = base.Factory.CreateRibbonButton();
            this.Save = base.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.tab1.ControlId.ControlIdType = RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            this.group1.Items.Add(this.SaveAs);
            this.group1.Items.Add(this.Save);
            this.group1.Label = "Journaliser";
            this.group1.Name = "group1";
            this.group1.Position = base.Factory.RibbonPosition.AfterOfficeId("GroupClipboard");
            this.SaveAs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SaveAs.Image = Resources.VismaCase32x32;
            this.SaveAs.Label = "SaveAs";
            this.SaveAs.Name = "SaveAs";
            this.SaveAs.ShowImage = true;
            this.SaveAs.Click += new RibbonControlEventHandler(this.SaveAs_Click);
            this.Save.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Save.Image = Resources.VismaCase32x32;
            this.Save.Label = "Save";
            this.Save.Name = "Save";
            this.Save.ShowImage = true;
            this.Save.Click += new RibbonControlEventHandler(this.Save_Click);
            base.Name = "OpenESDHRibbon";
            base.RibbonType = "Microsoft.PowerPoint.Presentation";
            base.Tabs.Add(this.tab1);
            base.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OpenESDHRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
        }

        [DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)]
        private static extern IntPtr LockWindowUpdate(IntPtr hwndLock);
        private void OpenESDHRibbon_Load(object sender, Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs e)
        {
            if (this._presenter == null)
            {
                this.Initialize();
            }
        }

        private void Save_Click(object sender, RibbonControlEventArgs e)
        {
            this._presenter.Save((dynamic) e.Control.Context);
        }

        private void SaveAs_Click(object sender, RibbonControlEventArgs e)
        {
            this._presenter.SaveAs((dynamic) e.Control.Context);
        }

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, bool wParam, int lParam);

        public bool SaveAsEnabled
        {
            get
            {
                return this.SaveAs.Enabled;
            }
            set
            {
                this.SaveAs.Enabled = value;
            }
        }

        public bool SaveEnabled
        {
            get
            {
                return this.Save.Enabled;
            }
            set
            {
                this.Save.Enabled = value;
            }
        }

        public bool ViewIsLocked
        {
            get
            {
                return this._viewIsLocked;
            }
            set
            {
                this._viewIsLocked = value;
                if (value)
                {
                    try
                    {
                        if ((Process.GetCurrentProcess() != null) && (1 != 0))
                        {
                            LockWindowUpdate(Process.GetCurrentProcess().MainWindowHandle);
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    catch
                    {
                    }
                }
                else
                {
                    LockWindowUpdate(IntPtr.Zero);
                    System.Windows.Forms.Application.DoEvents();
                }
            }
        }
    }
}


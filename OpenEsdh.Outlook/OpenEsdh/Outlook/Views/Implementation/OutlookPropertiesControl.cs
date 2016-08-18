namespace OpenEsdh.Outlook.Views.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Views.Implementation.Utilities;
    using OpenEsdh.Outlook.Views.Interface;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class OutlookPropertiesControl : UserControl, IDisplayRegionControl
    {
        private AlfrescoBrowser alfrescoBrowser;
        private IContainer components = null;

        public OutlookPropertiesControl()
        {
            this.InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.alfrescoBrowser = new AlfrescoBrowser();
            base.SuspendLayout();
            this.alfrescoBrowser.Dock = DockStyle.Fill;
            this.alfrescoBrowser.Location = new Point(0, 0);
            this.alfrescoBrowser.Name = "alfrescoBrowser";
            this.alfrescoBrowser.Size = new Size(0x420, 0x1e7);
            this.alfrescoBrowser.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.Controls.Add(this.alfrescoBrowser);
            base.Name = "OutlookPropertiesControl";
            base.Size = new Size(0x420, 0x1e7);
            base.ResumeLayout(false);
        }

        public void Show(string url)
        {
            IOutlookConfiguration configuration = TypeResolver.Current.Create<IOutlookConfiguration>();
            this.alfrescoBrowser.RunRequests(configuration, new Uri(url), "");
        }
    }
}


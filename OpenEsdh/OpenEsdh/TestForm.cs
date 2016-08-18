namespace OpenEsdh
{
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class TestForm : Form
    {
        private IContainer components = null;
        private WebBrowser webBrowser1;

        public TestForm()
        {
            this.InitializeComponent();
            this.webBrowser1.Navigate("http://10.170.12.135:8081/share/page/dp/ws/office-dialog", "_top");
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
            this.webBrowser1 = new WebBrowser();
            base.SuspendLayout();
            this.webBrowser1.Dock = DockStyle.Fill;
            this.webBrowser1.Location = new Point(0, 0);
            this.webBrowser1.MinimumSize = new Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.ScriptErrorsSuppressed = true;
            this.webBrowser1.ScrollBarsEnabled = false;
            this.webBrowser1.Size = new Size(0x480, 0x1a5);
            this.webBrowser1.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x480, 0x1a5);
            base.Controls.Add(this.webBrowser1);
            base.Name = "TestForm";
            this.Text = "TestForm";
            base.ResumeLayout(false);
        }
    }
}


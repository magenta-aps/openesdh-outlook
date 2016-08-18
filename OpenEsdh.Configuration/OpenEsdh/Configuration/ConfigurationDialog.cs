namespace OpenEsdh.Configuration
{
    using Microsoft.Win32;
    using System;
    using System.ComponentModel;
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Web;
    using System.Windows.Forms;
    using System.Xml.Linq;

    public class ConfigurationDialog : Form
    {
        private Button button1;
        private Button button2;
        private CheckBox checkBox1;
        private IContainer components = null;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private GroupBox groupBox4;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label label5;
        private Label label6;
        private RadioButton radioButton1;
        private RadioButton radioButton2;
        private TextBox textBox1;
        private TextBox textBox2;
        private TextBox textBox3;
        private TextBox textBox4;
        private TextBox textBox5;
        private TextBox textBox6;

        public ConfigurationDialog()
        {
            this.InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Exception exception;
            bool flag = false;
            if (this.radioButton2.Checked)
            {
                try
                {
                    this.ModifyAllFiles(true);
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    this.WriteLine("An error occured");
                    this.WriteLine(exception.Message);
                    flag = true;
                }
            }
            else
            {
                try
                {
                    this.GetZipFile();
                    this.ModifyAllFiles(false);
                }
                catch (Exception exception2)
                {
                    exception = exception2;
                    this.WriteLine("An error occured");
                    this.WriteLine(exception.Message);
                    flag = true;
                }
            }
            if (!flag)
            {
                this.button1.Enabled = false;
                this.WriteLine("Done configuring");
                this.button2.Text = "Done";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox1.Checked)
            {
                this.textBox1.Enabled = false;
                this.textBox2.Enabled = false;
                this.textBox3.Enabled = false;
                this.label1.Enabled = false;
                this.label2.Enabled = false;
                this.label3.Enabled = false;
            }
            else
            {
                this.textBox1.Enabled = true;
                this.textBox2.Enabled = true;
                this.textBox3.Enabled = true;
                this.label1.Enabled = true;
                this.label2.Enabled = true;
                this.label3.Enabled = true;
            }
        }

        private void ConfigurationDialog_Load(object sender, EventArgs e)
        {
            this.radioButton2_CheckedChanged(this, new EventArgs());
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        public void GetZipFile()
        {
            string directoryName = "";
            string str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
            if (string.IsNullOrEmpty(str2))
            {
                str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
            }
            if (!string.IsNullOrEmpty(str2))
            {
                directoryName = Path.GetDirectoryName(str2);
            }
            else
            {
                directoryName = Directory.GetCurrentDirectory();
            }
            this.WriteLine("Fetching configurations from " + this.textBox4.Text);
            WebClient client = new WebClient();
            using (MemoryStream stream = new MemoryStream(client.DownloadData(this.textBox4.Text)))
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        this.WriteLine("Extracting " + entry.FullName);
                        string path = Path.Combine(directoryName, entry.FullName);
                        if (System.IO.File.Exists(path))
                        {
                            System.IO.File.Delete(path);
                        }
                        entry.ExtractToFile(path);
                    }
                }
            }
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(ConfigurationDialog));
            this.groupBox1 = new GroupBox();
            this.label6 = new Label();
            this.label3 = new Label();
            this.textBox3 = new TextBox();
            this.label2 = new Label();
            this.textBox2 = new TextBox();
            this.label1 = new Label();
            this.textBox1 = new TextBox();
            this.checkBox1 = new CheckBox();
            this.groupBox2 = new GroupBox();
            this.label5 = new Label();
            this.textBox5 = new TextBox();
            this.label4 = new Label();
            this.textBox4 = new TextBox();
            this.radioButton2 = new RadioButton();
            this.radioButton1 = new RadioButton();
            this.groupBox3 = new GroupBox();
            this.button2 = new Button();
            this.button1 = new Button();
            this.groupBox4 = new GroupBox();
            this.textBox6 = new TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            base.SuspendLayout();
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Location = new Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new Size(500, 0x9b);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Login Information";
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.Red;
            this.label6.Location = new Point(0x162, 0x6d);
            this.label6.Name = "label6";
            this.label6.Size = new Size(0x86, 13);
            this.label6.TabIndex = 7;
            this.label6.Text = "* Passwords must be equal";
            this.label6.Visible = false;
            this.label3.AutoSize = true;
            this.label3.Location = new Point(6, 0x6d);
            this.label3.Name = "label3";
            this.label3.Size = new Size(0x5b, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Confirm Password";
            this.textBox3.Location = new Point(0x6c, 0x6a);
            this.textBox3.Name = "textBox3";
            this.textBox3.PasswordChar = '*';
            this.textBox3.Size = new Size(0xef, 20);
            this.textBox3.TabIndex = 5;
            this.textBox3.TextChanged += new EventHandler(this.textBox2_TextChanged);
            this.label2.AutoSize = true;
            this.label2.Location = new Point(6, 0x53);
            this.label2.Name = "label2";
            this.label2.Size = new Size(0x35, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Password";
            this.textBox2.Location = new Point(0x6c, 80);
            this.textBox2.Name = "textBox2";
            this.textBox2.PasswordChar = '*';
            this.textBox2.Size = new Size(0xef, 20);
            this.textBox2.TabIndex = 3;
            this.textBox2.TextChanged += new EventHandler(this.textBox2_TextChanged);
            this.label1.AutoSize = true;
            this.label1.Location = new Point(6, 0x39);
            this.label1.Name = "label1";
            this.label1.Size = new Size(0x37, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Username";
            this.textBox1.Location = new Point(0x6c, 0x36);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new Size(0xef, 20);
            this.textBox1.TabIndex = 1;
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new Point(9, 0x1f);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new Size(0x8e, 0x11);
            this.checkBox1.TabIndex = 0;
            this.checkBox1.Text = "Integrated Login (NTLM)";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new EventHandler(this.checkBox1_CheckedChanged);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.textBox5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.radioButton2);
            this.groupBox2.Controls.Add(this.radioButton1);
            this.groupBox2.Location = new Point(12, 0xae);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new Size(500, 0x87);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Alfresco Server";
            this.label5.AutoSize = true;
            this.label5.Location = new Point(0x17, 0x61);
            this.label5.Name = "label5";
            this.label5.Size = new Size(0x4a, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "OpenE Server";
            this.textBox5.Location = new Point(0x6c, 0x5e);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new Size(0x182, 20);
            this.textBox5.TabIndex = 8;
            this.textBox5.Text = "http://Your.Alfresco.Server";
            this.label4.AutoSize = true;
            this.label4.Location = new Point(0x17, 0x2e);
            this.label4.Name = "label4";
            this.label4.Size = new Size(0x36, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Server Url";
            this.textBox4.Location = new Point(0x6c, 0x2b);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new Size(0x182, 20);
            this.textBox4.TabIndex = 2;
            this.textBox4.Text = "http://Your.Configuration.Server/Configuration.zip";
            this.radioButton2.AutoSize = true;
            this.radioButton2.Checked = true;
            this.radioButton2.Location = new Point(9, 0x47);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new Size(0x7d, 0x11);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Manual Configuration";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new EventHandler(this.radioButton2_CheckedChanged);
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new Point(9, 0x13);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new Size(190, 0x11);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.Text = "Retrieve Configuration From Server";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new EventHandler(this.radioButton2_CheckedChanged);
            this.groupBox3.Controls.Add(this.button2);
            this.groupBox3.Controls.Add(this.button1);
            this.groupBox3.Location = new Point(12, 0x1af);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new Size(500, 0x2b);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.button2.Location = new Point(0x151, 10);
            this.button2.Name = "button2";
            this.button2.Size = new Size(0x4b, 0x17);
            this.button2.TabIndex = 1;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new EventHandler(this.button2_Click);
            this.button1.Location = new Point(0x1a2, 10);
            this.button1.Name = "button1";
            this.button1.Size = new Size(0x4b, 0x17);
            this.button1.TabIndex = 0;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new EventHandler(this.button1_Click);
            this.groupBox4.Controls.Add(this.textBox6);
            this.groupBox4.Location = new Point(12, 0x13b);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new Size(0x1f3, 0x74);
            this.groupBox4.TabIndex = 3;
            this.groupBox4.TabStop = false;
            this.textBox6.Dock = DockStyle.Fill;
            this.textBox6.Location = new Point(3, 0x10);
            this.textBox6.Multiline = true;
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.ScrollBars = ScrollBars.Vertical;
            this.textBox6.Size = new Size(0x1ed, 0x61);
            this.textBox6.TabIndex = 0;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(520, 0x1e6);
            base.Controls.Add(this.groupBox4);
            base.Controls.Add(this.groupBox3);
            base.Controls.Add(this.groupBox2);
            base.Controls.Add(this.groupBox1);
            base.FormBorderStyle = FormBorderStyle.FixedDialog;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.MaximizeBox = false;
            base.MinimizeBox = false;
            base.Name = "ConfigurationDialog";
            base.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "OpenE Configuration";
            base.Load += new EventHandler(this.ConfigurationDialog_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            base.ResumeLayout(false);
        }

        private void ModifyAllFiles(bool ModifyHost)
        {
            string path = "";
            string str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
            if (string.IsNullOrEmpty(str2))
            {
                str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
            }
            if (!string.IsNullOrEmpty(str2))
            {
                path = Path.GetDirectoryName(str2);
            }
            else
            {
                path = Directory.GetCurrentDirectory();
            }
            if (Directory.Exists(path))
            {
                foreach (string str3 in Directory.GetFiles(path, "*.config"))
                {
                    this.ModifyFile(str3, ModifyHost);
                }
            }
        }

        private void ModifyFile(string filename, bool ModifyHost)
        {
            if (System.IO.File.Exists(filename))
            {
                FileStream stream;
                XDocument document = null;
                using (stream = new FileStream(filename, FileMode.Open))
                {
                    document = XDocument.Load(stream);
                }
                this.WriteLine("Loading " + filename);
                XElement officeElement = document.Root.Element("Office");
                if (officeElement == null)
                {
                    officeElement = document.Root.Element("Outlook");
                }
                if (officeElement != null)
                {
                    XElement element2 = officeElement.Element("PreAuthentication");
                    if (element2 != null)
                    {
                        if (!this.checkBox1.Checked)
                        {
                            this.WriteLine("Setting Username");
                            if (element2.Attribute("Username") == null)
                            {
                                element2.Add(new XAttribute("Username", this.textBox1.Text));
                            }
                            else
                            {
                                element2.Attribute("Username").Value = this.textBox1.Text;
                            }
                            this.WriteLine("Setting Password");
                            if (element2.Attribute("Password") == null)
                            {
                                element2.Add(new XAttribute("Password", this.textBox2.Text));
                            }
                            else
                            {
                                element2.Attribute("Password").Value = this.textBox2.Text;
                            }
                            this.WriteLine("Setting Preauthenticate");
                            if (officeElement.Attribute("PreAuthenticate") == null)
                            {
                                officeElement.Add(new XAttribute("PreAuthenticate", "true"));
                            }
                            else
                            {
                                officeElement.Attribute("PreAuthenticate").Value = "true";
                            }
                            this.WriteLine("Setting UseConfigCredentials");
                            if (element2.Attribute("UseConfigCredentials") == null)
                            {
                                element2.Add(new XAttribute("UseConfigCredentials", "true"));
                            }
                            else
                            {
                                element2.Attribute("UseConfigCredentials").Value = "true";
                            }
                        }
                        else
                        {
                            this.WriteLine("Setting UseConfigCredentials");
                            if (element2.Attribute("UseConfigCredentials") == null)
                            {
                                element2.Add(new XAttribute("UseConfigCredentials", "false"));
                            }
                            else
                            {
                                element2.Attribute("UseConfigCredentials").Value = "false";
                            }
                            this.WriteLine("Setting Preauthenticate");
                            if (officeElement.Attribute("PreAuthenticate") == null)
                            {
                                officeElement.Add(new XAttribute("PreAuthenticate", "false"));
                            }
                            else
                            {
                                officeElement.Attribute("PreAuthenticate").Value = "false";
                            }
                        }
                        if (ModifyHost)
                        {
                            this.SetUrlHost(officeElement, "SaveAsDialogUrl");
                            this.SetUrlHost(officeElement, "SaveDialogUrl");
                            this.SetUrlHost(officeElement, "UploadEndPoint");
                            this.SetUrlHost(officeElement, "EndUploadEndpoint");
                            this.SetUrlHost(element2, "AuthenticationUrl");
                            XElement element3 = officeElement.Element("DisplayRegion");
                            if (element3 != null)
                            {
                                this.SetUrlHost(element3, "DisplayDialogUrl");
                            }
                        }
                    }
                }
                using (stream = new FileStream(filename, FileMode.Truncate))
                {
                    document.Save(stream);
                }
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radioButton2.Checked)
            {
                this.label4.Enabled = false;
                this.textBox4.Enabled = false;
                this.label5.Enabled = true;
                this.textBox5.Enabled = true;
            }
            else
            {
                this.label4.Enabled = true;
                this.textBox4.Enabled = true;
                this.label5.Enabled = false;
                this.textBox5.Enabled = false;
            }
        }

        private void SetUrlHost(XElement officeElement, string p)
        {
            if (officeElement.Attribute(p) != null)
            {
                this.WriteLine("Setting " + p + " to server " + this.textBox5.Text);
                string uriString = officeElement.Attribute(p).Value;
                Uri uri = new Uri(uriString);
                string str2 = "";
                if (uriString.Contains("#"))
                {
                    str2 = uriString.Substring(uriString.IndexOf("#"));
                }
                string text = this.textBox5.Text;
                if (!text.EndsWith("/"))
                {
                    text = text + "/";
                }
                for (int i = 1; i < uri.Segments.Length; i++)
                {
                    text = text + HttpUtility.UrlDecode(uri.Segments[i]);
                }
                if (!string.IsNullOrEmpty(str2))
                {
                    if (!(text.EndsWith("/") || str2.StartsWith("/")))
                    {
                        str2 = "/" + str2;
                    }
                    text = text + str2;
                }
                officeElement.Attribute(p).Value = text;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (!this.checkBox1.Checked)
            {
                if (this.textBox2.Text != this.textBox3.Text)
                {
                    this.label6.Visible = true;
                    this.button1.Enabled = false;
                }
                else
                {
                    this.label6.Visible = false;
                    this.button1.Enabled = true;
                }
            }
        }

        private void WriteLine(string s)
        {
            this.textBox6.Text = this.textBox6.Text + "\r\n" + s;
            this.textBox6.SelectionStart = this.textBox6.TextLength;
            this.textBox6.ScrollToCaret();
        }
    }
}


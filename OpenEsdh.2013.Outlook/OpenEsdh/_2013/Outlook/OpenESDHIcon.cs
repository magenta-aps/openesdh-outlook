namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Outlook;
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Threading;
    using System.Windows.Forms;

    [ToolboxItem(false)]
    internal class OpenESDHIcon : FormRegionBase
    {
        private IContainer components;

        public OpenESDHIcon(Microsoft.Office.Interop.Outlook.FormRegion formRegion) : base((Microsoft.Office.Tools.Outlook.Factory) Globals.Factory, (Microsoft.Office.Interop.Outlook.FormRegion) formRegion)
        {
            this.components = null;
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
            base.SuspendLayout();
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.Name = "OpenESDHIcon";
            base.FormRegionShowing += new EventHandler(this.OpenESDHIcon_FormRegionShowing);
            base.FormRegionClosed += new EventHandler(this.OpenESDHIcon_FormRegionClosed);
            base.ResumeLayout(false);
        }

        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(OpenESDHIcon));
            manifest.ExactMessageClass = true;
            manifest.FormRegionName = "OpenESDHIcon";
            manifest.FormRegionType = FormRegionType.Replacement;
            manifest.Hidden = true;
            manifest.Icons.Default = (Icon) manager.GetObject("OpenESDHIcon.Manifest.Icons.Default");
            manifest.LoadLegacyForm = true;
            manifest.ShowInspectorCompose = false;
            manifest.ShowInspectorRead = false;
            manifest.ShowReadingPane = false;
            manifest.Title = "OpenESDHIcon";
        }

        private void OpenESDHIcon_FormRegionClosed(object sender, EventArgs e)
        {
        }

        private void OpenESDHIcon_FormRegionShowing(object sender, EventArgs e)
        {
        }

        [FormRegionMessageClass("IPM.Note.OpenESDH"), FormRegionName("OpenEsdh.2013.Outlook.OpenESDHIcon")]
        public class OpenESDHIconFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest = Globals.Factory.CreateFormRegionManifest();

            public event FormRegionInitializingEventHandler FormRegionInitializing;

            [DebuggerNonUserCode]
            public OpenESDHIconFactory()
            {
                OpenESDHIcon.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new FormRegionInitializingEventHandler(this.OpenESDHIconFactory_FormRegionInitializing);
            }

            [DebuggerNonUserCode]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                OpenESDHIcon icon = new OpenESDHIcon(formRegion) {
                    Factory = (Microsoft.Office.Tools.Outlook.IFormRegionFactory) this
                };
                return (Microsoft.Office.Tools.Outlook.IFormRegion) icon;
            }

            [DebuggerNonUserCode]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new NotSupportedException();
            }

            [DebuggerNonUserCode]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    FormRegionInitializingEventArgs e = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, (Microsoft.Office.Interop.Outlook.OlFormRegionMode) formRegionMode, (Microsoft.Office.Interop.Outlook.OlFormRegionSize) formRegionSize, false);
                    this.FormRegionInitializing(this, e);
                    return !e.Cancel;
                }
                return true;
            }

            private void OpenESDHIconFactory_FormRegionInitializing(object sender, FormRegionInitializingEventArgs e)
            {
            }

            [DebuggerNonUserCode]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [DebuggerNonUserCode]
            FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }
}


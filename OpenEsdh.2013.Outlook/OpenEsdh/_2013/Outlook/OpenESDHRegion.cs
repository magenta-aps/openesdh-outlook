namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools.Outlook;
    using OpenEsdh._2013.Outlook.Model;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Interface;
    using System;
    using System.Collections;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Threading;
    using System.Windows.Forms;

    [ToolboxItem(false)]
    internal class OpenESDHRegion : FormRegionBase, IDisplayRegion
    {
        private IDisplayRegionPresenter _presenter;
        private IContainer components;

        public OpenESDHRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion) : base((Microsoft.Office.Tools.Outlook.Factory) Globals.Factory, (Microsoft.Office.Interop.Outlook.FormRegion) formRegion)
        {
            this._presenter = null;
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
            base.Name = "OpenESDHRegion";
            base.Size = new Size(0x57b, 0xee);
            base.FormRegionShowing += new EventHandler(this.OpenESDHRegion_FormRegionShowing);
            base.FormRegionClosed += new EventHandler(this.OpenESDHRegion_FormRegionClosed);
            base.Load += new EventHandler(this.OpenESDHRegion_Load);
            base.ResumeLayout(false);
        }

        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(OpenESDHRegion));
            manifest.FormRegionName = "Alfresco Integration";
            manifest.FormRegionType = FormRegionType.Adjoining;
            manifest.Icons.Default = (Icon) manager.GetObject("OpenESDHRegion.Manifest.Icons.Default");
            manifest.LoadLegacyForm = true;
        }

        private void OpenESDHRegion_FormRegionClosed(object sender, EventArgs e)
        {
        }

        private void OpenESDHRegion_FormRegionShowing(object sender, EventArgs e)
        {
            this._presenter = TypeResolver.Current.Create<IDisplayRegionPresenter>(this);
            MailItem outlookItem = base.OutlookItem as MailItem;
            if (outlookItem != null)
            {
                this._presenter.Show(outlookItem.ToMailDescriptor());
            }
        }

        private void OpenESDHRegion_Load(object sender, EventArgs e)
        {
        }

        public IList FormControlCollection
        {
            get
            {
                return base.Controls;
            }
        }

        [FormRegionMessageClass("IPM.Note.OpenESDH"), FormRegionName("OpenEsdh.2013.Outlook.OpenESDHRegion")]
        public class OpenESDHRegionFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest = Globals.Factory.CreateFormRegionManifest();

            public event FormRegionInitializingEventHandler FormRegionInitializing;

            [DebuggerNonUserCode]
            public OpenESDHRegionFactory()
            {
                OpenESDHRegion.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new FormRegionInitializingEventHandler(this.OpenESDHRegionFactory_FormRegionInitializing);
                try
                {
                    this._Manifest.FormRegionName = ResourceResolver.Current.GetString("ViewRegionTitle");
                }
                catch
                {
                }
            }

            [DebuggerNonUserCode]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                OpenESDHRegion region = new OpenESDHRegion(formRegion) {
                    Factory = (Microsoft.Office.Tools.Outlook.IFormRegionFactory) this
                };
                return (Microsoft.Office.Tools.Outlook.IFormRegion) region;
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

            private void OpenESDHRegionFactory_FormRegionInitializing(object sender, FormRegionInitializingEventArgs e)
            {
                try
                {
                    this.Manifest.FormRegionName = ResourceResolver.Current.GetString("ViewRegionTitle");
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                }
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


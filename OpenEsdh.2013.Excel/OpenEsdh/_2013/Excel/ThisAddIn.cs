namespace OpenEsdh._2013.Excel
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using Microsoft.Office.Tools.Excel;
    using Microsoft.VisualStudio.Tools.Applications.Runtime;
    using OpenEsdh._2013.Excel.Presentation.Implementation;
    using OpenEsdh._2013.Excel.Presentation.Interface;
    using OpenEsdh.Outlook.Model.BrowserVersion;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Security.Permissions;
    using System.Windows.Forms;

    [StartupObject(0), PermissionSet(SecurityAction.Demand, Name="FullTrust")]
    public sealed class ThisAddIn : AddInBase
    {
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        internal Microsoft.Office.Interop.Excel.Application Application;
        internal CustomTaskPaneCollection CustomTaskPanes;
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private object missing;
        internal SmartTagCollection VstoSmartTags;

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider) : base((Microsoft.Office.Tools.Factory) factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            this.missing = System.Type.Missing;
            Globals.Factory = factory;
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void BeginInitialization()
        {
            this.BeginInit();
            this.CustomTaskPanes.BeginInit();
            this.VstoSmartTags.BeginInit();
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        private void BindToData()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        private void EndInitialization()
        {
            this.VstoSmartTags.EndInit();
            this.CustomTaskPanes.EndInit();
            this.EndInit();
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never)]
        protected override void Initialize()
        {
            base.Initialize();
            this.Application = base.GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), "Application");
            Globals.ThisAddIn = this;
            System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void InitializeCachedData()
        {
            if ((base.DataHost != null) && base.DataHost.IsCacheInitialized)
            {
                base.DataHost.FillCachedData(this);
            }
        }

        [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeComponents()
        {
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            this.VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        private void InitializeData()
        {
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        protected override void InitializeDataBindings()
        {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }

        private void InternalStartup()
        {
            base.Startup += new EventHandler(this.ThisAddIn_Startup);
            base.Shutdown += new EventHandler(this.ThisAddIn_Shutdown);
        }

        [EditorBrowsable(EditorBrowsableState.Advanced), DebuggerNonUserCode]
        private bool IsCached(string MemberName)
        {
            return base.DataHost.IsCached(this, MemberName);
        }

        [EditorBrowsable(EditorBrowsableState.Advanced), DebuggerNonUserCode]
        private bool NeedsFill(string MemberName)
        {
            return base.DataHost.NeedsFill(this, MemberName);
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        protected override void OnShutdown()
        {
            this.VstoSmartTags.Dispose();
            this.CustomTaskPanes.Dispose();
            base.OnShutdown();
        }

        [EditorBrowsable(EditorBrowsableState.Advanced), DebuggerNonUserCode]
        private void StartCaching(string MemberName)
        {
            base.DataHost.StartCaching(this, MemberName);
        }

        [EditorBrowsable(EditorBrowsableState.Advanced), DebuggerNonUserCode]
        private void StopCaching(string MemberName)
        {
            base.DataHost.StopCaching(this, MemberName);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (TypeResolver.Current != null)
            {
                TypeResolver.Current.Dispose();
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            InternetExplorerBrowserEmulation.SetBrowserEmulationVersion(BrowserEmulationVersion.Version11Edge);
            Logger.Current.LogInformation("Application Startup", "");
            TypeResolver.Current = new WordResolver(typeof(ThisAddIn));
            TypeResolver.Current.AddComponentWithParam<IExcelPresenter>(delegate (object view) {
                IExcelView view2 = view as IExcelView;
                if (view2 != null)
                {
                    return new ExcelPresenter(view2);
                }
                return null;
            });
            Globals.Ribbons.OpenESDHRibbon.Initialize();
        }
    }
}


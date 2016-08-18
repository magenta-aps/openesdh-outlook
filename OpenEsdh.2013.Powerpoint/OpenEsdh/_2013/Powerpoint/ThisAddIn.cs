namespace OpenEsdh._2013.Powerpoint
{
    using Microsoft.Office.Interop.PowerPoint;
    using Microsoft.Office.Tools;
    using Microsoft.VisualStudio.Tools.Applications.Runtime;
    using OpenEsdh._2013.Powerpoint.Presentation.Implementation;
    using OpenEsdh._2013.Powerpoint.Presentation.Interface;
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
        internal Microsoft.Office.Interop.PowerPoint.Application Application;
        internal CustomTaskPaneCollection CustomTaskPanes;
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private object missing;

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        public ThisAddIn(Microsoft.Office.Tools.Factory factory, IServiceProvider serviceProvider) : base((Microsoft.Office.Tools.Factory) factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            this.missing = System.Type.Missing;
            Globals.Factory = factory;
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        private void BeginInitialization()
        {
            this.BeginInit();
            this.CustomTaskPanes.BeginInit();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        private void BindToData()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        private void EndInitialization()
        {
            this.CustomTaskPanes.EndInit();
            this.EndInit();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        protected override void Initialize()
        {
            base.Initialize();
            this.Application = base.GetHostItem<Microsoft.Office.Interop.PowerPoint.Application>(typeof(Microsoft.Office.Interop.PowerPoint.Application), "Application");
            Globals.ThisAddIn = this;
            System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }

        [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeCachedData()
        {
            if ((base.DataHost != null) && base.DataHost.IsCacheInitialized)
            {
                base.DataHost.FillCachedData(this);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        private void InitializeComponents()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void InitializeData()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
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

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName)
        {
            return base.DataHost.NeedsFill(this, MemberName);
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        protected override void OnShutdown()
        {
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
            TypeResolver.Current.AddComponentWithParam<IPowerpointPresenter>(delegate (object view) {
                IPowerpointView view2 = view as IPowerpointView;
                if (view2 != null)
                {
                    return new PowerpointPresenter(view2);
                }
                return null;
            });
            Globals.Ribbons.OpenESDHRibbon.Initialize();
        }
    }
}


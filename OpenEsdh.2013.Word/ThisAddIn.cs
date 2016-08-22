namespace OpenEsdh._2013.Word
{
    using Microsoft.Office.Interop.Word;
    using Microsoft.Office.Tools;
    using Microsoft.Office.Tools.Word;
    using Microsoft.VisualStudio.Tools.Applications.Runtime;
    using OpenEsdh._2013.Word.Presentation.Implementation;
    using OpenEsdh._2013.Word.Presentation.Interface;
    using OpenEsdh.Outlook.Model.BrowserVersion;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Security.Permissions;
    using forms = System.Windows.Forms;
    using sys = System;

    [StartupObject(0), PermissionSet(SecurityAction.Demand, Name="FullTrust")]
    public sealed class ThisAddIn : AddInBase
    {
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        internal Microsoft.Office.Interop.Word.Application Application;
        internal CustomTaskPaneCollection CustomTaskPanes;
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private object missing;
        internal SmartTagCollection VstoSmartTags;

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        public ThisAddIn(ApplicationFactory factory, IServiceProvider serviceProvider) : base((Microsoft.Office.Tools.Factory) factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            sys.Diagnostics.Debugger.Launch();
            this.missing = Type.Missing;
            Globals.Factory = factory;
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        private void BeginInitialization()
        {
            this.BeginInit();
            this.CustomTaskPanes.BeginInit();
            this.VstoSmartTags.BeginInit();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        private void BindToData()
        {
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void EndInitialization()
        {
            this.VstoSmartTags.EndInit();
            this.CustomTaskPanes.EndInit();
            this.EndInit();
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        protected override void Initialize()
        {
            base.Initialize();
            this.Application = base.GetHostItem<Microsoft.Office.Interop.Word.Application>(typeof(Microsoft.Office.Interop.Word.Application), "Application");
            Globals.ThisAddIn = this;
            forms.Application.EnableVisualStyles();
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

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void InitializeComponents()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            this.VstoSmartTags = Globals.Factory.CreateSmartTagCollection(null, null, "VstoSmartTags", "VstoSmartTags", this);
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeData()
        {
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
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

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName)
        {
            return base.DataHost.IsCached(this, MemberName);
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName)
        {
            return base.DataHost.NeedsFill(this, MemberName);
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
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
            TypeResolver.Current.AddComponentWithParam<IWordPresenter>(delegate (object view) {
                IWordView view2 = view as IWordView;
                if (view2 != null)
                {
                    return new WordPresenter(view2);
                }
                return null;
            });
            Globals.Ribbons.OpenESDHRibbon.Initialize();
        }
    }
}


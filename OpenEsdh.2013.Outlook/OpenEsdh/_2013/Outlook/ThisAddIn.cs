namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.Office.Tools;
    using Microsoft.Office.Tools.Outlook;
    using Microsoft.VisualStudio.Tools.Applications.Runtime;
    using OpenEsdh._2013.Outlook.Model;
    using OpenEsdh._2013.Outlook.Presentation.Implementation;
    using OpenEsdh.Outlook.Model.BrowserVersion;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.CodeDom.Compiler;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Security.Permissions;
    using System.ServiceModel;
    using System.ServiceModel.Description;
    using System.Windows.Forms;

    [StartupObject(0), PermissionSet(SecurityAction.Demand, Name="FullTrust")]
    public sealed class ThisAddIn : OutlookAddInBase
    {
        private Explorers _explorers;
        private ServiceHost _host;
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        internal Microsoft.Office.Interop.Outlook.Application Application;
        internal CustomTaskPaneCollection CustomTaskPanes;
        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private object missing;

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        public ThisAddIn(Microsoft.Office.Tools.Outlook.Factory factory, IServiceProvider serviceProvider) : base((Microsoft.Office.Tools.Outlook.Factory) factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            this._explorers = null;
            this._host = null;
            this.missing = System.Type.Missing;
            Globals.Factory = factory;
        }

        private void _explorers_NewExplorer(Microsoft.Office.Interop.Outlook.Explorer Explorer)
        {
            try
            {
                if (Explorer != null)
                {
                    try
                    {
                        new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "SelectionChange").RemoveEventHandler(Explorer, new ExplorerEvents_10_SelectionChangeEventHandler(this.ThisAddIn_SelectionChange));
                        new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "FolderSwitch").RemoveEventHandler(Explorer, new ExplorerEvents_10_FolderSwitchEventHandler(this.ThisAddIn_FolderSwitch));
                    }
                    catch
                    {
                    }
                    new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "SelectionChange").AddEventHandler(Explorer, new ExplorerEvents_10_SelectionChangeEventHandler(this.ThisAddIn_SelectionChange));
                    new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "FolderSwitch").AddEventHandler(Explorer, new ExplorerEvents_10_FolderSwitchEventHandler(this.ThisAddIn_FolderSwitch));
                }
            }
            catch
            {
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void BeginInitialization()
        {
            this.BeginInit();
            this.CustomTaskPanes.BeginInit();
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        private void BindToData()
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OpenEsdh._2013.Outlook.ContextMenu();
        }

        [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never)]
        private void EndInitialization()
        {
            this.CustomTaskPanes.EndInit();
            this.EndInit();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode]
        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode]
        protected override void Initialize()
        {
            base.Initialize();
            this.Application = base.GetHostItem<Microsoft.Office.Interop.Outlook.Application>(typeof(Microsoft.Office.Interop.Outlook.Application), "Application");
            Globals.ThisAddIn = this;
            System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }

        [GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Never)]
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

        [DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0"), EditorBrowsable(EditorBrowsableState.Never)]
        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
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

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "12.0.0.0")]
        protected override void OnShutdown()
        {
            this.CustomTaskPanes.Dispose();
            base.OnShutdown();
        }

        private void SetVisibility()
        {
            string message = "";
            ThisRibbonCollection ribbons = Globals.Ribbons[Globals.ThisAddIn.Application.ActiveExplorer()];
            if ((ribbons != null) && (ribbons.OpenESDHRibbon != null))
            {
                try
                {
                    if (((this.Application.ActiveExplorer().CurrentFolder.DefaultItemType == OlItemType.olMailItem) && (this.Application.ActiveExplorer().CurrentFolder.Items.Count > 0)) || (this.Application.ActiveExplorer().CurrentFolder.Store.ExchangeStoreType == OlExchangeStoreType.olExchangePublicFolder))
                    {
                    }
                }
                catch (Exception exception)
                {
                    message = exception.Message;
                }
            }
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName)
        {
            base.DataHost.StartCaching(this, MemberName);
        }

        [DebuggerNonUserCode, EditorBrowsable(EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName)
        {
            base.DataHost.StopCaching(this, MemberName);
        }

        private void ThisAddIn_FolderSwitch()
        {
            this.SetVisibility();
        }

        private void ThisAddIn_SelectionChange()
        {
            this.SetVisibility();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (this._host != null)
            {
                try
                {
                    this._host.Close();
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                }
            }
            Logger.Current.LogInformation("Application Shutdown", "");
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            InternetExplorerBrowserEmulation.SetBrowserEmulationVersion(BrowserEmulationVersion.Version11Edge);
            Logger.Current.LogInformation("Application Startup", "");
            TypeResolver.Current = new OutlookResolver(typeof(ThisAddIn));
            try
            {
                this._explorers = this.Application.Explorers;
                new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorersEvents_Event), "NewExplorer").AddEventHandler(this._explorers, new ExplorersEvents_NewExplorerEventHandler(this._explorers_NewExplorer));
                if (this.Application.ActiveExplorer() != null)
                {
                    new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "SelectionChange").AddEventHandler(this.Application.ActiveExplorer(), new ExplorerEvents_10_SelectionChangeEventHandler(this.ThisAddIn_SelectionChange));
                    new ComAwareEventInfo(typeof(Microsoft.Office.Interop.Outlook.ExplorerEvents_10_Event), "FolderSwitch").AddEventHandler(this.Application.ActiveExplorer(), new ExplorerEvents_10_FolderSwitchEventHandler(this.ThisAddIn_FolderSwitch));
                }
            }
            catch
            {
            }
            TypeResolver.Current.AddComponent<ISaveEmailPresenter>(() => new SaveEmailPresenter());
            TypeResolver.Current.AddComponent<IAttachFilePresenter>(() => new AttachFilePresenter());
            try
            {
                Uri uri = new Uri("http://localhost:8086/Attach");
                this._host = new ServiceHost(typeof(AttachmentService), new Uri[] { uri });
                ServiceMetadataBehavior item = new ServiceMetadataBehavior {
                    HttpGetEnabled = true,
                    MetadataExporter = { PolicyVersion = PolicyVersion.Policy15 }
                };
                this._host.Description.Behaviors.Add(item);
                this._host.Open();
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                this._host = null;
            }
        }
    }
}


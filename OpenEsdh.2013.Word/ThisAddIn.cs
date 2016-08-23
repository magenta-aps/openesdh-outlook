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

    [PermissionSet(SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class ThisAddIn
    {
        private void InternalStartup()
        {
            base.Startup += new EventHandler(this.ThisAddIn_Startup);
            base.Shutdown += new EventHandler(this.ThisAddIn_Shutdown);
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


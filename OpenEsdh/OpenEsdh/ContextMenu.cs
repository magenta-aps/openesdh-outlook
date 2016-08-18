namespace OpenEsdh
{
    using Microsoft.Win32;
    using OpenEsdh.Outlook.Model.BrowserVersion;
    using OpenEsdh.Outlook.Model.Configuration.Implementation;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Presentation.Implementation;
    using OpenEsdh.Properties;
    using SharpShell.Attributes;
    using SharpShell.SharpContextMenu;
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using System.Xml.Linq;

    [ComVisible(true), COMServerAssociation(AssociationType.AllFiles, new string[] {  })]
    public class ContextMenu : SharpShell.SharpContextMenu.SharpContextMenu
    {
        public ContextMenu()
        {
            Logger.Current.LogInformation("Application Startup", "");
            TypeResolver.Current = new WordResolver(typeof(OpenEsdh.ContextMenu));
            TypeResolver.Current.AddComponent<IExplorerPresenter>(() => new ExplorerPresenter());
            TypeResolver.Current.Replace<IWordConfiguration>(delegate {
                try
                {
                    if (TypeResolver.Current.Singletons.ContainsKey(typeof(IWordConfiguration)))
                    {
                        return TypeResolver.Current.Singletons[typeof(IWordConfiguration)];
                    }
                    string localPath = new Uri(Assembly.GetAssembly(typeof(OpenEsdh.ContextMenu)).CodeBase).LocalPath;
                    WordConfiguration configuration = new WordConfiguration();
                    if (string.IsNullOrEmpty(localPath) || !File.Exists(localPath + ".config"))
                    {
                        string str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
                        if (string.IsNullOrEmpty(str2))
                        {
                            str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
                        }
                        if (!string.IsNullOrEmpty(str2))
                        {
                            localPath = str2;
                        }
                    }
                    else
                    {
                        localPath = localPath + ".config";
                    }
                    if (!string.IsNullOrEmpty(localPath) && File.Exists(localPath))
                    {
                        using (FileStream stream = new FileStream(localPath, FileMode.Open, FileAccess.Read))
                        {
                            XElement element = XDocument.Load(stream).Root.Element("Office");
                            if (element != null)
                            {
                                configuration.SaveAsDialogUrl = (element.Attribute("SaveAsDialogUrl") != null) ? element.Attribute("SaveAsDialogUrl").Value : "";
                                configuration.SaveDialogUrl = (element.Attribute("SaveDialogUrl") != null) ? element.Attribute("SaveDialogUrl").Value : "";
                                configuration.UploadEndPoint = (element.Attribute("UploadEndPoint") != null) ? element.Attribute("UploadEndPoint").Value : "";
                                configuration.PreAuthenticate = (element.Attribute("PreAuthenticate") != null) ? bool.Parse(element.Attribute("PreAuthenticate").Value) : false;
                                configuration.MaxRedirectRetries = (element.Attribute("MaxRedirectRetries") != null) ? int.Parse(element.Attribute("MaxRedirectRetries").Value) : 0;
                                configuration.IgnoreCertificateErrors = (element.Attribute("IgnoreCertificateErrors") != null) ? bool.Parse(element.Attribute("IgnoreCertificateErrors").Value) : false;
                                configuration.EndUploadEndpoint = (element.Attribute("EndUploadEndpoint") != null) ? element.Attribute("EndUploadEndpoint").Value : "";
                                configuration.PreAuthentication_internal = new PreAuthenticationConfiguration();
                                if (element.Element("PreAuthentication") != null)
                                {
                                    XElement element2 = element.Element("PreAuthentication");
                                    configuration.PreAuthentication_internal.Username = (element2.Attribute("Username") != null) ? element2.Attribute("Username").Value : "";
                                    configuration.PreAuthentication_internal.Password = (element2.Attribute("Password") != null) ? element2.Attribute("Password").Value : "";
                                    configuration.PreAuthentication_internal.Domain = (element2.Attribute("Domain") != null) ? element2.Attribute("Domain").Value : "";
                                    configuration.PreAuthentication_internal.PreAuthenticateParameterName = (element2.Attribute("PreAuthenticateParameterName") != null) ? element2.Attribute("PreAuthenticateParameterName").Value : "";
                                    configuration.PreAuthentication_internal.UseConfigCredentials = (element2.Attribute("UseConfigCredentials") != null) ? bool.Parse(element2.Attribute("UseConfigCredentials").Value) : false;
                                    configuration.PreAuthentication_internal.AuthenticationUrl = (element2.Attribute("AuthenticationUrl") != null) ? element2.Attribute("AuthenticationUrl").Value : "";
                                    configuration.PreAuthentication_internal.AuthenticationPackageFormat = (element2.Attribute("AuthenticationPackageFormat") != null) ? element2.Attribute("AuthenticationPackageFormat").Value : "";
                                    configuration.PreAuthentication_internal.AdditionalRequestHeaders = (element2.Attribute("AdditionalRequestHeaders") != null) ? element2.Attribute("AdditionalRequestHeaders").Value : "";
                                }
                                configuration.DialogExtend_Internal = new ExtendDialog();
                                if (element.Element("DialogExtend") != null)
                                {
                                    XElement element3 = element.Element("DialogExtend");
                                    configuration.DialogExtend_Internal.X = (element3.Attribute("Y") != null) ? int.Parse(element3.Attribute("X").Value) : 100;
                                    configuration.DialogExtend_Internal.Y = (element3.Attribute("Y") != null) ? int.Parse(element3.Attribute("Y").Value) : 100;
                                    configuration.DialogExtend_Internal.MaxHeight = (element3.Attribute("MaxHeight") != null) ? int.Parse(element3.Attribute("MaxHeight").Value) : 800;
                                    configuration.DialogExtend_Internal.MaxWidth = (element3.Attribute("MaxWidth") != null) ? int.Parse(element3.Attribute("MaxWidth").Value) : 0x4b0;
                                }
                                configuration.CommunicationConfiguration_Internal = new CommunicationConfiguration();
                                if (element.Element("CommunicationConfiguration") != null)
                                {
                                    XElement element4 = element.Element("CommunicationConfiguration");
                                    configuration.CommunicationConfiguration_Internal.SendMethod_Internal = (element4.Attribute("SendMethod") != null) ? element4.Attribute("SendMethod").Value : "";
                                    configuration.CommunicationConfiguration_Internal.JavaScriptMethodName = (element4.Attribute("JavaScriptMethodName") != null) ? element4.Attribute("JavaScriptMethodName").Value : "";
                                    configuration.CommunicationConfiguration_Internal.PostMethodName = (element4.Attribute("PostMethodName") != null) ? element4.Attribute("PostMethodName").Value : "";
                                }
                            }
                            TypeResolver.Current.Singletons.Add(typeof(IWordConfiguration), configuration);
                        }
                    }
                    return configuration;
                }
                catch (Exception)
                {
                    return new WordConfiguration();
                }
            });
        }

        protected override bool CanShowMenu()
        {
            return true;
        }

        public void ClickElement(string p)
        {
            try
            {
                TypeResolver.Current.Create<IExplorerPresenter>().SaveAs(p);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        protected override ContextMenuStrip CreateMenu()
        {
            ContextMenuStrip strip = new ContextMenuStrip();
            ToolStripMenuItem item = new ToolStripMenuItem {
                Text = "Gem i Visma Case",
                Image = Resources.VismaCase16x16
            };
            item.Click += delegate (object sender, EventArgs args) {
                this.SaveToOpenEsdh();
            };
            strip.Items.Add(item);
            return strip;
        }

        private void SaveToOpenEsdh()
        {
            InternetExplorerBrowserEmulation.SetBrowserEmulationVersion(BrowserEmulationVersion.Version11Edge);
            try
            {
                if ((base.SelectedItemPaths != null) && (base.SelectedItemPaths.Count<string>() > 0))
                {
                    TypeResolver.Current.Create<IExplorerPresenter>().SaveAs(base.SelectedItemPaths.First<string>());
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }
    }
}


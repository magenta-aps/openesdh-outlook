namespace OpenEsdh
{
    using Microsoft.Win32;
    using OpenEsdh.Outlook.Model.BrowserVersion;
    using OpenEsdh.Outlook.Model.Logging;
    using SharpShell.Diagnostics;
    using SharpShell.ServerRegistration;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Configuration.Install;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;

    [RunInstaller(true)]
    public class ExplorerInstaller : Installer
    {
        private IContainer components = null;
        private const bool InstallExplorerIntegration = true;

        public ExplorerInstaller()
        {
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

        private void ExplorerInstaller_Committed(object sender, InstallEventArgs e)
        {
            string path = "";
            string str2 = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Visma Consulting\ExplorerIntegration", "ConfigFile", "") as string;
            if (!string.IsNullOrEmpty(str2))
            {
                path = Path.Combine(Path.GetDirectoryName(str2), "OpenEsdh.Configuration.exe");
            }
            if (File.Exists(path))
            {
                SetForegroundWindow(Process.Start(path).MainWindowHandle);
            }
            Logger.Current.LogInformation("LocalPath:" + path, "");
        }

        private void InitializeComponent()
        {
            base.Committed += new InstallEventHandler(this.ExplorerInstaller_Committed);
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            try
            {
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("Word.exe", BrowserEmulationVersion.Version11Edge);
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("Excel.exe", BrowserEmulationVersion.Version11Edge);
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("Powerpnt.exe", BrowserEmulationVersion.Version11Edge);
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("Outlook.exe", BrowserEmulationVersion.Version11Edge);
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("Explorer.exe", BrowserEmulationVersion.Version11Edge);
                InternetExplorerBrowserEmulation.SetBrowserEmulationVersion("ServerManager.exe", BrowserEmulationVersion.Version11Edge);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        protected override void OnAfterInstall(IDictionary savedState)
        {
            base.OnAfterInstall(savedState);
            try
            {
                string location = "";
                if (string.IsNullOrEmpty(location))
                {
                    location = Assembly.GetExecutingAssembly().Location;
                }
                IEnumerable<ServerEntry> enumerable = ServerManagerApi.LoadServers(location);
                using (IEnumerator<ServerEntry> enumerator = enumerable.GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        SafeFunc func = null;
                        SafeFunc func2 = null;
                        SafeFunc func3 = null;
                        SafeFunc func4 = null;
                        ServerEntry entry = enumerator.Current;
                        if (func == null)
                        {
                            func = () => ServerRegistrationManager.InstallServer(entry.Server, RegistrationType.OS32Bit, true);
                        }
                        this.SafeInitialize(func);
                        if (func2 == null)
                        {
                            func2 = () => ServerRegistrationManager.InstallServer(entry.Server, RegistrationType.OS64Bit, true);
                        }
                        this.SafeInitialize(func2);
                        if (func3 == null)
                        {
                            func3 = () => ServerRegistrationManager.RegisterServer(entry.Server, RegistrationType.OS32Bit);
                        }
                        this.SafeInitialize(func3);
                        if (func4 == null)
                        {
                            func4 = () => ServerRegistrationManager.RegisterServer(entry.Server, RegistrationType.OS64Bit);
                        }
                        this.SafeInitialize(func4);
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        protected override void OnBeforeUninstall(IDictionary savedState)
        {
            base.OnBeforeUninstall(savedState);
            try
            {
                string str = "";
                if (string.IsNullOrEmpty(str))
                {
                    str = str = Assembly.GetExecutingAssembly().Location;
                }
                IEnumerable<ServerEntry> enumerable = ServerManagerApi.LoadServers(str);
                using (IEnumerator<ServerEntry> enumerator = enumerable.GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        SafeFunc func = null;
                        SafeFunc func2 = null;
                        SafeFunc func3 = null;
                        SafeFunc func4 = null;
                        ServerEntry entry = enumerator.Current;
                        if (func == null)
                        {
                            func = () => ServerRegistrationManager.UnregisterServer(entry.Server, RegistrationType.OS32Bit);
                        }
                        this.SafeInitialize(func);
                        if (func2 == null)
                        {
                            func2 = () => ServerRegistrationManager.UnregisterServer(entry.Server, RegistrationType.OS64Bit);
                        }
                        this.SafeInitialize(func2);
                        if (func3 == null)
                        {
                            func3 = () => ServerRegistrationManager.UninstallServer(entry.Server, RegistrationType.OS32Bit);
                        }
                        this.SafeInitialize(func3);
                        if (func4 == null)
                        {
                            func4 = () => ServerRegistrationManager.UninstallServer(entry.Server, RegistrationType.OS64Bit);
                        }
                        this.SafeInitialize(func4);
                        this.SafeInitialize(() => ExplorerManager.RestartExplorer());
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        private void SafeInitialize(SafeFunc func)
        {
            try
            {
                func();
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        [DllImport("User32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
        }

        private delegate void SafeFunc();
    }
}


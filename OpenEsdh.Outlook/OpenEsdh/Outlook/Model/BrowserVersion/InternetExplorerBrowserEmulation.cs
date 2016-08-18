namespace OpenEsdh.Outlook.Model.BrowserVersion
{
    using Microsoft.Win32;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Security;

    public static class InternetExplorerBrowserEmulation
    {
        private const string BrowserEmulationKey = @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
        private const string InternetExplorerRootKey = @"Software\Microsoft\Internet Explorer";

        public static BrowserEmulationVersion GetBrowserEmulationVersion()
        {
            BrowserEmulationVersion version = BrowserEmulationVersion.Default;
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                if (key == null)
                {
                    return version;
                }
                string fileName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);
                object obj2 = key.GetValue(fileName, null);
                if (obj2 != null)
                {
                    version = (BrowserEmulationVersion) Convert.ToInt32(obj2);
                }
            }
            catch (SecurityException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return version;
        }

        public static int GetInternetExplorerMajorVersion()
        {
            int result = 0;
            try
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Internet Explorer");
                if (key == null)
                {
                    return result;
                }
                object obj2 = key.GetValue("svcVersion", null) ?? key.GetValue("Version", null);
                if (obj2 == null)
                {
                    return result;
                }
                string str = obj2.ToString();
                int index = str.IndexOf('.');
                if (index != -1)
                {
                    int.TryParse(str.Substring(0, index), out result);
                }
            }
            catch (SecurityException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return result;
        }

        public static bool IsBrowserEmulationSet()
        {
            return (GetBrowserEmulationVersion() != BrowserEmulationVersion.Default);
        }

        public static bool SetBrowserEmulationVersion()
        {
            BrowserEmulationVersion version;
            int internetExplorerMajorVersion = GetInternetExplorerMajorVersion();
            if (internetExplorerMajorVersion >= 11)
            {
                version = BrowserEmulationVersion.Version11;
            }
            else
            {
                switch (internetExplorerMajorVersion)
                {
                    case 8:
                        version = BrowserEmulationVersion.Version8;
                        goto Label_0056;

                    case 9:
                        version = BrowserEmulationVersion.Version9;
                        goto Label_0056;

                    case 10:
                        version = BrowserEmulationVersion.Version10;
                        goto Label_0056;
                }
                version = BrowserEmulationVersion.Version7;
            }
        Label_0056:
            return SetBrowserEmulationVersion(version);
        }

        public static bool SetBrowserEmulationVersion(BrowserEmulationVersion browserEmulationVersion)
        {
            bool flag = false;
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                if (key == null)
                {
                    key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                }
                if (key == null)
                {
                    return flag;
                }
                string fileName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);
                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = Process.GetCurrentProcess().MainModule.ModuleName;
                }
                if (browserEmulationVersion != BrowserEmulationVersion.Default)
                {
                    key.SetValue(fileName, (int) browserEmulationVersion, RegistryValueKind.DWord);
                }
                else
                {
                    key.DeleteValue(fileName, false);
                }
                flag = true;
            }
            catch (SecurityException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return flag;
        }

        public static bool SetBrowserEmulationVersion(string programName, BrowserEmulationVersion browserEmulationVersion)
        {
            bool flag = false;
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                if (key == null)
                {
                    key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", true);
                }
                if (key == null)
                {
                    return flag;
                }
                if (browserEmulationVersion != BrowserEmulationVersion.Default)
                {
                    key.SetValue(programName, (int) browserEmulationVersion, RegistryValueKind.DWord);
                }
                else
                {
                    key.DeleteValue(programName, false);
                }
                flag = true;
            }
            catch (SecurityException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
            return flag;
        }
    }
}


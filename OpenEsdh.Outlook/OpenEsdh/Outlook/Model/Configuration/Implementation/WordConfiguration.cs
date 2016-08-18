namespace OpenEsdh.Outlook.Model.Configuration.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.Configuration;

    public class WordConfiguration : ConfigurationSection, IWordConfiguration, IConfiguration
    {
        public ICommunicationConfiguration CommunicationConfiguration
        {
            get
            {
                ICommunicationConfiguration configuration;
                try
                {
                    configuration = this.CommunicationConfiguration_Internal;
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return configuration;
            }
        }

        [ConfigurationProperty("CommunicationConfiguration", IsRequired=false)]
        public OpenEsdh.Outlook.Model.Configuration.Implementation.CommunicationConfiguration CommunicationConfiguration_Internal
        {
            get
            {
                OpenEsdh.Outlook.Model.Configuration.Implementation.CommunicationConfiguration configuration;
                try
                {
                    object obj2 = base["CommunicationConfiguration"];
                    if (obj2 != null)
                    {
                        return (OpenEsdh.Outlook.Model.Configuration.Implementation.CommunicationConfiguration) obj2;
                    }
                    configuration = new OpenEsdh.Outlook.Model.Configuration.Implementation.CommunicationConfiguration();
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return configuration;
            }
            set
            {
                base["CommunicationConfiguration"] = value;
            }
        }

        public IExtendDialog DialogExtend
        {
            get
            {
                return this.DialogExtend_Internal;
            }
            set
            {
                ExtendDialog dialog = new ExtendDialog {
                    X = value.X,
                    Y = value.Y
                };
                this.DialogExtend_Internal = dialog;
            }
        }

        [ConfigurationProperty("DialogExtend")]
        public ExtendDialog DialogExtend_Internal
        {
            get
            {
                ExtendDialog dialog;
                try
                {
                    dialog = (ExtendDialog) base["DialogExtend"];
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return dialog;
            }
            set
            {
                base["DialogExtend"] = value;
            }
        }

        public IDisplayRegionConfiguration DisplayRegion
        {
            get
            {
                IDisplayRegionConfiguration configuration;
                try
                {
                    configuration = this.DisplayRegion_internal;
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return configuration;
            }
        }

        [ConfigurationProperty("DisplayRegion", IsRequired=false)]
        public DisplayRegionConfiguration DisplayRegion_internal
        {
            get
            {
                DisplayRegionConfiguration configuration;
                try
                {
                    object obj2 = base["DisplayRegion"];
                    if (obj2 != null)
                    {
                        return (DisplayRegionConfiguration) obj2;
                    }
                    configuration = new DisplayRegionConfiguration();
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return configuration;
            }
            set
            {
                base["DisplayRegion"] = value;
            }
        }

        [ConfigurationProperty("EndUploadEndpoint", DefaultValue="", IsRequired=false)]
        public string EndUploadEndpoint
        {
            get
            {
                return (string) base["EndUploadEndpoint"];
            }
            set
            {
                base["EndUploadEndpoint"] = value;
            }
        }

        [ConfigurationProperty("EndUploadPackage", DefaultValue="{\"status\": \"final\"}", IsRequired=false)]
        public string EndUploadPackage
        {
            get
            {
                return (string) base["EndUploadPackage"];
            }
            set
            {
                base["EndUploadPackage"] = value;
            }
        }

        [ConfigurationProperty("GetFileEndPoint", DefaultValue="", IsRequired=false)]
        public string GetFileEndPoint
        {
            get
            {
                return (string) base["GetFileEndPoint"];
            }
            set
            {
                base["GetFileEndPoint"] = value;
            }
        }

        [ConfigurationProperty("IgnoreCertificateErrors", DefaultValue="false", IsRequired=false)]
        public bool IgnoreCertificateErrors
        {
            get
            {
                return bool.Parse((base["IgnoreCertificateErrors"] != null) ? base["IgnoreCertificateErrors"].ToString() : "false");
            }
            set
            {
                base["IgnoreCertificateErrors"] = value;
            }
        }

        [ConfigurationProperty("LoginDialogTagToFind", DefaultValue="slingshot-login_x0023_default", IsRequired=false)]
        public string LoginIdToFind
        {
            get
            {
                return (string) base["LoginDialogTagToFind"];
            }
            set
            {
                base["LoginDialogTagToFind"] = value;
            }
        }

        [ConfigurationProperty("LoginTagToFind", DefaultValue="div", IsRequired=false)]
        public string LoginTagToFind
        {
            get
            {
                return (string) base["LoginTagToFind"];
            }
            set
            {
                base["LoginTagToFind"] = value;
            }
        }

        [ConfigurationProperty("MaxRedirectRetries", DefaultValue="1", IsRequired=false)]
        public int MaxRedirectRetries
        {
            get
            {
                return (int) base["MaxRedirectRetries"];
            }
            set
            {
                base["MaxRedirectRetries"] = value;
            }
        }

        [ConfigurationProperty("PreAuthenticate", DefaultValue="false", IsRequired=false)]
        public bool PreAuthenticate
        {
            get
            {
                return ((base["PreAuthenticate"] != null) ? bool.Parse(base["PreAuthenticate"].ToString()) : false);
            }
            set
            {
                base["PreAuthenticate"] = value;
            }
        }

        public IPreAuthenticateConfiguration PreAuthentication
        {
            get
            {
                IPreAuthenticateConfiguration configuration;
                try
                {
                    configuration = this.PreAuthentication_internal;
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
                return configuration;
            }
        }

        [ConfigurationProperty("PreAuthentication", IsRequired=false)]
        public PreAuthenticationConfiguration PreAuthentication_internal
        {
            get
            {
                object obj2 = base["PreAuthentication"];
                if (obj2 != null)
                {
                    return (PreAuthenticationConfiguration) obj2;
                }
                return new PreAuthenticationConfiguration();
            }
            set
            {
                base["PreAuthentication"] = value;
            }
        }

        [ConfigurationProperty("SaveAsDialogUrl", DefaultValue="http://www.google.dk", IsRequired=false)]
        public string SaveAsDialogUrl
        {
            get
            {
                return (string) base["SaveAsDialogUrl"];
            }
            set
            {
                base["SaveAsDialogUrl"] = value;
            }
        }

        [ConfigurationProperty("SaveDialogUrl", DefaultValue="http://www.google.dk", IsRequired=false)]
        public string SaveDialogUrl
        {
            get
            {
                return (string) base["SaveDialogUrl"];
            }
            set
            {
                base["SaveDialogUrl"] = value;
            }
        }

        [ConfigurationProperty("UploadEndPoint", DefaultValue="https://alfresco.dk.vsw.datakraftverk.no:8443/alfresco/service/dk-openesdh-aoi-save", IsRequired=false)]
        public string UploadEndPoint
        {
            get
            {
                return (string) base["UploadEndPoint"];
            }
            set
            {
                base["UploadEndPoint"] = value;
            }
        }

        [ConfigurationProperty("UseRedirectJavascript", DefaultValue="true", IsRequired=false)]
        public bool UseRedirectJavascript
        {
            get
            {
                return (bool) base["UseRedirectJavascript"];
            }
            set
            {
                base["UseRedirectJavascript"] = value;
            }
        }
    }
}


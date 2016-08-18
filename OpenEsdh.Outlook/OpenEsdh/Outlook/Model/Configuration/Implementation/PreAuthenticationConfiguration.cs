namespace OpenEsdh.Outlook.Model.Configuration.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using System;
    using System.Configuration;

    public class PreAuthenticationConfiguration : ConfigurationElement, IPreAuthenticateConfiguration
    {
        [ConfigurationProperty("AdditionalRequestHeaders", DefaultValue="", IsRequired=false)]
        public string AdditionalRequestHeaders
        {
            get
            {
                return (string) base["AdditionalRequestHeaders"];
            }
            set
            {
                base["AdditionalRequestHeaders"] = value;
            }
        }

        [ConfigurationProperty("AuthenticationPackageFormat", DefaultValue="AuthenticationPackageFormat", IsRequired=false)]
        public string AuthenticationPackageFormat
        {
            get
            {
                return (string) base["AuthenticationPackageFormat"];
            }
            set
            {
                base["AuthenticationPackageFormat"] = value;
            }
        }

        [ConfigurationProperty("AuthenticationUrl", DefaultValue="", IsRequired=false)]
        public string AuthenticationUrl
        {
            get
            {
                return (string) base["AuthenticationUrl"];
            }
            set
            {
                base["AuthenticationUrl"] = value;
            }
        }

        [ConfigurationProperty("Domain", DefaultValue="", IsRequired=false)]
        public string Domain
        {
            get
            {
                return (string) base["Domain"];
            }
            set
            {
                base["Domain"] = value;
            }
        }

        [ConfigurationProperty("Password", DefaultValue="", IsRequired=false)]
        public string Password
        {
            get
            {
                return (string) base["Password"];
            }
            set
            {
                base["Password"] = value;
            }
        }

        [ConfigurationProperty("PreAuthenticateParameterName", DefaultValue="alt_ticket", IsRequired=false)]
        public string PreAuthenticateParameterName
        {
            get
            {
                return (string) base["PreAuthenticateParameterName"];
            }
            set
            {
                base["PreAuthenticateParameterName"] = value;
            }
        }

        [ConfigurationProperty("ReauthenticateOnEachRequest", DefaultValue=true, IsRequired=false)]
        public bool ReauthenticateOnEachRequest
        {
            get
            {
                return (bool) base["ReauthenticateOnEachRequest"];
            }
            set
            {
                base["ReauthenticateOnEachRequest"] = value;
            }
        }

        [ConfigurationProperty("UseConfigCredentials", DefaultValue="false", IsRequired=false)]
        public bool UseConfigCredentials
        {
            get
            {
                return bool.Parse((base["UseConfigCredentials"] != null) ? base["UseConfigCredentials"].ToString() : "false");
            }
            set
            {
                base["UseConfigCredentials"] = value;
            }
        }

        [ConfigurationProperty("Username", DefaultValue="", IsRequired=false)]
        public string Username
        {
            get
            {
                return (string) base["Username"];
            }
            set
            {
                base["Username"] = value;
            }
        }
    }
}


namespace OpenEsdh.Outlook.Model.Configuration.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using System;
    using System.Configuration;

    public class DisplayRegionConfiguration : ConfigurationElement, IDisplayRegionConfiguration
    {
        [ConfigurationProperty("DisplayDialogUrl", DefaultValue="www.google.com", IsRequired=false)]
        public string DisplayDialogUrl
        {
            get
            {
                return (string) base["DisplayDialogUrl"];
            }
            set
            {
                base["DisplayDialogUrl"] = value;
            }
        }

        [ConfigurationProperty("RequestParameter", DefaultValue="q", IsRequired=false)]
        public string RequestParameter
        {
            get
            {
                return (string) base["RequestParameter"];
            }
            set
            {
                base["RequestParameter"] = value;
            }
        }
    }
}


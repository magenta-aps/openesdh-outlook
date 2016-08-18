namespace OpenEsdh.Outlook.Model.Configuration.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using System;
    using System.Configuration;

    public class ExtendDialog : ConfigurationElement, IExtendDialog
    {
        [ConfigurationProperty("MaxHeight", DefaultValue="510", IsRequired=false)]
        public int MaxHeight
        {
            get
            {
                return (int) base["MaxHeight"];
            }
            set
            {
                base["MaxHeight"] = value;
            }
        }

        [ConfigurationProperty("MaxWidth", DefaultValue="280", IsRequired=false)]
        public int MaxWidth
        {
            get
            {
                return (int) base["MaxWidth"];
            }
            set
            {
                base["MaxWidth"] = value;
            }
        }

        [ConfigurationProperty("X", DefaultValue="60", IsRequired=false)]
        public int X
        {
            get
            {
                return (int) base["X"];
            }
            set
            {
                base["X"] = value;
            }
        }

        [ConfigurationProperty("Y", DefaultValue="10", IsRequired=false)]
        public int Y
        {
            get
            {
                return (int) base["Y"];
            }
            set
            {
                base["Y"] = value;
            }
        }
    }
}


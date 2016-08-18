namespace OpenEsdh.Outlook.Model.Configuration.Implementation
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using System;
    using System.Configuration;

    public class CommunicationConfiguration : ConfigurationElement, ICommunicationConfiguration
    {
        public CommunicationConfiguration()
        {
            this.JavaScriptMethodName = "OpenESDHInitialize";
            this.PostMethodName = "Initialize";
            this.SendMethod_Internal = "GET";
            this.DelayUntilJavaMethodCall = 0;
        }

        [ConfigurationProperty("DelayUntilJavaMethodCall", DefaultValue="0", IsRequired=false)]
        public int DelayUntilJavaMethodCall
        {
            get
            {
                return (int) base["DelayUntilJavaMethodCall"];
            }
            set
            {
                base["DelayUntilJavaMethodCall"] = value;
            }
        }

        [ConfigurationProperty("JavaScriptMethodName", DefaultValue="OpenESDHInitialize", IsRequired=false)]
        public string JavaScriptMethodName
        {
            get
            {
                return (string) base["JavaScriptMethodName"];
            }
            set
            {
                base["JavaScriptMethodName"] = value;
            }
        }

        [ConfigurationProperty("PostMethodName", DefaultValue="Initialize", IsRequired=false)]
        public string PostMethodName
        {
            get
            {
                return (string) base["PostMethodName"];
            }
            set
            {
                base["PostMethodName"] = value;
            }
        }

        public SendDataMethod SendMethod
        {
            get
            {
                switch (this.SendMethod_Internal.ToUpper())
                {
                    case "GET":
                        return SendDataMethod.JavascriptMethod;

                    case "POST":
                        return SendDataMethod.Post;
                }
                return SendDataMethod.JavascriptMethod;
            }
            set
            {
                switch (value)
                {
                    case SendDataMethod.JavascriptMethod:
                        this.SendMethod_Internal = "GET";
                        break;

                    case SendDataMethod.Post:
                        this.SendMethod_Internal = "POST";
                        break;

                    default:
                        this.SendMethod_Internal = "GET";
                        break;
                }
            }
        }

        [ConfigurationProperty("SendMethod", DefaultValue="GET", IsRequired=false)]
        public string SendMethod_Internal
        {
            get
            {
                return (string) base["SendMethod"];
            }
            set
            {
                base["SendMethod"] = value;
            }
        }
    }
}


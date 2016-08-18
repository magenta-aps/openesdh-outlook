namespace OpenEsdh.Outlook.Model
{
    using System;
    using System.Collections.Specialized;

    public interface IPreAuthenticator
    {
        string AddAuthentificationUrl(string url);
        string GetParameters();

        NameValueCollection AdditionalParameters { get; }
    }
}


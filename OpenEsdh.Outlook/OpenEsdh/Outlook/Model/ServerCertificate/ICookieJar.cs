namespace OpenEsdh.Outlook.Model.ServerCertificate
{
    using System;
    using System.Collections.Generic;

    internal interface ICookieJar
    {
        void Add(string cookie);
        void AddCookiesForUri(Uri uri);
        string GetCookieString();
        string GetParamString();

        Dictionary<string, string> Cookies { get; }
    }
}


namespace OpenEsdh.Outlook.Model.ServerCertificate
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Runtime.InteropServices;
    using System.Text;

    public class CookieJar : ICookieJar
    {
        private Dictionary<string, string> _cookies = new Dictionary<string, string>();
        private const int INTERNET_COOKIE_HTTPONLY = 0x2000;

        public void Add(string cookie)
        {
            if (!string.IsNullOrEmpty(cookie))
            {
                string[] strArray = cookie.Split(new char[] { ';' });
                foreach (string str in strArray)
                {
                    string[] strArray2 = str.Trim().Split(new char[] { '=' });
                    string key = "";
                    string str3 = "";
                    if (strArray2.Length >= 1)
                    {
                        key = strArray2[0];
                        if (strArray2.Length >= 2)
                        {
                            str3 = strArray2[1];
                        }
                        if (!this._cookies.ContainsKey(key))
                        {
                            this._cookies.Add(key, str3);
                        }
                        else
                        {
                            this._cookies[key] = str3;
                        }
                    }
                }
            }
        }

        public void AddCookiesForUri(Uri uri)
        {
            CookieContainer uriCookieContainer = this.GetUriCookieContainer(uri);
            if (uriCookieContainer != null)
            {
                foreach (Cookie cookie in uriCookieContainer.GetCookies(uri))
                {
                    if (this._cookies.ContainsKey(cookie.Name))
                    {
                        this._cookies[cookie.Name] = cookie.Value;
                    }
                    else
                    {
                        this._cookies.Add(cookie.Name, cookie.Value);
                    }
                }
            }
            uriCookieContainer = this.GetUriCookieContainerEx(uri);
            if (uriCookieContainer != null)
            {
                foreach (Cookie cookie in uriCookieContainer.GetCookies(uri))
                {
                    if (this._cookies.ContainsKey(cookie.Name))
                    {
                        this._cookies[cookie.Name] = cookie.Value;
                    }
                    else
                    {
                        this._cookies.Add(cookie.Name, cookie.Value);
                    }
                }
            }
        }

        public string GetCookieString()
        {
            string str = "";
            string str2 = "";
            foreach (string str3 in this.Cookies.Keys)
            {
                str = str + str2 + str3 + this.Cookies[str3];
                str2 = "; ";
            }
            return str;
        }

        public string GetParamString()
        {
            string str = "";
            string str2 = "";
            foreach (string str3 in this.Cookies.Keys)
            {
                str = str + str2 + str3 + this.Cookies[str3];
                str2 = "&";
            }
            return str;
        }

        private CookieContainer GetUriCookieContainer(Uri uri)
        {
            CookieContainer container = null;
            int capacity = 0x100;
            StringBuilder cookieData = new StringBuilder(capacity);
            if (!InternetGetCookie(uri.ToString(), null, cookieData, ref capacity))
            {
                if (capacity < 0)
                {
                    return null;
                }
                cookieData = new StringBuilder(capacity);
                if (!InternetGetCookie(uri.ToString(), null, cookieData, ref capacity))
                {
                    return null;
                }
            }
            if (cookieData.Length > 0)
            {
                container = new CookieContainer();
                container.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return container;
        }

        public CookieContainer GetUriCookieContainerEx(Uri uri)
        {
            CookieContainer container = null;
            int capacity = 0x200;
            StringBuilder cookieData = new StringBuilder(capacity);
            if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref capacity, 0x2000, IntPtr.Zero))
            {
                if (capacity < 0)
                {
                    return null;
                }
                cookieData = new StringBuilder(capacity);
                if (!InternetGetCookieEx(uri.ToString(), null, cookieData, ref capacity, 0x2000, IntPtr.Zero))
                {
                    return null;
                }
            }
            if (cookieData.Length > 0)
            {
                container = new CookieContainer();
                container.SetCookies(uri, cookieData.ToString().Replace(';', ','));
            }
            return container;
        }

        [DllImport("wininet.dll", SetLastError=true)]
        private static extern bool InternetGetCookie(string url, string cookieName, StringBuilder cookieData, ref int size);
        [DllImport("wininet.dll", SetLastError=true)]
        private static extern bool InternetGetCookieEx(string url, string cookieName, StringBuilder cookieData, ref int size, int flags, IntPtr pReserved);
        [DllImport("wininet.dll", CharSet=CharSet.Auto, SetLastError=true)]
        private static extern bool InternetSetCookie(string url, string name, string data);
        [DllImport("wininet.dll")]
        private static extern InternetCookieState InternetSetCookieEx(string lpszURL, string lpszCookieName, string lpszCookieData, int dwFlags, IntPtr dwReserved);
        public InternetCookieState SetCookie(string url, string name, string data)
        {
            return InternetSetCookieEx(url, name, data, 0x2000, IntPtr.Zero);
        }

        public Dictionary<string, string> Cookies
        {
            get
            {
                return this._cookies;
            }
        }

        public enum InternetCookieState
        {
            COOKIE_STATE_ACCEPT = 1,
            COOKIE_STATE_DOWNGRADE = 4,
            COOKIE_STATE_LEASH = 3,
            COOKIE_STATE_MAX = 5,
            COOKIE_STATE_PROMPT = 2,
            COOKIE_STATE_REJECT = 5,
            COOKIE_STATE_UNKNOWN = 0
        }
    }
}


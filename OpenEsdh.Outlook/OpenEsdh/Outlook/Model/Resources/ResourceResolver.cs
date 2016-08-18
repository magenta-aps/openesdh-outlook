namespace OpenEsdh.Outlook.Model.Resources
{
    using OpenEsdh.Outlook.Resources;
    using System;
    using System.Resources;

    public class ResourceResolver
    {
        private static ResourceResolver _current = null;
        private static object _lock = new object();
        private ResourceManager _mgr;

        public ResourceResolver()
        {
            this._mgr = new ResourceManager("OpenEsdh.Outlook", base.GetType().Assembly);
        }

        public string GetString(string Key)
        {
            try
            {
                return OpenEsdh_Outlook.ResourceManager.GetString(Key);
            }
            catch
            {
                return ("(" + Key + ") Ikke Fundet");
            }
        }

        public static ResourceResolver Current
        {
            get
            {
                if (_current == null)
                {
                    lock (_lock)
                    {
                        if (_current == null)
                        {
                            _current = new ResourceResolver();
                        }
                    }
                }
                return _current;
            }
        }
    }
}


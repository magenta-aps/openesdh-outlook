namespace OpenEsdh.Outlook.Model
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using System.Threading;

    public class ApplicationDescriptor
    {
        public ApplicationDescriptor()
        {
            this.Name = "";
            this.Title = "";
            this.ID = string.Empty;
            this.Author = Thread.CurrentPrincipal.Identity.Name;
            this.MetaData = new List<KeyValuePair<string, string>>();
        }

        public string Author { get; set; }

        public string ID { get; set; }

        public IList<KeyValuePair<string, string>> MetaData { get; set; }

        public string Name { get; set; }

        public string Title { get; set; }
    }
}


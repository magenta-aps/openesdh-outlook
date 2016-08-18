namespace OpenEsdh.Outlook.Model
{
    using System;
    using System.Runtime.CompilerServices;

    public class SelectableAttachment : Attachment
    {
        public SelectableAttachment()
        {
            this.Selected = false;
        }

        public SelectableAttachment(string name, string mimeType) : this()
        {
        }

        public SelectableAttachment(string name, string mimeType, bool forceImport) : this()
        {
        }

        public bool Selected { get; set; }
    }
}


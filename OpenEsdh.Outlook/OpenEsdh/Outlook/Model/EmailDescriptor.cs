namespace OpenEsdh.Outlook.Model
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;

    public class EmailDescriptor
    {
        public EmailDescriptor()
        {
            this.From = new EmailAddress();
            this.To = new List<EmailAddress>();
            this.CC = new List<EmailAddress>();
            this.BCC = new List<EmailAddress>();
            this.MetaData = new List<KeyValuePair<string, string>>();
            this.Attachments = new List<Attachment>();
            this.Subject = "";
            this.BodyText = "";
            this.BodyHtml = "";
        }

        public IList<Attachment> Attachments { get; set; }

        public IList<EmailAddress> BCC { get; set; }

        public string BodyHtml { get; set; }

        public string BodyText { get; set; }

        public IList<EmailAddress> CC { get; set; }

        public EmailAddress From { get; set; }

        public IList<KeyValuePair<string, string>> MetaData { get; set; }

        public string Subject { get; set; }

        public IList<EmailAddress> To { get; set; }
    }
}


namespace OpenEsdh.Outlook.Model
{
    using System;
    using System.Runtime.CompilerServices;

    public class EmailAddress
    {
        public EmailAddress()
        {
            this.EMail = "";
            this.Name = "";
        }

        public EmailAddress(string email) : this()
        {
            this.EMail = email;
            this.Name = email;
        }

        public EmailAddress(string email, string name) : this()
        {
            this.EMail = email;
            this.Name = name;
        }

        public string EMail { get; set; }

        public string Name { get; set; }
    }
}


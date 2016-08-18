namespace OpenEsdh.Outlook.Model
{
    using System;

    public interface IAttachEmail
    {
        void AddAttachmentConfiguration(string[] ConfigurationSettings, SetMailPropertyDelegate SetProperty, AddMailPropertyDelegate AddProperty, AddFileDelegate AddFile);
    }
}


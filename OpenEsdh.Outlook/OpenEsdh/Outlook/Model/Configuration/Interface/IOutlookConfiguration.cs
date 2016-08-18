namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface IOutlookConfiguration : IConfiguration
    {
        string AttachFileEndPoint { get; }

        string RevieveMessageClass { get; }

        string SendMessageClass { get; }
    }
}


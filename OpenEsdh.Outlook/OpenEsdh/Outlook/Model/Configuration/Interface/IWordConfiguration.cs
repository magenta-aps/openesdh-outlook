namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface IWordConfiguration : IConfiguration
    {
        string GetFileEndPoint { get; }
    }
}


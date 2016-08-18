namespace OpenEsdh.Outlook.Views.Interface
{
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using System;

    public interface ISaveAsBrowserView
    {
        void CancelOpenEsdh();
        void InitializeOpendEsdhPost(IOutlookConfiguration config, string payload);
        void InitializeOpenEsdh(IOutlookConfiguration config, string jsonEmail);
        void SaveAsOpenEsdh(string unknownJson, string attachmentSelectedJson);
    }
}


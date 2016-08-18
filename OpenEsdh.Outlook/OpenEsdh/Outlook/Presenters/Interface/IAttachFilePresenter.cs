namespace OpenEsdh.Outlook.Presenters.Interface
{
    using OpenEsdh.Outlook.Model;
    using System;

    public interface IAttachFilePresenter
    {
        void AttachFile(string AttachmentConfiguration);
        void Cancel();
        void Initialize(EmailDescriptor descriptor, AttachFileCallback callback);
    }
}


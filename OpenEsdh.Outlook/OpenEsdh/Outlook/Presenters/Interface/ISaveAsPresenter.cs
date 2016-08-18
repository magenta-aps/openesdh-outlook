namespace OpenEsdh.Outlook.Presenters.Interface
{
    using OpenEsdh.Outlook.Model;
    using System;

    public interface ISaveAsPresenter
    {
        event SaveAttachmentDelegate SaveAttachment;

        event SaveMailBodyDelegate SaveMailBody;

        event SetMessageClassDelegate SetMessageClass;

        event SetMessageIDDelegate SetMessageID;

        void Cancel();
        void SaveAs(string unknown, SelectableAttachment[] SelectedAttachments);
        void Show(EmailDescriptor Email);
        bool ShowAndSend(EmailDescriptor Email);
    }
}


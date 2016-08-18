namespace OpenEsdh._2013.Outlook.Presentation.Interface
{
    using Microsoft.Office.Interop.Outlook;
    using System;

    public interface IAttachFilePresenter
    {
        void AttachFileClick(MailItem item);
    }
}


namespace OpenEsdh._2013.Outlook.Presentation.Interface
{
    using Microsoft.Office.Interop.Outlook;
    using System;
    using System.Runtime.CompilerServices;

    internal interface ISaveEmailPresenter
    {
        void Load([Dynamic] object Context);
        bool SaveEmailAndSend(MailItem item, Action sendOperation);
        void SaveEmailClick(MailItem item);

        ISaveEmailButtonView View { get; set; }
    }
}


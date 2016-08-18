namespace OpenEsdh.Outlook.Presenters.Interface
{
    using OpenEsdh.Outlook.Model;
    using System;

    public interface IApplicationSaveAsPresenter
    {
        event SaveDocumentDelegate SaveDocument;

        event SetDocumentIDDelegate SetDocumentID;

        void Cancel();
        void SaveAs(string unknown);
        bool Show(ApplicationDescriptor document);
    }
}


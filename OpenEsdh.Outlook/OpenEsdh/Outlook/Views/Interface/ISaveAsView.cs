namespace OpenEsdh.Outlook.Views.Interface
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;

    public interface ISaveAsView
    {
        void Cancel();
        void Initialize(string uri, EmailDescriptor Email);
        void ShowView();

        ISaveAsPresenter Presenter { get; set; }
    }
}


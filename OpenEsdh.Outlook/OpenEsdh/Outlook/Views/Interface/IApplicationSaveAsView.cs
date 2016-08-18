namespace OpenEsdh.Outlook.Views.Interface
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;

    public interface IApplicationSaveAsView
    {
        void Cancel();
        void Initialize(string uri, ApplicationDescriptor Email);
        void ShowView();

        IApplicationSaveAsPresenter Presenter { get; set; }
    }
}


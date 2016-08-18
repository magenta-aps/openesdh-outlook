namespace OpenEsdh.Outlook.Views.Interface
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;

    public interface IAttachFileView
    {
        void Cancel();
        void Initialize(string uri, EmailDescriptor Email);
        void ShowView();

        IAttachFilePresenter Presenter { get; set; }
    }
}


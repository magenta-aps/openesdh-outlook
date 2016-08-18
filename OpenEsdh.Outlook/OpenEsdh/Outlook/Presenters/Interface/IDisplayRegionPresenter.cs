namespace OpenEsdh.Outlook.Presenters.Interface
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Views.Interface;
    using System;

    public interface IDisplayRegionPresenter
    {
        void Show(EmailDescriptor email);

        IDisplayRegion DisplayRegion { get; }
    }
}


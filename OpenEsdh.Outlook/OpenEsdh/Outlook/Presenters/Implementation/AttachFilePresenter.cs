namespace OpenEsdh.Outlook.Presenters.Implementation
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Interface;
    using System;

    public class AttachFilePresenter : IAttachFilePresenter
    {
        private AttachFileCallback _callback;
        private IOutlookConfiguration _configuration;
        private IAttachFileView _view;

        public AttachFilePresenter() : this(TypeResolver.Current.Create<IAttachFileView>())
        {
        }

        public AttachFilePresenter(IAttachFileView view)
        {
            this._callback = null;
            this._configuration = TypeResolver.Current.Create<IOutlookConfiguration>();
            this._view = view;
        }

        public void AttachFile(string AttachmentConfiguration)
        {
            if (this._callback != null)
            {
                this._callback(AttachmentConfiguration);
            }
        }

        public void Cancel()
        {
            this._view.Cancel();
        }

        public void Initialize(EmailDescriptor descriptor, AttachFileCallback callback)
        {
            this._callback = callback;
            this._view.Initialize(this._configuration.AttachFileEndPoint, descriptor);
        }
    }
}


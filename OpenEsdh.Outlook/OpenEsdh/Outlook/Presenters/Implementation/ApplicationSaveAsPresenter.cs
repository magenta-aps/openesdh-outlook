namespace OpenEsdh.Outlook.Presenters.Implementation
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Alfresco;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Interface;
    using System;
    using System.IO;
    using System.Threading;

    public class ApplicationSaveAsPresenter : IApplicationSaveAsPresenter
    {
        private bool _canceled;
        private bool _inOperation;
        private IApplicationSaveAsView _view;

        public event SaveDocumentDelegate SaveDocument;

        public event SetDocumentIDDelegate SetDocumentID;

        public ApplicationSaveAsPresenter()
        {
            this._canceled = false;
            this._inOperation = false;
            this._view = TypeResolver.Current.Create<IApplicationSaveAsView>();
            this._view.Presenter = this;
        }

        public ApplicationSaveAsPresenter(IApplicationSaveAsView view)
        {
            this._canceled = false;
            this._inOperation = false;
            this._view = view;
            this._view.Presenter = this;
        }

        public void Cancel()
        {
            if (!this._inOperation)
            {
                try
                {
                    this._inOperation = true;
                    this._canceled = true;
                    this._view.Cancel();
                }
                finally
                {
                    this._inOperation = false;
                }
            }
        }

        protected void DoSaveDocument(string unknown, bool SetDocumentId)
        {
            if (this.SaveDocument != null)
            {
                IAlfrescoFilePost Upload = TypeResolver.Current.Create<IAlfrescoFilePost>();
                if (SetDocumentId)
                {
                    this.DoSetDocumentID(unknown);
                }
                this.SaveDocument(delegate (string filename, string name) {
                    Exception exception;
                    string path = Path.GetTempFileName();
                    using (Stream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (FileStream stream2 = File.OpenWrite(path))
                        {
                            stream.CopyTo(stream2);
                        }
                    }
                    try
                    {
                        Upload.UploadFile(unknown, path, name);
                    }
                    catch (Exception exception1)
                    {
                        exception = exception1;
                        Logger.Current.LogException(exception, "");
                    }
                    try
                    {
                        File.Delete(path);
                    }
                    catch (Exception exception2)
                    {
                        exception = exception2;
                        Logger.Current.LogException(exception, "");
                    }
                });
                try
                {
                    new EndUploadRequest(TypeResolver.Current.Create<IWordConfiguration>()).EndUpload(unknown);
                }
                catch (Exception exception)
                {
                    Logger.Current.LogException(exception, "");
                }
            }
        }

        protected void DoSetDocumentID(string unknown)
        {
            try
            {
                if (this.SetDocumentID != null)
                {
                    this.SetDocumentID(unknown);
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public virtual void SaveAs(string unknown)
        {
            try
            {
                this.DoSaveDocument(unknown, true);
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }

        public virtual bool Show(ApplicationDescriptor document)
        {
            this._canceled = false;
            IWordConfiguration configuration = TypeResolver.Current.Create<IWordConfiguration>();
            string saveAsDialogUrl = configuration.SaveAsDialogUrl;
            if (configuration.PreAuthenticate)
            {
                saveAsDialogUrl = TypeResolver.Current.Create<IPreAuthenticator>().AddAuthentificationUrl(saveAsDialogUrl);
            }
            this._view.Initialize(saveAsDialogUrl, document);
            this._view.ShowView();
            return this._canceled;
        }
    }
}


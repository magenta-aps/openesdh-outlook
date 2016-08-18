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
    using System.Threading;
    using System.Web.Script.Serialization;

    public class SaveAsPresenter : ISaveAsPresenter
    {
        private readonly IOutlookConfiguration _configuration;
        private bool _inOperation;
        private bool _result;
        private readonly ISaveAsView _view;

        public event SaveAttachmentDelegate SaveAttachment;

        public event SaveMailBodyDelegate SaveMailBody;

        public event SetMessageClassDelegate SetMessageClass;

        public event SetMessageIDDelegate SetMessageID;

        public SaveAsPresenter() : this(TypeResolver.Current.Create<ISaveAsView>())
        {
        }

        public SaveAsPresenter(ISaveAsView view)
        {
            this._result = false;
            this._inOperation = false;
            this._view = view;
            this._configuration = TypeResolver.Current.Create<IOutlookConfiguration>();
        }

        public void Cancel()
        {
            if (!this._inOperation)
            {
                try
                {
                    this._inOperation = true;
                    this._result = false;
                    this._view.Cancel();
                }
                finally
                {
                    this._inOperation = false;
                }
            }
        }

        private void DoSetMessageClass(string messageClass)
        {
            if (this.SetMessageClass != null)
            {
                this.SetMessageClass(messageClass);
            }
        }

        private void DoSetMessageId(string messageId)
        {
            if (this.SetMessageID != null)
            {
                this.SetMessageID(messageId);
            }
        }

        public void SaveAs(string unknown, SelectableAttachment[] SelectedAttachments)
        {
            Exception exception;
            try
            {
                UploadMailFileDelegate uploadFunction = null;
                UploadMailFileDelegate delegate3 = null;
                IAlfrescoFilePost Upload = TypeResolver.Current.Create<IAlfrescoFilePost>();
                if (this.SaveMailBody != null)
                {
                    if (uploadFunction == null)
                    {
                        uploadFunction = delegate (string filename, string name) {
                            try
                            {
                                Upload.UploadFile(unknown, filename, name);
                            }
                            catch (Exception exception0)
                            {
                                exception = exception0;
                                Logger.Current.LogException(exception, "");
                                throw exception;
                            }
                        };
                    }
                    this.SaveMailBody(uploadFunction);
                }
                if (this.SaveAttachment != null)
                {
                    foreach (SelectableAttachment attachment in SelectedAttachments)
                    {
                        if (delegate3 == null)
                        {
                            delegate3 = delegate (string filename, string name) {
                                try
                                {
                                    Upload.UploadFile(unknown, filename, name);
                                }
                                catch (Exception exception0)
                                {
                                    exception = exception0;
                                    Logger.Current.LogException(exception, "");
                                    throw exception;
                                }
                            };
                        }
                        this.SaveAttachment(attachment.Name, delegate3);
                    }
                }
                try
                {
                    UploadTicket ticket = new JavaScriptSerializer().Deserialize<UploadTicket>(unknown);
                    this.SetMessageClass(this._configuration.RevieveMessageClass);
                    this.SetMessageID(unknown);
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                }
                try
                {
                    new EndUploadRequest(this._configuration).EndUpload(unknown);
                }
                catch (Exception exception2)
                {
                    exception = exception2;
                    Logger.Current.LogException(exception, "");
                }
                this._result = true;
            }
            catch (Exception exception3)
            {
                exception = exception3;
                Logger.Current.LogException(exception, "");
            }
        }

        public void Show(EmailDescriptor Email)
        {
            try
            {
                if (this._configuration == null)
                {
                    Logger.Current.LogInformation("Configuration is null", "");
                }
                if (string.IsNullOrEmpty(this._configuration.SaveAsDialogUrl))
                {
                    Logger.Current.LogInformation("SaveAsDialogUrl is null", "");
                }
                if (Email == null)
                {
                    Logger.Current.LogInformation("Email is null", "");
                }
                if (this._view == null)
                {
                    Logger.Current.LogInformation("View is null", "");
                }
                string saveAsDialogUrl = this._configuration.SaveAsDialogUrl;
                if (this._configuration.PreAuthenticate)
                {
                    saveAsDialogUrl = TypeResolver.Current.Create<IPreAuthenticator>().AddAuthentificationUrl(saveAsDialogUrl);
                }
                this._view.Initialize(saveAsDialogUrl, Email);
                this._view.ShowView();
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public bool ShowAndSend(EmailDescriptor Email)
        {
            this._result = false;
            this.Show(Email);
            return this._result;
        }
    }
}


namespace OpenEsdh.Outlook.Presenters.Implementation
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Alfresco;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;

    public class ApplicationSavePresenter : ApplicationSaveAsPresenter, IApplicationSavePresenter, IApplicationSaveAsPresenter
    {
        public override void SaveAs(string unknown)
        {
        }

        public override bool Show(ApplicationDescriptor document)
        {
            try
            {
                if (!string.IsNullOrEmpty(document.ID))
                {
                    IAlfrescoFilePost post = TypeResolver.Current.Create<IAlfrescoFilePost>();
                    base.DoSaveDocument(document.ID, true);
                    return true;
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
            return false;
        }
    }
}


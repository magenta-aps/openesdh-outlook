namespace OpenEsdh._2013.Outlook.Model
{
    using OpenEsdh.Outlook.Model.Logging;
    using System;

    public class AttachmentService : IAttachmentService
    {
        public void AttachDocument(string AttachmentContent)
        {
            Logger.Current.LogInformation("Attach:" + AttachmentContent, "");
        }
    }
}


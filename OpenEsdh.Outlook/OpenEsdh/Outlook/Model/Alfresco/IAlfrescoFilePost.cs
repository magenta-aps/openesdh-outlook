namespace OpenEsdh.Outlook.Model.Alfresco
{
    using System;

    public interface IAlfrescoFilePost
    {
        string UploadFile(string unknownJson, string fileName, string name);
        string UploadFile(string address, string unknownJson, string fileName, string name);
    }
}


namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface IConfiguration
    {
        ICommunicationConfiguration CommunicationConfiguration { get; }

        IExtendDialog DialogExtend { get; }

        IDisplayRegionConfiguration DisplayRegion { get; }

        string EndUploadEndpoint { get; }

        string EndUploadPackage { get; }

        bool IgnoreCertificateErrors { get; }

        string LoginIdToFind { get; }

        string LoginTagToFind { get; }

        int MaxRedirectRetries { get; }

        bool PreAuthenticate { get; }

        IPreAuthenticateConfiguration PreAuthentication { get; }

        string SaveAsDialogUrl { get; }

        string UploadEndPoint { get; }

        bool UseRedirectJavascript { get; }
    }
}


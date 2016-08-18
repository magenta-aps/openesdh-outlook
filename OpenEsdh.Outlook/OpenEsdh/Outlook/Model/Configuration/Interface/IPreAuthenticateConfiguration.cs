namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface IPreAuthenticateConfiguration
    {
        string AdditionalRequestHeaders { get; }

        string AuthenticationPackageFormat { get; }

        string AuthenticationUrl { get; }

        string Domain { get; }

        string Password { get; }

        string PreAuthenticateParameterName { get; }

        bool ReauthenticateOnEachRequest { get; set; }

        bool UseConfigCredentials { get; }

        string Username { get; }
    }
}


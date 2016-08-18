namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface ICommunicationConfiguration
    {
        int DelayUntilJavaMethodCall { get; }

        string JavaScriptMethodName { get; }

        string PostMethodName { get; }

        SendDataMethod SendMethod { get; }
    }
}


namespace OpenEsdh.Outlook.Model.Container
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Alfresco;
    using OpenEsdh.Outlook.Model.Configuration.Implementation;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using OpenEsdh.Outlook.Presenters.Implementation;
    using OpenEsdh.Outlook.Views.Implementation.OfficeApplications;
    using OpenEsdh.Outlook.Views.ServerCertificate;
    using OpenEsdh.Outlook.Views.Interface;
    using Presenters.Interface;
    using System;
    using System.Configuration;
    using System.Net;
    using System.Net.Security;
    using System.Reflection;
    using System.Security.Cryptography.X509Certificates;
        
    public class WordResolver : TypeResolver, IDisposable
    {
        private Type _configurationFileOwner = null;
        private static bool CertificateAccepterInitialized = false;

        public WordResolver(Type ConfigurationFileOwner)
        {
            this._configurationFileOwner = ConfigurationFileOwner;
        }

        protected override void BuildComponents()
        {
            base.AddComponent<IAttachEmail>(() => new AttachEmail());
            base.AddComponent<IWordConfiguration>(delegate {
                try
                {
                    if (base._singletons.ContainsKey(typeof(IWordConfiguration)))
                    {
                        return base._singletons[typeof(IWordConfiguration)];
                    }
                    WordConfiguration section = (WordConfiguration) ConfigurationManager.OpenExeConfiguration(new Uri(Assembly.GetAssembly(this._configurationFileOwner).CodeBase).LocalPath).GetSection("Office");
                    base._singletons.Add(typeof(IWordConfiguration), section);
                    if (!(!section.IgnoreCertificateErrors || CertificateAccepterInitialized))
                    {
                        WindowsInterop.Hook();
                        ServicePointManager.ServerCertificateValidationCallback = (param0, param1, param2, param3) => true;
                        CertificateAccepterInitialized = true;
                    }
                    return section;
                }
                catch (Exception)
                {
                    return new WordConfiguration();
                }
            });
            base.AddComponent<IPreAuthenticator>(delegate {
                if (base._singletons.ContainsKey(typeof(IPreAuthenticator)))
                {
                    return base._singletons[typeof(IPreAuthenticator)];
                }
                PreAuthenticator authenticator = new PreAuthenticator(base.Create<IWordConfiguration>().PreAuthentication);
                base._singletons.Add(typeof(IPreAuthenticator), authenticator);
                return authenticator;
            });
            base.AddComponent<IApplicationSaveAsView>(() => new SaveAs());
            base.AddComponent<IAlfrescoFilePost>(delegate {
                IWordConfiguration configuration = base.Create<IWordConfiguration>();
                string url = configuration.UploadEndPoint;
                if (configuration.PreAuthenticate)
                {
                    url = TypeResolver.Current.Create<IPreAuthenticator>().AddAuthentificationUrl(url);
                }
                return new UploadFileApplication(url, true);
            });
            base.AddComponent<ICookieJar>(delegate {
                if (base._singletons.ContainsKey(typeof(ICookieJar)))
                {
                    return base._singletons[typeof(ICookieJar)];
                }
                CookieJar jar = new CookieJar();
                base._singletons.Add(typeof(ICookieJar), jar);
                return jar;
            });
            base.AddComponent<IApplicationSaveAsPresenter>(() => new ApplicationSaveAsPresenter());
            base.AddComponent<IApplicationSavePresenter>(() => new ApplicationSavePresenter());
        }

        public override void Dispose()
        {
            if (CertificateAccepterInitialized)
            {
                WindowsInterop.Unhook();
                CertificateAccepterInitialized = false;
            }
        }
    }
}


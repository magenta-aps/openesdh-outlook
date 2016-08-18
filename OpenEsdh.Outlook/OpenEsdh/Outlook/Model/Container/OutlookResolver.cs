namespace OpenEsdh.Outlook.Model.Container
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Alfresco;
    using OpenEsdh.Outlook.Model.Configuration.Implementation;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.ServerCertificate;
    using OpenEsdh.Outlook.Presenters.Implementation;
    using OpenEsdh.Outlook.Views.Implementation;
    using OpenEsdh.Outlook.Views.Interface;
    using OpenEsdh.Outlook.Views.ServerCertificate;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;
    using System.Configuration;
    using System.Net;
    using System.Net.Security;
    using System.Reflection;
    using System.Security.Cryptography.X509Certificates;

    public class OutlookResolver : TypeResolver, IDisposable
    {
        private Type _configurationFileOwner = null;
        private static bool CertificateAccepterInitialized = false;

        public OutlookResolver(Type ConfigurationFileOwner)
        {
            this._configurationFileOwner = ConfigurationFileOwner;
        }

        protected override void BuildComponents()
        {
            base.AddComponent<IAttachEmail>(() => new AttachEmail());
            base.AddComponent<IOutlookConfiguration>(delegate {
                try
                {
                    if (base._singletons.ContainsKey(typeof(IOutlookConfiguration)))
                    {
                        return base._singletons[typeof(IOutlookConfiguration)];
                    }
                    OutlookConfiguration section = (OutlookConfiguration) ConfigurationManager.OpenExeConfiguration(new Uri(Assembly.GetAssembly(this._configurationFileOwner).CodeBase).LocalPath).GetSection("Outlook");
                    base._singletons.Add(typeof(IOutlookConfiguration), section);
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
                    return new OutlookConfiguration();
                }
            });
            base.AddComponent<IPreAuthenticator>(delegate {
                if (base._singletons.ContainsKey(typeof(IPreAuthenticator)))
                {
                    return base._singletons[typeof(IPreAuthenticator)];
                }
                PreAuthenticator authenticator = new PreAuthenticator(base.Create<IOutlookConfiguration>().PreAuthentication);
                base._singletons.Add(typeof(IPreAuthenticator), authenticator);
                return authenticator;
            });
            base.AddComponent<ISaveAsView>(() => new OutlookSaveView());
            base.AddComponent<ISaveAsPresenter>(delegate {
                ISaveAsView view = base.Create<ISaveAsView>();
                SaveAsPresenter presenter = new SaveAsPresenter(view);
                view.Presenter = presenter;
                return presenter;
            });
            base.AddComponent<IDisplayRegionControl>(() => new OutlookPropertiesControl());
            base.AddComponentWithParam<IDisplayRegionPresenter>(inputParam => new DisplayRegionPresenter(inputParam as IDisplayRegion));
            base.AddComponent<IAlfrescoFilePost>(delegate {
                IOutlookConfiguration configuration = base.Create<IOutlookConfiguration>();
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
            base.AddComponent<IAttachFilePresenter>(() => new AttachFilePresenter());
            base.AddComponent<IAttachFileView>(() => new AttachFileView());
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


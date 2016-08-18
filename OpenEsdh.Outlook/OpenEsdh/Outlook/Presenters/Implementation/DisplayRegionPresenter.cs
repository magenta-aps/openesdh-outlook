namespace OpenEsdh.Outlook.Presenters.Implementation
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Alfresco;
    using OpenEsdh.Outlook.Model.Configuration.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Outlook.Views.Interface;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Forms;

    public class DisplayRegionPresenter : IDisplayRegionPresenter
    {
        private readonly IOutlookConfiguration _configuration;
        private readonly IDisplayRegion _displayRegion;
        private readonly IDisplayRegionControl _displayRegionControl;

        public DisplayRegionPresenter(IDisplayRegion displayRegion)
        {
            this._displayRegion = displayRegion;
            this._displayRegionControl = TypeResolver.Current.Create<IDisplayRegionControl>();
            this._configuration = TypeResolver.Current.Create<IOutlookConfiguration>();
            if (this._displayRegionControl is Control)
            {
                Control control = this._displayRegionControl as Control;
                this.DisplayRegion.FormControlCollection.Add(control);
                control.Dock = DockStyle.Fill;
            }
        }

        public void Show(EmailDescriptor email)
        {
            string displayDialogUrl = this._configuration.DisplayRegion.DisplayDialogUrl;
            string url = displayDialogUrl;
            IEnumerable<KeyValuePair<string, string>> source = from metadata in email.MetaData
                where metadata.Key == "OpenESDHID"
                select metadata;
            if (source.Any<KeyValuePair<string, string>>())
            {
                string package = source.First<KeyValuePair<string, string>>().Value;
                UrlTokenReplacer replacer = new UrlTokenReplacer(displayDialogUrl, package);
                url = replacer.Url;
            }
            if (this._configuration.PreAuthenticate)
            {
                url = TypeResolver.Current.Create<IPreAuthenticator>().AddAuthentificationUrl(url);
            }
            this._displayRegionControl.Show(url);
        }

        public IDisplayRegion DisplayRegion
        {
            get
            {
                return this._displayRegion;
            }
        }
    }
}


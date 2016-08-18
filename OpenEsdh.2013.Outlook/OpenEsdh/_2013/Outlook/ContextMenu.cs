namespace OpenEsdh._2013.Outlook
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Outlook;
    using OpenEsdh._2013.Outlook.Presentation.Interface;
    using OpenEsdh._2013.Outlook.Properties;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Model.Resources;
    using System;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    public class ContextMenu : Microsoft.Office.Core.IRibbonExtensibility, ISaveEmailButtonView
    {
        private Microsoft.Office.Core.IRibbonUI ribbon;

        public void btnAttachFile(Microsoft.Office.Core.IRibbonControl control)
        {
            Exception exception;
            try
            {
                MailItem item = ((dynamic) control.Context).CurrentItem as MailItem;
                IAttachFilePresenter presenter = TypeResolver.Current.Create<IAttachFilePresenter>();
                try
                {
                    presenter.AttachFileClick(item);
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
            }
            catch (Exception exception2)
            {
                exception = exception2;
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public void btnSaveAsSend(Microsoft.Office.Core.IRibbonControl control)
        {
            Exception exception;
            try
            {
                Action sendOperation = null;
                ISaveEmailPresenter presenter = TypeResolver.Current.Create<ISaveEmailPresenter>();
                presenter.View = this;
                MailItem item = ((dynamic) control.Context).CurrentItem as MailItem;
                try
                {
                    if (sendOperation == null)
                    {
                        sendOperation = () => item.Send();
                    }
                    bool flag = presenter.SaveEmailAndSend(item, sendOperation);
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
            }
            catch (Exception exception2)
            {
                exception = exception2;
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public void btnSaveFile(Microsoft.Office.Core.IRibbonControl control)
        {
            Exception exception;
            try
            {
                ISaveEmailPresenter presenter = TypeResolver.Current.Create<ISaveEmailPresenter>();
                presenter.View = this;
                try
                {
                    MailItem item = ((dynamic) control.Context).CurrentItem as MailItem;
                    if (item != null)
                    {
                        presenter.SaveEmailClick(item);
                        try
                        {
                            Microsoft.Office.Interop.Outlook.Inspector getInspector = item.GetInspector;
                            if (getInspector != null)
                            {
                                getInspector.ShowFormPage(ResourceResolver.Current.GetString("ViewRegionTitle"));
                                getInspector.Display(Missing.Value);
                            }
                        }
                        catch
                        {
                        }
                    }
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                    throw exception;
                }
            }
            catch (Exception exception2)
            {
                exception = exception2;
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public string GetAttachFileLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            return ResourceResolver.Current.GetString("AttachFileBtn");
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OpenEsdh._2013.Outlook.ContextMenu.xml");
        }

        public bool GetEnabled(Microsoft.Office.Core.IRibbonControl control)
        {
            bool flag = false;
            try
            {
                MailItem item = ((dynamic) control.Context).CurrentItem as MailItem;
                if (item != null)
                {
                    flag = true;
                }
            }
            catch
            {
            }
            return flag;
        }

        public string GetGroupLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            return ResourceResolver.Current.GetString("OpenESDHButtonGroup");
        }

        public Image getImage(Microsoft.Office.Core.IRibbonControl control)
        {
            return Resources.VismaCase16x16;
        }

        public Image getImageLarge(Microsoft.Office.Core.IRibbonControl control)
        {
            return Resources.VismaCase32x32;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            string[] manifestResourceNames = executingAssembly.GetManifestResourceNames();
            for (int i = 0; i < manifestResourceNames.Length; i++)
            {
                if (string.Compare(resourceName, manifestResourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader reader = new StreamReader(executingAssembly.GetManifestResourceStream(manifestResourceNames[i])))
                    {
                        if (reader != null)
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public string GetSaveFileLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            return ResourceResolver.Current.GetString("SaveBtn");
        }

        public string GetSaveSendLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            return ResourceResolver.Current.GetString("SaveSendBtn");
        }

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void SaveToOpenESDH(Microsoft.Office.Core.IRibbonControl control)
        {
            try
            {
                if ((Globals.ThisAddIn.Application.ActiveExplorer() != null) && (((Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.DefaultItemType == OlItemType.olMailItem) && (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Items.Count > 0)) || (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Store.ExchangeStoreType == OlExchangeStoreType.olExchangePublicFolder)))
                {
                    ISaveEmailPresenter presenter = TypeResolver.Current.Create<ISaveEmailPresenter>();
                    presenter.View = this;
                    MailItem item = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1] as MailItem;
                    presenter.SaveEmailClick(item);
                }
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                throw exception;
            }
        }

        public bool Visible
        {
            get
            {
                return true;
            }
            set
            {
            }
        }
    }
}


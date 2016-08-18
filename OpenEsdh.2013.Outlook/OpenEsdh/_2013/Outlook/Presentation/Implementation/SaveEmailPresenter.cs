namespace OpenEsdh._2013.Outlook.Presentation.Implementation
{
    using Microsoft.CSharp.RuntimeBinder;
    using Microsoft.Office.Interop.Outlook;
    using OpenEsdh._2013.Outlook.Model;
    using OpenEsdh._2013.Outlook.Presentation.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;
    using System.IO;
    using System.Linq.Expressions;
    using System.Runtime.CompilerServices;

    public class SaveEmailPresenter : ISaveEmailPresenter
    {
        public void Load([Dynamic] object Context)
        {
            this.View.Visible = true;
            if (<Load>o__SiteContainer10.<>p__Site11 == null)
            {
                <Load>o__SiteContainer10.<>p__Site11 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(SaveEmailPresenter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
            }
            if (<Load>o__SiteContainer10.<>p__Site12 == null)
            {
                <Load>o__SiteContainer10.<>p__Site12 = CallSite<Func<CallSite, object, object, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof(SaveEmailPresenter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.Constant, null) }));
            }
            if (<Load>o__SiteContainer10.<>p__Site11.Target(<Load>o__SiteContainer10.<>p__Site11, <Load>o__SiteContainer10.<>p__Site12.Target(<Load>o__SiteContainer10.<>p__Site12, ((dynamic) Context).CurrentItem, null)))
            {
                MailItem item = ((dynamic) Context).CurrentItem as MailItem;
                if (item == null)
                {
                    this.View.Visible = false;
                }
            }
        }

        public bool SaveEmailAndSend(MailItem item, Action SendOperation)
        {
            SetMessageClassDelegate delegate2 = null;
            SetMessageIDDelegate delegate3 = null;
            SaveAttachmentDelegate delegate4 = null;
            if (item != null)
            {
                ISaveAsPresenter presenter = TypeResolver.Current.Create<ISaveAsPresenter>();
                if (item != null)
                {
                    if (delegate2 == null)
                    {
                        delegate2 = delegate (string messageClass) {
                            item.MessageClass = messageClass;
                            item.Save();
                        };
                    }
                    presenter.SetMessageClass += delegate2;
                    if (delegate3 == null)
                    {
                        delegate3 = delegate (string messageID) {
                            try
                            {
                                if (item.ItemProperties["OpenESDHID"] == null)
                                {
                                    item.ItemProperties.Add("OpenESDHID", OlUserPropertyType.olText, Type.Missing, Type.Missing);
                                }
                                item.ItemProperties["OpenESDHID"].Value = messageID;
                            }
                            catch (Exception exception)
                            {
                                Logger.Current.LogException(exception, "");
                            }
                            item.Save();
                        };
                    }
                    presenter.SetMessageID += delegate3;
                    if (delegate4 == null)
                    {
                        delegate4 = delegate (string name, UploadMailFileDelegate Upload) {
                            string path = "";
                            string tempFileName = Path.GetTempFileName();
                            for (int j = 1; j <= item.Attachments.Count; j++)
                            {
                                Attachment attachment = item.Attachments[j];
                                if (attachment.DisplayName == name)
                                {
                                    string extension = Path.GetExtension(attachment.FileName);
                                    if (!extension.StartsWith("."))
                                    {
                                        extension = "." + extension;
                                    }
                                    path = tempFileName + extension;
                                    attachment.SaveAsFile(path);
                                    Upload(path, attachment.FileName);
                                    break;
                                }
                            }
                            if (!string.IsNullOrEmpty(path))
                            {
                                try
                                {
                                    File.Delete(path);
                                }
                                catch (Exception exception)
                                {
                                    Logger.Current.LogException(exception, "");
                                }
                            }
                        };
                    }
                    presenter.SaveAttachment += delegate4;
                    bool flag = presenter.ShowAndSend(item.ToMailDescriptor());
                    if (flag && (SendOperation != null))
                    {
                        SendOperation();
                    }
                    return flag;
                }
            }
            return false;
        }

        public void SaveEmailClick(MailItem item)
        {
            SetMessageClassDelegate delegate2 = null;
            SetMessageIDDelegate delegate3 = null;
            SaveAttachmentDelegate delegate4 = null;
            if (item != null)
            {
                ISaveAsPresenter presenter = TypeResolver.Current.Create<ISaveAsPresenter>();
                if (item != null)
                {
                    if (delegate2 == null)
                    {
                        delegate2 = messageClass => item.MessageClass = messageClass;
                    }
                    presenter.SetMessageClass += delegate2;
                    if (delegate3 == null)
                    {
                        delegate3 = delegate (string messageID) {
                            try
                            {
                                if (item.ItemProperties["OpenESDHID"] == null)
                                {
                                    item.ItemProperties.Add("OpenESDHID", OlUserPropertyType.olText, Type.Missing, Type.Missing);
                                }
                                item.ItemProperties["OpenESDHID"].Value = messageID;
                            }
                            catch (Exception exception)
                            {
                                Logger.Current.LogException(exception, "");
                            }
                            item.Save();
                        };
                    }
                    presenter.SetMessageID += delegate3;
                    if (delegate4 == null)
                    {
                        delegate4 = delegate (string name, UploadMailFileDelegate Upload) {
                            string path = "";
                            string tempFileName = Path.GetTempFileName();
                            for (int j = 1; j <= item.Attachments.Count; j++)
                            {
                                Attachment attachment = item.Attachments[j];
                                if (attachment.DisplayName == name)
                                {
                                    string extension = Path.GetExtension(attachment.FileName);
                                    if (!extension.StartsWith("."))
                                    {
                                        extension = "." + extension;
                                    }
                                    path = tempFileName + extension;
                                    attachment.SaveAsFile(path);
                                    Upload(path, attachment.FileName);
                                    break;
                                }
                            }
                            if (!string.IsNullOrEmpty(path))
                            {
                                try
                                {
                                    File.Delete(path);
                                }
                                catch (Exception exception)
                                {
                                    Logger.Current.LogException(exception, "");
                                }
                            }
                        };
                    }
                    presenter.SaveAttachment += delegate4;
                    presenter.Show(item.ToMailDescriptor());
                }
            }
        }

        public ISaveEmailButtonView View { get; set; }
    }
}


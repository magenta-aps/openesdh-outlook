namespace OpenEsdh._2013.Outlook.Presentation.Implementation
{
    using Microsoft.Office.Interop.Outlook;
    using OpenEsdh._2013.Outlook.Model;
    using OpenEsdh._2013.Outlook.Presentation.Interface;
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Container;
    using System;
    using System.Reflection;

    public class AttachFilePresenter : IAttachFilePresenter
    {
        public void AttachFileClick(MailItem _item)
        {
            TypeResolver.Current.Create<IAttachFilePresenter>().Initialize(_item.ToMailDescriptor(), delegate (string descriptor) {
                IAttachEmail email = TypeResolver.Current.Create<IAttachEmail>();
                string[] configurationSettings = descriptor.Replace("\r", "").Split(new char[] { '\n' });
                email.AddAttachmentConfiguration(configurationSettings, delegate (string PropertyName, string value) {
                    switch (PropertyName.ToLower())
                    {
                        case "to":
                            _item.To = value;
                            break;

                        case "cc":
                            _item.CC = value;
                            break;

                        case "bcc":
                            _item.BCC = value;
                            break;

                        case "subject":
                            _item.Subject = value;
                            break;

                        case "htmlbody":
                            _item.HTMLBody = value;
                            break;

                        case "body":
                            _item.Body = value;
                            break;
                    }
                }, delegate (string PropertyName, string value) {
                    switch (PropertyName.ToLower())
                    {
                        case "htmlbody+":
                            if (!string.IsNullOrEmpty(_item.HTMLBody))
                            {
                                _item.HTMLBody = _item.HTMLBody + "\r\n";
                            }
                            _item.HTMLBody = _item.HTMLBody + value;
                            break;

                        case "body+":
                            if (!string.IsNullOrEmpty(_item.Body))
                            {
                                _item.Body = _item.Body + "\r\n";
                            }
                            _item.Body = _item.Body + value;
                            break;
                    }
                }, FileName => _item.Attachments.Add(FileName, Missing.Value, Missing.Value, Missing.Value));
            });
        }
    }
}


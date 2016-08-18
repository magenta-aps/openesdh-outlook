namespace OpenEsdh._2013.Outlook.Model
{
    using Microsoft.CSharp.RuntimeBinder;
    using Microsoft.Office.Interop.Outlook;
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Runtime.CompilerServices;

    public static class MailConverter
    {
        public static EmailDescriptor ToMailDescriptor(this MailItem mailItem)
        {
            EmailDescriptor descriptor = new EmailDescriptor {
                Subject = mailItem.Subject,
                BodyText = mailItem.Body,
                BodyHtml = mailItem.HTMLBody
            };
            if (mailItem.Sender != null)
            {
                descriptor.From = new EmailAddress(mailItem.Sender.Address);
            }
            else
            {
                descriptor.From = new EmailAddress();
            }
            if (mailItem.To != null)
            {
                foreach (string str in mailItem.To.Split(new char[] { ';' }))
                {
                    descriptor.To.Add(new EmailAddress(str));
                }
            }
            if (mailItem.CC != null)
            {
                foreach (string str in mailItem.CC.Split(new char[] { ';' }))
                {
                    descriptor.CC.Add(new EmailAddress(str));
                }
            }
            if (mailItem.BCC != null)
            {
                foreach (string str in mailItem.BCC.Split(new char[] { ';' }))
                {
                    descriptor.BCC.Add(new EmailAddress(str));
                }
            }
            if (mailItem.Attachments != null)
            {
                foreach (Microsoft.Office.Interop.Outlook.Attachment attachment in mailItem.Attachments)
                {
                    descriptor.Attachments.Add(new OpenEsdh.Outlook.Model.Attachment(attachment.DisplayName, OpenEsdh.Outlook.Model.Attachment.GetMimeType(attachment.FileName)));
                }
            }
            if (mailItem.ItemProperties != null)
            {
                foreach (ItemProperty property in mailItem.ItemProperties)
                {
                    try
                    {
                        if (<ToMailDescriptor>o__SiteContainer0.<>p__Site1 == null)
                        {
                            <ToMailDescriptor>o__SiteContainer0.<>p__Site1 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(MailConverter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                        }
                        if (<ToMailDescriptor>o__SiteContainer0.<>p__Site2 == null)
                        {
                            <ToMailDescriptor>o__SiteContainer0.<>p__Site2 = CallSite<Func<CallSite, object, object, object>>.Create(Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof(MailConverter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.Constant, null) }));
                        }
                        if (<ToMailDescriptor>o__SiteContainer0.<>p__Site1.Target(<ToMailDescriptor>o__SiteContainer0.<>p__Site1, <ToMailDescriptor>o__SiteContainer0.<>p__Site2.Target(<ToMailDescriptor>o__SiteContainer0.<>p__Site2, property.Value, null)))
                        {
                            string str2 = (string) ((dynamic) property.Value).ToString();
                            descriptor.MetaData.Add(new KeyValuePair<string, string>(property.Name, str2));
                        }
                    }
                    catch (Exception exception)
                    {
                        Logger.Current.LogException(exception, "");
                    }
                }
            }
            return descriptor;
        }
    }
}


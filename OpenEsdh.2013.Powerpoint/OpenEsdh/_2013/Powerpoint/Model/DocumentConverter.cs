namespace OpenEsdh._2013.Powerpoint.Model
{
    using Microsoft.CSharp.RuntimeBinder;
    using Microsoft.Office.Interop.PowerPoint;
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq.Expressions;
    using System.Runtime.CompilerServices;
    using System.Threading;

    public static class DocumentConverter
    {
        public static ApplicationDescriptor ToDescriptor(this Presentation document)
        {
            try
            {
                ApplicationDescriptor descriptor = new ApplicationDescriptor {
                    Author = Thread.CurrentPrincipal.Identity.Name,
                    Name = document.Path
                };
                foreach (dynamic obj2 in (IEnumerable) document.BuiltInDocumentProperties)
                {
                    try
                    {
                        if (<ToDescriptor>o__SiteContainer0.<>p__Site2 == null)
                        {
                            <ToDescriptor>o__SiteContainer0.<>p__Site2 = CallSite<Func<CallSite, Type, object, object, KeyValuePair<string, string>>>.Create(Binder.InvokeConstructor(CSharpBinderFlags.None, typeof(DocumentConverter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.IsStaticType | CSharpArgumentInfoFlags.UseCompileTimeType, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                        }
                        descriptor.MetaData.Add(<ToDescriptor>o__SiteContainer0.<>p__Site2.Target(<ToDescriptor>o__SiteContainer0.<>p__Site2, typeof(KeyValuePair<string, string>), obj2.Name, obj2.Value.ToString()));
                    }
                    catch
                    {
                    }
                }
                foreach (dynamic obj2 in (IEnumerable) document.CustomDocumentProperties)
                {
                    try
                    {
                        if (<ToDescriptor>o__SiteContainer0.<>p__Site7 == null)
                        {
                            <ToDescriptor>o__SiteContainer0.<>p__Site7 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(DocumentConverter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                        }
                        if (<ToDescriptor>o__SiteContainer0.<>p__Site7.Target(<ToDescriptor>o__SiteContainer0.<>p__Site7, obj2.Name == "OpenESDHID"))
                        {
                            descriptor.ID = (string) obj2.Value.ToString;
                        }
                        if (<ToDescriptor>o__SiteContainer0.<>p__Sited == null)
                        {
                            <ToDescriptor>o__SiteContainer0.<>p__Sited = CallSite<Func<CallSite, Type, object, object, KeyValuePair<string, string>>>.Create(Binder.InvokeConstructor(CSharpBinderFlags.None, typeof(DocumentConverter), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.IsStaticType | CSharpArgumentInfoFlags.UseCompileTimeType, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                        }
                        descriptor.MetaData.Add(<ToDescriptor>o__SiteContainer0.<>p__Sited.Target(<ToDescriptor>o__SiteContainer0.<>p__Sited, typeof(KeyValuePair<string, string>), obj2.Name, obj2.Value.ToString()));
                    }
                    catch
                    {
                    }
                }
                return descriptor;
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                return null;
            }
        }
    }
}


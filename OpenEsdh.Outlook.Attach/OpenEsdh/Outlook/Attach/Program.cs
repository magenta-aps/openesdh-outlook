namespace OpenEsdh.Outlook.Attach
{
    using Microsoft.CSharp.RuntimeBinder;
    using Microsoft.Office.Interop.Outlook;
    using Microsoft.VisualBasic;
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;

    internal class Program
    {
        [MTAThread]
        private static void Main(string[] args)
        {
            TextWriter @out = Console.Out;
            try
            {
                SetMailPropertyDelegate setProperty = null;
                AddMailPropertyDelegate addProperty = null;
                AddFileDelegate addFile = null;
                Logger.Current.LogInformation("Application Startup:" + args[0], "");
                TypeResolver.Current = new WordResolver(typeof(Program));
                Application activeObject = null;
                MailItem _item = null;
                if ((args != null) && (args.Length > 0))
                {
                    Exception exception;
                    try
                    {
                        Process[] processesByName = Process.GetProcessesByName("OUTLOOK");
                        if (processesByName.Count<Process>() > 0)
                        {
                            Process process = processesByName[0];
                            try
                            {
                                List<object> runningInstances = Utils.GetRunningInstances(new string[] { "Outlook.Application", "Outlook.Application.15" });
                                foreach (object obj2 in runningInstances)
                                {
                                    string name = obj2.GetType().Name;
                                }
                            }
                            catch
                            {
                                activeObject = Interaction.CreateObject("Outlook.Application", "") as Application;
                            }
                            activeObject = (Application) Marshal.GetActiveObject("Outlook.Application");
                        }
                    }
                    catch (Exception exception1)
                    {
                        exception = exception1;
                        Logger.Current.LogException(exception, "");
                    }
                    if (activeObject == null)
                    {
                        activeObject = (Application) Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("0006F03A-0000-0000-C000-000000000046")));
                    }
                    try
                    {
                        if (activeObject != null)
                        {
                            if (<Main>o__SiteContainer0.<>p__Site1 == null)
                            {
                                <Main>o__SiteContainer0.<>p__Site1 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsTrue, typeof(Program), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                            }
                            if (<Main>o__SiteContainer0.<>p__Site2 == null)
                            {
                                <Main>o__SiteContainer0.<>p__Site2 = CallSite<Func<CallSite, object, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.None, ExpressionType.NotEqual, typeof(Program), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.Constant, null) }));
                            }
                            object obj3 = <Main>o__SiteContainer0.<>p__Site2.Target(<Main>o__SiteContainer0.<>p__Site2, activeObject.ActiveWindow(), null);
                            if (<Main>o__SiteContainer0.<>p__Site3 == null)
                            {
                                <Main>o__SiteContainer0.<>p__Site3 = CallSite<Func<CallSite, object, bool>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.UnaryOperation(CSharpBinderFlags.None, ExpressionType.IsFalse, typeof(Program), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                            }
                            if (!<Main>o__SiteContainer0.<>p__Site3.Target(<Main>o__SiteContainer0.<>p__Site3, obj3) && (<Main>o__SiteContainer0.<>p__Site4 == null))
                            {
                                <Main>o__SiteContainer0.<>p__Site4 = CallSite<Func<CallSite, object, object, object>>.Create(Microsoft.CSharp.RuntimeBinder.Binder.BinaryOperation(CSharpBinderFlags.BinaryOperationLogical, ExpressionType.And, typeof(Program), new CSharpArgumentInfo[] { CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null), CSharpArgumentInfo.Create(CSharpArgumentInfoFlags.None, null) }));
                            }
                            if (<Main>o__SiteContainer0.<>p__Site1.Target(<Main>o__SiteContainer0.<>p__Site1, (<Main>o__SiteContainer0.<>p__Site5 != null) ? obj3 : <Main>o__SiteContainer0.<>p__Site4.Target(<Main>o__SiteContainer0.<>p__Site4, obj3, <Main>o__SiteContainer0.<>p__Site5.Target(<Main>o__SiteContainer0.<>p__Site5, activeObject.ActiveWindow(), activeObject.ActiveInspector()))))
                            {
                                _item = activeObject.ActiveInspector().CurrentItem as MailItem;
                            }
                            if (_item == null)
                            {
                                _item = (MailItem) activeObject.CreateItem(OlItemType.olMailItem);
                                _item.Display(Missing.Value);
                            }
                            string[] configurationSettings = File.ReadAllText(args[0]).Replace("##", "\x00a4").Split(new char[] { '\x00a4' });
                            if (setProperty == null)
                            {
                                setProperty = delegate (string PropertyName, string value) {
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
                                };
                            }
                            if (addProperty == null)
                            {
                                addProperty = delegate (string PropertyName, string value) {
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
                                };
                            }
                            if (addFile == null)
                            {
                                addFile = FileName => _item.Attachments.Add(FileName, Missing.Value, Missing.Value, Missing.Value);
                            }
                            TypeResolver.Current.Create<IAttachEmail>().AddAttachmentConfiguration(configurationSettings, setProperty, addProperty, addFile);
                        }
                    }
                    catch (Exception exception2)
                    {
                        exception = exception2;
                        Logger.Current.LogException(exception, "");
                    }
                }
            }
            finally
            {
                Logger.Current.LogInformation("Application Stopping", "");
                Console.SetOut(@out);
            }
        }
    }
}


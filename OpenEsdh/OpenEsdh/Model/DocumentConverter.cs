namespace OpenEsdh.Model
{
    using OpenEsdh.Outlook.Model;
    using OpenEsdh.Outlook.Model.Logging;
    using System;
    using System.IO;
    using System.Threading;

    public class DocumentConverter
    {
        public static ApplicationDescriptor ToDescriptor(string fileName)
        {
            try
            {
                return new ApplicationDescriptor { 
                    Name = fileName,
                    Author = Thread.CurrentPrincipal.Identity.Name,
                    Title = Path.GetFileName(fileName)
                };
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
                return null;
            }
        }
    }
}


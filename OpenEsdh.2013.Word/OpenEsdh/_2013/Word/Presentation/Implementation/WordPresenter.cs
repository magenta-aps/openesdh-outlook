namespace OpenEsdh._2013.Word.Presentation.Implementation
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Word;
    using OpenEsdh._2013.Word.Model;
    using OpenEsdh._2013.Word.Presentation.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;
    using System.IO;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class WordPresenter : IWordPresenter
    {
        public WordPresenter(IWordView view)
        {
            this.View = view;
        }

        public void Load(Microsoft.Office.Interop.Word.Document document)
        {
            if (document != null)
            {
                this.SetSaveEnabled(document);
            }
            else
            {
                this.View.SaveAsEnabled = false;
                this.View.SaveEnabled = false;
            }
        }

        private string ReadDocumentProperty(DocumentProperties properties, string propertyName)
        {
            foreach (DocumentProperty property in properties)
            {
                if (property.Name == propertyName)
                {
                    return (string) ((dynamic) property.Value).ToString();
                }
            }
            return null;
        }

        public void Save(dynamic Context)
        {
            Exception exception;
            try
            {
                try
                {
                    SaveDocumentDelegate delegate2 = null;
                    SetDocumentIDDelegate delegate3 = null;
                    this.View.ViewIsLocked = true;
                    IApplicationSavePresenter presenter = TypeResolver.Current.Create<IApplicationSavePresenter>();
                    Microsoft.Office.Interop.Word.Document document = (Microsoft.Office.Interop.Word.Document) ((dynamic) Context).Document;
                    if (document != null)
                    {
                        if (delegate2 == null)
                        {
                            delegate2 = delegate (UploadMailFileDelegate Upload) {
                                if (string.IsNullOrEmpty(document.Path))
                                {
                                    string str = Path.GetTempFileName() + ".docx";
                                    object fileName = str;
                                    object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;
                                    document.SaveAs2(ref fileName, ref fileFormat);
                                }
                                else
                                {
                                    document.Save();
                                }
                                Upload(document.Path, document.Name);
                            };
                        }
                        presenter.SaveDocument += delegate2;
                        if (delegate3 == null)
                        {
                            delegate3 = delegate (string ID) {
                                DocumentProperties properties = document.CustomDocumentProperties as DocumentProperties;
                                if (properties != null)
                                {
                                    if (this.ReadDocumentProperty(properties, "OpenESDHID") != null)
                                    {
                                        properties["OpenESDHID"].Delete();
                                    }
                                    properties.Add("OpenESDHID", false, MsoDocProperties.msoPropertyTypeString, ID, Missing.Value);
                                }
                            };
                        }
                        presenter.SetDocumentID += delegate3;
                    }
                    presenter.Show(document.ToDescriptor());
                }
                catch (Exception exception1)
                {
                    exception = exception1;
                    Logger.Current.LogException(exception, "");
                }
                finally
                {
                    this.View.ViewIsLocked = false;
                }
            }
            catch (Exception exception2)
            {
                exception = exception2;
                Logger.Current.LogException(exception, "");
            }
            finally
            {
                this.View.ViewIsLocked = false;
            }
        }

        public void SaveAs(dynamic Context)
        {
            try
            {
                SaveDocumentDelegate delegate2 = null;
                SetDocumentIDDelegate delegate3 = null;
                this.View.ViewIsLocked = true;
                IApplicationSaveAsPresenter presenter = TypeResolver.Current.Create<IApplicationSaveAsPresenter>();
                Microsoft.Office.Interop.Word.Document document = (Microsoft.Office.Interop.Word.Document) ((dynamic) Context).Document;
                if (document != null)
                {
                    if (delegate2 == null)
                    {
                        delegate2 = delegate (UploadMailFileDelegate Upload) {
                            if (string.IsNullOrEmpty(document.Path))
                            {
                                string str = Path.ChangeExtension(Path.GetTempFileName(), "docx");
                                object fileName = str;
                                object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
                                document.SaveAs(ref fileName, ref fileFormat);
                            }
                            else
                            {
                                document.Save();
                            }
                            string fullPathName = Path.Combine(document.Path, document.Name);
                            string name = document.Name;
                            Microsoft.Office.Interop.Word.Application application = document.Application;
                            try
                            {
                                Upload(fullPathName, name);
                            }
                            finally
                            {
                                document.Activate();
                                document.Save();
                            }
                        };
                    }
                    presenter.SaveDocument += delegate2;
                    if (delegate3 == null)
                    {
                        delegate3 = delegate (string ID) {
                            DocumentProperties properties = document.CustomDocumentProperties as DocumentProperties;
                            if (properties != null)
                            {
                                if (this.ReadDocumentProperty(properties, "OpenESDHID") != null)
                                {
                                    properties["OpenESDHID"].Delete();
                                }
                                properties.Add("OpenESDHID", false, MsoDocProperties.msoPropertyTypeString, ID, Missing.Value);
                            }
                        };
                    }
                    presenter.SetDocumentID += delegate3;
                }
                presenter.Show(document.ToDescriptor());
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
            finally
            {
                this.View.ViewIsLocked = false;
            }
        }

        private void SetSaveEnabled(Microsoft.Office.Interop.Word.Document document)
        {
            if (document != null)
            {
                DocumentProperties customDocumentProperties = document.CustomDocumentProperties as DocumentProperties;
                if ((!string.IsNullOrEmpty(document.Path) && (customDocumentProperties != null)) && (this.ReadDocumentProperty(customDocumentProperties, "OpenESDHID") != null))
                {
                    this.View.SaveEnabled = true;
                }
                else
                {
                    this.View.SaveEnabled = false;
                }
                if (document.ProtectionType != Microsoft.Office.Interop.Word.WdProtectionType.wdNoProtection)
                {
                    this.View.SaveAsEnabled = false;
                }
                else
                {
                    this.View.SaveAsEnabled = true;
                }
            }
        }

        public IWordView View { get; set; }
    }
}


namespace OpenEsdh._2013.Excel.Presentation.Implementation
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;
    using OpenEsdh._2013.Excel.Model;
    using OpenEsdh._2013.Excel.Presentation.Interface;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using System;
    using System.IO;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class ExcelPresenter : IExcelPresenter
    {
        public ExcelPresenter(IExcelView view)
        {
            this.View = view;
        }

        public void Load(Microsoft.Office.Interop.Excel.Workbook document)
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

        public void Save([Dynamic] object Context)
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
                    Microsoft.Office.Interop.Excel.Workbook document = (Microsoft.Office.Interop.Excel.Workbook) ((dynamic) Context).Parent;
                    if (document != null)
                    {
                        if (delegate2 == null)
                        {
                            delegate2 = delegate (UploadMailFileDelegate Upload) {
                                if (string.IsNullOrEmpty(document.Path))
                                {
                                    string filename = Path.ChangeExtension(Path.GetTempFileName(), "xlsx");
                                    document.SaveAs(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                                }
                                else
                                {
                                    document.Save();
                                }
                                string fullPathName = Path.Combine(document.Path, document.Name);
                                string name = document.Name;
                                Microsoft.Office.Interop.Excel.Application application = document.Application;
                                try
                                {
                                    document.Close(Missing.Value, Missing.Value, Missing.Value);
                                    Upload(fullPathName, name);
                                }
                                finally
                                {
                                    document = application.Workbooks.Open(fullPathName, Missing.Value, false, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                                    document.Activate();
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

        public void SaveAs([Dynamic] object Context)
        {
            try
            {
                SaveDocumentDelegate delegate2 = null;
                SetDocumentIDDelegate delegate3 = null;
                this.View.ViewIsLocked = true;
                IApplicationSaveAsPresenter presenter = TypeResolver.Current.Create<IApplicationSaveAsPresenter>();
                Microsoft.Office.Interop.Excel.Workbook document = (Microsoft.Office.Interop.Excel.Workbook) ((dynamic) Context).Parent;
                if (document != null)
                {
                    if (delegate2 == null)
                    {
                        delegate2 = delegate (UploadMailFileDelegate Upload) {
                            if (string.IsNullOrEmpty(document.Path))
                            {
                                string filename = Path.ChangeExtension(Path.GetTempFileName(), "xlsx");
                                document.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            }
                            else
                            {
                                document.Save();
                            }
                            string fullPathName = Path.Combine(document.Path, document.Name);
                            string name = document.Name;
                            Microsoft.Office.Interop.Excel.Application application = document.Application;
                            try
                            {
                                Upload(fullPathName, name);
                            }
                            finally
                            {
                                document.Activate();
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

        private void SetSaveEnabled(Microsoft.Office.Interop.Excel.Workbook document)
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
                this.View.SaveAsEnabled = true;
            }
        }

        public IExcelView View { get; set; }
    }
}


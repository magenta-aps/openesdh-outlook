namespace OpenEsdh.Presentation.Implementation
{
    using OpenEsdh.Model;
    using OpenEsdh.Outlook.Model.Container;
    using OpenEsdh.Outlook.Model.Logging;
    using OpenEsdh.Outlook.Presenters.Interface;
    using OpenEsdh.Presentation.Interface;
    using System;
    using System.IO;

    public class ExplorerPresenter : IExplorerPresenter
    {
        private IExplorerView _view;

        public ExplorerPresenter()
        {
        }

        public ExplorerPresenter(IExplorerView view)
        {
            this._view = view;
        }

        public void SaveAs(string filename)
        {
            SaveDocumentDelegate delegate2 = null;
            try
            {
                IApplicationSaveAsPresenter presenter = TypeResolver.Current.Create<IApplicationSaveAsPresenter>();
                if (delegate2 == null)
                {
                    delegate2 = Upload => Upload(filename, Path.GetFileName(filename));
                }
                presenter.SaveDocument += delegate2;
                presenter.SetDocumentID += delegate (string ID) {
                };
                presenter.Show(DocumentConverter.ToDescriptor(filename));
            }
            catch (Exception exception)
            {
                Logger.Current.LogException(exception, "");
            }
        }
    }
}


namespace OpenEsdh._2013.Word.Presentation.Interface
{
    using Microsoft.Office.Interop.Word;
    using System;
    using System.Runtime.CompilerServices;

    public interface IWordPresenter
    {
        void Load(Microsoft.Office.Interop.Word.Document document);
        void Save(dynamic Context);
        void SaveAs(dynamic Context);

        IWordView View { get; set; }
    }
}


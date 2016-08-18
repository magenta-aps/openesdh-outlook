namespace OpenEsdh._2013.Powerpoint.Presentation.Interface
{
    using Microsoft.Office.Interop.PowerPoint;
    using System;
    using System.Runtime.CompilerServices;

    public interface IPowerpointPresenter
    {
        void Load(Presentation document);
        void Save([Dynamic] object Context);
        void SaveAs([Dynamic] object Context);

        IPowerpointView View { get; set; }
    }
}


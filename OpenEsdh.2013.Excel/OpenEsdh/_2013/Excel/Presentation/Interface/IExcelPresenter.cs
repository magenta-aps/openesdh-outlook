namespace OpenEsdh._2013.Excel.Presentation.Interface
{
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Runtime.CompilerServices;

    public interface IExcelPresenter
    {
        void Load(Microsoft.Office.Interop.Excel.Workbook document);
        void Save([Dynamic] object Context);
        void SaveAs([Dynamic] object Context);

        IExcelView View { get; set; }
    }
}


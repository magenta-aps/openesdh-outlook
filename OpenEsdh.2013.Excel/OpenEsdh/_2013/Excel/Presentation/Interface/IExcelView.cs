namespace OpenEsdh._2013.Excel.Presentation.Interface
{
    using System;

    public interface IExcelView
    {
        bool SaveAsEnabled { get; set; }

        bool SaveEnabled { get; set; }

        bool ViewIsLocked { get; set; }
    }
}


namespace OpenEsdh._2013.Word.Presentation.Interface
{
    using System;

    public interface IWordView
    {
        bool SaveAsEnabled { get; set; }

        bool SaveEnabled { get; set; }

        bool ViewIsLocked { get; set; }
    }
}


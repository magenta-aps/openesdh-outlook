namespace OpenEsdh._2013.Powerpoint.Presentation.Interface
{
    using System;

    public interface IPowerpointView
    {
        bool SaveAsEnabled { get; set; }

        bool SaveEnabled { get; set; }

        bool ViewIsLocked { get; set; }
    }
}


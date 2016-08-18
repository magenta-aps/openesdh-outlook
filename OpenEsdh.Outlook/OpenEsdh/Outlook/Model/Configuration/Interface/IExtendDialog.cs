namespace OpenEsdh.Outlook.Model.Configuration.Interface
{
    using System;

    public interface IExtendDialog
    {
        int MaxHeight { get; }

        int MaxWidth { get; }

        int X { get; }

        int Y { get; }
    }
}


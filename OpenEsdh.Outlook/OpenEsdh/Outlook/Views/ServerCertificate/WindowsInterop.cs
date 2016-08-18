namespace OpenEsdh.Outlook.Views.ServerCertificate
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;

    internal static class WindowsInterop
    {
        private static IntPtr _pWH_CALLWNDPROCRET = IntPtr.Zero;
        private static HookProcedureDelegate _WH_CALLWNDPROCRET_PROC = new HookProcedureDelegate(WindowsInterop.WH_CALLWNDPROCRET_PROC);
        private const int WM_COMMAND = 0x111;
        private const int WM_INITDIALOG = 0x110;

        internal static  event GenericDelegate<string, string, bool> ConnectToDialogWillBeShown;

        internal static  event GenericDelegate<bool, bool> SecurityAlertDialogWillBeShown;

        [DllImport("user32.dll")]
        private static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumerateWindowDelegate callback, IntPtr data);
        private static bool enumWindowsCallback(IntPtr hwnd, IntPtr p)
        {
            GCHandle handle = GCHandle.FromIntPtr(p);
            List<IntPtr> target = handle.Target as List<IntPtr>;
            if (target == null)
            {
                throw new InvalidCastException("GCHandle target is not expected type");
            }
            target.Add(hwnd);
            return true;
        }

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern int GetDlgCtrlID(IntPtr hwndCtl);
        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);
        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        internal static void Hook()
        {
            if (_pWH_CALLWNDPROCRET == IntPtr.Zero)
            {
                _pWH_CALLWNDPROCRET = SetWindowsHookEx(HookType.WH_CALLWNDPROCRET, _WH_CALLWNDPROCRET_PROC, IntPtr.Zero, (uint) AppDomain.GetCurrentThreadId());
            }
        }

        private static List<IntPtr> listChildWindows(IntPtr p)
        {
            List<IntPtr> list = new List<IntPtr>();
            GCHandle handle = GCHandle.Alloc(list);
            try
            {
                EnumChildWindows(p, new EnumerateWindowDelegate(WindowsInterop.enumWindowsCallback), GCHandle.ToIntPtr(handle));
            }
            finally
            {
                if (handle.IsAllocated)
                {
                    handle.Free();
                }
            }
            return list;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(HookType hooktype, HookProcedureDelegate callback, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")]
        private static extern bool SetWindowText(IntPtr hWnd, string lpString);
        internal static void Unhook()
        {
            if (_pWH_CALLWNDPROCRET != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_pWH_CALLWNDPROCRET);
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr UnhookWindowsHookEx(IntPtr hhk);
        private static int WH_CALLWNDPROCRET_PROC(int iCode, IntPtr pWParam, IntPtr pLParam)
        {
            if (iCode >= 0)
            {
                CWPRETSTRUCT cwpretstruct = (CWPRETSTRUCT) Marshal.PtrToStructure(pLParam, typeof(CWPRETSTRUCT));
                if (cwpretstruct.message == 0x110)
                {
                    int windowTextLength;
                    StringBuilder builder2;
                    int dlgCtrlID;
                    StringBuilder text = new StringBuilder(GetWindowTextLength(cwpretstruct.hwnd) + 1);
                    GetWindowText(cwpretstruct.hwnd, text, text.Capacity);
                    if ("Security Alert".Equals(text.ToString(), StringComparison.InvariantCultureIgnoreCase))
                    {
                        bool param = true;
                        IntPtr zero = IntPtr.Zero;
                        foreach (IntPtr ptr2 in listChildWindows(cwpretstruct.hwnd))
                        {
                            windowTextLength = GetWindowTextLength(ptr2);
                            if (windowTextLength > 0)
                            {
                                builder2 = new StringBuilder(windowTextLength + 1);
                                GetWindowText(ptr2, builder2, builder2.Capacity);
                                if ("You are about to leave a secure Internet connection.  It will be possible for others to view information you send.".Equals(builder2.ToString(), StringComparison.InvariantCultureIgnoreCase) || "You are about to view pages over a secure connection.".Equals(builder2.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                {
                                    param = false;
                                }
                                if ("&Yes".Equals(builder2.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                {
                                    zero = ptr2;
                                }
                            }
                        }
                        if ((SecurityAlertDialogWillBeShown != null) && (SecurityAlertDialogWillBeShown(param) && (zero != IntPtr.Zero)))
                        {
                            dlgCtrlID = GetDlgCtrlID(zero);
                            SendMessage(cwpretstruct.hwnd, 0x111, new IntPtr(dlgCtrlID), zero);
                            return 1;
                        }
                    }
                    else if (text.ToString().StartsWith("Connect to") || text.ToString().StartsWith("Windows Security"))
                    {
                        IntPtr hWnd = IntPtr.Zero;
                        IntPtr ptr4 = IntPtr.Zero;
                        IntPtr hwndCtl = IntPtr.Zero;
                        foreach (IntPtr ptr2 in listChildWindows(cwpretstruct.hwnd))
                        {
                            builder2 = new StringBuilder(0xff);
                            if ((GetClassName(ptr2, builder2, builder2.Capacity) != 0) && !string.IsNullOrEmpty(builder2.ToString()))
                            {
                                if (builder2.ToString().Equals("SysCredential"))
                                {
                                    foreach (IntPtr ptr6 in listChildWindows(ptr2))
                                    {
                                        StringBuilder lpClassName = new StringBuilder(0xff);
                                        if ((GetClassName(ptr6, lpClassName, lpClassName.Capacity) != 0) && !string.IsNullOrEmpty(lpClassName.ToString()))
                                        {
                                            if ("ComboBoxEx32".Equals(lpClassName.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                hWnd = ptr6;
                                            }
                                            if ("Edit".Equals(lpClassName.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                ptr4 = ptr6;
                                            }
                                        }
                                    }
                                }
                                if ("Button".Equals(builder2.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                {
                                    windowTextLength = GetWindowTextLength(ptr2);
                                    if (windowTextLength > 0)
                                    {
                                        StringBuilder builder4 = new StringBuilder(windowTextLength + 1);
                                        GetWindowText(ptr2, builder4, builder4.Capacity);
                                        if ("OK".Equals(builder4.ToString(), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            hwndCtl = ptr2;
                                        }
                                    }
                                }
                            }
                        }
                        if (ConnectToDialogWillBeShown != null)
                        {
                            string str = null;
                            string str2 = null;
                            if ((((ConnectToDialogWillBeShown(ref str, ref str2) && (str != null)) && ((str2 != null) && (hwndCtl != IntPtr.Zero))) && (hWnd != IntPtr.Zero)) && (ptr4 != IntPtr.Zero))
                            {
                                SetWindowText(hWnd, str);
                                SetWindowText(ptr4, str2);
                                dlgCtrlID = GetDlgCtrlID(hwndCtl);
                                SendMessage(cwpretstruct.hwnd, 0x111, new IntPtr(dlgCtrlID), hwndCtl);
                                return 1;
                            }
                        }
                    }
                }
            }
            return CallNextHookEx(_pWH_CALLWNDPROCRET, iCode, pWParam, pLParam);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CWPRETSTRUCT
        {
            public IntPtr lResult;
            public IntPtr lParam;
            public IntPtr wParam;
            public uint message;
            public IntPtr hwnd;
        }

        private delegate bool EnumerateWindowDelegate(IntPtr pHwnd, IntPtr pParam);

        private delegate int HookProcedureDelegate(int iCode, IntPtr pWParam, IntPtr pLParam);

        private enum HookType
        {
            WH_JOURNALRECORD,
            WH_JOURNALPLAYBACK,
            WH_KEYBOARD,
            WH_GETMESSAGE,
            WH_CALLWNDPROC,
            WH_CBT,
            WH_SYSMSGFILTER,
            WH_MOUSE,
            WH_HARDWARE,
            WH_DEBUG,
            WH_SHELL,
            WH_FOREGROUNDIDLE,
            WH_CALLWNDPROCRET,
            WH_KEYBOARD_LL,
            WH_MOUSE_LL
        }
    }
}


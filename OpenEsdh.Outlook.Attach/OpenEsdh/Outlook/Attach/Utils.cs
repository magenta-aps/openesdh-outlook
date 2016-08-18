namespace OpenEsdh.Outlook.Attach
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Text;

    public static class Utils
    {
        private const int ERROR_NO_MORE_ITEMS = 0x103;
        private const int MAXSIZE = 0x4000;
        public const int MIB_TCP_STATE_CLOSE_WAIT = 8;
        public const int MIB_TCP_STATE_CLOSED = 1;
        public const int MIB_TCP_STATE_CLOSING = 9;
        public const int MIB_TCP_STATE_DELETE_TCB = 12;
        public const int MIB_TCP_STATE_ESTAB = 5;
        public const int MIB_TCP_STATE_FIN_WAIT1 = 6;
        public const int MIB_TCP_STATE_FIN_WAIT2 = 7;
        public const int MIB_TCP_STATE_LAST_ACK = 10;
        public const int MIB_TCP_STATE_LISTEN = 2;
        public const int MIB_TCP_STATE_SYN_RCVD = 4;
        public const int MIB_TCP_STATE_SYN_SENT = 3;
        public const int MIB_TCP_STATE_TIME_WAIT = 11;
        public const int NO_ERROR = 0;
        public const int TOKEN_QUERY = 8;

        [DllImport("kernel32")]
        private static extern bool CloseHandle(IntPtr IntPtr);
        [DllImport("advapi32", CharSet=CharSet.Auto)]
        public static extern bool ConvertSidToStringSid(IntPtr pSID, [In, Out, MarshalAs(UnmanagedType.LPTStr)] ref string pStringSid);
        [DllImport("advapi32", CharSet=CharSet.Auto)]
        private static extern bool ConvertStringSidToSid([In, MarshalAs(UnmanagedType.LPTStr)] string pStringSid, ref IntPtr pSID);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);
        public static bool DumpUserInfo(IntPtr pToken, out IntPtr SID)
        {
            int desiredAccess = 8;
            IntPtr zero = IntPtr.Zero;
            bool flag = false;
            SID = IntPtr.Zero;
            try
            {
                if (OpenProcessToken(pToken, desiredAccess, ref zero))
                {
                    flag = ProcessTokenToSid(zero, out SID);
                    CloseHandle(zero);
                }
                return flag;
            }
            catch (Exception)
            {
                return true;
            }
        }

        [DllImport("advapi32.dll", CharSet=CharSet.Auto, SetLastError=true)]
        public static extern int DuplicateToken(IntPtr hToken, int impersonationLevel, ref IntPtr hNewToken);
        public static string ExGetProcessInfoByPID(int PID, out string SID)
        {
            IntPtr zero = IntPtr.Zero;
            SID = string.Empty;
            try
            {
                Process processById = Process.GetProcessById(PID);
                if (DumpUserInfo(processById.Handle, out zero))
                {
                    ConvertSidToStringSid(zero, ref SID);
                }
                return processById.ProcessName;
            }
            catch
            {
                return "Unknown";
            }
        }

        [DllImport("kernel32")]
        private static extern IntPtr GetCurrentProcess();
        public static List<object> GetRunningInstances(string[] progIds)
        {
            List<string> list = new List<string>();
            foreach (string str in progIds)
            {
                Type typeFromProgID = Type.GetTypeFromProgID(str);
                if (typeFromProgID != null)
                {
                    list.Add(typeFromProgID.GUID.ToString().ToUpper());
                }
            }
            IRunningObjectTable prot = null;
            GetRunningObjectTable(0, out prot);
            if (prot == null)
            {
                return null;
            }
            IEnumMoniker ppenumMoniker = null;
            prot.EnumRunning(out ppenumMoniker);
            if (ppenumMoniker == null)
            {
                return null;
            }
            ppenumMoniker.Reset();
            List<object> list2 = new List<object>();
            IntPtr pceltFetched = new IntPtr();
            IMoniker[] rgelt = new IMoniker[1];
            while (ppenumMoniker.Next(1, rgelt, pceltFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);
                if (ctx != null)
                {
                    string str2;
                    rgelt[0].GetDisplayName(ctx, null, out str2);
                    foreach (string str3 in list)
                    {
                        if (str2.ToUpper().IndexOf(str3) > 0)
                        {
                            object obj2;
                            prot.GetObject(rgelt[0], out obj2);
                            if (obj2 != null)
                            {
                                list2.Add(obj2);
                                break;
                            }
                        }
                    }
                }
            }
            return list2;
        }

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
        [DllImport("advapi32", CharSet=CharSet.Auto)]
        private static extern bool GetTokenInformation(IntPtr hToken, TOKEN_INFORMATION_CLASS tokenInfoClass, IntPtr TokenInformation, int tokeInfoLength, ref int reqLength);
        [DllImport("advapi32", CharSet=CharSet.Auto)]
        private static extern bool LookupAccountSid([In, MarshalAs(UnmanagedType.LPTStr)] string lpSystemName, IntPtr pSid, StringBuilder Account, ref int cbName, StringBuilder DomainName, ref int cbDomainName, ref int peUse);
        [DllImport("advapi32")]
        private static extern bool OpenProcessToken(IntPtr ProcessIntPtr, int DesiredAccess, ref IntPtr TokenIntPtr);
        private static bool ProcessTokenToSid(IntPtr token, out IntPtr SID)
        {
            bool flag2;
            IntPtr tokenInformation = Marshal.AllocHGlobal(0x100);
            bool flag = false;
            SID = IntPtr.Zero;
            try
            {
                int tokeInfoLength = 0x100;
                flag = GetTokenInformation(token, TOKEN_INFORMATION_CLASS.TokenUser, tokenInformation, tokeInfoLength, ref tokeInfoLength);
                if (flag)
                {
                    TOKEN_USER token_user = (TOKEN_USER) Marshal.PtrToStructure(tokenInformation, typeof(TOKEN_USER));
                    SID = token_user.User.Sid;
                }
                flag2 = flag;
            }
            catch (Exception)
            {
                flag2 = false;
            }
            finally
            {
                Marshal.FreeHGlobal(tokenInformation);
            }
            return flag2;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct _SID_AND_ATTRIBUTES
        {
            public IntPtr Sid;
            public int Attributes;
        }

        private enum TOKEN_INFORMATION_CLASS
        {
            TokenDefaultDacl = 6,
            TokenGroups = 2,
            TokenImpersonationLevel = 9,
            TokenOwner = 4,
            TokenPrimaryGroup = 5,
            TokenPrivileges = 3,
            TokenRestrictedSids = 11,
            TokenSessionId = 12,
            TokenSource = 7,
            TokenStatistics = 10,
            TokenType = 8,
            TokenUser = 1
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct TOKEN_USER
        {
            public Utils._SID_AND_ATTRIBUTES User;
        }
    }
}


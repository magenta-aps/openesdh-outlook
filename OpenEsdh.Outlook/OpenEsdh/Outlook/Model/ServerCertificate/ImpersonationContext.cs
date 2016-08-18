namespace OpenEsdh.Outlook.Model.ServerCertificate
{
    using System;
    using System.Runtime.InteropServices;
    using System.Security.Principal;

    public class ImpersonationContext : IDisposable
    {
        private bool impersonating;
        private WindowsImpersonationContext impersonationContext;
        private const int LOGON32_LOGON_INTERACTIVE = 2;
        private const int LOGON32_PROVIDER_DEFAULT = 0;

        public ImpersonationContext()
        {
            this.impersonating = false;
        }

        public ImpersonationContext(string userName, string password, string domain)
        {
            this.BeginImpersonationContext(userName, password, domain);
        }

        public void BeginImpersonationContext(string userName, string password, string domain)
        {
            IntPtr zero = IntPtr.Zero;
            IntPtr hNewToken = IntPtr.Zero;
            if ((RevertToSelf() && (LogonUserA(userName, domain, password, 2, 0, ref zero) != 0)) && (DuplicateToken(zero, 2, ref hNewToken) != 0))
            {
                using (WindowsIdentity identity = new WindowsIdentity(hNewToken))
                {
                    this.impersonationContext = identity.Impersonate();
                    this.impersonating = true;
                }
            }
            if (zero != IntPtr.Zero)
            {
                CloseHandle(zero);
            }
            if (hNewToken != IntPtr.Zero)
            {
                CloseHandle(hNewToken);
            }
        }

        [DllImport("kernel32.dll", CharSet=CharSet.Auto)]
        private static extern bool CloseHandle(IntPtr handle);
        public void Dispose()
        {
            this.Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing && (this.impersonationContext != null))
            {
                this.impersonationContext.Undo();
                this.impersonationContext.Dispose();
            }
        }

        [DllImport("advapi32.dll", CharSet=CharSet.Auto, SetLastError=true)]
        private static extern int DuplicateToken(IntPtr hToken, int impersonationLevel, ref IntPtr hNewToken);
        public void EndImpersonationContext()
        {
            if (this.impersonationContext != null)
            {
                this.impersonationContext.Undo();
                this.impersonationContext.Dispose();
            }
            this.impersonating = false;
        }

        ~ImpersonationContext()
        {
            this.Dispose(false);
        }

        [DllImport("advapi32.dll")]
        private static extern int LogonUserA(string lpszUserName, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);
        [DllImport("advapi32.dll", CharSet=CharSet.Auto, SetLastError=true)]
        private static extern bool RevertToSelf();

        public bool Impersonating
        {
            get
            {
                return this.impersonating;
            }
        }
    }
}


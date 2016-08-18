namespace OpenEsdh
{
    using SharpShell;
    using System;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class ServerEntry
    {
        public string GetSecurityStatus()
        {
            byte[] publicKey = AssemblyName.GetAssemblyName(this.ServerPath).GetPublicKey();
            return (((publicKey != null) && (publicKey.Length > 0)) ? "Signed" : "Not Signed");
        }

        public Guid ClassId { get; set; }

        public bool IsInvalid { get; set; }

        public ISharpShellServer Server { get; set; }

        public string ServerName { get; set; }

        public string ServerPath { get; set; }

        public SharpShell.ServerType ServerType { get; set; }
    }
}

